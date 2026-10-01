#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Автоматизированный генератор графика отгрузок (РЦ Черная Грязь)
и модуль персональной почтовой рассылки через Microsoft Outlook.

Ключевые принципы:
1. Вход: 'График_отгрузки_филиалов_неделя_2.xlsm' (или любой .xlsm в папке).
2. Выход: Excel с 7 листами по дням недели (28.09 Пн, 29.09 Вт, ...).
3. Ограничение емкости: строго 2 машины в час.
4. Конфиденциальность: каждому перевозчику уходит ТОЛЬКО его расписание!
5. Интеграция с Outlook: прямая отправка без SMTP и паролей (pywin32).
"""

import sys
import os
import re
import glob
import json
import argparse
from datetime import datetime, timedelta, time
from collections import Counter, defaultdict
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side

RUSSIAN_MONTHS = {
    1: 'Январь', 2: 'Февраль', 3: 'Март', 4: 'Апрель', 5: 'Май', 6: 'Июнь',
    7: 'Июль', 8: 'Август', 9: 'Сентябрь', 10: 'Октябрь', 11: 'Ноябрь', 12: 'Декабрь'
}

DAYS_MAPPING = [
    ('Пн', 'пн', 0),
    ('Вт', 'вт', 1),
    ('Ср', 'ср', 2),
    ('Чт', 'чт', 3),
    ('Пт', 'пт', 4),
    ('Сб', 'сб', 5),
    ('Вс', 'вс', 6)
]

# Логистический порядок подачи направлений под погрузку на ворота РЦ
MASTER_CITY_ORDER = {
    'Пн': ['Ярославль', 'Казань', 'Воронеж', 'Краснодар', 'Волгоград', 'Самара', 'Екатеринбург', 'Иркутск', 'Санкт-Петербург'],
    'Вт': ['Ярославль', 'Казань', 'Воронеж', 'Краснодар', 'Самара', 'Красноярск', 'Санкт-Петербург', 'Брянск'],
    'Ср': ['Ярославль', 'Казань', 'Воронеж', 'Краснодар', 'Волгоград', 'Самара', 'Екатеринбург', 'Новосибирск', 'Красноярск', 'Иркутск', 'Хабаровск', 'Брянск'],
    'Чт': ['Ярославль', 'Воронеж', 'Краснодар', 'Екатеринбург', 'Новосибирск', 'Хабаровск', 'Санкт-Петербург', 'Брянск'],
    'Пт': ['Ярославль', 'Краснодар', 'Екатеринбург', 'Новосибирск', 'Красноярск', 'Иркутск', 'Хабаровск', 'Санкт-Петербург'],
    'Сб': ['Воронеж', 'Краснодар', 'Волгоград', 'Самара', 'Хабаровск'],
    'Вс': ['Казань', 'Воронеж', 'Краснодар', 'Волгоград', 'Самара', 'Екатеринбург', 'Санкт-Петербург', 'Брянск']
}

START_HOURS = {
    'Пн': 9, 'Вт': 9, 'Ср': 9, 'Чт': 9, 'Пт': 9,
    'Сб': 5, # Суббота: ранние окна
    'Вс': 9
}

DELIVERY_TIMINGS = {
    'брянск': {'transit_hours': 8, 'tz_diff': 0},
    'волгоград': {'transit_hours': 36, 'tz_diff': 0},       # 1.5 суток
    'воронеж': {'transit_hours': 24, 'tz_diff': 0},         # 1 сутки
    'екатеринбург': {'transit_hours': 72, 'tz_diff': 2},    # 3 суток, +2ч
    'иркутск': {'transit_hours': 168, 'tz_diff': 5},        # 7 суток, +5ч
    'казань': {'transit_hours': 36, 'tz_diff': 0},          # 1.5 суток
    'краснодар': {'transit_hours': 48, 'tz_diff': 0},       # 2 суток
    'красноярск': {'transit_hours': 144, 'tz_diff': 4},     # 6 суток, +4ч
    'новосибирск': {'transit_hours': 120, 'tz_diff': 3},    # 5 суток, +3ч
    'самара': {'transit_hours': 48, 'tz_diff': 1},          # 2 суток, +1ч
    'санкт-петербург': {'transit_hours': 24, 'tz_diff': 0}, # 1 сутки
    'спб': {'transit_hours': 24, 'tz_diff': 0},
    'ярославль': {'transit_hours': 5, 'tz_diff': 0},        # 5 часов
    'хабаровск': {'transit_hours': 276, 'tz_diff': 8},      # 11.5 суток, +8ч
    # Резервные/дополнительные направления
    'ростов-на-дону': {'transit_hours': 48, 'tz_diff': 0},
    'ростов': {'transit_hours': 48, 'tz_diff': 0},
    'уфа': {'transit_hours': 48, 'tz_diff': 2},
}

def get_delivery_info(city):
    key = str(city).strip().lower()
    for k, v in DELIVERY_TIMINGS.items():
        if k in key or key in k:
            return v
    return {'transit_hours': 24, 'tz_diff': 0}

def calculate_delivery_datetime(dep_date, time_str, city):
    """
    Рассчитывает дату и время доставки на основе времени погрузки,
    количества дней/часов в пути и разницы во времени.
    """
    if isinstance(dep_date, datetime):
        base_dt = datetime(dep_date.year, dep_date.month, dep_date.day)
    elif hasattr(dep_date, 'year'):
        base_dt = datetime(dep_date.year, dep_date.month, dep_date.day)
    else:
        try:
            base_dt = datetime.strptime(str(dep_date).strip(), '%d.%m.%Y')
        except Exception:
            base_dt = datetime.now()
            
    parts = [int(p) for p in str(time_str).strip().split(':') if p.isdigit()]
    h = parts[0] if len(parts) > 0 else 9
    m = parts[1] if len(parts) > 1 else 0
    dep_dt = base_dt.replace(hour=h, minute=m, second=0)
    
    info = get_delivery_info(city)
    total_hours = info['transit_hours'] + info['tz_diff']
    delivery_dt = dep_dt + timedelta(hours=total_hours)
    return delivery_dt


def find_tariffs_file(base_dir=None):
    """Ищет файл ТАРИФЫ РК ТАБЛИЦА.xlsx в проекте, Загрузках или на Рабочем столе"""
    candidates = []
    if base_dir:
        candidates.append(os.path.join(base_dir, 'ТАРИФЫ РК ТАБЛИЦА.xlsx'))
    user_home = os.path.expanduser('~')
    candidates.extend([
        os.path.join(os.path.dirname(os.path.abspath(__file__)), 'ТАРИФЫ РК ТАБЛИЦА.xlsx'),
        os.path.join(user_home, 'Downloads', 'ТАРИФЫ РК ТАБЛИЦА.xlsx'),
        os.path.join(user_home, 'Загрузки', 'ТАРИФЫ РК ТАБЛИЦА.xlsx'),
        os.path.join(user_home, 'Desktop', 'ТАРИФЫ РК ТАБЛИЦА.xlsx'),
        os.path.join(user_home, 'Рабочий стол', 'ТАРИФЫ РК ТАБЛИЦА.xlsx'),
        os.path.join(user_home, 'OneDrive', 'Desktop', 'ТАРИФЫ РК ТАБЛИЦА.xlsx'),
        os.path.join(user_home, 'OneDrive', 'Рабочий стол', 'ТАРИФЫ РК ТАБЛИЦА.xlsx'),
    ])
    for c in candidates:
        if os.path.exists(c):
            return c
    return None

def load_tariffs_matrix(base_dir=None):
    """
    Загружает актуальную матрицу тарифов из листа 'Актуальные ' файла 'ТАРИФЫ РК ТАБЛИЦА.xlsx'.
    Возвращает словарь: {город: {перевозчик: тариф}}
    """
    tariff_file = find_tariffs_file(base_dir)
    tariffs = {}
    
    carriers_norm = {
        'Сияние': 'ТК Сияние',
        'Наткар': 'АО Национальный',
        'АЗИМУТ': 'Азимут',
        'Примум': 'Примум',
        'Буш': 'Буш-Авто',
        'Норд Лайн': 'НОРДЛАЙН',
        'Караван': 'Караван',
        'Виллайн': 'Виллайн',
        'ЕманТрансАвто': 'ЕманТрансАвто',
        'Олимп': 'Олимп',
        'Агро-Авто': 'Агро-Авто',
        'Веб-Логистика': 'Веб-Логистика',
        'Буш-Авто': 'Буш-Авто',
        'Азимут': 'Азимут',
        'ИП Коршунов': 'ИП Коршунов',
        'ИП Мельник': 'ИП Мельник',
        'ИП Гусманов': 'ИП Гусманов'
    }
    
    city_norm = {
        'СПБ': 'Санкт-Петербург',
        'Санкт-Петербург': 'Санкт-Петербург',
        'Брянск': 'Брянск',
        'Ярославль': 'Ярославль',
        'Воронеж': 'Воронеж',
        'Казань': 'Казань',
        'Волгоград': 'Волгоград',
        'Краснодар': 'Краснодар',
        'Самара': 'Самара',
        'Екатеринбург': 'Екатеринбург',
        'Новосибирск': 'Новосибирск',
        'Красноярск': 'Красноярск',
        'Иркутск': 'Иркутск',
        'Хабаровск': 'Хабаровск',
        'Уфа': 'Уфа',
        'УФА': 'Уфа'
    }
    
    if tariff_file and os.path.exists(tariff_file):
        try:
            wb = openpyxl.load_workbook(tariff_file, data_only=True)
            ws = wb['Актуальные '] if 'Актуальные ' in wb.sheetnames else wb.active
            for r in range(5, ws.max_row + 1):
                c_raw = ws.cell(r, 1).value
                if not c_raw or '-экспресс' in str(c_raw):
                    continue
                c_clean = str(c_raw).strip()
                c = city_norm.get(c_clean, c_clean)
                if c not in tariffs:
                    tariffs[c] = {}
                    
                for col in range(2, ws.max_column + 1):
                    car_raw = ws.cell(3, col).value
                    val = ws.cell(r, col).value
                    if car_raw in carriers_norm and isinstance(val, (int, float)) and val > 0:
                        car = carriers_norm[car_raw]
                        if car not in tariffs[c] or val < tariffs[c][car]:
                            tariffs[c][car] = int(val)
            wb.close()
        except Exception as e:
            print(f"Ошибка загрузки тарифов из {tariff_file}: {e}")
            
    # Резервная встроенная сетка тарифов (если файл отсутствует)
    if not tariffs:
        tariffs = {
            'Брянск': {'НОРДЛАЙН': 51000, 'ЕманТрансАвто': 51500, 'Олимп': 51500, 'ТК Сияние': 52500, 'Агро-Авто': 73000},
            'Ярославль': {'ИП Гусманов': 47000, 'Буш-Авто': 48800, 'ИП Коршунов': 52000, 'ИП Мельник': 52000, 'ТК Сияние': 52500, 'ЕманТрансАвто': 56500, 'Олимп': 56500, 'Караван': 65000, 'Агро-Авто': 77000},
            'Воронеж': {'Буш-Авто': 57950, 'ТК Сияние': 65000, 'ЕманТрансАвто': 65000, 'Олимп': 65000, 'Караван': 70000, 'Агро-Авто': 85000},
            'Санкт-Петербург': {'НОРДЛАЙН': 86700, 'Караван': 94000, 'Буш-Авто': 96583, 'ЕманТрансАвто': 97000, 'Олимп': 97000, 'ТК Сияние': 97650, 'Агро-Авто': 133000, 'АО Национальный': 135000},
            'Казань': {'Буш-Авто': 116917, 'НОРДЛАЙН': 118320, 'Караван': 128000, 'ЕманТрансАвто': 133800, 'Олимп': 133800, 'ТК Сияние': 136500, 'АО Национальный': 143000, 'Азимут': 150000, 'Агро-Авто': 157000, 'Веб-Логистика': 160000},
            'Волгоград': {'ЕманТрансАвто': 134000, 'Олимп': 134000, 'ТК Сияние': 136500, 'АО Национальный': 155000, 'Агро-Авто': 157000},
            'Краснодар': {'Караван': 175000, 'Буш-Авто': 177917, 'ЕманТрансАвто': 180000, 'Олимп': 180000, 'ТК Сияние': 189000, 'Веб-Логистика': 195000, 'АО Национальный': 201600, 'Азимут': 265000, 'Агро-Авто': 287000},
            'Самара': {'ЕманТрансАвто': 134200, 'Олимп': 134200, 'ТК Сияние': 137550, 'Азимут': 150000, 'Агро-Авто': 162000, 'АО Национальный': 166000},
            'Екатеринбург': {'НОРДЛАЙН': 245000, 'Олимп': 247000, 'Веб-Логистика': 250000, 'ТК Сияние': 252000, 'Азимут': 260000, 'АО Национальный': 261450, 'Буш-Авто': 274500, 'ЕманТрансАвто': 280000, 'Агро-Авто': 290000, 'Примум': 300000},
            'Новосибирск': {'Примум': 360000, 'ЕманТрансАвто': 433000, 'Олимп': 433000, 'Виллайн': 450000, 'Веб-Логистика': 450000, 'ТК Сияние': 451500, 'Агро-Авто': 455000, 'Азимут': 457000, 'АО Национальный': 495000},
            'Хабаровск': {'Примум': 800000, 'Азимут': 940000, 'ЕманТрансАвто': 972000, 'Олимп': 972000, 'Агро-Авто': 997000, 'ТК Сияние': 997500, 'Виллайн': 1050000},
            'Красноярск': {'Примум': 420000, 'Виллайн': 530000, 'ЕманТрансАвто': 547000, 'Олимп': 547000, 'Агро-Авто': 550000, 'Веб-Логистика': 550000, 'ТК Сияние': 556500, 'Азимут': 620000},
            'Иркутск': {'Примум': 500000, 'Агро-Авто': 636000, 'Азимут': 650000, 'ЕманТрансАвто': 660500, 'Олимп': 660500, 'ТК Сияние': 682500},
            'Уфа': {'Буш-Авто': 162667, 'ТК Сияние': 189000, 'Азимут': 200000, 'АО Национальный': 225000, 'Агро-Авто': 225000}
        }
    return tariffs

def plan_weekly_carrier_assignments(plan_trips, tariffs):
    """
    Распределяет перевозчиков и тарифы по рейсам согласно официальным правилам
    из файла 'Исключения для алгоритма распределения транспорта.docx':

    Иерархия приоритетов:
    1. Направление -> 2. Количество машин -> 3. Номер машины -> 4. Фиксированный перевозчик
    -> 5. Доступность ТС -> 6. Исключение/отказ -> 7. Резервный перевозчик -> 8. Экономическая оптимизация.
    Фиксированные правила имеют наивысший приоритет над оптимизацией по тарифу.
    """
    assigned = {}
    carrier_counts = defaultdict(int)
    city_carrier_counts = defaultdict(lambda: defaultdict(int))

    def assign_trip(d_name, city, tr_i, carrier):
        key = (d_name, city, tr_i)
        tar = tariffs.get(city, {}).get(carrier, 0)
        if tar == 0 and carrier == 'Авангард':
            tar = tariffs.get(city, {}).get('Азимут', 940000)
        assigned[key] = (carrier, tar)
        carrier_counts[carrier] += 1
        city_carrier_counts[city][carrier] += 1

    def pick_balanced(city, allowed):
        # Балансировка: приоритет перевозчику с наименьшим накопленным количеством рейсов за неделю
        sorted_cars = sorted(
            allowed,
            key=lambda c: (
                carrier_counts[c],
                city_carrier_counts[city][c],
                tariffs.get(city, {}).get(c, 9999999)
            )
        )
        return sorted_cars[0]

    def pick_cheapest(city, allowed=None, exclude=None):
        c_tar = tariffs.get(city, {})
        candidates = allowed if allowed else list(c_tar.keys())
        if exclude:
            candidates = [c for c in candidates if c not in exclude]
        if not candidates:
            return 'ТК Сияние'
        sorted_cars = sorted(
            candidates,
            key=lambda c: (
                c_tar.get(c, 9999999),
                carrier_counts[c]
            )
        )
        return sorted_cars[0]

    # Обрабатываем дни недели последовательно (Пн -> Вс)
    for d_name, d_short, offset in DAYS_MAPPING:
        d_plan = plan_trips.get(d_name, {})
        city_order = list(MASTER_CITY_ORDER.get(d_name, []))
        for c in d_plan:
            if c not in city_order:
                city_order.append(c)

        for city in city_order:
            cnt = d_plan.get(city, 0)
            for tr_i in range(1, cnt + 1):

                # 2. ЯРОСЛАВЛЬ
                # 1-я машина: чередование по дням (Пн/Ср/Пт - Мельник, Вт/Чт - Коршунов)
                # 2-я и последующие - всегда ИП Гусманов
                if city == 'Ярославль':
                    if tr_i == 1:
                        carrier = 'ИП Мельник' if d_name in ['Пн', 'Ср', 'Пт', 'Вс'] else 'ИП Коршунов'
                    else:
                        carrier = 'ИП Гусманов'

                # 3. БРЯНСК
                # 1-я машина — Норд Лайн, 2-я машина — Сияние ТК, 3-я+ — по рентабельному тарифу
                elif city == 'Брянск':
                    if tr_i == 1:
                        carrier = 'НОРДЛАЙН'
                    elif tr_i == 2:
                        carrier = 'ТК Сияние'
                    else:
                        carrier = pick_cheapest(city, exclude={'НОРДЛАЙН', 'ТК Сияние'})

                # 4. КАЗАНЬ
                # 1-я машина — Сияние ТК; 2-я+ между Норд Лайн, АО Национальный, ООО Агро-Авто
                elif city == 'Казань':
                    if tr_i == 1:
                        carrier = 'ТК Сияние'
                    else:
                        carrier = pick_balanced(city, ['НОРДЛАЙН', 'АО Национальный', 'Агро-Авто'])

                # 5. УФА
                # 1-я машина — Сияние ТК; 2-я+ — АО Национальный и/или ООО Агро-Авто
                elif city == 'Уфа':
                    if tr_i == 1:
                        carrier = 'ТК Сияние'
                    else:
                        carrier = pick_balanced(city, ['АО Национальный', 'Агро-Авто'])

                # 6. ВОРОНЕЖ
                # Направление Воронеж — ТОЛЬКО Сияние ТК. Другие перевозчики не назначаются.
                elif city == 'Воронеж':
                    carrier = 'ТК Сияние'

                # 7. РОСТОВ-НА-ДОНУ
                # Основной перевозчик — Сияние ТК. Резервные: АО Национальный -> Агро-Авто -> Буш-Авто
                elif city == 'Ростов-на-Дону':
                    carrier = 'ТК Сияние'

                # 8. КРАСНОДАР
                # 1-я: Сияние ТК, 2-я: АО Национальный, 3-я: ООО Буш-Автопром, 4-я: ООО Агро-Авто
                elif city == 'Краснодар':
                    kras_seq = ['ТК Сияние', 'АО Национальный', 'Буш-Авто', 'Агро-Авто']
                    if tr_i <= len(kras_seq):
                        carrier = kras_seq[tr_i - 1]
                    else:
                        carrier = kras_seq[(tr_i - 1) % len(kras_seq)]

                # 9. ВОЛГОГРАД
                # Направление Волгоград — ТОЛЬКО Сияние ТК.
                elif city == 'Волгоград':
                    carrier = 'ТК Сияние'

                # 10. САМАРА
                # 1-я машина — Сияние ТК, 2-я и последующие — преимущественно АО Национальный перевозчик
                elif city == 'Самара':
                    if tr_i == 1:
                        carrier = 'ТК Сияние'
                    else:
                        carrier = 'АО Национальный'

                # 11. ЕКАТЕРИНБУРГ
                # 1-я: Сияние ТК, 2-я: АО Национальный, 3-я и последующие — по рентабельному тарифу
                elif city == 'Екатеринбург':
                    if tr_i == 1:
                        carrier = 'ТК Сияние'
                    elif tr_i == 2:
                        carrier = 'АО Национальный'
                    else:
                        carrier = pick_cheapest(city, exclude={'ТК Сияние', 'АО Национальный'})

                # 12. НОВОСИБИРСК
                # Допустимые: ЕманТрансАвто, Виллайн, ВЕБ ЛОГИСТИКА. Равномерный баланс за неделю.
                elif city == 'Новосибирск':
                    carrier = pick_balanced(city, ['ЕманТрансАвто', 'Виллайн', 'Веб-Логистика'])

                # 13. КРАСНОЯРСК
                # Допустимые: ЕманТрансАвто, Виллайн, ВЕБ ЛОГИСТИКА. Равномерный баланс за неделю.
                elif city == 'Красноярск':
                    carrier = pick_balanced(city, ['ЕманТрансАвто', 'Виллайн', 'Веб-Логистика'])

                # 14. ИРКУТСК
                # Допустимые: Азимут ТК, ЕманТрансАвто. Равномерный баланс за неделю.
                elif city == 'Иркутск':
                    carrier = pick_balanced(city, ['Азимут', 'ЕманТрансАвто'])

                # 15. ХАБАРОВСК
                # 1-я и 2-я машины — Азимут ТК; 3-я и последующие — Авангард
                elif city == 'Хабаровск':
                    if tr_i in (1, 2):
                        carrier = 'Азимут'
                    else:
                        carrier = 'Авангард'

                # 16. САНКТ-ПЕТЕРБУРГ
                # 1-я машина — Сияние ТК; 2-я и последующие — Норд Лайн / АО Национальный
                elif city == 'Санкт-Петербург':
                    if tr_i == 1:
                        carrier = 'ТК Сияние'
                    else:
                        carrier = pick_balanced(city, ['НОРДЛАЙН', 'АО Национальный'])

                # Прочие направления по умолчанию
                else:
                    if tr_i == 1:
                        carrier = 'ТК Сияние'
                    else:
                        carrier = pick_cheapest(city)

                assign_trip(d_name, city, tr_i, carrier)

    return assigned

def find_default_plan_file(base_dir):
    candidates = [
        os.path.join(base_dir, 'График_отгрузки_филиалов_неделя_2.xlsm'),
        *glob.glob(os.path.join(base_dir, '*График*отгрузки*.xlsm')),
        *glob.glob(os.path.join(base_dir, '*.xlsm'))
    ]
    for c in candidates:
        if os.path.exists(c):
            return c
    return None

def parse_plan_file(plan_path):
    wb = openpyxl.load_workbook(plan_path, data_only=True)
    sheet_name = wb.sheetnames[0]
    for s in wb.sheetnames:
        if '.' in s and '-' in s:
            sheet_name = s
            break
    ws = wb[sheet_name]
    
    header_val = str(ws['A1'].value or '')
    match = re.search(r'(\d{1,2}\.\d{2})\s*-\s*(\d{1,2}\.\d{2})', header_val)
    if match:
        start_d_str = match.group(1)
        cur_year = datetime.now().year
        try:
            start_date = datetime.strptime(f"{start_d_str}.{cur_year}", "%d.%m.%Y")
        except Exception:
            start_date = datetime(2026, 9, 28)
    else:
        start_date = datetime(2026, 9, 28)
        
    days_cols = ['Пн', 'Вт', 'Ср', 'Чт', 'Пт', 'Сб', 'Вс']
    plan_trips = defaultdict(lambda: defaultdict(int))
    
    for r in range(2, ws.max_row+1):
        city = ws.cell(r, 1).value
        if not city or not str(city).strip() or str(city).strip().lower() in ['итого', 'всего']:
            continue
        city_clean = str(city).strip()
        for col_idx in range(2, 9):
            d_name = days_cols[col_idx-2]
            cnt = ws.cell(r, col_idx).value or 0
            try:
                cnt_int = int(cnt)
            except Exception:
                cnt_int = 0
            if cnt_int > 0:
                plan_trips[d_name][city_clean] = cnt_int
                
    wb.close()
    return plan_trips, start_date

def parse_final_schedule_file(excel_path):
    """
    Считывает готовый график отгрузок (Недельные графики.xlsx), 
    в который сотрудник мог внести ручные правки в Excel 
    (добавил/удалил машины, изменил время, перевозчика или тариф).
    Поддерживает как новый формат (Месяц, Дата, Время, День, Направление, Номер, Перевозчик, Паллеты, Дата доставки, Тариф),
    так и прежний формат (Дата, День, Месяц, Номер, Направление, Паллеты, Перевозчик, Тариф, Время).
    """
    wb = openpyxl.load_workbook(excel_path, data_only=True)
    all_trips = []
    
    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        # Проверяем, что это дневной лист графика
        h_c = str(ws['C1'].value or '').strip().lower()
        h_d = str(ws['D1'].value or '').strip().lower()
        h_g = str(ws['G1'].value or '').strip().lower()
        
        if 'месяц' not in h_c and 'дата' not in h_c and 'дата' not in h_d and 'направление' not in h_g:
            continue
            
        is_new_format = ('месяц' in h_c) or ('время' in str(ws['E1'].value or '').lower())
        
        for row in range(2, ws.max_row + 1):
            if is_new_format:
                month_val = ws[f'C{row}'].value or ''
                date_val = ws[f'D{row}'].value
                time_val = ws[f'E{row}'].value
                day_val = ws[f'F{row}'].value or ''
                city_val = ws[f'G{row}'].value
                truck_num = ws[f'H{row}'].value or 1
                carrier_val = ws[f'I{row}'].value
                pallets = ws[f'J{row}'].value or 33
                deliv_val = ws[f'K{row}'].value
                tariff_val = ws[f'L{row}'].value
            else:
                date_val = ws[f'C{row}'].value
                day_val = ws[f'D{row}'].value or ''
                month_val = ws[f'E{row}'].value or ''
                truck_num = ws[f'F{row}'].value or 1
                city_val = ws[f'G{row}'].value
                pallets = ws[f'H{row}'].value or 33
                carrier_val = ws[f'I{row}'].value
                tariff_val = ws[f'J{row}'].value
                time_val = ws[f'K{row}'].value
                deliv_val = None
            
            if not city_val or str(city_val).strip() == '' or 'Итого' in str(city_val):
                continue
            if not date_val:
                continue
                
            tariff_num = 0
            if tariff_val is not None:
                try:
                    clean_t = str(tariff_val).replace(' ', '').replace('₽', '').replace('\xa0', '').replace(',', '.')
                    tariff_num = int(float(clean_t))
                except Exception:
                    tariff_num = 0
            
            if isinstance(date_val, datetime):
                dt_obj = date_val
                date_str = dt_obj.strftime('%d.%m.%Y')
            else:
                date_str = str(date_val).strip()
                try:
                    dt_obj = datetime.strptime(date_str, '%d.%m.%Y')
                except Exception:
                    dt_obj = None
                    
            if isinstance(time_val, (datetime, time)):
                time_str = time_val.strftime('%H:%M:%S')
            else:
                time_str = str(time_val or '09:00:00').strip()
                if len(time_str) == 5:
                    time_str += ':00'
                    
            city_str = str(city_val).strip()
            if deliv_val is not None and str(deliv_val).strip():
                if isinstance(deliv_val, datetime):
                    delivery_dt = deliv_val
                    delivery_str = delivery_dt.strftime('%d.%m.%Y %H:%M')
                else:
                    delivery_str = str(deliv_val).strip()
                    try:
                        delivery_dt = datetime.strptime(delivery_str, '%d.%m.%Y %H:%M')
                    except Exception:
                        delivery_dt = None
            else:
                delivery_dt = calculate_delivery_datetime(dt_obj or date_str, time_str, city_str)
                delivery_str = delivery_dt.strftime('%d.%m.%Y %H:%M')
                    
            all_trips.append({
                'date': date_str,
                'dt': dt_obj,
                'day': str(day_val).strip(),
                'month': str(month_val).strip(),
                'city': city_str,
                'truck_num': truck_num,
                'carrier': str(carrier_val or 'ТК Сияние').strip(),
                'pallets': pallets,
                'tariff': tariff_num,
                'time': time_str,
                'delivery_dt': delivery_dt,
                'delivery_str': delivery_str
            })
            
    wb.close()
    return all_trips


def build_schedule_from_plan(plan_trips, start_date, output_path):
    wb_out = openpyxl.Workbook()
    wb_out.remove(wb_out.active)
    
    tariffs = load_tariffs_matrix(os.path.dirname(output_path))
    weekly_assignments = plan_weekly_carrier_assignments(plan_trips, tariffs)
    
    font_header = Font(name='Arial', size=10, bold=True)
    font_regular = Font(name='Arial', size=10)
    font_bold_center = Font(name='Times New Roman', size=11, bold=True)
    font_city = Font(name='Times New Roman', size=11, bold=True)
    font_carrier = Font(name='Times New Roman', size=11)
    font_key = Font(name='Arial', size=10, color='7F7F7F')

    border_thin = Border(
        left=Side(style='thin', color='D9D9D9'),
        right=Side(style='thin', color='D9D9D9'),
        top=Side(style='thin', color='D9D9D9'),
        bottom=Side(style='thin', color='D9D9D9')
    )
    border_header = Border(
        left=Side(style='thin', color='A6A6A6'),
        right=Side(style='thin', color='A6A6A6'),
        top=Side(style='thin', color='A6A6A6'),
        bottom=Side(style='medium', color='000000')
    )
    border_summary_header = Border(
        left=Side(style='thin', color='BFBFBF'),
        right=Side(style='thin', color='BFBFBF'),
        top=Side(style='thin', color='BFBFBF'),
        bottom=Side(style='thin', color='000000')
    )
    border_total = Border(
        top=Side(style='thin', color='000000'),
        bottom=Side(style='double', color='000000')
    )
    fill_40p = PatternFill(start_color='E2EFDA', end_color='E2EFDA', fill_type='solid')

    col_widths = {
        'A': 28, 'B': 5, 'C': 12, 'D': 14, 'E': 16, 'F': 6,
        'G': 18, 'H': 15, 'I': 24, 'J': 15, 'K': 18, 'L': 14,
        'M': 5, 'N': 22, 'O': 10, 'P': 10
    }

    all_scheduled_trips = []

    for d_name, d_short, offset in DAYS_MAPPING:
        cur_date = start_date + timedelta(days=offset)
        d_str = cur_date.strftime('%d.%m.%Y')
        d_month = RUSSIAN_MONTHS[cur_date.month]
        sheet_title = f"{cur_date.strftime('%d.%m')} {d_name}"
        
        ws = wb_out.create_sheet(title=sheet_title)
        ws.views.sheetView[0].showGridLines = True
        
        for col_l, w in col_widths.items():
            ws.column_dimensions[col_l].width = w
        ws.row_dimensions[1].height = 26.0

        headers = {
            'C': 'Месяц', 'D': 'Дата', 'E': 'Время погрузки', 'F': '',
            'G': 'Направление', 'H': 'Номер Машины', 'I': 'Перевозчик',
            'J': 'Кол-во паллет ', 'K': 'Дата доставки', 'L': 'Тариф'
        }
        for col_l, h_text in headers.items():
            cell = ws[f'{col_l}1']
            cell.value = h_text
            cell.font = font_header
            cell.border = border_header
            cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)

        day_plan = plan_trips.get(d_name, {})
        city_order = list(MASTER_CITY_ORDER.get(d_name, []))
        for c in day_plan:
            if c not in city_order:
                city_order.append(c)
        
        ordered_trips = []
        city_truck_counters = Counter()
        for city in city_order:
            count = day_plan.get(city, 0)
            pal = 40 if city in ['Хабаровск', 'Ярославль'] else 33
            for _ in range(count):
                city_truck_counters[city] += 1
                tr_num = city_truck_counters[city]
                carrier, tariff = weekly_assignments.get((d_name, city, tr_num), ('ТК Сияние', 0))
                ordered_trips.append({
                    'date': d_str,
                    'dt': cur_date,
                    'day': d_short,
                    'month': d_month,
                    'city': city,
                    'truck_num': tr_num,
                    'carrier': carrier,
                    'pallets': pal,
                    'tariff': tariff
                })
                
        # Назначаем время: СТРОГО 2 МАШИНЫ В ЧАС!
        start_h = START_HOURS.get(d_name, 9)
        for i, tr in enumerate(ordered_trips):
            slot_h = start_h + (i // 2)
            time_str = f"{slot_h:02d}:00:00"
            tr['time'] = time_str
            deliv_dt = calculate_delivery_datetime(tr['dt'], time_str, tr['city'])
            tr['delivery_dt'] = deliv_dt
            tr['delivery_str'] = deliv_dt.strftime('%d.%m.%Y %H:%M')
            all_scheduled_trips.append(tr)

        row_idx = 2
        for tr in ordered_trips:
            c_a = ws[f'A{row_idx}']
            c_a.value = f'=G{row_idx}&I{row_idx}&J{row_idx}'
            c_a.font = font_key
            c_a.alignment = Alignment(horizontal='left', vertical='center')
            
            c_c = ws[f'C{row_idx}']
            c_c.value = d_month
            c_c.font = font_regular
            c_c.alignment = Alignment(horizontal='center', vertical='center')
            c_c.border = border_thin
            
            c_d = ws[f'D{row_idx}']
            c_d.value = d_str
            c_d.font = font_regular
            c_d.alignment = Alignment(horizontal='center', vertical='center')
            c_d.border = border_thin
            
            c_e = ws[f'E{row_idx}']
            c_e.value = tr['time']
            c_e.font = font_regular
            c_e.alignment = Alignment(horizontal='center', vertical='center')
            c_e.border = border_thin
            
            c_f = ws[f'F{row_idx}']
            c_f.value = d_short
            c_f.font = font_header
            c_f.alignment = Alignment(horizontal='center', vertical='center')
            c_f.border = border_thin
            
            c_g = ws[f'G{row_idx}']
            c_g.value = tr['city']
            c_g.font = font_city
            c_g.alignment = Alignment(horizontal='left', vertical='center')
            c_g.border = border_thin
            
            c_h = ws[f'H{row_idx}']
            c_h.value = tr['truck_num']
            c_h.font = font_bold_center
            c_h.alignment = Alignment(horizontal='center', vertical='center')
            c_h.border = border_thin
            
            c_i = ws[f'I{row_idx}']
            c_i.value = tr['carrier']
            c_i.font = font_carrier
            c_i.alignment = Alignment(horizontal='center', vertical='center')
            c_i.border = border_thin
            
            c_j = ws[f'J{row_idx}']
            c_j.value = tr['pallets']
            c_j.font = font_regular
            c_j.alignment = Alignment(horizontal='center', vertical='center')
            c_j.border = border_thin
            
            c_k = ws[f'K{row_idx}']
            c_k.value = tr['delivery_str']
            c_k.font = font_regular
            c_k.alignment = Alignment(horizontal='center', vertical='center')
            c_k.border = border_thin
            
            c_l = ws[f'L{row_idx}']
            c_l.value = tr['tariff']
            c_l.font = font_regular
            c_l.number_format = '#,##0 ₽'
            c_l.alignment = Alignment(horizontal='right', vertical='center')
            c_l.border = border_thin
            
            if tr['pallets'] == 40:
                for col_l in ['C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L']:
                    ws[f'{col_l}{row_idx}'].fill = fill_40p
            
            ws.row_dimensions[row_idx].height = 18.0
            row_idx += 1

        last_data_row = row_idx - 1
        if last_data_row >= 2:
            ws.auto_filter.ref = f'C1:L{last_data_row}'

        carrier_counts = Counter([t['carrier'] for t in ordered_trips if t.get('carrier')])
        if carrier_counts:
            ws['N3'] = 'Перевозчик'
            ws['N3'].font = font_header
            ws['N3'].border = border_summary_header
            ws['O3'] = 'Рейсы'
            ws['O3'].font = font_header
            ws['O3'].alignment = Alignment(horizontal='center')
            ws['O3'].border = border_summary_header
            ws['P3'] = '%'
            ws['P3'].font = font_header
            ws['P3'].alignment = Alignment(horizontal='center')
            ws['P3'].border = border_summary_header
            
            s_row = 4
            for car, cnt in carrier_counts.most_common():
                ws[f'N{s_row}'] = car
                ws[f'N{s_row}'].font = font_regular
                ws[f'N{s_row}'].border = border_thin
                
                ws[f'O{s_row}'] = f'=COUNTIF(I$2:I${last_data_row}, N{s_row})'
                ws[f'O{s_row}'].font = font_regular
                ws[f'O{s_row}'].alignment = Alignment(horizontal='center')
                ws[f'O{s_row}'].border = border_thin
                
                total_target_row = 4 + len(carrier_counts)
                ws[f'P{s_row}'] = f'=O{s_row}/O{total_target_row}'
                ws[f'P{s_row}'].font = font_regular
                ws[f'P{s_row}'].number_format = '0.0%'
                ws[f'P{s_row}'].alignment = Alignment(horizontal='right')
                ws[f'P{s_row}'].border = border_thin
                s_row += 1
                
            ws[f'N{s_row}'] = 'Всего'
            ws[f'N{s_row}'].font = font_header
            ws[f'N{s_row}'].border = border_total
            
            ws[f'O{s_row}'] = f'=SUM(O4:O{s_row-1})'
            ws[f'O{s_row}'].font = font_header
            ws[f'O{s_row}'].alignment = Alignment(horizontal='center')
            ws[f'O{s_row}'].border = border_total
            
            ws[f'P{s_row}'] = '100%'
            ws[f'P{s_row}'].font = font_header
            ws[f'P{s_row}'].alignment = Alignment(horizontal='right')
            ws[f'P{s_row}'].border = border_total

    wb_out.save(output_path)
    print(f"Таблица успешно создана: {output_path}")
    return all_scheduled_trips

def refresh_schedule_file_summary(excel_path):
    """
    Обновляет сводные таблицы (колонки N, O, P) на всех листах Excel-файла графика,
    а также актуализирует формулы колонки A (=G&I&J) и дату доставки (колонка K),
    если пользователь добавил/удалил строки или изменил перевозчиков вручную в Excel.
    """
    if not excel_path or not os.path.exists(excel_path):
        return False
        
    try:
        wb = openpyxl.load_workbook(excel_path)
    except Exception as e:
        print(f"Ошибка открытия файла для обновления сводки: {e}")
        return False
        
    modified = False
    
    font_header = Font(name='Arial', size=10, bold=True)
    font_regular = Font(name='Arial', size=10)
    font_key = Font(name='Arial', size=10, color='7F7F7F')
    border_thin = Border(
        left=Side(style='thin', color='D9D9D9'),
        right=Side(style='thin', color='D9D9D9'),
        top=Side(style='thin', color='D9D9D9'),
        bottom=Side(style='thin', color='D9D9D9')
    )
    border_summary_header = Border(
        left=Side(style='thin', color='BFBFBF'),
        right=Side(style='thin', color='BFBFBF'),
        top=Side(style='thin', color='BFBFBF'),
        bottom=Side(style='thin', color='000000')
    )
    border_total = Border(
        top=Side(style='thin', color='000000'),
        bottom=Side(style='double', color='000000')
    )
    
    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        h_c = str(ws['C1'].value or '').strip().lower()
        h_d = str(ws['D1'].value or '').strip().lower()
        h_g = str(ws['G1'].value or '').strip().lower()
        if 'месяц' not in h_c and 'дата' not in h_c and 'дата' not in h_d and 'направление' not in h_g:
            continue
            
        is_new_format = ('месяц' in h_c) or ('время' in str(ws['E1'].value or '').lower())
        
        # 1. Сканируем строки данных
        data_rows = []
        for r in range(2, ws.max_row + 1):
            if is_new_format:
                city = ws[f'G{r}'].value
                carrier = ws[f'I{r}'].value
                date_val = ws[f'D{r}'].value
            else:
                city = ws[f'G{r}'].value
                carrier = ws[f'I{r}'].value
                date_val = ws[f'C{r}'].value
                
            if not city or not str(city).strip() or 'итого' in str(city).lower() or 'всего' in str(city).lower():
                continue
            if not carrier or not str(carrier).strip():
                continue
            if not date_val:
                continue
                
            data_rows.append(r)
            
            # Актуализируем формулу в колонке A и дату доставки в K (для нового формата)
            if is_new_format:
                ws[f'A{r}'].value = f'=G{r}&I{r}&J{r}'
                ws[f'A{r}'].font = font_key
                # Если дата доставки пустая, рассчитываем её
                if not ws[f'K{r}'].value or str(ws[f'K{r}'].value).strip() == '':
                    dep_d = date_val
                    dep_t = str(ws[f'E{r}'].value or '09:00:00').strip()
                    deliv_dt = calculate_delivery_datetime(dep_d, dep_t, str(city).strip())
                    ws[f'K{r}'].value = deliv_dt.strftime('%d.%m.%Y %H:%M')
                    ws[f'K{r}'].font = font_regular
                    ws[f'K{r}'].alignment = Alignment(horizontal='center', vertical='center')
                    ws[f'K{r}'].border = border_thin
                    
        if not data_rows:
            continue
            
        last_data_row = max(data_rows)
        if is_new_format:
            ws.auto_filter.ref = f'C1:L{last_data_row}'
            
        # 2. Очищаем старую сводную таблицу в N..P
        max_clear = max(ws.max_row + 10, 35)
        for r in range(3, max_clear):
            for col_l in ['N', 'O', 'P']:
                cell = ws[f'{col_l}{r}']
                cell.value = None
                cell.border = None
                
        # 3. Подсчитываем перевозчиков по актуальным данным строк
        carriers_in_sheet = [str(ws[f'I{r}'].value).strip() for r in data_rows if ws[f'I{r}'].value]
        carrier_counts = Counter(carriers_in_sheet)
        
        if carrier_counts:
            ws['N3'] = 'Перевозчик'
            ws['N3'].font = font_header
            ws['N3'].border = border_summary_header
            ws['O3'] = 'Рейсы'
            ws['O3'].font = font_header
            ws['O3'].alignment = Alignment(horizontal='center')
            ws['O3'].border = border_summary_header
            ws['P3'] = '%'
            ws['P3'].font = font_header
            ws['P3'].alignment = Alignment(horizontal='center')
            ws['P3'].border = border_summary_header
            
            s_row = 4
            for car, cnt in carrier_counts.most_common():
                ws[f'N{s_row}'] = car
                ws[f'N{s_row}'].font = font_regular
                ws[f'N{s_row}'].border = border_thin
                
                ws[f'O{s_row}'] = f'=COUNTIF(I$2:I${last_data_row}, N{s_row})'
                ws[f'O{s_row}'].font = font_regular
                ws[f'O{s_row}'].alignment = Alignment(horizontal='center')
                ws[f'O{s_row}'].border = border_thin
                
                total_target_row = 4 + len(carrier_counts)
                ws[f'P{s_row}'] = f'=O{s_row}/O{total_target_row}'
                ws[f'P{s_row}'].font = font_regular
                ws[f'P{s_row}'].number_format = '0.0%'
                ws[f'P{s_row}'].alignment = Alignment(horizontal='right')
                ws[f'P{s_row}'].border = border_thin
                s_row += 1
                
            ws[f'N{s_row}'] = 'Всего'
            ws[f'N{s_row}'].font = font_header
            ws[f'N{s_row}'].border = border_total
            
            ws[f'O{s_row}'] = f'=SUM(O4:O{s_row-1})'
            ws[f'O{s_row}'].font = font_header
            ws[f'O{s_row}'].alignment = Alignment(horizontal='center')
            ws[f'O{s_row}'].border = border_total
            
            ws[f'P{s_row}'] = '100%'
            ws[f'P{s_row}'].font = font_header
            ws[f'P{s_row}'].alignment = Alignment(horizontal='right')
            ws[f'P{s_row}'].border = border_total
            
        modified = True
        
    if modified:
        try:
            wb.save(excel_path)
            print(f"Сводные таблицы успешно обновлены в файле: {excel_path}")
        except Exception as e:
            print(f"Предупреждение: не удалось сохранить файл {excel_path}: {e}")
            
    wb.close()
    return True

def generate_carrier_html(carrier_name, trips, start_date_str, end_date_str):
    """Генерирует индивидуальное конфиденциальное HTML-письмо для конкретного перевозчика"""
    rows_html = ""
    for tr in trips:
        bg_color = "#e2efda" if tr.get('pallets') == 40 else "#ffffff"
        deliv_str = tr.get('delivery_str') or tr.get('deliv_date') or ''
        day_str = tr.get('day') or ''
        city_str = tr.get('city') or tr.get('route') or ''
        time_str = str(tr.get('time', ''))[:5]
        truck_str = str(tr.get('truck_num', ''))
        pal_str = str(tr.get('pallets', ''))
        rows_html += f"""
        <tr style="background-color: {bg_color}; text-align: center;">
            <td style="padding: 8px; border: 1px solid #d9d9d9;">{tr.get('date', '')}</td>
            <td style="padding: 8px; border: 1px solid #d9d9d9; font-weight: bold;">{day_str}</td>
            <td style="padding: 8px; border: 1px solid #d9d9d9; font-weight: bold; color: #1f4e79;">{time_str}</td>
            <td style="padding: 8px; border: 1px solid #d9d9d9; text-align: left; font-weight: bold;">{city_str}</td>
            <td style="padding: 8px; border: 1px solid #d9d9d9;">№ {truck_str}</td>
            <td style="padding: 8px; border: 1px solid #d9d9d9;">{pal_str} пал.</td>
            <td style="padding: 8px; border: 1px solid #d9d9d9; font-weight: bold; color: #1f4e79;">{deliv_str}</td>
        </tr>
        """
        
    period_str = f"{start_date_str} – {end_date_str}" if start_date_str != end_date_str else start_date_str

    html = f"""
    <!DOCTYPE html>
    <html>
    <head>
        <meta charset="utf-8">
        <style>
            body {{ font-family: Arial, sans-serif; font-size: 14px; color: #333333; line-height: 1.5; }}
            table {{ border-collapse: collapse; width: 100%; max-width: 720px; margin-top: 15px; margin-bottom: 20px; }}
            th {{ background-color: #2f5597; color: #ffffff; padding: 10px; border: 1px solid #2f5597; text-align: center; }}
            .notice {{ background-color: #fff2cc; border-left: 4px solid #d6b656; padding: 12px; margin: 15px 0; font-size: 13px; }}
            .footer {{ font-size: 12px; color: #7f7f7f; margin-top: 25px; border-top: 1px solid #e0e0e0; padding-top: 10px; }}
        </style>
    </head>
    <body>
        <p>Добрый день!</p>
        <p>Направляем согласованный график погрузки на период <b>{period_str}</b>.<br>
        Просим ознакомиться с графиком и обеспечить его соблюдение.</p>
        
        <p><b>Важные условия регламента:</b><br>
        • Погрузка на складе осуществляется строго из расчёта 2 машины в час.<br>
        • Просьба обеспечить своевременное прибытие ТС к назначенному тайм-слоту, без опозданий.<br>
        • В случае возникновения задержек в пути необходимо оперативно информировать ответственных лиц, чтобы своевременно скорректировать дальнейшее планирование.</p>
        
        <table>
            <thead>
                <tr>
                    <th>Дата погрузки</th>
                    <th>День</th>
                    <th>Время</th>
                    <th>Направление</th>
                    <th>№ ТС</th>
                    <th>Паллеты</th>
                    <th>Дата доставки</th>
                </tr>
            </thead>
            <tbody>
                {rows_html}
            </tbody>
        </table>
        
        <p>С уважением,<br>
        <b>Отдел логистики РЦ «Черная Грязь»</b><br>
        ФК «ПУЛЬС»</p>
        
        <div class="footer">
            Данное сообщение сформировано автоматически и предназначено исключительно для перевозчика {carrier_name}. Конфиденциально.
        </div>
    </body>
    </html>
    """
    return html

def send_emails_via_outlook(all_trips, target_email="n.rozhkov@puls.ru", draft_mode=False):
    """
    Отправляет индивидуальные письма через установленный Microsoft Outlook.
    Каждый перевозчик получает ТОЛЬКО свое расписание!
    В тестовом режиме все письма отправляются на адрес target_email.
    """
    try:
        import win32com.client as win32
    except ImportError:
        print("\n[ВНИМАНИЕ] Библиотека 'pywin32' не установлена.")
        print("Для отправки через Outlook выполните: pip install pywin32")
        print("\nПоказываем текстовый предпросмотр писем:\n")
        preview_emails_console(all_trips, target_email)
        return

    # Группируем рейсы по перевозчикам
    carrier_trips = defaultdict(list)
    for tr in all_trips:
        car = tr['carrier']
        carrier_trips[car].append(tr)

    dates = [tr['date'] for tr in all_trips]
    start_d = dates[0] if dates else ''
    end_d = dates[-1] if dates else ''

    print("\n" + "="*70)
    print(f"ПОДКЛЮЧЕНИЕ К OUTLOOK: Подготовка {len(carrier_trips)} писем (получатель: {target_email})")
    print("="*70)

    try:
        outlook = win32.Dispatch('outlook.application')
    except Exception as e:
        print(f"Ошибка подключения к приложению Outlook: {e}")
        print("Убедитесь, что Microsoft Outlook запущен на компьютере.")
        return

    base_dir = os.path.dirname(os.path.abspath(__file__))
    contacts = {}
    for fname in ['contacts_schedules.json', 'carriers_contacts.json']:
        contacts_file = os.path.join(base_dir, fname)
        if os.path.exists(contacts_file):
            try:
                with open(contacts_file, 'r', encoding='utf-8') as f:
                    contacts = json.load(f)
                    break
            except Exception:
                pass

    sent_count = 0
    for carrier, c_trips in sorted(carrier_trips.items()):
        # Сортируем рейсы перевозчика по хронологии
        sorted_trips = sorted(c_trips, key=lambda x: (x.get('dt') or datetime.strptime(x['date'], '%d.%m.%Y'), x['time']))
        
        recipient = contacts.get(carrier, target_email)
        mail = outlook.CreateItem(0) # 0 = olMailItem
        mail.To = recipient
        mail.Subject = f"[{carrier}] График погрузки РЦ Черная Грязь ({start_d} - {end_d})"
        mail.HTMLBody = generate_carrier_html(carrier, sorted_trips, start_d, end_d)
        
        if draft_mode:
            mail.Save() # Сохранить в папку "Черновики" Outlook
            print(f"  [ЧЕРНОВИК СОХРАНЕН] -> {carrier:25s} ({len(sorted_trips)} рейсов)")
        else:
            mail.Send() # Отправить
            print(f"  [ОТПРАВЛЕНО В OUTLOOK] -> {carrier:25s} ({len(sorted_trips)} рейсов)")
            
        sent_count += 1

    mode_text = "сохранены в Черновиках" if draft_mode else "успешно отправлены"
    print("="*70)
    print(f"ИТОГО: {sent_count} индивидуальных писем {mode_text} через Outlook на адрес {target_email}!")
    print("Каждое письмо содержит ТОЛЬКО рейсы соответствующего перевозчика.")

def preview_emails_console(all_trips, target_email):
    carrier_trips = defaultdict(list)
    for tr in all_trips:
        carrier_trips[tr['carrier']].append(tr)

    dates = [tr['date'] for tr in all_trips]
    start_d = dates[0] if dates else ''
    end_d = dates[-1] if dates else ''

    for carrier, c_trips in sorted(carrier_trips.items()):
        print(f"\n[ТЕМА]: [{carrier}] График погрузки РЦ Черная Грязь ({start_d} - {end_d})")
        print(f"[КОМУ]: {target_email} (в боевом режиме: диспетчер {carrier})")
        print(f"Рейсы перевозчика ({len(c_trips)} шт.):")
        for tr in c_trips:
            deliv_info = f" -> Доставка: {tr.get('delivery_str')}" if tr.get('delivery_str') else ""
            print(f"  • {tr['date']} ({tr['day']}) в {tr['time'][:5]} -> {tr['city']:15s} (ТС #{tr['truck_num']}, {tr['pallets']} пал.){deliv_info}")


def get_schedules_dir():
    """
    Возвращает путь к папке 'Готовые недельные графики'.
    Приоритет: корпоративный сетевой/облачный диск L:
    L:\\ОТДЕЛЫ\\ТРАНСПОРТНЫЙ ОТДЕЛ\\Рожков\\РК недельные графики + заявки\\Готовые недельные графики
    Резерв: Рабочий стол / Готовые недельные графики
    """
    network_base = r"L:\ОТДЕЛЫ\ТРАНСПОРТНЫЙ ОТДЕЛ\Рожков\РК недельные графики + заявки\Готовые недельные графики"
    try:
        drive = os.path.splitdrive(network_base)[0]
        if drive and os.path.exists(drive + "\\"):
            os.makedirs(network_base, exist_ok=True)
            return network_base
        elif os.path.exists(network_base):
            return network_base
    except Exception:
        pass

    user_home = os.path.expanduser("~")
    desktop_candidates = [
        os.path.join(user_home, "OneDrive", "Рабочий стол"),
        os.path.join(user_home, "OneDrive", "Desktop"),
        os.path.join(user_home, "Рабочий стол"),
        os.path.join(user_home, "Desktop"),
    ]
    for d in desktop_candidates:
        if os.path.exists(d):
            folder = os.path.join(d, "Готовые недельные графики")
            os.makedirs(folder, exist_ok=True)
            return folder

    fallback = os.path.join(os.path.dirname(os.path.abspath(__file__)), "Готовые недельные графики")
    os.makedirs(fallback, exist_ok=True)
    return fallback

if __name__ == '__main__':
    base_dir = os.path.dirname(os.path.abspath(__file__))
    default_plan = find_default_plan_file(base_dir)
    schedules_dir = get_schedules_dir()
    default_out = os.path.join(schedules_dir, 'Недельные графики (готовый).xlsx')
    
    parser = argparse.ArgumentParser(description='Генератор графика отгрузок из плана .xlsm и отправка в Outlook')
    parser.add_argument('--plan', default=default_plan, help='Путь к файлу График_отгрузки_филиалов_неделя_2.xlsm')
    parser.add_argument('--out', default=default_out, help='Путь к результирующему файлу Excel')
    parser.add_argument('--send-outlook', action='store_true', help='Отправить индивидуальные письма через Outlook')
    parser.add_argument('--draft', action='store_true', help='Сохранить письма в черновики Outlook вместо отправки')
    parser.add_argument('--email', default='n.rozhkov@puls.ru', help='Тестовый адрес получателя писем')
    
    args = parser.parse_args()
    
    if not args.plan or not os.path.exists(args.plan):
        print(f"ОШИБКА: Файл плана отгрузок (*.xlsm) не найден в папке {base_dir}!")
        sys.exit(1)
        
    print(f"Обработка плана отгрузок: {args.plan}")
    plan_trips, start_date = parse_plan_file(args.plan)
    all_trips = build_schedule_from_plan(plan_trips, start_date, args.out)
    
    # Если указан флаг отправки
    if args.send_outlook or args.draft:
        send_emails_via_outlook(all_trips, target_email=args.email, draft_mode=args.draft)
    else:
        print("\n[ПОДСКАЗКА]:")
        print(f"1. Чтобы сохранить черновики в Outlook для проверки:")
        print(f"   python generate_schedule.py --draft")
        print(f"2. Чтобы сразу отправить письма на {args.email}:")
        print(f"   python generate_schedule.py --send-outlook")
