#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Автоматизированный генератор графика отгрузок (РЦ Черная Грязь).

Основной входной файл: 'График_отгрузки_филиалов_неделя_2.xlsm' (или любой .xlsm в текущей папке).
Опциональный входной файл: 'Филиалы...xlsx' (если есть оперативные ручные правки).

Что делает скрипт:
1. Автоматически находит файл плана (*.xlsm) в текущей директории.
2. Определяет даты недели из заголовка (например, 28.09-04.10).
3. Разворачивает матрицу отгрузок по дням недели в линейный реестр.
4. Выстраивает правильный логистический порядок подачи городов (ближние -> дальние).
5. Строго соблюдает ограничение: РОВНО 2 МАШИНЫ В ЧАС (09:00 - 2 машины, 10:00 - 2 машины и т.д.).
6. Назначает проверенных перевозчиков по направлениям.
7. Подсвечивает 40-паллетные машины мягким зеленым цветом (#E2EFDA).
8. Создает отдельный лист на каждый день недели (28.09 Пн, 29.09 Вт, ...).
9. Выводит сводную аналитику долей ТК справа (M:O).
"""

import sys
import os
import re
import glob
import argparse
from datetime import datetime, timedelta
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

# Стартовый час начала погрузки по дням
START_HOURS = {
    'Пн': 9,
    'Вт': 9,
    'Ср': 9,
    'Чт': 9,
    'Пт': 9,
    'Сб': 5, # Суббота: ранние утренние окна
    'Вс': 9
}

# Базовое закрепление проверенных перевозчиков по направлениям
def get_carrier_for_trip(city, truck_num, day_name):
    if city == 'Ярославль':
        if day_name in ['Вт', 'Чт']:
            return 'ИП Мельник' if truck_num == 1 else 'ИП Гусманов'
        else:
            return 'ИП Коршунов' if truck_num == 1 else 'ИП Гусманов'
    elif city == 'Казань':
        if day_name == 'Ср':
            return ['НОРДЛАЙН', 'ТК Сияние', 'Агро-Авто'][min(truck_num-1, 2)]
        elif day_name == 'Вт':
            return 'ТК Сияние' if truck_num == 1 else 'Агро-Авто'
        else:
            return 'ТК Сияние'
    elif city == 'Краснодар':
        carriers = ['ТК Сияние', 'АО Национальный', 'Буш-Авто', 'Агро-Авто Отрада']
        return carriers[min(truck_num-1, len(carriers)-1)]
    elif city == 'Екатеринбург':
        carriers = ['ТК Сияние', 'АО Национальный', 'Веб-Логистика', 'Азимут']
        return carriers[min(truck_num-1, len(carriers)-1)]
    elif city == 'Новосибирск':
        return 'Виллайн' if truck_num == 1 else 'ЕманТрансАвто'
    elif city == 'Красноярск':
        return 'ТК Сияние' if day_name == 'Пт' else 'Виллайн'
    elif city == 'Иркутск':
        return 'Азимут' if day_name == 'Пт' else 'ЕманТрансАвто'
    elif city == 'Санкт-Петербург':
        return 'ТК Сияние' if truck_num == 1 else 'НОРДЛАЙН'
    elif city == 'Хабаровск':
        return 'Азимут'
    elif city == 'Брянск':
        return 'НОРДЛАЙН'
    elif city in ['Волгоград', 'Воронеж', 'Самара']:
        return 'ТК Сияние'
    return 'ТК Сияние'

def find_default_plan_file(base_dir):
    """Ищет файл плана .xlsm в текущей или указанной папке"""
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
    """Считывает план отгрузок из .xlsm файла"""
    wb = openpyxl.load_workbook(plan_path, data_only=True)
    # Ищем лист с планом (обычно первый или с датами)
    sheet_name = wb.sheetnames[0]
    for s in wb.sheetnames:
        if '.' in s and '-' in s:
            sheet_name = s
            break
    ws = wb[sheet_name]
    
    header_val = str(ws['A1'].value or '')
    start_date = None
    match = re.search(r'(\d{1,2}\.\d{2})\s*-\s*(\d{1,2}\.\d{2})', header_val)
    if match:
        start_d_str = match.group(1)
        # Год определяем текущий или 2026
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

def build_schedule_from_plan(plan_trips, start_date, output_path):
    """Формирует итоговый Excel файл на 7 листов по дням строго по 2 машины в час"""
    wb_out = openpyxl.Workbook()
    wb_out.remove(wb_out.active)
    
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
        'A': 28, 'B': 5, 'C': 13, 'D': 6, 'E': 12, 'F': 15,
        'G': 20, 'H': 15, 'I': 26, 'J': 12, 'K': 16, 'L': 5,
        'M': 25, 'N': 10, 'O': 10
    }

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
            'C': 'Дата', 'D': '', 'E': 'Месяц', 'F': 'Номер Машины',
            'G': 'Направление', 'H': 'Кол-во паллет ', 'I': 'Перевозчик',
            'J': 'Тариф', 'K': 'Время погрузки'
        }
        for col_l, h_text in headers.items():
            cell = ws[f'{col_l}1']
            cell.value = h_text
            cell.font = font_header
            cell.border = border_header
            cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)

        # Собираем список рейсов дня в правильном логистическом порядке
        day_plan = plan_trips.get(d_name, {})
        city_order = MASTER_CITY_ORDER.get(d_name, [])
        
        # Добавляем города по порядку
        ordered_trips = []
        city_truck_counters = Counter()
        for city in city_order:
            count = day_plan.get(city, 0)
            pal = 40 if city in ['Хабаровск', 'Ярославль'] else 33
            for _ in range(count):
                city_truck_counters[city] += 1
                tr_num = city_truck_counters[city]
                carrier = get_carrier_for_trip(city, tr_num, d_name)
                ordered_trips.append({
                    'city': city,
                    'truck_num': tr_num,
                    'carrier': carrier,
                    'pallets': pal
                })
                
        # Назначаем время: СТРОГО 2 МАШИНЫ В ЧАС!
        start_h = START_HOURS.get(d_name, 9)
        for i, tr in enumerate(ordered_trips):
            slot_h = start_h + (i // 2)
            tr['time'] = f"{slot_h:02d}:00:00"

        # Записываем строки в лист
        row_idx = 2
        for tr in ordered_trips:
            c_a = ws[f'A{row_idx}']
            c_a.value = f'=G{row_idx}&I{row_idx}&H{row_idx}'
            c_a.font = font_key
            c_a.alignment = Alignment(horizontal='left', vertical='center')
            
            c_c = ws[f'C{row_idx}']
            c_c.value = d_str
            c_c.font = font_regular
            c_c.alignment = Alignment(horizontal='center', vertical='center')
            c_c.border = border_thin
            
            c_d = ws[f'D{row_idx}']
            c_d.value = d_short
            c_d.font = font_header
            c_d.alignment = Alignment(horizontal='center', vertical='center')
            c_d.border = border_thin
            
            c_e = ws[f'E{row_idx}']
            c_e.value = d_month
            c_e.font = font_regular
            c_e.alignment = Alignment(horizontal='center', vertical='center')
            c_e.border = border_thin
            
            c_f = ws[f'F{row_idx}']
            c_f.value = tr['truck_num']
            c_f.font = font_bold_center
            c_f.alignment = Alignment(horizontal='center', vertical='center')
            c_f.border = border_thin
            
            c_g = ws[f'G{row_idx}']
            c_g.value = tr['city']
            c_g.font = font_city
            c_g.alignment = Alignment(horizontal='left', vertical='center')
            c_g.border = border_thin
            
            c_h = ws[f'H{row_idx}']
            c_h.value = tr['pallets']
            c_h.font = font_regular
            c_h.alignment = Alignment(horizontal='center', vertical='center')
            c_h.border = border_thin
            
            c_i = ws[f'I{row_idx}']
            c_i.value = tr['carrier']
            c_i.font = font_carrier
            c_i.alignment = Alignment(horizontal='center', vertical='center')
            c_i.border = border_thin
            
            c_j = ws[f'J{row_idx}']
            c_j.value = None # Тариф пустой
            c_j.border = border_thin
            
            c_k = ws[f'K{row_idx}']
            c_k.value = tr['time']
            c_k.font = font_regular
            c_k.alignment = Alignment(horizontal='center', vertical='center')
            c_k.border = border_thin
            
            if tr['pallets'] == 40:
                for col_l in ['C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K']:
                    ws[f'{col_l}{row_idx}'].fill = fill_40p
            
            ws.row_dimensions[row_idx].height = 18.0
            row_idx += 1

        last_data_row = row_idx - 1
        if last_data_row >= 2:
            ws.auto_filter.ref = f'C1:K{last_data_row}'

        # Сводная таблица ТК справа
        carrier_counts = Counter([t['carrier'] for t in ordered_trips if t.get('carrier')])
        if carrier_counts:
            ws['M3'] = 'Перевозчик'
            ws['M3'].font = font_header
            ws['M3'].border = border_summary_header
            ws['N3'] = 'Рейсы'
            ws['N3'].font = font_header
            ws['N3'].alignment = Alignment(horizontal='center')
            ws['N3'].border = border_summary_header
            ws['O3'] = '%'
            ws['O3'].font = font_header
            ws['O3'].alignment = Alignment(horizontal='center')
            ws['O3'].border = border_summary_header
            
            s_row = 4
            for car, cnt in carrier_counts.most_common():
                ws[f'M{s_row}'] = car
                ws[f'M{s_row}'].font = font_regular
                ws[f'M{s_row}'].border = border_thin
                
                ws[f'N{s_row}'] = cnt
                ws[f'N{s_row}'].font = font_regular
                ws[f'N{s_row}'].alignment = Alignment(horizontal='center')
                ws[f'N{s_row}'].border = border_thin
                
                total_target_row = 4 + len(carrier_counts)
                ws[f'O{s_row}'] = f'=N{s_row}/N{total_target_row}'
                ws[f'O{s_row}'].font = font_regular
                ws[f'O{s_row}'].number_format = '0.0%'
                ws[f'O{s_row}'].alignment = Alignment(horizontal='right')
                ws[f'O{s_row}'].border = border_thin
                s_row += 1
                
            ws[f'M{s_row}'] = 'Всего'
            ws[f'M{s_row}'].font = font_header
            ws[f'M{s_row}'].border = border_total
            
            ws[f'N{s_row}'] = f'=SUM(N4:N{s_row-1})'
            ws[f'N{s_row}'].font = font_header
            ws[f'N{s_row}'].alignment = Alignment(horizontal='center')
            ws[f'N{s_row}'].border = border_total
            
            ws[f'O{s_row}'] = '100%'
            ws[f'O{s_row}'].font = font_header
            ws[f'O{s_row}'].alignment = Alignment(horizontal='right')
            ws[f'O{s_row}'].border = border_total

    wb_out.save(output_path)
    print(f"Готово! Файл успешно создан: {output_path}")

if __name__ == '__main__':
    base_dir = os.path.dirname(os.path.abspath(__file__))
    default_plan = find_default_plan_file(base_dir)
    default_out = os.path.join(base_dir, 'Недельные графики (готовый).xlsx')
    
    parser = argparse.ArgumentParser(description='Генератор графика отгрузок из файла плана .xlsm')
    parser.add_argument('--plan', default=default_plan, help='Путь к файлу График_отгрузки_филиалов_неделя_2.xlsm')
    parser.add_argument('--out', default=default_out, help='Путь к результирующему файлу Excel')
    
    args = parser.parse_args()
    
    if not args.plan or not os.path.exists(args.plan):
        print(f"ОШИБКА: Файл плана отгрузок не найден в папке {base_dir}!")
        print("Пожалуйста, убедитесь, что файл 'График_отгрузки_филиалов_неделя_2.xlsm' лежит в папке проекта.")
        sys.exit(1)
        
    print(f"Обработка плана: {args.plan}")
    plan_trips, start_date = parse_plan_file(args.plan)
    build_schedule_from_plan(plan_trips, start_date, args.out)
