#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
generate_orders.py — Модуль формирования транспортных заявок для перевозчиков.

Логика:
1. Загрузка исходного реестра (например, Книга4.xlsx).
2. Определение перевозчика по водителю из База перевозчиков.xlsx (лист 'Лист_1').
3. Суммирование паллет для одинаковых 'Адрес ссылка' у одного перевозчика.
4. Создание сводной таблицы Excel с добавленной колонкой 'Перевозчик' и объединенными паллетами.
5. Заполнение Word-шаблона 'Шаблон.docx' для каждого перевозчика:
   - Кому: Перевозчик
   - Дата: текущая дата (например, 30 сентября 2026)
   - Тип транспортного средства: сумма паллет, РЕФ и температурный режим
   - Адрес погрузки: Адрес ссылка, Адрес
   - Дата и время погрузки: Дата отгрузки к 9:00
   - Дата и время выгрузки: Дата отгрузки, по прибытию
6. Сохранение каждого файла в папку Заявки:
   Основной путь: L:\\ОТДЕЛЫ\\ТРАНСПОРТНЫЙ ОТДЕЛ\\Рожков\\РК недельные графики + заявки\\Заявки
   Формат имени: <Перевозчик> <ДД.ММ>.docx
"""

import os
import sys
import re
import datetime
import zipfile
from collections import defaultdict
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import docx
from docx.enum.text import WD_ALIGN_PARAGRAPH

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

RUSSIAN_MONTHS = [
    "января", "февраля", "марта", "апреля", "мая", "июня",
    "июля", "августа", "сентября", "октября", "ноября", "декабря"
]

def get_orders_dir():
    """
    Возвращает путь к целевой папке 'Заявки'.
    Приоритет: сетевой диск L:
    L:\\ОТДЕЛЫ\\ТРАНСПОРТНЫЙ ОТДЕЛ\\Рожков\\РК недельные графики + заявки\\Заявки
    Резерв: Рабочий стол / Заявки или локальная папка.
    """
    network_base = r"L:\ОТДЕЛЫ\ТРАНСПОРТНЫЙ ОТДЕЛ\Рожков\РК недельные графики + заявки\Заявки"
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
            folder = os.path.join(d, "Заявки")
            os.makedirs(folder, exist_ok=True)
            return folder

    fallback = os.path.join(BASE_DIR, "Заявки")
    os.makedirs(fallback, exist_ok=True)
    return fallback

def sanitize_filename(name: str) -> str:
    """Очищает строку от запрещенных символов в именах файлов ОС Windows и Mac"""
    for ch in ['"', '<', '>', ':', '/', '\\', '|', '?', '*']:
        name = name.replace(ch, '')
    return re.sub(r'\s+', ' ', name).strip()

def format_pallets(val) -> str:
    """Форматирует число паллет (20.0 -> '20', 20.8 -> '20.8')"""
    try:
        f = float(val)
        if f.is_integer():
            return str(int(f))
        return f"{f:g}"
    except (ValueError, TypeError):
        return str(val) if val is not None else "0"

def get_russian_date_str(dt: datetime.date = None) -> str:
    """Возвращает дату в формате '30 сентября 2026'"""
    if dt is None:
        dt = datetime.date.today()
    month_name = RUSSIAN_MONTHS[dt.month - 1]
    return f"{dt.day} {month_name} {dt.year}"

def load_drivers_database(db_path: str = None) -> dict:
    """
    Загружает базу перевозчиков из База перевозчиков.xlsx (лист Лист_1).
    Возвращает словарь { 'ФИО Водителя': 'Перевозчик' }.
    При наличии дубликатов выбирается перевозчик с наибольшим числом записей (мода).
    """
    if db_path is None:
        db_path = os.path.join(BASE_DIR, "База перевозчиков.xlsx")

    if not os.path.exists(db_path):
        raise FileNotFoundError(f"Файл базы перевозчиков не найден: {db_path}")

    wb = openpyxl.load_workbook(db_path, data_only=True)
    ws = wb["Лист_1"] if "Лист_1" in wb.sheetnames else wb.active

    driver_counts = defaultdict(lambda: defaultdict(int))
    for row in ws.iter_rows(values_only=True):
        if not row or len(row) < 2:
            continue
        driver_val = row[0]
        carrier_val = row[1]
        if not driver_val or not carrier_val:
            continue
        driver_clean = re.sub(r'\s+', ' ', str(driver_val)).strip()
        carrier_clean = re.sub(r'\s+', ' ', str(carrier_val)).strip()
        if driver_clean.lower() == "водитель" or carrier_clean.lower() == "перевозчик":
            continue
        driver_counts[driver_clean][carrier_clean] += 1

    driver_to_carrier = {}
    for driver, counts in driver_counts.items():
        # Выбираем перевозчика с максимальным числом совпадений
        best_carrier = max(counts.items(), key=lambda x: x[1])[0]
        driver_to_carrier[driver] = best_carrier

    return driver_to_carrier

def parse_orders_excel(input_path: str, driver_to_carrier: dict) -> list:
    """
    Читает строки из суточного файла отгрузок (Книга4.xlsx),
    сопоставляет водителей с перевозчиками и нормализует типы данных.
    """
    wb = openpyxl.load_workbook(input_path, data_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        return []

    header = [str(c).strip() if c is not None else "" for c in rows[0]]

    def find_idx(candidates, exclude_idx=-1):
        for cand in candidates:
            for i, h in enumerate(header):
                if i != exclude_idx and cand.lower() in h.lower():
                    return i
        return -1

    idx_ref = find_idx(["адрес ссылка", "ссылка", "клиент"])
    idx_addr = find_idx(["адрес"], exclude_idx=idx_ref)
    idx_pal = find_idx(["паллет", "кол-во", "количество"])
    idx_temp = find_idx(["режим", "термо"])
    idx_date = find_idx(["дата отгрузки", "дата"])
    idx_driver = find_idx(["водитель", "фио"])
    idx_tc = find_idx(["тс", "машина", "автомобиль"])

    raw_items = []
    for r in rows[1:]:
        if not any(r):
            continue

        driver_raw = str(r[idx_driver]).strip() if idx_driver >= 0 and r[idx_driver] else ""
        driver_clean = re.sub(r'\s+', ' ', driver_raw).strip()

        carrier = driver_to_carrier.get(driver_clean)
        if not carrier:
            # Попробуем нестрогий поиск без учета регистра
            for d_name, c_name in driver_to_carrier.items():
                if d_name.lower() == driver_clean.lower():
                    carrier = c_name
                    break
        if not carrier:
            carrier = "Неизвестный перевозчик"

        pal_val = 0.0
        if idx_pal >= 0 and r[idx_pal] is not None:
            try:
                pal_val = float(r[idx_pal])
            except (ValueError, TypeError):
                pal_val = 0.0

        date_val = r[idx_date] if idx_date >= 0 else None
        if isinstance(date_val, (datetime.date, datetime.datetime)):
            date_str = date_val.strftime("%d.%m.%Y")
        elif date_val:
            date_str = str(date_val).strip()
        else:
            date_str = datetime.date.today().strftime("%d.%m.%Y")

        item = {
            "addr_ref": str(r[idx_ref]).strip() if idx_ref >= 0 and r[idx_ref] else "",
            "addr": str(r[idx_addr]).strip() if idx_addr >= 0 and r[idx_addr] else "",
            "pallets": pal_val,
            "temp": str(r[idx_temp]).strip() if idx_temp >= 0 and r[idx_temp] else "",
            "date": date_str,
            "driver": driver_clean,
            "tc": str(r[idx_tc]).strip() if idx_tc >= 0 and r[idx_tc] else "",
            "carrier": carrier
        }
        raw_items.append(item)

    return raw_items

def aggregate_orders(raw_items: list) -> dict:
    """
    Группирует данные:
    1. По перевозчику.
    2. Внутри перевозчика схлопывает одинаковые 'Адрес ссылка', суммируя 'Количество паллет'.
    Возвращает:
      {
         carrier_name: [
             {
                 'addr_ref': ...,
                 'addr': ...,
                 'pallets': sum_pallets,
                 'temp': ...,
                 'date': ...,
                 'drivers': [...],
                 'tcs': [...]
             },
             ...
         ]
      }
    """
    carriers_data = defaultdict(lambda: defaultdict(lambda: {
        "addr_ref": "",
        "addr": "",
        "pallets": 0.0,
        "temps": set(),
        "dates": set(),
        "drivers": set(),
        "tcs": set()
    }))

    for item in raw_items:
        c = item["carrier"]
        ref = item["addr_ref"]
        group = carriers_data[c][ref]
        group["addr_ref"] = ref
        if item["addr"] and not group["addr"]:
            group["addr"] = item["addr"]
        group["pallets"] += item["pallets"]
        if item["temp"]:
            group["temps"].add(item["temp"])
        if item["date"]:
            group["dates"].add(item["date"])
        if item["driver"]:
            group["drivers"].add(item["driver"])
        if item["tc"]:
            group["tcs"].add(item["tc"])

    result = {}
    for carrier, refs in carriers_data.items():
        carrier_items = []
        for ref, g in refs.items():
            carrier_items.append({
                "addr_ref": g["addr_ref"],
                "addr": g["addr"],
                "pallets": g["pallets"],
                "temp": ", ".join(sorted(g["temps"])) if g["temps"] else "+15+25",
                "date": sorted(list(g["dates"]))[0] if g["dates"] else datetime.date.today().strftime("%d.%m.%Y"),
                "all_dates": sorted(list(g["dates"])),
                "drivers": sorted(list(g["drivers"])),
                "tcs": sorted(list(g["tcs"]))
            })
        result[carrier] = carrier_items

    return result

def save_aggregated_excel(aggregated_orders: dict, output_path: str):
    """
    Формирует и сохраняет красивый Excel-файл со сводной таблицей заявок:
    Колонки: Адрес ссылка, Адрес, Количество паллет, Режим термоперевозки, Дата отгрузки, Водитель, ТС, Перевозчик
    """
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Сводка заявок"

    headers = [
        "Адрес ссылка",
        "Адрес",
        "Количество паллет",
        "Режим термоперевозки",
        "Дата отгрузки",
        "Водитель",
        "ТС",
        "Перевозчик"
    ]
    ws.append(headers)

    header_font = Font(name="Arial", size=11, bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")
    align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    align_left = Alignment(horizontal="left", vertical="center", wrap_text=True)
    align_right = Alignment(horizontal="right", vertical="center")
    
    thin_border = Border(
        left=Side(style='thin', color='CBD5E1'),
        right=Side(style='thin', color='CBD5E1'),
        top=Side(style='thin', color='CBD5E1'),
        bottom=Side(style='thin', color='CBD5E1')
    )

    for col_idx in range(1, len(headers) + 1):
        cell = ws.cell(row=1, column=col_idx)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = align_center
        cell.border = thin_border
    ws.row_dimensions[1].height = 28

    current_row = 2
    for carrier, items in sorted(aggregated_orders.items()):
        for it in items:
            drivers_str = ", ".join(it["drivers"])
            tcs_str = ", ".join(it["tcs"])
            ws.append([
                it["addr_ref"],
                it["addr"],
                it["pallets"],
                it["temp"],
                it["date"],
                drivers_str,
                tcs_str,
                carrier
            ])
            
            row_cells = ws[current_row]
            for idx, c in enumerate(row_cells):
                c.border = thin_border
                c.font = Font(name="Arial", size=10)
                if idx in (0, 1, 5, 6, 7):
                    c.alignment = align_left
                elif idx in (2,):
                    c.alignment = align_right
                    c.number_format = '#,##0.#'
                else:
                    c.alignment = align_center
            ws.row_dimensions[current_row].height = 22
            current_row += 1

    # Автоподбор ширины колонок
    for col in ws.columns:
        col_letter = get_column_letter(col[0].column)
        max_len = 0
        for cell in col:
            val_str = str(cell.value or "")
            if len(val_str) > max_len:
                max_len = len(val_str)
        ws.column_dimensions[col_letter].width = min(max(max_len + 3, 12), 50)

    wb.save(output_path)

def fill_order_docx(
    template_path: str,
    carrier: str,
    items: list,
    output_docx_path: str,
    today_date_str: str = None
):
    """
    Заполняет Шаблон.docx для конкретного перевозчика.
    """
    if today_date_str is None:
        today_date_str = get_russian_date_str()

    doc = docx.Document(template_path)

    # 1. Сумма паллет
    total_pallets = sum(it["pallets"] for it in items)
    pallets_formatted = format_pallets(total_pallets)

    # 2. Температурный режим
    temps = set()
    for it in items:
        if it["temp"]:
            for t in it["temp"].split(","):
                t_clean = t.strip()
                if t_clean:
                    temps.add(t_clean)
    temp_str = ", ".join(sorted(temps)) if temps else "+15+25"
    type_ts_value = f"{pallets_formatted} паллет, РЕФ {temp_str}"

    # 3. Адрес погрузки
    if len(items) == 1:
        load_address_value = f"{items[0]['addr_ref']}, {items[0]['addr']}"
    else:
        addr_lines = []
        for i, it in enumerate(items):
            pal_s = format_pallets(it['pallets'])
            addr_lines.append(f"{i+1}. {it['addr_ref']}, {it['addr']} ({pal_s} пал.)")
        load_address_value = "; ".join(addr_lines)

    # 4. Даты погрузки и выгрузки
    dates = set()
    for it in items:
        if it.get("all_dates"):
            dates.update(it["all_dates"])
        elif it.get("date"):
            dates.add(it["date"])
    date_ship_str = ", ".join(sorted(dates)) if dates else datetime.date.today().strftime("%d.%m.%Y")

    load_datetime_value = f"{date_ship_str} к 9:00"
    unload_datetime_value = f"{date_ship_str}, по прибытию"

    def set_field(paragraph, label_bold: str, val_text: str, prefix_spaces: str = "", align=None):
        paragraph.text = ""
        if prefix_spaces:
            r_pre = paragraph.add_run(prefix_spaces)
            r_pre.bold = True
            r_pre.font.name = "Arial"
        r_lbl = paragraph.add_run(label_bold)
        r_lbl.bold = True
        r_lbl.font.name = "Arial"
        r_val = paragraph.add_run(val_text)
        r_val.bold = False
        r_val.font.name = "Arial"
        if align is not None:
            paragraph.alignment = align

    for p in doc.paragraphs:
        txt = p.text.strip()
        if "Дата:" in txt and any(y in txt for y in ["августа", "202", "2025", "2026"]):
            # Шапка Дата
            set_field(p, "Дата: ", today_date_str, prefix_spaces="                                                                                                                ")
        elif txt.startswith("Кому:"):
            # Шапка Кому
            set_field(p, "Кому: ", carrier, align=WD_ALIGN_PARAGRAPH.RIGHT)
        elif txt.startswith("Тип транспортного средства:"):
            set_field(p, "Тип транспортного средства: ", type_ts_value)
        elif txt.startswith("Адрес погрузки:"):
            set_field(p, "Адрес погрузки: ", load_address_value)
        elif txt.startswith("Дата и время погрузки:"):
            set_field(p, "Дата и время погрузки: ", load_datetime_value)
        elif txt.startswith("Дата и время выгрузки:"):
            set_field(p, "Дата и время выгрузки: ", unload_datetime_value)

    doc.save(output_docx_path)

def process_daily_orders(
    input_excel_path: str,
    carriers_db_path: str = None,
    template_docx_path: str = None,
    output_dir: str = None
) -> dict:
    """
    Основная точка входа для генерации заявок:
    - Считывает Excel
    - Сопоставляет перевозчиков
    - Суммирует одинаковые Адрес ссылка
    - Создает сводный файл Excel
    - Создает документы Word для каждого перевозчика
    - Создает общий ZIP-архив
    Возвращает отчет со списком файлов и статистикой.
    """
    if carriers_db_path is None:
        carriers_db_path = os.path.join(BASE_DIR, "База перевозчиков.xlsx")
    if template_docx_path is None:
        template_docx_path = os.path.join(BASE_DIR, "Шаблон.docx")
    if output_dir is None:
        output_dir = get_orders_dir()

    os.makedirs(output_dir, exist_ok=True)

    # 1. Загрузка базы
    driver_to_carrier = load_drivers_database(carriers_db_path)

    # 2. Чтение файла
    raw_items = parse_orders_excel(input_excel_path, driver_to_carrier)
    if not raw_items:
        raise ValueError("В загруженном файле нет данных для формирования заявок")

    # 3. Агрегация
    aggregated = aggregate_orders(raw_items)

    # Определяем суточную дату отгрузки для имен файлов (ДД.ММ)
    all_dates = set()
    for items in aggregated.values():
        for it in items:
            for d in it.get("all_dates", []):
                all_dates.add(d)

    date_tag = ""
    if all_dates:
        # Берем минимальную/основную дату
        primary_date = sorted(list(all_dates))[0]
        # первичная дата в формате 30.09.2026 -> 30.09
        parts = primary_date.split(".")
        if len(parts) >= 2:
            date_tag = f"{parts[0]}.{parts[1]}"
    if not date_tag:
        date_tag = datetime.date.today().strftime("%d.%m")

    # 4. Сохранение сводной таблицы Excel
    summary_excel_name = f"Сводный реестр заявок ({date_tag}).xlsx"
    summary_excel_path = os.path.join(output_dir, summary_excel_name)
    save_aggregated_excel(aggregated, summary_excel_path)

    # 5. Заполнение Word заявок
    generated_docs = []
    total_pallets_all = 0.0

    today_str = get_russian_date_str()

    for carrier, items in sorted(aggregated.items()):
        total_carrier_pallets = sum(it["pallets"] for it in items)
        total_pallets_all += total_carrier_pallets

        # Имя файла: Перевозчик ДД.ММ.docx (например: ИП Гончаров Иван Николаевич 30.09.docx)
        clean_carrier = sanitize_filename(carrier)
        docx_filename = f"{clean_carrier} {date_tag}.docx"
        docx_path = os.path.join(output_dir, docx_filename)

        fill_order_docx(
            template_path=template_docx_path,
            carrier=carrier,
            items=items,
            output_docx_path=docx_path,
            today_date_str=today_str
        )

        generated_docs.append({
            "carrier": carrier,
            "filename": docx_filename,
            "path": docx_path,
            "pallets": total_carrier_pallets,
            "pallets_str": format_pallets(total_carrier_pallets),
            "temp": items[0]["temp"] if items else "+15+25",
            "items_count": len(items),
            "date": items[0]["date"] if items else ""
        })

    # 6. Создание ZIP-архива со всеми сформированными файлами
    zip_filename = f"Заявки_{date_tag}.zip"
    zip_path = os.path.join(output_dir, zip_filename)
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.write(summary_excel_path, arcname=summary_excel_name)
        for doc_info in generated_docs:
            zf.write(doc_info["path"], arcname=doc_info["filename"])

    return {
        "status": "success",
        "output_dir": output_dir,
        "date_tag": date_tag,
        "total_carriers": len(generated_docs),
        "total_pallets": format_pallets(total_pallets_all),
        "summary_excel": summary_excel_name,
        "zip_file": zip_filename,
        "orders": generated_docs
    }

if __name__ == "__main__":
    test_in = os.path.join(BASE_DIR, "Книга4.xlsx")
    if os.path.exists(test_in):
        print(f"Тестовый запуск на {test_in}...")
        res = process_daily_orders(test_in)
        print(f"Готово! Сформировано заявок: {res['total_carriers']}")
        print(f"Всего паллет: {res['total_pallets']}")
        print(f"Папка сохранения: {res['output_dir']}")
        print(f"Сводка Excel: {res['summary_excel']}")
        print(f"Архив ZIP: {res['zip_file']}")
        for ord_info in res["orders"]:
            print(f"  • {ord_info['filename']} ({ord_info['pallets_str']} пал.)")
