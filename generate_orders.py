#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
generate_orders.py — Модуль формирования транспортных заявок для перевозчиков.

Логика:
1. Загрузка исходного реестра со столбцами:
   Адрес ссылка, Адрес, Количество паллет, Режим термоперевозки, Дата отгрузки, Перевозчик.
2. Формирование раздельных заявок для каждого клиента (Адрес ссылка) у каждого перевозчика.
   Если у одного перевозчика несколько строк с одинаковым 'Адрес ссылка', количество паллет суммируется.
3. Расчет количества требуемых транспортных средств:
   - до 33 паллет = 1 машина;
   - от 34 до 66 = 2 машины;
   - от 67 до 99 = 3 машины и т.д. (каждые 33 паллета сверху добавляют машину).
4. Формирование строки Маршрут: <город/деревня погрузки> – д. Черная Грязь.
5. Заполнение Word-шаблона 'Шаблон.docx' для каждого заказа:
   - Кому: Перевозчик
   - Дата: текущая дата (например, 30 сентября 2026)
   - Маршрут: <город/деревня> – д. Черная Грязь
   - Тип транспортного средства: <N> х <кол-во> паллет, РЕФ <режим>
   - Адрес погрузки: Адрес ссылка, Адрес
   - Дата и время погрузки: Дата отгрузки к 9:00
   - Дата и время выгрузки: Дата отгрузки, по прибытию
6. Сохранение каждого файла в папку Заявки:
   Основной путь: L:\\ОТДЕЛЫ\\ТРАНСПОРТНЫЙ ОТДЕЛ\\Рожков\\РК недельные графики + заявки\\Заявки
   Формат имени: <Перевозчик> (<Адрес ссылка>) <ДД.ММ>.docx
7. Подготовка и отправка писем с вложенными заявками через Outlook.
"""

import os
import sys
import re
import math
import datetime
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
    Приоритет: сетевой/облачный диск L:
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

def extract_settlement(addr: str) -> str:
    """
    Извлекает название города / деревни / села / пгт из полного адреса для строки 'Маршрут:'.
    Примеры:
    - 'Московская обл, Истринский р-н, Лешково с, д. стр.244...' -> 'Лешково'
    - '143500, Московская область, ..., деревня Давыдовское, ...' -> 'д. Давыдовское'
    - '141400 Московская область, г. Химки, ...' -> 'г. Химки'
    - '143080, Московская область, ..., пгт. Лесной Городок, ...' -> 'пгт. Лесной Городок'
    - '142153, Московская область, ..., д Новоселки, ...' -> 'д. Новоселки'
    """
    if not addr:
        return ""

    # 1. Поиск деревни: 'деревня Название' или 'д. Название' или 'д Название'
    m = re.search(r'\bдеревня\s+([А-Яа-яЁё\-]+)', addr, re.IGNORECASE)
    if m:
        return f"д. {m.group(1).strip()}"

    m = re.search(r'(?:^|[\s,])д\.?\s+([А-Яа-яЁё\-]+)(?!\s*\d)', addr, re.IGNORECASE)
    if m:
        val = m.group(1).strip()
        if val.lower() not in ('стр', 'корп', 'влд', 'к'):
            return f"д. {val}"

    # 2. Поиск пгт / поселка
    m = re.search(r'\bпгт\.?\s+([А-Яа-яЁё\-]+(?:\s+[А-Яа-яЁё\-]+)?)', addr, re.IGNORECASE)
    if m:
        return f"пгт. {m.group(1).strip()}"
    m = re.search(r'\bпос(?:елок|\.)?\s+([А-Яа-яЁё\-]+(?:\s+[А-Яа-яЁё\-]+)?)', addr, re.IGNORECASE)
    if m:
        return f"пос. {m.group(1).strip()}"

    # 3. Поиск села: 'Лешково с' или 'с. Лешково' или 'село Лешково'
    m = re.search(r'\bсело\s+([А-Яа-яЁё\-]+)', addr, re.IGNORECASE)
    if m:
        return f"{m.group(1).strip()}"
    m = re.search(r'\b([А-Яа-яЁё\-]+)\s+с\b', addr, re.IGNORECASE)
    if m:
        return f"{m.group(1).strip()}"
    m = re.search(r'(?:^|[\s,])с\.?\s+([А-Яа-яЁё\-]+)', addr, re.IGNORECASE)
    if m:
        return f"{m.group(1).strip()}"

    # 4. Поиск города: 'г. Химки', 'г Пушкино', 'г. Домодедово', 'Санкт-Петербург г', 'Домодедово г'
    m = re.search(r'(?:^|[\s,])г\.(?!о\b)\s*([А-Яа-яЁё\-]+(?:\s+[А-Яа-яЁё\-]+)?)', addr, re.IGNORECASE)
    if m:
        val = m.group(1).strip()
        if not val.lower().startswith('о '):
            return f"г. {val}"

    m = re.search(r'(?:^|[\s,])г\s+([А-Яа-яЁё\-]+(?:\s+[А-Яа-яЁё\-]+)?)', addr, re.IGNORECASE)
    if m:
        val = m.group(1).strip()
        if val.lower() not in ('о', 'о.'):
            return f"г. {val}"

    m = re.search(r'\b([А-Яа-яЁё\-]+)\s+г\b', addr, re.IGNORECASE)
    if m:
        return f"г. {m.group(1).strip()}"

    # 5. Поиск городского округа: 'г.о. Подольск' -> 'г. Подольск'
    m = re.search(r'\bг\.?о\.?\s+([А-Яа-яЁё\-]+)', addr, re.IGNORECASE)
    if m:
        return f"г. {m.group(1).strip()}"

    parts = [p.strip() for p in addr.split(',') if p.strip()]
    for p in parts:
        if not re.match(r'^\d+$', p) and 'область' not in p.lower() and 'обл' not in p.lower():
            return p
    return addr

def parse_orders_excel(input_path: str) -> list:
    """
    Считывает строки из файла отгрузок.
    Колонка 'Перевозчик' берется напрямую из файла.
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
    idx_carrier = find_idx(["перевозчик", "подрядчик", "компания"])

    if idx_carrier == -1 and len(header) >= 6:
        idx_carrier = 5

    raw_items = []
    for r in rows[1:]:
        if not any(r):
            continue

        carrier_val = str(r[idx_carrier]).strip() if idx_carrier >= 0 and len(r) > idx_carrier and r[idx_carrier] else ""
        if not carrier_val:
            carrier_val = "Неизвестный перевозчик"

        pal_val = 0.0
        if idx_pal >= 0 and len(r) > idx_pal and r[idx_pal] is not None:
            try:
                pal_val = float(r[idx_pal])
            except (ValueError, TypeError):
                pal_val = 0.0

        date_val = r[idx_date] if idx_date >= 0 and len(r) > idx_date else None
        if isinstance(date_val, (datetime.date, datetime.datetime)):
            date_str = date_val.strftime("%d.%m.%Y")
        elif date_val:
            date_str = str(date_val).strip()
        else:
            date_str = datetime.date.today().strftime("%d.%m.%Y")

        addr_str = str(r[idx_addr]).strip() if idx_addr >= 0 and len(r) > idx_addr and r[idx_addr] else ""

        item = {
            "addr_ref": str(r[idx_ref]).strip() if idx_ref >= 0 and len(r) > idx_ref and r[idx_ref] else "",
            "addr": addr_str,
            "pallets": pal_val,
            "temp": str(r[idx_temp]).strip() if idx_temp >= 0 and len(r) > idx_temp and r[idx_temp] else "+15+25",
            "date": date_str,
            "carrier": carrier_val
        }
        raw_items.append(item)

    return raw_items

def aggregate_orders(raw_items: list) -> list:
    """
    Группирует данные по паре (Перевозчик, Адрес ссылка).
    Для каждой пары схлопывает количество паллет.
    Возвращает список агрегированных заказов:
    [
        {
            'carrier': ...,
            'addr_ref': ...,
            'addr': ...,
            'pallets': sum_pallets,
            'temp': ...,
            'date': ...,
            'trucks': ...
        },
        ...
    ]
    """
    groups = defaultdict(lambda: {
        "carrier": "",
        "addr_ref": "",
        "addr": "",
        "pallets": 0.0,
        "temps": set(),
        "dates": set()
    })

    for item in raw_items:
        key = (item["carrier"], item["addr_ref"])
        g = groups[key]
        g["carrier"] = item["carrier"]
        g["addr_ref"] = item["addr_ref"]
        if item["addr"] and not g["addr"]:
            g["addr"] = item["addr"]
        g["pallets"] += item["pallets"]
        if item["temp"]:
            g["temps"].add(item["temp"])
        if item["date"]:
            g["dates"].add(item["date"])

    aggregated = []
    for (carrier, ref), g in groups.items():
        pallets = g["pallets"]
        trucks = max(1, math.ceil(pallets / 33.0)) if pallets > 0 else 1
        temp_val = ", ".join(sorted(g["temps"])) if g["temps"] else "+15+25"
        date_val = sorted(list(g["dates"]))[0] if g["dates"] else datetime.date.today().strftime("%d.%m.%Y")
        
        aggregated.append({
            "carrier": carrier,
            "addr_ref": ref,
            "addr": g["addr"],
            "pallets": pallets,
            "pallets_str": format_pallets(pallets),
            "temp": temp_val,
            "date": date_val,
            "all_dates": sorted(list(g["dates"])),
            "trucks": trucks
        })

    # Сортировка по перевозчику, затем по клиенту
    aggregated.sort(key=lambda x: (x["carrier"], x["addr_ref"]))
    return aggregated

def save_aggregated_excel(aggregated_orders: list, output_path: str):
    """
    Формирует и сохраняет сводный Excel-файл со всеми заказами.
    Колонки: Адрес ссылка, Адрес, Количество паллет, Режим термоперевозки, Дата отгрузки, Перевозчик
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
    for it in aggregated_orders:
        ws.append([
            it["addr_ref"],
            it["addr"],
            it["pallets"],
            it["temp"],
            it["date"],
            it["carrier"]
        ])
        
        row_cells = ws[current_row]
        for idx, c in enumerate(row_cells):
            c.border = thin_border
            c.font = Font(name="Arial", size=10)
            if idx in (0, 1, 5):
                c.alignment = align_left
            elif idx in (2,):
                c.alignment = align_right
                c.number_format = '#,##0.#'
            else:
                c.alignment = align_center
        ws.row_dimensions[current_row].height = 22
        current_row += 1

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
    addr_ref: str,
    addr: str,
    pallets: float,
    trucks: int,
    temp: str,
    date_ship: str,
    output_docx_path: str,
    today_date_str: str = None
):
    """
    Заполняет Шаблон.docx для конкретного заказа:
    - Кому: Перевозчик
    - Дата: сегодняшняя дата
    - Маршрут: <город/деревня> – д. Черная Грязь
    - Тип транспортного средства: <N> х <кол-во> паллет, РЕФ <режим>
    - Адрес погрузки: Адрес ссылка, Адрес
    - Дата и время погрузки: Дата отгрузки к 9:00
    - Дата и время выгрузки: Дата отгрузки, по прибытию
    """
    if today_date_str is None:
        today_date_str = get_russian_date_str()

    doc = docx.Document(template_path)

    pallets_formatted = format_pallets(pallets)
    type_ts_value = f"{trucks} х {pallets_formatted} паллет, РЕФ {temp}"

    settlement = extract_settlement(addr)
    route_from = settlement if settlement else "г. Москва"
    route_value = f"{route_from} – д. Черная Грязь"

    load_address_value = f"{addr_ref}, {addr}"
    load_datetime_value = f"{date_ship} к 9:00"
    unload_datetime_value = f"{date_ship}, по прибытию"

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
        elif txt.startswith("Маршрут:"):
            # Маршрут
            set_field(p, "Маршрут: ", route_value)
        elif txt.startswith("Тип транспортного средства:"):
            # Тип транспортного средства
            set_field(p, "Тип транспортного средства: ", type_ts_value)
        elif txt.startswith("Адрес погрузки:"):
            # Адрес погрузки
            set_field(p, "Адрес погрузки: ", load_address_value)
        elif txt.startswith("Дата и время погрузки:"):
            # Дата и время погрузки
            set_field(p, "Дата и время погрузки: ", load_datetime_value)
        elif txt.startswith("Дата и время выгрузки:"):
            # Дата и время выгрузки
            set_field(p, "Дата и время выгрузки: ", unload_datetime_value)

    doc.save(output_docx_path)

def process_daily_orders(
    input_excel_path: str,
    template_docx_path: str = None,
    output_dir: str = None
) -> dict:
    """
    Основная точка входа для генерации заявок:
    - Считывает Excel
    - Группирует по (Перевозчик, Адрес ссылка)
    - Создает отдельный документ Word для каждого клиента каждого перевозчика
    - Создает сводный Excel-файл
    - Сохраняет все файлы в output_dir
    """
    if template_docx_path is None:
        template_docx_path = os.path.join(BASE_DIR, "Шаблон.docx")
    if output_dir is None:
        output_dir = get_orders_dir()

    os.makedirs(output_dir, exist_ok=True)

    # 1. Чтение файла
    raw_items = parse_orders_excel(input_excel_path)
    if not raw_items:
        raise ValueError("В загруженном файле нет данных для формирования заявок")

    # 2. Агрегация по (Перевозчик, Адрес ссылка)
    aggregated_orders = aggregate_orders(raw_items)

    # Определение даты отгрузки (ДД.ММ)
    all_dates = set()
    for it in aggregated_orders:
        for d in it.get("all_dates", []):
            all_dates.add(d)

    date_tag = ""
    primary_date = ""
    if all_dates:
        primary_date = sorted(list(all_dates))[0]
        parts = primary_date.split(".")
        if len(parts) >= 2:
            date_tag = f"{parts[0]}.{parts[1]}"
    if not date_tag:
        date_tag = datetime.date.today().strftime("%d.%m")
    if not primary_date:
        primary_date = datetime.date.today().strftime("%d.%m.%Y")

    # 3. Сохранение сводной таблицы Excel
    summary_excel_name = f"Сводный реестр заявок ({date_tag}).xlsx"
    summary_excel_path = os.path.join(output_dir, summary_excel_name)
    save_aggregated_excel(aggregated_orders, summary_excel_path)

    # 4. Заполнение Word заявок
    generated_docs = []
    total_pallets_all = 0.0
    today_str = get_russian_date_str()

    carrier_orders_map = defaultdict(list)

    for it in aggregated_orders:
        carrier = it["carrier"]
        addr_ref = it["addr_ref"]
        addr = it["addr"]
        pallets = it["pallets"]
        trucks = it["trucks"]
        temp = it["temp"]
        date_ship = it["date"]

        total_pallets_all += pallets

        settlement = extract_settlement(addr)
        route_str = f"{settlement} – д. Черная Грязь" if settlement else "д. Черная Грязь"

        # Имя файла: Перевозчик (Адрес ссылка) ДД.ММ.docx
        clean_carrier = sanitize_filename(carrier)
        clean_ref = sanitize_filename(addr_ref)
        if clean_ref:
            docx_filename = f"{clean_carrier} ({clean_ref}) {date_tag}.docx"
        else:
            docx_filename = f"{clean_carrier} {date_tag}.docx"
            
        docx_path = os.path.join(output_dir, docx_filename)

        fill_order_docx(
            template_path=template_docx_path,
            carrier=carrier,
            addr_ref=addr_ref,
            addr=addr,
            pallets=pallets,
            trucks=trucks,
            temp=temp,
            date_ship=date_ship,
            output_docx_path=docx_path,
            today_date_str=today_str
        )

        doc_info = {
            "carrier": carrier,
            "addr_ref": addr_ref,
            "addr": addr,
            "filename": docx_filename,
            "path": docx_path,
            "trucks": trucks,
            "route": route_str,
            "pallets": pallets,
            "pallets_str": format_pallets(pallets),
            "temp": temp,
            "date": date_ship
        }
        generated_docs.append(doc_info)
        carrier_orders_map[carrier].append(doc_info)

    # Список уникальных перевозчиков со статистикой для интерфейса
    carriers_summary = []
    for c_name, c_orders in sorted(carrier_orders_map.items()):
        c_pallets = sum(o["pallets"] for o in c_orders)
        carriers_summary.append({
            "name": c_name,
            "orders_count": len(c_orders),
            "total_pallets": c_pallets,
            "pallets_str": format_pallets(c_pallets),
            "files": [o["filename"] for o in c_orders]
        })

    return {
        "status": "success",
        "output_dir": output_dir,
        "date_tag": date_tag,
        "primary_date": primary_date,
        "total_carriers": len(carriers_summary),
        "total_orders": len(generated_docs),
        "total_pallets": format_pallets(total_pallets_all),
        "summary_excel": summary_excel_name,
        "carriers": carriers_summary,
        "orders": generated_docs
    }

def generate_order_email_html(carrier: str, orders: list, date_str: str) -> str:
    """
    Генерирует HTML-письмо для отправки заявок конкретному перевозчику.
    """
    rows_html = ""
    for o in orders:
        rows_html += f"""
        <tr>
            <td style="border: 1px solid #cbd5e1; padding: 9px 12px; font-weight: 700; color: #1e3a8a;">{o.get('addr_ref', '')}</td>
            <td style="border: 1px solid #cbd5e1; padding: 9px 12px; text-align: center; font-weight: 700; color: #174ea6;">{o.get('trucks', 1)}</td>
            <td style="border: 1px solid #cbd5e1; padding: 9px 12px; text-align: right; font-weight: 700;">{o.get('pallets_str', '')}</td>
            <td style="border: 1px solid #cbd5e1; padding: 9px 12px; text-align: center;">{o.get('temp', '+15+25')}</td>
            <td style="border: 1px solid #cbd5e1; padding: 9px 12px;">{o.get('route', '')}</td>
            <td style="border: 1px solid #cbd5e1; padding: 9px 12px; font-size: 11px; color: #475569;">{o.get('addr', '')}</td>
        </tr>
        """

    html = f"""
    <div style="font-family: Arial, sans-serif; font-size: 14px; color: #1e293b; line-height: 1.5;">
        <p>Здравствуйте!</p>
        <p>Во вложении направляем транспортную заявку на организацию автоперевозки со склада РЦ Черная Грязь (дата отгрузки: <b>{date_str}</b>).</p>
        
        <table style="border-collapse: collapse; width: 100%; margin: 16px 0; font-size: 13px;">
            <tr style="background: #f1f5f9; color: #334155;">
                <th style="border: 1px solid #cbd5e1; padding: 9px 12px; text-align: left;">Клиент / Точка</th>
                <th style="border: 1px solid #cbd5e1; padding: 9px 12px; text-align: center;">Машин</th>
                <th style="border: 1px solid #cbd5e1; padding: 9px 12px; text-align: right;">Паллеты</th>
                <th style="border: 1px solid #cbd5e1; padding: 9px 12px; text-align: center;">Режим</th>
                <th style="border: 1px solid #cbd5e1; padding: 9px 12px; text-align: left;">Маршрут</th>
                <th style="border: 1px solid #cbd5e1; padding: 9px 12px; text-align: left;">Адрес погрузки</th>
            </tr>
            {rows_html}
        </table>
        
        <p><b>Дата и время погрузки:</b> {date_str} к 9:00<br>
        <b>Дата и время выгрузки:</b> {date_str}, по прибытию<br>
        <b>Адрес выгрузки:</b> ООО «ФК ПУЛЬС», Московская обл, Солнечногорский р-н, Черная Грязь д, Сходненская ул, д. стр.1, «НЛП»</p>
        
        <p style="margin-top: 15px;">Официальный бланк заявки прикреплен к письму в формате Word (.docx).<br>
        Просьба подтвердить получение и заблаговременно предоставить данные на водителя и транспортное средство.</p>
        
        <br>
        <p style="color: #64748b; font-size: 12px; border-top: 1px solid #e2e8f0; padding-top: 12px;">
            С уважением,<br>
            <b>Транспортный отдел ООО «ФК ПУЛЬС»</b><br>
            Тел. 8 (495) 665-76-20 (доб. 1231)<br>
            E-mail: transport_opt@puls.ru
        </p>
    </div>
    """
    return html

if __name__ == "__main__":
    test_in = os.path.join(BASE_DIR, "Тест_Заявки_Скриншот.xlsx")
    if not os.path.exists(test_in):
        test_in = os.path.join(BASE_DIR, "Книга4.xlsx")
    if os.path.exists(test_in):
        print(f"Тестовый запуск на {test_in}...")
        res = process_daily_orders(test_in)
        print(f"Готово! Перевозчиков: {res['total_carriers']}, Всего заявок: {res['total_orders']}")
        print(f"Всего паллет: {res['total_pallets']}")
        print(f"Папка сохранения: {res['output_dir']}")
        print(f"Сводка Excel: {res['summary_excel']}")
        for ord_info in res["orders"]:
            print(f"  • {ord_info['filename']} | Машин: {ord_info['trucks']} | {ord_info['pallets_str']} пал. | Маршрут: {ord_info['route']}")
