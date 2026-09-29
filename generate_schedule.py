#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Автоматизированный генератор графика отгрузок (РЦ Черная Грязь).
Строго соблюдает условие: МАКСИМУМ 2 МАШИНЫ В ЧАС.
"""

import sys
import os
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

START_HOURS = {
    'Пн': 9,
    'Вт': 9,
    'Ср': 9,
    'Чт': 9,
    'Пт': 9,
    'Сб': 5, # Суббота: ранние окна
    'Вс': 9
}

def parse_oper_file(oper_path):
    wb = openpyxl.load_workbook(oper_path, data_only=True)
    ws = wb['Лист1']
    
    city_map = {
        7: 'Ярославль', 9: 'Казань', 10: 'Казань', 13: 'Воронеж', 14: 'Воронеж',
        16: 'Краснодар', 17: 'Краснодар', 18: 'Краснодар', 20: 'Волгоград',
        22: 'Самара', 24: 'Екатеринбург', 25: 'Екатеринбург', 26: 'Екатеринбург',
        28: 'Новосибирск', 29: 'Новосибирск', 30: 'Красноярск', 32: 'Иркутск',
        33: 'Хабаровск', 34: 'Санкт-Петербург', 35: 'Санкт-Петербург', 36: 'Брянск'
    }
    
    days_cols = [
        ('Пн', 2, [3, 4]),
        ('Вт', 5, [6, 7]),
        ('Ср', 8, [9, 10]),
        ('Чт', 11, [12, 13]),
        ('Пт', 14, [15, 16]),
        ('Сб', 17, [18]),
        ('Вс', 19, [20])
    ]
    
    trips_by_day = defaultdict(list)
    for d_name, time_col, carrier_cols in days_cols:
        start_h = START_HOURS.get(d_name, 9)
        day_raw = []
        for r in sorted(city_map.keys()):
            city = city_map[r]
            for c_col in carrier_cols:
                c = ws.cell(r, c_col).value
                if c and str(c).strip():
                    pal = 40 if city in ['Хабаровск', 'Ярославль'] else 33
                    day_raw.append({
                        'city': city,
                        'carrier': str(c).strip(),
                        'pallets': pal,
                        'row_idx': r
                    })
        
        # СТРОГОЕ СОБЛЮДЕНИЕ ПРАВИЛА: РОВНО 2 МАШИНЫ В ЧАС!
        # Каждые 2 машины получают свой час: start_h + (i // 2)
        for i, tr in enumerate(day_raw):
            h = start_h + (i // 2)
            tr['time'] = f"{h:02d}:00:00"
            trips_by_day[d_name].append(tr)
            
    wb.close()
    return trips_by_day

def build_excel_schedule(trips_by_day, output_path, start_date_str='2026-09-28'):
    start_date = datetime.strptime(start_date_str, '%Y-%m-%d')
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

        day_trips = trips_by_day.get(d_name, [])

        city_truck_count = Counter()
        row_idx = 2
        for trip in day_trips:
            city = trip['city']
            carrier = trip['carrier']
            pallets = trip['pallets']
            time_val = trip['time']
            
            city_truck_count[city] += 1
            truck_num = city_truck_count[city]
            
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
            c_f.value = truck_num
            c_f.font = font_bold_center
            c_f.alignment = Alignment(horizontal='center', vertical='center')
            c_f.border = border_thin
            
            c_g = ws[f'G{row_idx}']
            c_g.value = city
            c_g.font = font_city
            c_g.alignment = Alignment(horizontal='left', vertical='center')
            c_g.border = border_thin
            
            c_h = ws[f'H{row_idx}']
            c_h.value = pallets
            c_h.font = font_regular
            c_h.alignment = Alignment(horizontal='center', vertical='center')
            c_h.border = border_thin
            
            c_i = ws[f'I{row_idx}']
            c_i.value = carrier
            c_i.font = font_carrier
            c_i.alignment = Alignment(horizontal='center', vertical='center')
            c_i.border = border_thin
            
            c_j = ws[f'J{row_idx}']
            c_j.value = None
            c_j.border = border_thin
            
            c_k = ws[f'K{row_idx}']
            c_k.value = time_val
            c_k.font = font_regular
            c_k.alignment = Alignment(horizontal='center', vertical='center')
            c_k.border = border_thin
            
            if pallets == 40:
                for col_l in ['C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K']:
                    ws[f'{col_l}{row_idx}'].fill = fill_40p
            
            ws.row_dimensions[row_idx].height = 18.0
            row_idx += 1

        last_data_row = row_idx - 1
        if last_data_row >= 2:
            ws.auto_filter.ref = f'C1:K{last_data_row}'

        # Сводная таблица распределения долей ТК справа (M:O)
        carrier_counts = Counter([t['carrier'] for t in day_trips if t.get('carrier')])
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
    print(f"Таблица успешно сохранена: {output_path}")

if __name__ == '__main__':
    base_dir = os.path.dirname(os.path.abspath(__file__))
    def_oper = os.path.join(base_dir, 'Филиалы 2026_09_28_04.xlsx')
    def_out = os.path.join(base_dir, 'Недельные графики (готовый).xlsx')
    
    parser = argparse.ArgumentParser(description='Генератор графика отгрузок филиалов')
    parser.add_argument('--oper', default=def_oper, help='Путь к файлу Филиалы...xlsx')
    parser.add_argument('--out', default=def_out, help='Путь к результирующему файлу')
    parser.add_argument('--date', default='2026-09-28', help='Дата понедельника недели YYYY-MM-DD')
    
    args = parser.parse_args()
    
    if os.path.exists(args.oper):
        trips_by_day = parse_oper_file(args.oper)
        build_excel_schedule(trips_by_day, args.out, start_date_str=args.date)
    else:
        print(f"Файл {args.oper} не найден!")
