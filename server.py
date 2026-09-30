#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
server.py — Локальный веб-сервис транспортного отдела (РЦ Черная Грязь).
Два основных раздела:
1. Графики отгрузки филиалов:
   - Формирование недельных графиков отгрузок из .xlsm
   - Назначение перевозчиков по актуальным тарифам и исключениям
   - Сохранение на корпоративном диске L: (Готовые недельные графики)
   - Персональная рассылка перевозчикам через Microsoft Outlook
2. Заявки перевозчикам:
   - Загрузка суточного реестра отгрузок (Книга4.xlsx)
   - Привязка перевозчиков по базе водителей (База перевозчиков.xlsx, Лист_1)
   - Суммирование паллет при одинаковом 'Адрес ссылка' у одного перевозчика
   - Генерация документов Word (Шаблон.docx) и сводного реестра Excel
   - Сохранение на корпоративном диске L: (Заявки)
"""

import os
import sys
import shutil
import tempfile
import webbrowser
import socket
import json
from datetime import datetime, timedelta
from typing import Optional, List
from collections import Counter, defaultdict

from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.responses import HTMLResponse, FileResponse, JSONResponse
from pydantic import BaseModel
import uvicorn

# Импортируем модули бизнес-логики
import generate_schedule as core
import generate_orders as orders_core

def get_local_ip():
    """Определяет локальный IP-адрес компьютера в корпоративной сети / Wi-Fi"""
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.settimeout(0.5)
        s.connect(("8.8.8.8", 80))
        ip = s.getsockname()[0]
        s.close()
        return ip
    except Exception:
        try:
            return socket.gethostbyname(socket.gethostname())
        except Exception:
            return "127.0.0.1"

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

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
    for c in desktop_candidates:
        if os.path.exists(c):
            folder = os.path.join(c, "Готовые недельные графики")
            os.makedirs(folder, exist_ok=True)
            return folder
            
    folder = os.path.join(user_home, "Desktop", "Готовые недельные графики")
    os.makedirs(folder, exist_ok=True)
    return folder

# Для совместимости со старыми вызовами
get_desktop_dir = get_schedules_dir

def get_orders_dir():
    """Возвращает путь к целевой папке 'Заявки'"""
    return orders_core.get_orders_dir()

def load_carriers_contacts():
    """Загружает справочник email-адресов перевозчиков из carriers_contacts.json"""
    contacts_file = os.path.join(BASE_DIR, "carriers_contacts.json")
    if os.path.exists(contacts_file):
        try:
            with open(contacts_file, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception as e:
            print(f"Ошибка чтения {contacts_file}: {e}")
    return {}

def get_carrier_email(carrier_name, fallback_email="n.rozhkov@puls.ru"):
    """Возвращает email перевозчика из carriers_contacts.json (или дефолтный)"""
    contacts = load_carriers_contacts()
    return contacts.get(carrier_name, fallback_email)

app = FastAPI(title="Транспортный отдел — РЦ Черная Грязь")

# Состояние графиков в памяти
CURRENT_STATE = {
    "all_trips": [],
    "week_range": "",
    "total_trips": 0,
    "carriers": [],
    "last_output_path": None,
    "last_output_filename": None
}

# Состояние заявок в памяти
CURRENT_ORDERS_STATE = {
    "output_dir": None,
    "summary_excel": None,
    "zip_file": None,
    "orders": [],
    "total_carriers": 0,
    "total_pallets": "0",
    "date_tag": ""
}

class SelectFileRequest(BaseModel):
    filename: str

class EmailRequest(BaseModel):
    mode: str # 'draft' или 'send'
    carrier: Optional[str] = None # None = всем перевозчикам, строка = конкретному
    email: Optional[str] = None

@app.get("/", response_class=HTMLResponse)
def index_page():
    html_content = """<!DOCTYPE html>
<html lang="ru">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Транспортный отдел — Графики и Заявки</title>
  <style>
    :root {
      --primary: #174ea6;
      --primary-hover: #123d82;
      --primary-light: #e8f0fe;
      --success: #15803d;
      --success-hover: #166534;
      --success-bg: #dcfce7;
      --warning: #b45309;
      --surface: #ffffff;
      --bg: #f8fafc;
      --border: #cbd5e1;
      --border-light: #e2e8f0;
      --text: #0f172a;
      --text-muted: #64748b;
      --radius: 12px;
      --shadow-sm: 0 1px 3px rgba(0,0,0,0.05);
      --shadow-md: 0 4px 20px -2px rgba(23, 78, 166, 0.08), 0 2px 8px -1px rgba(0,0,0,0.04);
      --shadow-lg: 0 10px 25px -3px rgba(23, 78, 166, 0.12), 0 4px 10px -2px rgba(0,0,0,0.05);
    }

    * { box-sizing: border-box; margin: 0; padding: 0; font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif; }
    body { background: linear-gradient(180deg, #edf2f7 0%, #f8fafc 100%); color: var(--text); min-height: 100vh; padding: 25px 20px; }
    .container { max-width: 960px; margin: 0 auto; }
    
    .header { background: var(--surface); padding: 20px 28px; border-radius: var(--radius); box-shadow: var(--shadow-sm); border-bottom: 3px solid var(--primary); margin-bottom: 18px; }
    .header-title h1 { font-size: 22px; font-weight: 700; color: var(--primary); display: flex; align-items: center; gap: 10px; }
    .header-title p { font-size: 13px; color: var(--text-muted); margin-top: 4px; }
    
    /* Табы навигации */
    .tabs-nav { display: flex; gap: 8px; background: var(--surface); padding: 6px; border-radius: var(--radius); border: 1px solid var(--border-light); box-shadow: var(--shadow-sm); margin-bottom: 22px; }
    .tab-btn { flex: 1; padding: 12px 18px; border: none; background: transparent; color: var(--text-muted); font-size: 14px; font-weight: 700; border-radius: 8px; cursor: pointer; transition: all 0.2s ease; display: flex; align-items: center; justify-content: center; gap: 8px; }
    .tab-btn:hover { background: #f1f5f9; color: var(--text); }
    .tab-btn.active { background: var(--primary); color: #fff; box-shadow: var(--shadow-sm); }
    
    .tab-content { display: none; }
    .tab-content.active { display: block; animation: fadeIn 0.2s ease-in-out; }
    @keyframes fadeIn { from { opacity: 0; transform: translateY(4px); } to { opacity: 1; transform: translateY(0); } }

    .card { background: var(--surface); border-radius: var(--radius); padding: 24px 28px; box-shadow: var(--shadow-sm); border: 1px solid var(--border-light); margin-bottom: 22px; transition: all 0.2s ease; }
    .card-title { font-size: 16px; font-weight: 600; margin-bottom: 14px; color: var(--text); display: flex; align-items: center; gap: 8px; }
    
    .dropzone { border: 2px dashed var(--border); border-radius: var(--radius); padding: 32px 20px; text-align: center; cursor: pointer; transition: all 0.2s; background: #fafafa; }
    .dropzone:hover, .dropzone.dragover { border-color: var(--primary); background: var(--primary-light); }
    .dropzone-icon { font-size: 38px; margin-bottom: 8px; }
    .dropzone-text { font-size: 15px; font-weight: 600; color: var(--text); }
    .dropzone-subtext { font-size: 12px; color: var(--text-muted); margin-top: 4px; }
    
    .path-badge { background: #f8fafc; border: 1px solid var(--border-light); border-radius: 8px; padding: 10px 14px; font-size: 12px; color: var(--text-muted); margin-top: 12px; display: flex; align-items: center; gap: 8px; word-break: break-all; }
    .path-badge.network { background: #f0fdf4; border-color: #bbf7d0; color: #166534; }
    .path-badge b { color: var(--text); }
    
    .btn { display: inline-flex; align-items: center; justify-content: center; gap: 8px; font-size: 13px; font-weight: 600; padding: 9px 18px; border-radius: 8px; border: none; cursor: pointer; transition: all 0.2s; text-decoration: none; }
    .btn-primary { background: var(--primary); color: #fff; }
    .btn-primary:hover { background: var(--primary-hover); box-shadow: var(--shadow-sm); }
    .btn-success { background: var(--success); color: #fff; }
    .btn-success:hover { background: var(--success-hover); }
    .btn-outline { background: transparent; border: 1px solid var(--border); color: var(--text); }
    .btn-outline:hover { background: #f1f5f9; border-color: #94a3b8; }
    .btn-sm { padding: 6px 12px; font-size: 12px; border-radius: 6px; }
    
    .actions-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 14px; margin-top: 14px; }
    
    .stats-row { display: flex; gap: 12px; margin: 14px 0; flex-wrap: wrap; }
    .stat-box { flex: 1; min-width: 140px; background: #f8fafc; border: 1px solid var(--border-light); padding: 12px 14px; border-radius: 8px; }
    .stat-label { font-size: 11px; text-transform: uppercase; color: var(--text-muted); font-weight: 600; }
    .stat-value { font-size: 18px; font-weight: 700; color: var(--primary); margin-top: 2px; }
    
    .carriers-list { display: flex; flex-wrap: wrap; gap: 8px; margin-top: 12px; }
    .carrier-tag { background: #f1f5f9; border: 1px solid var(--border-light); padding: 4px 10px; border-radius: 6px; font-size: 12px; display: inline-flex; align-items: center; gap: 6px; }
    .carrier-tag span { background: #e2e8f0; font-weight: 700; border-radius: 10px; padding: 1px 6px; font-size: 11px; }

    /* Таблицы */
    .table-responsive { overflow-x: auto; margin-top: 15px; border: 1px solid var(--border-light); border-radius: 8px; }
    .data-table { width: 100%; border-collapse: collapse; font-size: 13px; text-align: left; }
    .data-table th { background: #f8fafc; padding: 10px 12px; font-weight: 600; color: var(--text-muted); border-bottom: 1px solid var(--border-light); }
    .data-table td { padding: 10px 12px; border-bottom: 1px solid var(--border-light); vertical-align: middle; }
    .data-table tr:last-child td { border-bottom: none; }
    .data-table tr:hover td { background: #f8fafc; }

    /* Модальные окна */
    .modal-backdrop { display: none; position: fixed; top: 0; left: 0; width: 100%; height: 100%; background: rgba(15, 23, 42, 0.5); backdrop-filter: blur(4px); z-index: 999; align-items: center; justify-content: center; }
    .modal-backdrop.active { display: flex; }
    .modal { background: #fff; width: 95%; max-width: 540px; border-radius: var(--radius); box-shadow: var(--shadow-lg); overflow: hidden; animation: popIn 0.2s ease-out; }
    @keyframes popIn { from { transform: scale(0.95); opacity: 0; } to { transform: scale(1); opacity: 1; } }
    .modal-header { padding: 16px 22px; border-bottom: 1px solid var(--border-light); display: flex; justify-content: space-between; align-items: center; }
    .modal-header h3 { font-size: 16px; font-weight: 700; color: var(--text); }
    .modal-close { background: none; border: none; font-size: 20px; color: var(--text-muted); cursor: pointer; }
    .modal-body { padding: 18px 22px; }
    .modal-footer { padding: 14px 22px; background: #f8fafc; border-top: 1px solid var(--border-light); display: flex; justify-content: flex-end; gap: 10px; }

    .file-item { display: flex; align-items: center; gap: 12px; padding: 10px 12px; border-bottom: 1px solid var(--border-light); cursor: pointer; transition: background 0.15s; }
    .file-item:last-child { border-bottom: none; }
    .file-item:hover { background: #f8fafc; }
    .file-item.active { background: #eff6ff; }
    .file-item input[type="radio"] { accent-color: var(--primary); width: 18px; height: 18px; }

    .form-group { margin-bottom: 14px; }
    .form-label { display: block; font-size: 13px; font-weight: 600; margin-bottom: 6px; color: var(--text); }
    .form-control { width: 100%; padding: 8px 12px; border: 1px solid var(--border); border-radius: 6px; font-size: 13px; }
    .form-control:focus { outline: none; border-color: var(--primary); box-shadow: 0 0 0 3px var(--primary-light); }
    
    .choice-box { border: 1px solid var(--border); border-radius: 8px; padding: 12px; margin-bottom: 10px; cursor: pointer; transition: all 0.2s; display: flex; align-items: center; gap: 10px; }
    .choice-box:hover { border-color: var(--primary); background: #f8fafc; }
    .choice-box input[type="radio"] { margin-right: 5px; accent-color: var(--primary); width: 16px; height: 16px; }
    .choice-text b { display: block; font-size: 13px; color: var(--text); }
    .choice-text small { color: var(--text-muted); font-size: 12px; }
  </style>
</head>
<body>
  <div class="container">
    
    <div class="header">
      <div class="header-title">
        <h1>🚚 Транспортный отдел — РЦ Черная Грязь</h1>
        <p>Автоматизированная подготовка графиков филиалов, расчет тарифов, формирование транспортных заявок и рассылка через Outlook</p>
      </div>
    </div>

    <!-- Переключатель разделов -->
    <div class="tabs-nav">
      <button class="tab-btn active" id="tabBtnSchedules" onclick="switchTab('schedules')">
        🚚 Графики отгрузки филиалов
      </button>
      <button class="tab-btn" id="tabBtnOrders" onclick="switchTab('orders')">
        📝 Заявки перевозчикам
      </button>
    </div>

    <!-- ========================================== -->
    <!-- РАЗДЕЛ 1: ГРАФИКИ ОТГРУЗКИ ФИЛИАЛОВ        -->
    <!-- ========================================== -->
    <div id="tabContentSchedules" class="tab-content active">
      
      <!-- Шаг 1: Загрузка плана -->
      <div class="card">
        <div class="card-title">📁 1. Загрузите файл плана отгрузок (.xlsm)</div>
        <div class="dropzone" id="dropzone" onclick="document.getElementById('fileInput').click()">
          <div class="dropzone-icon">📥</div>
          <div class="dropzone-text" id="dropzoneText">Нажмите или перетащите сюда файл плана (.xlsm)</div>
          <div class="dropzone-subtext">Поддерживается базовый файл плана 'График_отгрузки_филиалов.xlsm'</div>
        </div>
        <input type="file" id="fileInput" accept=".xlsm,.xlsx" style="display:none" onchange="handleFileSelected(event)">
        
        <div class="path-badge" id="schedulesPathBadge">
          <span>📁 Папка сохранения:</span> <b id="schedulesPathText">L:\\ОТДЕЛЫ\\ТРАНСПОРТНЫЙ ОТДЕЛ\\Рожков\\РК недельные графики + заявки\\Готовые недельные графики</b>
        </div>

        <div style="margin-top: 15px; display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap; gap: 10px;">
          <button class="btn btn-outline" onclick="openFileSelectModal()" style="font-size: 13px; background: #fff;">
            📂 Выбрать готовый график из папки...
          </button>
          <button class="btn btn-primary" id="processBtn" style="display:none;" onclick="uploadAndProcess()">⚡ Обработать файл</button>
        </div>
      </div>

      <!-- Шаг 2: Результат обработки графика -->
      <div class="card" id="resultCard" style="display:none;">
        <div class="card-title">✅ 2. Готовый результат недели</div>
        
        <div class="stats-row">
          <div class="stat-box">
            <div class="stat-label">Период недели</div>
            <div class="stat-value" id="statWeek">-</div>
          </div>
          <div class="stat-box">
            <div class="stat-label">Всего рейсов</div>
            <div class="stat-value" id="statTotal">-</div>
          </div>
          <div class="stat-box">
            <div class="stat-label">Перевозчиков</div>
            <div class="stat-value" id="statCarriers">-</div>
          </div>
          <div class="stat-box">
            <div class="stat-label">Бюджет недели</div>
            <div class="stat-value" id="statCost" style="color: #15803d;">-</div>
          </div>
        </div>

        <div style="background: #f0fdf4; border: 1px solid #bbf7d0; border-radius: 8px; padding: 14px 18px; margin: 15px 0;">
          <div style="display: flex; align-items: center; justify-content: space-between; gap: 15px; flex-wrap: wrap;">
            <div style="flex: 1; min-width: 250px;">
              <div style="font-size: 14px; font-weight: 700; color: #166534; display: flex; align-items: center; gap: 8px;">
                📁 Активный файл недели
              </div>
              <div style="font-size: 12px; color: #1e293b; margin-top: 3px;" id="schedulesActiveFolderLabel">
                Папка: <b>L:\\ОТДЕЛЫ\\ТРАНСПОРТНЫЙ ОТДЕЛ\\Рожков\\РК недельные графики + заявки\\Готовые недельные графики</b>
              </div>
              <div style="font-size: 13px; font-weight: 600; color: var(--primary); margin-top: 4px;" id="currentFileLabel">
                Файл: -
              </div>
            </div>
            <div>
              <button class="btn btn-primary" onclick="openFileSelectModal()" style="gap: 6px;">
                🔄 Выбрать / Обновить из Excel
              </button>
            </div>
          </div>
        </div>
        <div id="reloadNotice" style="display:none; padding: 10px 14px; background: #e0f2fe; border: 1px solid #7dd3fc; border-radius: 8px; font-size: 13px; color: #0369a1; margin-bottom: 15px;"></div>

        <div style="font-size: 13px; font-weight: 600; margin-top: 18px; color: var(--text-muted);">
          Распределение рейсов по перевозчикам:
        </div>
        <div class="carriers-list" id="carriersContainer"></div>
      </div>

      <!-- Шаг 3: Рассылка в Outlook -->
      <div class="card" id="emailCard" style="display:none;">
        <div class="card-title">✉️ 3. Персональная отправка через Outlook</div>

        <div class="actions-grid">
          <button class="btn btn-primary" onclick="openSendModal(null)">
            ✉️ Отправить ВСЕМ перевозчикам
          </button>
          <button class="btn btn-outline" onclick="openSendModal('select')">
            👤 Выбрать конкретного перевозчика...
          </button>
        </div>
      </div>

    </div>

    <!-- ========================================== -->
    <!-- РАЗДЕЛ 2: ЗАЯВКИ ПЕРЕВОЗЧИКАМ              -->
    <!-- ========================================== -->
    <div id="tabContentOrders" class="tab-content">
      
      <div class="card">
        <div class="card-title">📝 Формирование транспортных заявок</div>

        <div class="dropzone" id="ordersDropzone" onclick="document.getElementById('ordersFileInput').click()">
          <div class="dropzone-icon">📋</div>
          <div class="dropzone-text" id="ordersDropzoneText">Нажмите или перетащите сюда суточный реестр отгрузок (.xlsx)</div>
          <div class="dropzone-subtext">Файл со столбцами: Адрес ссылка, Адрес, Количество паллет, Режим термоперевозки, Дата отгрузки, Перевозчик</div>
        </div>
        <input type="file" id="ordersFileInput" accept=".xlsx,.xlsm" style="display:none" onchange="handleOrdersFileSelected(event)">

        <div class="path-badge" id="ordersPathBadge">
          <span>📁 Папка сохранения заявок:</span> <b id="ordersPathText">L:\\ОТДЕЛЫ\\ТРАНСПОРТНЫЙ ОТДЕЛ\\Рожков\\РК недельные графики + заявки\\Заявки</b>
        </div>

        <div style="margin-top: 15px; display: flex; justify-content: flex-end;">
          <button class="btn btn-primary" id="processOrdersBtn" style="display:none;" onclick="uploadAndProcessOrders()">
            ⚡ Сформировать заявки
          </button>
        </div>
      </div>

      <!-- Результат генерации заявок -->
      <div class="card" id="ordersResultCard" style="display:none;">
        <div class="card-title">✅ Готовые транспортные заявки</div>

        <div class="stats-row">
          <div class="stat-box">
            <div class="stat-label">Перевозчиков</div>
            <div class="stat-value" id="statOrdersCarriers">-</div>
          </div>
          <div class="stat-box">
            <div class="stat-label">Всего паллет</div>
            <div class="stat-value" id="statOrdersPallets">-</div>
          </div>
          <div class="stat-box">
            <div class="stat-label">Дата отгрузки</div>
            <div class="stat-value" id="statOrdersDate">-</div>
          </div>
          <div class="stat-box">
            <div class="stat-label">Документов Word</div>
            <div class="stat-value" id="statOrdersDocsCount">-</div>
          </div>
        </div>

        <!-- Уведомление о сохранении в сетевую папку -->
        <div style="background: #f0fdf4; border: 1px solid #bbf7d0; border-radius: 8px; padding: 14px 18px; margin: 15px 0;">
          <div style="font-size: 14px; font-weight: 700; color: #166534; display: flex; align-items: center; gap: 8px;">
            ✅ Все файлы заявок успешно сформированы и сохранены
          </div>
          <div style="font-size: 12px; color: #1e293b; margin-top: 4px;" id="ordersSavedFolderLabel">
            Папка: <b>L:\\ОТДЕЛЫ\\ТРАНСПОРТНЫЙ ОТДЕЛ\\Рожков\\РК недельные графики + заявки\\Заявки</b>
          </div>
        </div>

        <!-- Таблица сформированных заявок -->
        <div class="table-responsive">
          <table class="data-table">
            <thead>
              <tr>
                <th style="width: 40px;">#</th>
                <th>Перевозчик</th>
                <th style="text-align: center; width: 75px;">Машин</th>
                <th style="text-align: right; width: 85px;">Паллеты</th>
                <th>Маршрут</th>
                <th style="text-align: center; width: 95px;">Режим</th>
                <th>Файл заявки (.docx)</th>
                <th style="text-align: center; width: 110px;">Статус</th>
              </tr>
            </thead>
            <tbody id="ordersTableBody">
            </tbody>
          </table>
        </div>

        <!-- Карточка отправки заявок через Outlook -->
        <div class="card" id="ordersEmailCard" style="display:none; margin-top: 20px;">
          <div class="card-title">✉️ Отправка заявок через Outlook</div>
          <p style="font-size: 13px; color: var(--text-muted); margin-bottom: 14px;">
            К каждому письму перевозчика будут автоматически прикреплены файлы сформированных для него транспортных заявок (.docx).
          </p>

          <div class="actions-grid">
            <button class="btn btn-primary" onclick="openOrdersSendModal(null)">
              ✉️ Отправить ВСЕМ перевозчикам
            </button>
            <button class="btn btn-outline" onclick="openOrdersSendModal('select')">
              👤 Выбрать конкретного перевозчика...
            </button>
          </div>
        </div>

      </div>

    </div>

  </div>

  <!-- Модальное окно подтверждения отправки Outlook -->
  <div class="modal-backdrop" id="sendModal">
    <div class="modal">
      <div class="modal-header">
        <h3 id="modalTitle">Отправка графика через Outlook</h3>
        <button class="modal-close" onclick="closeSendModal()">&times;</button>
      </div>
      <div class="modal-body">
        
        <div class="form-group" id="carrierSelectGroup" style="display:none;">
          <label class="form-label">Выберите перевозчика:</label>
          <select id="carrierSelect" class="form-control" onchange="updateModalTargetInfo()"></select>
        </div>

        <div style="background: #f1f5f9; padding: 12px; border-radius: 8px; margin-bottom: 16px; font-size: 13px;">
          <div id="modalTargetInfo"><b>Кому:</b> Все перевозчики</div>
        </div>

        <div class="form-group">
          <label class="form-label">Выберите режим отправки в Outlook:</label>
          
          <label class="choice-box">
            <input type="radio" name="sendMode" value="draft" checked>
            <div class="choice-text">
              <b>📝 Создать черновики (Рекомендуется)</b>
              <small>Письма сохранятся в папке «Черновики» Outlook. Вы сможете открыть и проверить каждое письмо перед отправкой.</small>
            </div>
          </label>

          <label class="choice-box">
            <input type="radio" name="sendMode" value="send">
            <div class="choice-text">
              <b>🚀 Отправить сразу</b>
              <small>Письма будут отправлены адресату в фоновом режиме через приложение Outlook.</small>
            </div>
          </label>
        </div>

        <div id="sendingStatus" style="display:none; text-align:center; padding: 10px; color: var(--primary); font-weight:600;">
          ⏳ Подключение к Outlook и формирование писем...
        </div>

      </div>
      <div class="modal-footer">
        <button class="btn btn-outline" onclick="closeSendModal()">Отмена</button>
        <button class="btn btn-primary" id="confirmSendBtn" onclick="executeEmailSend()">Подтвердить</button>
      </div>
    </div>
  </div>

  <!-- Модальное окно подтверждения отправки Заявок через Outlook -->
  <div class="modal-backdrop" id="ordersSendModal">
    <div class="modal">
      <div class="modal-header">
        <h3 id="ordersModalTitle">Отправка заявок через Outlook</h3>
        <button class="modal-close" onclick="closeOrdersSendModal()">&times;</button>
      </div>
      <div class="modal-body">
        
        <div class="form-group" id="ordersCarrierSelectGroup" style="display:none;">
          <label class="form-label">Выберите перевозчика:</label>
          <select id="ordersCarrierSelect" class="form-control" onchange="updateOrdersModalTargetInfo()"></select>
        </div>

        <div style="background: #f1f5f9; padding: 12px; border-radius: 8px; margin-bottom: 16px; font-size: 13px;">
          <div id="ordersModalTargetInfo"><b>Кому:</b> Все перевозчики</div>
          <div id="ordersModalFilesInfo" style="margin-top: 4px; color: var(--text-muted); font-size: 12px;"></div>
        </div>

        <div class="form-group">
          <label class="form-label">Выберите режим отправки в Outlook:</label>
          
          <label class="choice-box">
            <input type="radio" name="ordersSendMode" value="draft" checked>
            <div class="choice-text">
              <b>📝 Создать черновики (Рекомендуется)</b>
              <small>Письма с вложенными заявками (.docx) сохранятся в папке «Черновики» Outlook. Вы сможете проверить каждое письмо перед отправкой.</small>
            </div>
          </label>

          <label class="choice-box">
            <input type="radio" name="ordersSendMode" value="send">
            <div class="choice-text">
              <b>🚀 Отправить сразу</b>
              <small>Письма с файлами заявок будут сразу отправлены перевозчикам через приложение Outlook.</small>
            </div>
          </label>
        </div>

        <div id="ordersSendingStatus" style="display:none; text-align:center; padding: 10px; color: var(--primary); font-weight:600;">
          ⏳ Подключение к Outlook и прикрепление файлов заявок...
        </div>

      </div>
      <div class="modal-footer">
        <button class="btn btn-outline" onclick="closeOrdersSendModal()">Отмена</button>
        <button class="btn btn-primary" id="confirmOrdersSendBtn" onclick="executeOrdersEmailSend()">Подтвердить</button>
      </div>
    </div>
  </div>

  <!-- Модальное окно выбора файла из папки графика -->
  <div class="modal-backdrop" id="fileSelectModal">
    <div class="modal" style="max-width: 600px;">
      <div class="modal-header">
        <h3>🔄 Выбор файла графика</h3>
        <button class="modal-close" onclick="closeFileSelectModal()">&times;</button>
      </div>
      <div class="modal-body">
        <p style="font-size: 13px; color: var(--text-muted); margin-bottom: 12px;">
          Файлы отсортированы по дате изменения (новые — сверху). Выберите нужный файл для загрузки исправлений и отправки рассылки:
        </p>

        <div id="filesLoading" style="text-align: center; padding: 25px; color: var(--text-muted); font-size: 13px;">
          ⏳ Чтение файлов в папке...
        </div>

        <div id="filesContainer" style="display:none; max-height: 280px; overflow-y: auto; border: 1px solid var(--border-light); border-radius: 8px;">
        </div>

        <div id="noFilesMsg" style="display:none; text-align: center; padding: 20px; color: var(--text-muted); font-size: 13px;">
          В папке пока нет файлов Excel. Сначала загрузите и обработайте файл плана.
        </div>
      </div>
      <div class="modal-footer">
        <button class="btn btn-outline" onclick="closeFileSelectModal()">Отмена</button>
        <button class="btn btn-primary" id="confirmFileSelectBtn" onclick="confirmSelectedFile()" disabled>Выбрать и обновить данные</button>
      </div>
    </div>
  </div>

  <script>
    let selectedFile = null;
    let selectedOrdersFile = null;
    let availableCarriers = [];
    let availableOrdersCarriers = [];
    let currentOrdersSendTarget = null;

    // Переключение между вкладками
    function switchTab(tabName) {
      document.querySelectorAll('.tab-btn').forEach(btn => btn.classList.remove('active'));
      document.querySelectorAll('.tab-content').forEach(content => content.classList.remove('active'));
      
      if (tabName === 'schedules') {
        document.getElementById('tabBtnSchedules').classList.add('active');
        document.getElementById('tabContentSchedules').classList.add('active');
      } else {
        document.getElementById('tabBtnOrders').classList.add('active');
        document.getElementById('tabContentOrders').classList.add('active');
      }
    }

    // Загрузка путей при старте страницы
    async function loadStorageInfo() {
      try {
        const resp = await fetch('/api/storage-info');
        const data = await resp.json();
        if (data) {
          const schText = document.getElementById('schedulesPathText');
          if (schText) schText.innerText = data.schedules_dir;
          const schActive = document.getElementById('schedulesActiveFolderLabel');
          if (schActive) schActive.innerHTML = 'Папка: <b>' + data.schedules_dir + '</b>';
          const ordText = document.getElementById('ordersPathText');
          if (ordText) ordText.innerText = data.orders_dir;

          if (data.schedules_is_network) {
            document.getElementById('schedulesPathBadge').classList.add('network');
          }
          if (data.orders_is_network) {
            document.getElementById('ordersPathBadge').classList.add('network');
          }
        }
      } catch (e) {
        console.log('Storage info error:', e);
      }
    }
    loadStorageInfo();

    // ==========================================
    // ЛОГИКА ГРАФИКОВ (Раздел 1)
    // ==========================================
    function handleFileSelected(event) {
      if (event.target.files && event.target.files[0]) {
        selectedFile = event.target.files[0];
        document.getElementById('dropzoneText').innerText = 'Выбран: ' + selectedFile.name;
        document.getElementById('processBtn').style.display = 'inline-flex';
      }
    }

    const dropzone = document.getElementById('dropzone');
    dropzone.addEventListener('dragover', (e) => { e.preventDefault(); dropzone.classList.add('dragover'); });
    dropzone.addEventListener('dragleave', () => dropzone.classList.remove('dragover'));
    dropzone.addEventListener('drop', (e) => {
      e.preventDefault();
      dropzone.classList.remove('dragover');
      if (e.dataTransfer.files && e.dataTransfer.files[0]) {
        selectedFile = e.dataTransfer.files[0];
        document.getElementById('dropzoneText').innerText = 'Выбран: ' + selectedFile.name;
        document.getElementById('processBtn').style.display = 'inline-flex';
      }
    });

    async function uploadAndProcess() {
      if (!selectedFile) return;
      const processBtn = document.getElementById('processBtn');
      processBtn.disabled = true;
      processBtn.innerText = '⏳ Обработка...';

      const formData = new FormData();
      formData.append('file', selectedFile);

      try {
        const resp = await fetch('/api/upload', { method: 'POST', body: formData });
        const data = await resp.json();
        
        if (data.status === 'success') {
          document.getElementById('statWeek').innerText = data.week_range;
          document.getElementById('statTotal').innerText = data.total_trips + ' машин';
          document.getElementById('statCarriers').innerText = data.carriers.length;
          document.getElementById('statCost').innerText = data.total_cost || '-';
          document.getElementById('currentFileLabel').innerText = 'Файл: ' + data.filename;

          availableCarriers = data.carriers;
          const container = document.getElementById('carriersContainer');
          container.innerHTML = '';
          const select = document.getElementById('carrierSelect');
          select.innerHTML = '';

          data.carriers.forEach(c => {
            const costStr = c.cost_str ? ` • ${c.cost_str}` : '';
            container.innerHTML += `<div class="carrier-tag">${c.name} <span>${c.count} рейсов${costStr}</span></div>`;
            select.innerHTML += `<option value="${c.name}">${c.name} (${c.count} рейсов${costStr})</option>`;
          });

          document.getElementById('resultCard').style.display = 'block';
          document.getElementById('emailCard').style.display = 'block';
          processBtn.innerText = '✅ Обработано';
        } else {
          alert('Ошибка: ' + (data.message || 'Не удалось обработать'));
          processBtn.disabled = false;
          processBtn.innerText = '⚡ Обработать файл';
        }
      } catch (err) {
        alert('Ошибка связи с сервером: ' + err);
        processBtn.disabled = false;
        processBtn.innerText = '⚡ Обработать файл';
      }
    }

    async function openFileSelectModal() {
      const modal = document.getElementById('fileSelectModal');
      const container = document.getElementById('filesContainer');
      const loading = document.getElementById('filesLoading');
      const noFiles = document.getElementById('noFilesMsg');
      const confirmBtn = document.getElementById('confirmFileSelectBtn');

      modal.classList.add('active');
      container.style.display = 'none';
      noFiles.style.display = 'none';
      loading.style.display = 'block';
      confirmBtn.disabled = true;

      try {
        const resp = await fetch('/api/list-ready-files');
        const data = await resp.json();
        loading.style.display = 'none';

        if (data.status === 'success' && data.files && data.files.length > 0) {
          container.innerHTML = '';
          data.files.forEach((f, idx) => {
            const isChecked = f.is_current || (idx === 0);
            container.innerHTML += `
              <label class="file-item ${isChecked ? 'active' : ''}">
                <input type="radio" name="fileSelectRadio" value="${f.filename}" ${isChecked ? 'checked' : ''} onchange="onFileRadioChange(this)">
                <div style="flex:1;">
                  <div style="font-weight: 600; font-size: 13px; color: var(--text);">📊 ${f.filename}</div>
                  <div style="font-size: 11px; color: var(--text-muted); margin-top: 3px;">
                    Изменен: <b>${f.mtime_str}</b> &bull; Размер: ${f.size}
                  </div>
                </div>
              </label>
            `;
          });
          container.style.display = 'block';
          confirmBtn.disabled = false;
        } else {
          noFiles.style.display = 'block';
        }
      } catch (e) {
        loading.innerHTML = '❌ Ошибка чтения папки: ' + e;
      }
    }

    function onFileRadioChange(radio) {
      document.querySelectorAll('.file-item').forEach(el => el.classList.remove('active'));
      radio.closest('.file-item').classList.add('active');
      document.getElementById('confirmFileSelectBtn').disabled = false;
    }

    function closeFileSelectModal() {
      document.getElementById('fileSelectModal').classList.remove('active');
    }

    async function confirmSelectedFile() {
      const selected = document.querySelector('input[name="fileSelectRadio"]:checked');
      if (!selected) return;

      const filename = selected.value;
      const confirmBtn = document.getElementById('confirmFileSelectBtn');
      confirmBtn.disabled = true;
      confirmBtn.innerText = 'Загрузка...';

      try {
        const resp = await fetch('/api/select-file', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ filename })
        });
        const data = await resp.json();
        confirmBtn.disabled = false;
        confirmBtn.innerText = 'Выбрать и обновить данные';

        if (data.status === 'success') {
          document.getElementById('statWeek').innerText = data.week_range;
          document.getElementById('statTotal').innerText = data.total_trips + ' машин';
          document.getElementById('statCarriers').innerText = data.carriers.length;
          document.getElementById('statCost').innerText = data.total_cost || '-';
          document.getElementById('currentFileLabel').innerText = 'Файл: ' + data.filename;

          availableCarriers = data.carriers;
          const container = document.getElementById('carriersContainer');
          container.innerHTML = '';
          const select = document.getElementById('carrierSelect');
          select.innerHTML = '';

          data.carriers.forEach(c => {
            const costStr = c.cost_str ? ` • ${c.cost_str}` : '';
            container.innerHTML += `<div class="carrier-tag">${c.name} <span>${c.count} рейсов${costStr}</span></div>`;
            select.innerHTML += `<option value="${c.name}">${c.name} (${c.count} рейсов${costStr})</option>`;
          });

          document.getElementById('resultCard').style.display = 'block';
          document.getElementById('emailCard').style.display = 'block';

          const notice = document.getElementById('reloadNotice');
          notice.innerText = `✅ Успешно загружен актуальный файл '${filename}'. Все ручные правки применены!`;
          notice.style.display = 'block';

          closeFileSelectModal();
        } else {
          alert('Ошибка при выборе файла: ' + data.message);
        }
      } catch (e) {
        alert('Ошибка связи с сервером: ' + e);
        confirmBtn.disabled = false;
        confirmBtn.innerText = 'Выбрать и обновить данные';
      }
    }

    let currentSendTarget = null;
    function openSendModal(target) {
      currentSendTarget = target;
      const modal = document.getElementById('sendModal');
      const title = document.getElementById('modalTitle');
      const selectGroup = document.getElementById('carrierSelectGroup');
      const confirmBtn = document.getElementById('confirmSendBtn');
      const status = document.getElementById('sendingStatus');

      status.style.display = 'none';
      confirmBtn.disabled = false;

      if (target === 'select') {
        title.innerText = 'Отправка графика перевозчику';
        selectGroup.style.display = 'block';
      } else {
        title.innerText = 'Отправка графика ВСЕМ перевозчикам';
        selectGroup.style.display = 'none';
      }

      updateModalTargetInfo();
      modal.classList.add('active');
    }

    function closeSendModal() {
      document.getElementById('sendModal').classList.remove('active');
    }

    function updateModalTargetInfo() {
      const info = document.getElementById('modalTargetInfo');
      if (currentSendTarget === 'select') {
        const select = document.getElementById('carrierSelect');
        const carrier = select.value;
        info.innerHTML = `<b>Получатель:</b> ${carrier}`;
      } else {
        info.innerHTML = `<b>Получатели:</b> Все перевозчики (${availableCarriers.length} адресатов)`;
      }
    }

    async function executeEmailSend() {
      const mode = document.querySelector('input[name="sendMode"]:checked').value;
      const carrier = (currentSendTarget === 'select') ? document.getElementById('carrierSelect').value : null;

      const confirmBtn = document.getElementById('confirmSendBtn');
      confirmBtn.disabled = true;
      document.getElementById('sendingStatus').style.display = 'block';

      try {
        const resp = await fetch('/api/send-emails', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ mode, carrier })
        });
        const res = await resp.json();

        if (res.status === 'success') {
          alert('Успешно: ' + res.message);
          closeSendModal();
        } else {
          alert('Ошибка при отправке: ' + res.message);
          confirmBtn.disabled = false;
          document.getElementById('sendingStatus').style.display = 'none';
        }
      } catch (e) {
        alert('Ошибка связи с сервером: ' + e);
        confirmBtn.disabled = false;
        document.getElementById('sendingStatus').style.display = 'none';
      }
    }

    // ==========================================
    // ЛОГИКА ЗАЯВОК (Раздел 2)
    // ==========================================
    function handleOrdersFileSelected(event) {
      if (event.target.files && event.target.files[0]) {
        selectedOrdersFile = event.target.files[0];
        document.getElementById('ordersDropzoneText').innerText = 'Выбран: ' + selectedOrdersFile.name;
        document.getElementById('processOrdersBtn').style.display = 'inline-flex';
      }
    }

    const ordersDropzone = document.getElementById('ordersDropzone');
    ordersDropzone.addEventListener('dragover', (e) => { e.preventDefault(); ordersDropzone.classList.add('dragover'); });
    ordersDropzone.addEventListener('dragleave', () => ordersDropzone.classList.remove('dragover'));
    ordersDropzone.addEventListener('drop', (e) => {
      e.preventDefault();
      ordersDropzone.classList.remove('dragover');
      if (e.dataTransfer.files && e.dataTransfer.files[0]) {
        selectedOrdersFile = e.dataTransfer.files[0];
        document.getElementById('ordersDropzoneText').innerText = 'Выбран: ' + selectedOrdersFile.name;
        document.getElementById('processOrdersBtn').style.display = 'inline-flex';
      }
    });

    async function uploadAndProcessOrders() {
      if (!selectedOrdersFile) return;
      const btn = document.getElementById('processOrdersBtn');
      btn.disabled = true;
      btn.innerText = '⏳ Генерация заявок...';

      const formData = new FormData();
      formData.append('file', selectedOrdersFile);

      try {
        const resp = await fetch('/api/orders/generate', { method: 'POST', body: formData });
        const data = await resp.json();

        if (data.status === 'success') {
          document.getElementById('statOrdersCarriers').innerText = data.total_carriers;
          document.getElementById('statOrdersPallets').innerText = data.total_pallets;
          document.getElementById('statOrdersDate').innerText = data.date_tag;
          document.getElementById('statOrdersDocsCount').innerText = data.orders.length + ' Word + 1 Excel';

          const savedFolder = document.getElementById('ordersSavedFolderLabel');
          if (savedFolder && data.output_dir) {
            savedFolder.innerHTML = 'Папка: <b>' + data.output_dir + '</b>';
          }

          const tbody = document.getElementById('ordersTableBody');
          tbody.innerHTML = '';

          data.orders.forEach((ord, idx) => {
            tbody.innerHTML += `
              <tr>
                <td>${idx + 1}</td>
                <td>
                  <b>${ord.carrier}</b>
                  ${ord.addr_ref ? `<div style="font-size: 11px; color: #475569; margin-top: 2px;">Клиент: <b>${ord.addr_ref}</b></div>` : ''}
                </td>
                <td style="text-align: center; font-weight: 700; color: var(--primary);">${ord.trucks}</td>
                <td style="text-align: right; font-weight: 600;">${ord.pallets_str}</td>
                <td>${ord.route}</td>
                <td style="text-align: center;"><span style="background: #e0f2fe; color: #0369a1; padding: 2px 6px; border-radius: 4px; font-size: 11px;">${ord.temp}</span></td>
                <td>📄 ${ord.filename}</td>
                <td style="text-align: center;"><span style="color: #166534; font-weight: 600; font-size: 12px;">✅ Сохранен</span></td>
              </tr>
            `;
          });

          // Подготавливаем список перевозчиков для модального окна отправки писем
          availableOrdersCarriers = data.carriers || [];
          const selOrders = document.getElementById('ordersCarrierSelect');
          if (selOrders) {
            selOrders.innerHTML = '';
            availableOrdersCarriers.forEach(c => {
              const opt = document.createElement('option');
              opt.value = c.name;
              opt.textContent = `${c.name} (${c.orders_count} ${c.orders_count === 1 ? 'заявка' : 'заявок'}, ${c.pallets_str || c.total_pallets} пал.)`;
              selOrders.appendChild(opt);
            });
          }

          document.getElementById('ordersResultCard').style.display = 'block';
          document.getElementById('ordersEmailCard').style.display = 'block';
          btn.innerText = '✅ Заявки сформированы';
        } else {
          alert('Ошибка формирования заявок: ' + (data.message || 'Сбой'));
          btn.disabled = false;
          btn.innerText = '⚡ Сформировать заявки';
        }
      } catch (err) {
        alert('Ошибка связи с сервером: ' + err);
        btn.disabled = false;
        btn.innerText = '⚡ Сформировать заявки';
      }
    }

    // Обработчики модального окна отправки Заявок
    function openOrdersSendModal(target) {
      if (!availableOrdersCarriers || availableOrdersCarriers.length === 0) {
        alert('Сначала сформируйте заявки перевозчикам!');
        return;
      }
      currentOrdersSendTarget = target;
      const selectGroup = document.getElementById('ordersCarrierSelectGroup');
      const modalTitle = document.getElementById('ordersModalTitle');

      if (target === 'select') {
        modalTitle.innerText = 'Отправка заявки перевозчику через Outlook';
        selectGroup.style.display = 'block';
      } else {
        modalTitle.innerText = 'Отправка заявок ВСЕМ перевозчикам через Outlook';
        selectGroup.style.display = 'none';
      }

      updateOrdersModalTargetInfo();
      document.getElementById('confirmOrdersSendBtn').disabled = false;
      document.getElementById('ordersSendingStatus').style.display = 'none';
      document.getElementById('ordersSendModal').classList.add('active');
    }

    function closeOrdersSendModal() {
      document.getElementById('ordersSendModal').classList.remove('active');
    }

    function updateOrdersModalTargetInfo() {
      const info = document.getElementById('ordersModalTargetInfo');
      const filesInfo = document.getElementById('ordersModalFilesInfo');
      if (currentOrdersSendTarget === 'select') {
        const sel = document.getElementById('ordersCarrierSelect');
        const carrierName = sel ? sel.value : '';
        const carObj = availableOrdersCarriers.find(c => c.name === carrierName);
        const filesCount = carObj ? carObj.orders_count : 1;
        info.innerHTML = `<b>Получатель:</b> ${carrierName}`;
        filesInfo.innerHTML = `📎 Будет прикреплено файлов заявок: <b>${filesCount} шт.</b> (.docx)`;
      } else {
        const totalFiles = availableOrdersCarriers.reduce((acc, c) => acc + (c.orders_count || 1), 0);
        info.innerHTML = `<b>Получатели:</b> Все перевозчики (${availableOrdersCarriers.length} адресатов)`;
        filesInfo.innerHTML = `📎 Всего будет прикреплено файлов заявок: <b>${totalFiles} шт.</b> (.docx)`;
      }
    }

    async function executeOrdersEmailSend() {
      const mode = document.querySelector('input[name="ordersSendMode"]:checked').value;
      const carrier = (currentOrdersSendTarget === 'select') ? document.getElementById('ordersCarrierSelect').value : null;

      const confirmBtn = document.getElementById('confirmOrdersSendBtn');
      confirmBtn.disabled = true;
      document.getElementById('ordersSendingStatus').style.display = 'block';

      try {
        const resp = await fetch('/api/orders/send-emails', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ mode, carrier })
        });
        const res = await resp.json();

        if (res.status === 'success') {
          alert('Успешно: ' + res.message);
          closeOrdersSendModal();
        } else {
          alert('Ошибка при отправке: ' + res.message);
          confirmBtn.disabled = false;
          document.getElementById('ordersSendingStatus').style.display = 'none';
        }
      } catch (e) {
        alert('Ошибка связи с сервером: ' + e);
        confirmBtn.disabled = false;
        document.getElementById('ordersSendingStatus').style.display = 'none';
      }
    }
  </script>
</body>
</html>"""
    return HTMLResponse(content=html_content)

@app.get("/api/storage-info")
def get_storage_info():
    """Возвращает информацию о текущих целевых папках (сетевой диск / локальный)"""
    sch_dir = get_schedules_dir()
    ord_dir = get_orders_dir()
    return {
        "schedules_dir": sch_dir,
        "schedules_is_network": sch_dir.upper().startswith("L:"),
        "orders_dir": ord_dir,
        "orders_is_network": ord_dir.upper().startswith("L:")
    }

@app.post("/api/upload")
async def upload_file(file: UploadFile = File(...)):
    """Обработка плана отгрузок филиалов (Раздел 1)"""
    try:
        temp_dir = tempfile.gettempdir()
        temp_path = os.path.join(temp_dir, file.filename)
        with open(temp_path, "wb") as f:
            shutil.copyfileobj(file.file, f)
            
        base_name = os.path.splitext(file.filename)[0]
        schedules_dir = get_schedules_dir()
        
        # Проверяем: не является ли файл уже готовым скорректированным графиком (.xlsx с листами дней)?
        is_edited_schedule = False
        all_trips = []
        if file.filename.endswith(('.xlsx', '.xlsm')):
            try:
                test_trips = core.parse_final_schedule_file(temp_path)
                if len(test_trips) > 0:
                    is_edited_schedule = True
                    all_trips = test_trips
            except Exception:
                is_edited_schedule = False

        if is_edited_schedule and all_trips:
            # Загружен файл, который уже правили руками
            out_filename = file.filename
            target_path = os.path.join(schedules_dir, out_filename)
            shutil.copy2(temp_path, target_path)
        else:
            # Запускаем генерацию из базового плана
            plan_trips, start_date = core.parse_plan_file(temp_path)
            
            # Точные календарные границы недели
            end_date = start_date + timedelta(days=6)
            date_range_str = f"{start_date.strftime('%d.%m')} - {end_date.strftime('%d.%m.%Y')}"
            
            clean_name = base_name.replace(" (готовый)", "")
            out_filename = f"{clean_name} ({date_range_str}).xlsx"
            target_path = os.path.join(schedules_dir, out_filename)
            
            # Защита от блокировки в Excel: если файл уже открыт кем-то на компьютере
            try:
                all_trips = core.build_schedule_from_plan(plan_trips, start_date, target_path)
            except PermissionError:
                import time
                out_filename = f"{clean_name} ({date_range_str})_{int(time.time())}.xlsx"
                target_path = os.path.join(schedules_dir, out_filename)
                all_trips = core.build_schedule_from_plan(plan_trips, start_date, target_path)
        
        # Считаем статистику
        carrier_counts = Counter([t['carrier'] for t in all_trips if t.get('carrier')])
        dates = [t['date'] for t in all_trips]
        week_range = f"{dates[0]} - {dates[-1]}" if dates else ""
        
        total_cost = sum(t.get('tariff', 0) for t in all_trips)
        total_cost_str = f"{total_cost:,}".replace(",", " ") + " ₽"
        
        carrier_costs = defaultdict(int)
        for t in all_trips:
            car = t.get('carrier')
            if car:
                carrier_costs[car] += t.get('tariff', 0)
                
        carriers_list = [
            {
                "name": k,
                "count": v,
                "cost": carrier_costs[k],
                "cost_str": f"{carrier_costs[k]:,}".replace(",", " ") + " ₽"
            }
            for k, v in carrier_counts.most_common()
        ]
        
        CURRENT_STATE["all_trips"] = all_trips
        CURRENT_STATE["week_range"] = week_range
        CURRENT_STATE["total_trips"] = len(all_trips)
        CURRENT_STATE["total_cost"] = total_cost
        CURRENT_STATE["total_cost_str"] = total_cost_str
        CURRENT_STATE["carriers"] = carriers_list
        CURRENT_STATE["last_output_path"] = target_path
        CURRENT_STATE["last_output_filename"] = out_filename
        
        return {
            "status": "success",
            "week_range": week_range,
            "total_trips": len(all_trips),
            "total_cost": total_cost_str,
            "carriers": carriers_list,
            "filename": out_filename,
            "is_edited": is_edited_schedule
        }
    except Exception as e:
        return {"status": "error", "message": str(e)}

@app.get("/api/list-ready-files")
def list_ready_files():
    """Возвращает список готовых графиков Excel из папки Готовые недельные графики"""
    try:
        schedules_dir = get_schedules_dir()
        if not os.path.exists(schedules_dir):
            return {"status": "success", "files": []}
        
        files_info = []
        current_path = CURRENT_STATE.get("last_output_path")
        
        for fname in os.listdir(schedules_dir):
            if fname.startswith("~$") or not fname.endswith((".xlsx", ".xlsm")):
                continue
            full_path = os.path.join(schedules_dir, fname)
            if not os.path.isfile(full_path):
                continue
            
            stat = os.stat(full_path)
            mtime = stat.st_mtime
            mtime_str = datetime.fromtimestamp(mtime).strftime("%d.%m.%Y %H:%M:%S")
            
            size_bytes = stat.st_size
            if size_bytes < 1024 * 1024:
                size_str = f"{size_bytes / 1024:.1f} КБ"
            else:
                size_str = f"{size_bytes / (1024 * 1024):.2f} МБ"
                
            is_current = bool(current_path and os.path.abspath(full_path) == os.path.abspath(current_path))
            
            files_info.append({
                "filename": fname,
                "mtime": mtime,
                "mtime_str": mtime_str,
                "size": size_str,
                "is_current": is_current
            })
            
        # Сортировка: самые свежие сверху
        files_info.sort(key=lambda x: x["mtime"], reverse=True)
        return {"status": "success", "files": files_info}
    except Exception as e:
        return {"status": "error", "message": str(e), "files": []}

@app.post("/api/select-file")
def select_file(req: SelectFileRequest):
    """Выбор существующего файла графика из папки"""
    try:
        schedules_dir = get_schedules_dir()
        target_path = os.path.join(schedules_dir, req.filename)
        if not os.path.exists(target_path):
            return {"status": "error", "message": f"Файл не найден: {req.filename}"}
            
        trips = core.parse_final_schedule_file(target_path)
        if not trips:
            return {"status": "error", "message": "В выбранном файле не найдено строк с рейсами"}
            
        carrier_counts = Counter([t['carrier'] for t in trips if t.get('carrier')])
        dates = [t['date'] for t in trips]
        week_range = f"{dates[0]} - {dates[-1]}" if dates else ""
        
        total_cost = sum(t.get('tariff', 0) for t in trips)
        total_cost_str = f"{total_cost:,}".replace(",", " ") + " ₽"
        
        carrier_costs = defaultdict(int)
        for t in trips:
            car = t.get('carrier')
            if car:
                carrier_costs[car] += t.get('tariff', 0)
                
        carriers_list = [
            {
                "name": k,
                "count": v,
                "cost": carrier_costs[k],
                "cost_str": f"{carrier_costs[k]:,}".replace(",", " ") + " ₽"
            }
            for k, v in carrier_counts.most_common()
        ]
        
        CURRENT_STATE["all_trips"] = trips
        CURRENT_STATE["week_range"] = week_range
        CURRENT_STATE["total_trips"] = len(trips)
        CURRENT_STATE["total_cost"] = total_cost
        CURRENT_STATE["total_cost_str"] = total_cost_str
        CURRENT_STATE["carriers"] = carriers_list
        CURRENT_STATE["last_output_path"] = target_path
        CURRENT_STATE["last_output_filename"] = req.filename
        
        return {
            "status": "success",
            "filename": req.filename,
            "week_range": week_range,
            "total_trips": len(trips),
            "total_cost": total_cost_str,
            "carriers": carriers_list
        }
    except Exception as e:
        return {"status": "error", "message": f"Ошибка чтения файла: {str(e)}"}

@app.post("/api/send-emails")
def send_emails(req: EmailRequest):
    """Отправка индивидуальных писем через Microsoft Outlook"""
    target_file = CURRENT_STATE.get("last_output_path")
    if target_file and os.path.exists(target_file):
        try:
            disk_trips = core.parse_final_schedule_file(target_file)
            if disk_trips:
                CURRENT_STATE["all_trips"] = disk_trips
        except Exception as e:
            print(f"Используем данные из памяти: {e}")

    trips_to_send = CURRENT_STATE.get("all_trips", [])
    if not trips_to_send:
        return {"status": "error", "message": "Нет данных для отправки. Сначала загрузите и выберите файл плана!"}

    if req.carrier:
        target_trips = [t for t in trips_to_send if t['carrier'] == req.carrier]
        target_carriers = [req.carrier]
    else:
        target_trips = trips_to_send
        target_carriers = sorted(list(set(t['carrier'] for t in trips_to_send if t.get('carrier'))))

    if not target_trips:
        return {"status": "error", "message": "Рейсы для указанного перевозчика не найдены."}

    draft_mode = (req.mode == "draft")
    
    try:
        import win32com.client as win32
        outlook = win32.Dispatch('outlook.application')
        
        carrier_groups = defaultdict(list)
        for tr in target_trips:
            carrier_groups[tr['carrier']].append(tr)
            
        dates = [tr['date'] for tr in trips_to_send]
        start_d = dates[0] if dates else ''
        end_d = dates[-1] if dates else ''
        
        sent_count = 0
        for car, c_trips in carrier_groups.items():
            carrier_email = (req.email if req.email else None) or get_carrier_email(car)
            sorted_trips = sorted(c_trips, key=lambda x: (x.get('dt') or datetime.strptime(x['date'], '%d.%m.%Y'), x['time']))
            mail = outlook.CreateItem(0)
            mail.To = carrier_email
            mail.Subject = f"[{car}] График погрузки РЦ Черная Грязь ({start_d} - {end_d})"
            mail.HTMLBody = core.generate_carrier_html(car, sorted_trips, start_d, end_d)
            
            if draft_mode:
                mail.Save()
            else:
                mail.Send()
            sent_count += 1
            
        action_text = "сохранено в Черновиках" if draft_mode else "отправлено"
        msg = f"Успешно {action_text} писем перевозчикам: {sent_count} шт."
        return {"status": "success", "message": msg, "count": sent_count}
        
    except ImportError:
        action_text = "черновиков" if draft_mode else "отправки"
        msg = f"Библиотека pywin32 не установлена (требуется Windows + Outlook). Готово к работе {len(target_carriers)} писем {action_text}."
        return {"status": "success", "message": msg, "count": len(target_carriers)}
    except Exception as e:
        return {"status": "error", "message": f"Ошибка Outlook: {str(e)}"}

# ==========================================
# API ЭНДПОИНТЫ ДЛЯ ЗАЯВОК (Раздел 2)
# ==========================================
@app.post("/api/orders/generate")
async def generate_orders_api(file: UploadFile = File(...)):
    """Прием файла реестра (Книга4.xlsx) и генерация транспортных заявок Word и сводного Excel"""
    try:
        temp_dir = tempfile.gettempdir()
        temp_path = os.path.join(temp_dir, file.filename)
        with open(temp_path, "wb") as f:
            shutil.copyfileobj(file.file, f)

        res = orders_core.process_daily_orders(
            input_excel_path=temp_path,
            template_docx_path=os.path.join(BASE_DIR, "Шаблон.docx"),
            output_dir=get_orders_dir()
        )

        CURRENT_ORDERS_STATE.update(res)
        return res
    except Exception as e:
        return {"status": "error", "message": str(e)}

@app.get("/api/orders/download-file")
def download_order_file(filename: str):
    """Скачивание файла заявки (.docx), сводного реестра (.xlsx) или архива (.zip)"""
    orders_dir = get_orders_dir()
    clean_name = os.path.basename(filename)
    target_path = os.path.join(orders_dir, clean_name)
    if not os.path.exists(target_path):
        raise HTTPException(status_code=404, detail="Файл не найден")
    return FileResponse(target_path, filename=clean_name, media_type="application/octet-stream")

@app.get("/api/orders/list-files")
def list_orders_files():
    """Возвращает список файлов из папки Заявки"""
    try:
        orders_dir = get_orders_dir()
        if not os.path.exists(orders_dir):
            return {"status": "success", "files": []}
        files = []
        for f in os.listdir(orders_dir):
            if f.startswith("~$") or not f.endswith((".docx", ".xlsx", ".zip")):
                continue
            fp = os.path.join(orders_dir, f)
            if os.path.isfile(fp):
                stat = os.stat(fp)
                files.append({
                    "filename": f,
                    "mtime": stat.st_mtime,
                    "mtime_str": datetime.fromtimestamp(stat.st_mtime).strftime("%d.%m.%Y %H:%M"),
                    "size": f"{stat.st_size/1024:.1f} КБ" if stat.st_size < 1024*1024 else f"{stat.st_size/(1024*1024):.2f} МБ"
                })
        files.sort(key=lambda x: x["mtime"], reverse=True)
        return {"status": "success", "files": files, "output_dir": orders_dir}
    except Exception as e:
        return {"status": "error", "message": str(e), "files": []}

@app.post("/api/orders/send-emails")
def send_orders_emails(req: EmailRequest):
    """Отправка индивидуальных заявок перевозчикам через Microsoft Outlook с вложением файлов .docx"""
    orders_to_send = CURRENT_ORDERS_STATE.get("orders", [])
    if not orders_to_send:
        return {"status": "error", "message": "Нет сформированных заявок для отправки. Сначала сформируйте заявки из файла реестра!"}

    if req.carrier:
        target_orders = [o for o in orders_to_send if o.get('carrier') == req.carrier]
        target_carriers = [req.carrier]
    else:
        target_orders = orders_to_send
        target_carriers = sorted(list(set(o['carrier'] for o in orders_to_send if o.get('carrier'))))

    if not target_orders:
        return {"status": "error", "message": "Заявки для указанного перевозчика не найдены."}

    draft_mode = (req.mode == "draft")

    try:
        import win32com.client as win32
        outlook = win32.Dispatch('outlook.application')

        carrier_groups = defaultdict(list)
        for ord_info in target_orders:
            carrier_groups[ord_info['carrier']].append(ord_info)

        sent_count = 0
        orders_dir = get_orders_dir()

        for car, c_orders in carrier_groups.items():
            carrier_email = (req.email if req.email else None) or get_carrier_email(car)
            date_str = c_orders[0].get('date') or CURRENT_ORDERS_STATE.get("date_tag") or datetime.now().strftime("%d.%m.%Y")

            mail = outlook.CreateItem(0)
            mail.To = carrier_email
            mail.Subject = f"[{car}] Транспортная заявка на перевозку ({date_str})"
            mail.HTMLBody = orders_core.generate_order_email_html(car, c_orders, date_str)

            # Прикрепляем персональные файлы Word (.docx) для данного перевозчика
            for o in c_orders:
                fpath = o.get("path")
                if not fpath or not os.path.exists(fpath):
                    fpath = os.path.join(orders_dir, o.get("filename", ""))
                if fpath and os.path.exists(fpath):
                    mail.Attachments.Add(os.path.abspath(fpath))

            if draft_mode:
                mail.Save()
            else:
                mail.Send()
            sent_count += 1

        action_text = "сохранено в Черновиках" if draft_mode else "отправлено"
        msg = f"Успешно {action_text} писем с заявками перевозчикам: {sent_count} шт."
        return {"status": "success", "message": msg, "count": sent_count}

    except ImportError:
        action_text = "черновиков" if draft_mode else "отправки"
        msg = f"Библиотека pywin32 не установлена (требуется Windows + Outlook). Готово к работе {len(target_carriers)} писем {action_text} с прикрепленными заявками."
        return {"status": "success", "message": msg, "count": len(target_carriers)}
    except Exception as e:
        return {"status": "error", "message": f"Ошибка Outlook: {str(e)}"}

def find_available_port(start_port=8000, max_attempts=50):
    """Находит свободный сетевой порт, если стандартный (8000) уже занят другим приложением"""
    for p in range(start_port, start_port + max_attempts):
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            if s.connect_ex(('127.0.0.1', p)) != 0:
                return p
    return start_port

if __name__ == "__main__":
    local_ip = get_local_ip()
    port = find_available_port(8000)
    
    print("=" * 65)
    print("  ТРАНСПОРТНЫЙ ОТДЕЛ — ВЕБ-СЕРВИС ЗАПУЩЕН")
    print(f"  • Локальный адрес:   http://localhost:{port}")
    if local_ip != "127.0.0.1":
        print(f"  • Адрес в сети:      http://{local_ip}:{port}")
    print(f"  • Графики отгрузок:  {get_schedules_dir()}")
    print(f"  • Заявки:            {get_orders_dir()}")
    print("=" * 65)
    
    try:
        webbrowser.open(f"http://localhost:{port}")
    except Exception:
        pass

    uvicorn.run(app, host="0.0.0.0", port=port, log_level="warning")
