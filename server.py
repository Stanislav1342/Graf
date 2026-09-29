#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
server.py — Локальный веб-сервис для обработки графиков отгрузок (РЦ Черная Грязь)
и персональной рассылки перевозчикам через Microsoft Outlook.
"""

import os
import sys
import shutil
import tempfile
import webbrowser
from datetime import datetime
from typing import Optional, List
from collections import Counter, defaultdict

from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.responses import HTMLResponse, FileResponse, JSONResponse
from pydantic import BaseModel
import uvicorn

# Импортируем бизнес-логику из generate_schedule.py
import generate_schedule as core

app = FastAPI(title="Графики отгрузки филиалов — ФК ПУЛЬС")

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_FILE = os.path.join(BASE_DIR, "Недельные графики (готовый).xlsx")

# Состояние последней обработки в памяти
CURRENT_STATE = {
    "all_trips": [],
    "week_range": "",
    "total_trips": 0,
    "carriers": []
}

class EmailRequest(BaseModel):
    mode: str # 'draft' или 'send'
    carrier: Optional[str] = None # None = всем перевозчикам, строка = конкретному
    email: str = "n.rozhkov@puls.ru"

@app.get("/", response_class=HTMLResponse)
def index_page():
    html_content = """<!DOCTYPE html>
<html lang="ru">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Графики отгрузки филиалов — ФК ПУЛЬС</title>
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
    body { background: linear-gradient(180deg, #edf2f7 0%, #f8fafc 100%); color: var(--text); min-height: 100vh; padding: 30px 20px; }
    .container { max-width: 900px; margin: 0 auto; }
    
    .header { background: var(--surface); padding: 24px 30px; border-radius: var(--radius); box-shadow: var(--shadow-sm); border-bottom: 3px solid var(--primary); margin-bottom: 24px; display: flex; align-items: center; justify-content: space-between; }
    .header-title h1 { font-size: 22px; font-weight: 700; color: var(--primary); display: flex; align-items: center; gap: 10px; }
    .header-title p { font-size: 13px; color: var(--text-muted); margin-top: 4px; }
    .badge { background: var(--primary-light); color: var(--primary); padding: 5px 12px; border-radius: 20px; font-size: 12px; font-weight: 600; }
    
    .card { background: var(--surface); border-radius: var(--radius); padding: 24px 30px; box-shadow: var(--shadow-sm); border: 1px solid var(--border-light); margin-bottom: 24px; transition: all 0.2s ease; }
    .card-title { font-size: 16px; font-weight: 600; margin-bottom: 16px; color: var(--text); display: flex; align-items: center; gap: 8px; }
    
    .dropzone { border: 2px dashed var(--border); border-radius: var(--radius); padding: 36px 20px; text-align: center; cursor: pointer; transition: all 0.2s; background: #fafafa; }
    .dropzone:hover, .dropzone.dragover { border-color: var(--primary); background: var(--primary-light); }
    .dropzone-icon { font-size: 40px; margin-bottom: 10px; }
    .dropzone-text { font-size: 15px; font-weight: 600; color: var(--text); }
    .dropzone-subtext { font-size: 12px; color: var(--text-muted); margin-top: 5px; }
    
    .btn { display: inline-flex; align-items: center; justify-content: center; gap: 8px; font-size: 14px; font-weight: 600; padding: 10px 20px; border-radius: 8px; border: none; cursor: pointer; transition: all 0.2s; text-decoration: none; }
    .btn-primary { background: var(--primary); color: #fff; }
    .btn-primary:hover { background: var(--primary-hover); box-shadow: var(--shadow-sm); }
    .btn-success { background: var(--success); color: #fff; }
    .btn-success:hover { background: var(--success-hover); }
    .btn-outline { background: transparent; border: 1px solid var(--border); color: var(--text); }
    .btn-outline:hover { background: #f1f5f9; border-color: #94a3b8; }
    .btn-block { width: 100%; }
    
    .actions-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 15px; margin-top: 15px; }
    
    .stats-row { display: flex; gap: 15px; margin: 15px 0; }
    .stat-box { flex: 1; background: #f8fafc; border: 1px solid var(--border-light); padding: 12px 16px; border-radius: 8px; }
    .stat-label { font-size: 11px; text-transform: uppercase; color: var(--text-muted); font-weight: 600; }
    .stat-value { font-size: 18px; font-weight: 700; color: var(--primary); margin-top: 2px; }
    
    .carriers-list { display: flex; flex-wrap: wrap; gap: 8px; margin-top: 12px; }
    .carrier-tag { background: #f1f5f9; border: 1px solid var(--border-light); padding: 4px 10px; border-radius: 6px; font-size: 12px; display: inline-flex; align-items: center; gap: 6px; }
    .carrier-tag span { background: #e2e8f0; font-weight: 700; border-radius: 10px; padding: 1px 6px; font-size: 11px; }

    /* Modal styles */
    .modal-backdrop { display: none; position: fixed; top: 0; left: 0; width: 100%; height: 100%; background: rgba(15, 23, 42, 0.5); backdrop-filter: blur(4px); z-index: 999; align-items: center; justify-content: center; }
    .modal-backdrop.active { display: flex; }
    .modal { background: #fff; width: 95%; max-width: 520px; border-radius: var(--radius); box-shadow: var(--shadow-lg); overflow: hidden; animation: popIn 0.2s ease-out; }
    @keyframes popIn { from { transform: scale(0.95); opacity: 0; } to { transform: scale(1); opacity: 1; } }
    .modal-header { padding: 18px 24px; border-bottom: 1px solid var(--border-light); display: flex; justify-content: space-between; align-items: center; }
    .modal-header h3 { font-size: 17px; font-weight: 700; color: var(--text); }
    .modal-close { background: none; border: none; font-size: 20px; color: var(--text-muted); cursor: pointer; }
    .modal-body { padding: 20px 24px; }
    .modal-footer { padding: 16px 24px; background: #f8fafc; border-top: 1px solid var(--border-light); display: flex; justify-content: flex-end; gap: 10px; }

    .form-group { margin-bottom: 16px; }
    .form-label { display: block; font-size: 13px; font-weight: 600; margin-bottom: 6px; color: var(--text); }
    .form-control { width: 100%; padding: 9px 12px; border: 1px solid var(--border); border-radius: 6px; font-size: 14px; }
    .form-control:focus { outline: none; border-color: var(--primary); box-shadow: 0 0 0 3px var(--primary-light); }
    
    .alert { padding: 12px 16px; border-radius: 8px; font-size: 13px; margin-bottom: 15px; display: none; }
    .alert-success { background: var(--success-bg); color: var(--success); border: 1px solid #86efac; display: flex; align-items: center; gap: 8px; }
    .alert-info { background: var(--primary-light); color: var(--primary); border: 1px solid #bfdbfe; display: flex; align-items: center; gap: 8px; }
    
    .choice-box { border: 1px solid var(--border); border-radius: 8px; padding: 12px; margin-bottom: 10px; cursor: pointer; transition: all 0.2s; display: flex; align-items: center; gap: 10px; }
    .choice-box:hover { border-color: var(--primary); background: #f8fafc; }
    .choice-box input[type="radio"] { margin-right: 5px; accent-color: var(--primary); width: 16px; height: 16px; }
    .choice-text b { display: block; font-size: 14px; color: var(--text); }
    .choice-text small { color: var(--text-muted); font-size: 12px; }
  </style>
</head>
<body>
  <div class="container">
    
    <div class="header">
      <div class="header-title">
        <h1>🚚 Графики отгрузки филиалов</h1>
        <p>РЦ Черная Грязь • ФК «ПУЛЬС»</p>
      </div>
      <div class="badge">Регламент: 2 машины/час</div>
    </div>

    <!-- Шаг 1: Загрузка файла -->
    <div class="card">
      <div class="card-title">📁 1. Загрузите файл базового плана (.xlsm)</div>
      <div class="dropzone" id="dropzone" onclick="document.getElementById('fileInput').click()">
        <div class="dropzone-icon">📥</div>
        <div class="dropzone-text" id="dropzoneText">Нажмите или перетащите сюда файл плана</div>
        <div class="dropzone-subtext">Поддерживается 'График_отгрузки_филиалов_неделя_2.xlsm' (или любой .xlsm)</div>
      </div>
      <input type="file" id="fileInput" accept=".xlsm,.xlsx" style="display:none" onchange="handleFileSelected(event)">
      
      <div style="margin-top: 15px; display: flex; justify-content: flex-end;">
        <button class="btn btn-primary" id="processBtn" style="display:none;" onclick="uploadAndProcess()">⚡ Обработать файл</button>
      </div>
    </div>

    <!-- Шаг 2: Результат обработки -->
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
      </div>

      <div style="margin: 15px 0;">
        <a href="/api/download" class="btn btn-success btn-block" style="padding: 12px; font-size: 15px;">
          📥 Скачать Недельные графики (готовый).xlsx
        </a>
      </div>

      <div style="font-size: 13px; font-weight: 600; margin-top: 20px; color: var(--text-muted);">
        Распределение рейсов по перевозчикам:
      </div>
      <div class="carriers-list" id="carriersContainer"></div>
    </div>

    <!-- Шаг 3: Рассылка в Outlook -->
    <div class="card" id="emailCard" style="display:none;">
      <div class="card-title">✉️ 3. Персональная отправка через Outlook</div>
      <div class="alert alert-info" style="display:flex;">
        🔒 <b>Конфиденциальность:</b> Каждый перевозчик получает строго таблицу ТОЛЬКО со своими рейсами и тайм-слотами.
      </div>

      <div class="form-group" style="margin-top: 15px;">
        <label class="form-label">Тестовый email-получатель:</label>
        <input type="email" id="targetEmailInput" class="form-control" value="n.rozhkov@puls.ru">
      </div>

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

  <!-- Модальное окно подтверждения отправки -->
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
          <div style="margin-top: 4px;"><b>Куда (тест):</b> <span id="modalTargetEmail">n.rozhkov@puls.ru</span></div>
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

  <script>
    let selectedFile = null;
    let availableCarriers = [];

    function handleFileSelected(event) {
      if (event.target.files && event.target.files[0]) {
        selectedFile = event.target.files[0];
        document.getElementById('dropzoneText').innerText = 'Выбран: ' + selectedFile.name;
        document.getElementById('processBtn').style.display = 'inline-flex';
      }
    }

    // Drag and drop support
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

          availableCarriers = data.carriers;
          const container = document.getElementById('carriersContainer');
          container.innerHTML = '';
          const select = document.getElementById('carrierSelect');
          select.innerHTML = '';

          data.carriers.forEach(c => {
            container.innerHTML += `<div class="carrier-tag">${c.name} <span>${c.count}</span></div>`;
            select.innerHTML += `<option value="${c.name}">${c.name} (${c.count} рейсов)</option>`;
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
        alert('Ошибка при соединении с сервером: ' + err);
        processBtn.disabled = false;
        processBtn.innerText = '⚡ Обработать файл';
      }
    }

    let currentModalTarget = null; // null = all, 'select' = specific

    function openSendModal(target) {
      currentModalTarget = target;
      const email = document.getElementById('targetEmailInput').value;
      document.getElementById('modalTargetEmail').innerText = email;

      if (target === 'select') {
        document.getElementById('modalTitle').innerText = 'Выбор перевозчика для отправки';
        document.getElementById('carrierSelectGroup').style.display = 'block';
        updateModalTargetInfo();
      } else {
        document.getElementById('modalTitle').innerText = 'Отправка ВСЕМ перевозчикам';
        document.getElementById('carrierSelectGroup').style.display = 'none';
        document.getElementById('modalTargetInfo').innerHTML = `<b>Кому:</b> Все перевозчики (${availableCarriers.length} компаний)`;
      }
      document.getElementById('sendModal').classList.add('active');
    }

    function updateModalTargetInfo() {
      const select = document.getElementById('carrierSelect');
      const val = select.value;
      document.getElementById('modalTargetInfo').innerHTML = `<b>Кому:</b> ${val}`;
    }

    function closeSendModal() {
      document.getElementById('sendModal').classList.remove('active');
      document.getElementById('sendingStatus').style.display = 'none';
      document.getElementById('confirmSendBtn').disabled = false;
    }

    async function executeEmailSend() {
      const mode = document.querySelector('input[name="sendMode"]:checked').value;
      const email = document.getElementById('targetEmailInput').value;
      let carrier = null;
      if (currentModalTarget === 'select') {
        carrier = document.getElementById('carrierSelect').value;
      }

      const confirmBtn = document.getElementById('confirmSendBtn');
      confirmBtn.disabled = true;
      document.getElementById('sendingStatus').style.display = 'block';

      try {
        const resp = await fetch('/api/send-emails', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ mode: mode, carrier: carrier, email: email })
        });
        const res = await resp.json();
        if (res.status === 'success') {
          alert('Успешно! ' + res.message);
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
  </script>
</body>
</html>"""
    return HTMLResponse(content=html_content)

@app.post("/api/upload")
async def upload_file(file: UploadFile = File(...)):
    try:
        temp_dir = tempfile.gettempdir()
        temp_path = os.path.join(temp_dir, file.filename)
        with open(temp_path, "wb") as f:
            shutil.copyfileobj(file.file, f)
            
        # Запускаем парсинг и генерацию
        plan_trips, start_date = core.parse_plan_file(temp_path)
        all_trips = core.build_schedule_from_plan(plan_trips, start_date, OUTPUT_FILE)
        
        # Считаем статистику
        carrier_counts = Counter([t['carrier'] for t in all_trips if t.get('carrier')])
        dates = [t['date'] for t in all_trips]
        week_range = f"{dates[0]} - {dates[-1]}" if dates else "28.09 - 04.10"
        
        carriers_list = [{"name": k, "count": v} for k, v in carrier_counts.most_common()]
        
        CURRENT_STATE["all_trips"] = all_trips
        CURRENT_STATE["week_range"] = week_range
        CURRENT_STATE["total_trips"] = len(all_trips)
        CURRENT_STATE["carriers"] = carriers_list
        
        return {
            "status": "success",
            "week_range": week_range,
            "total_trips": len(all_trips),
            "carriers": carriers_list
        }
    except Exception as e:
        return {"status": "error", "message": str(e)}

@app.get("/api/download")
def download_schedule():
    if not os.path.exists(OUTPUT_FILE):
        raise HTTPException(status_code=404, detail="Файл еще не сформирован")
    return FileResponse(
        OUTPUT_FILE,
        filename="Недельные графики (готовый).xlsx",
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

@app.post("/api/send-emails")
def send_emails(req: EmailRequest):
    trips_to_send = CURRENT_STATE.get("all_trips", [])
    if not trips_to_send:
        # Если в памяти нет, пробуем перечитать дефолтный файл
        default_plan = core.find_default_plan_file(BASE_DIR)
        if default_plan and os.path.exists(default_plan):
            plan_trips, start_date = core.parse_plan_file(default_plan)
            trips_to_send = core.build_schedule_from_plan(plan_trips, start_date, OUTPUT_FILE)
            CURRENT_STATE["all_trips"] = trips_to_send
            
    if not trips_to_send:
        return {"status": "error", "message": "Нет данных для отправки. Сначала загрузите файл плана!"}

    # Фильтрация по перевозчику, если указан конкретный
    if req.carrier:
        target_trips = [t for t in trips_to_send if t['carrier'] == req.carrier]
        target_carriers = [req.carrier]
    else:
        target_trips = trips_to_send
        target_carriers = sorted(list(set(t['carrier'] for t in trips_to_send if t.get('carrier'))))

    if not target_trips:
        return {"status": "error", "message": "Рейсы для указанного перевозчика не найдены."}

    draft_mode = (req.mode == "draft")
    
    # Отправляем через Outlook
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
            sorted_trips = sorted(c_trips, key=lambda x: (x.get('dt') or datetime.strptime(x['date'], '%d.%m.%Y'), x['time']))
            mail = outlook.CreateItem(0)
            mail.To = req.email
            mail.Subject = f"[{car}] График погрузки РЦ Черная Грязь ({start_d} - {end_d})"
            mail.HTMLBody = core.generate_carrier_html(car, sorted_trips, start_d, end_d)
            
            if draft_mode:
                mail.Save()
            else:
                mail.Send()
            sent_count += 1
            
        action_text = "сохранено в Черновиках" if draft_mode else "отправлено"
        msg = f"Успешно {action_text} писем: {sent_count} шт. на адрес {req.email}"
        return {"status": "success", "message": msg, "count": sent_count}
        
    except ImportError:
        # Для Mac/Linux или если pywin32 не установлен
        action_text = "черновиков" if draft_mode else "отправки"
        msg = f"Библиотека pywin32 не установлена (требуется Windows + Outlook). Готово к работе {len(target_carriers)} писем {action_text}."
        return {"status": "success", "message": msg, "count": len(target_carriers)}
    except Exception as e:
        return {"status": "error", "message": f"Ошибка Outlook: {str(e)}"}

if __name__ == '__main__':
    port = 8000
    print(f"\n" + "="*70)
    print(f"🚀 Запуск локального сервера графиков отгрузок: http://localhost:{port}")
    print(f"="*70)
    try:
        webbrowser.open(f"http://localhost:{port}")
    except Exception:
        pass
    uvicorn.run(app, host="127.0.0.1", port=port)
