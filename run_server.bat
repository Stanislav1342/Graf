@echo off
chcp 65001 > nul
title Веб-сервер графиков отгрузок (ФК ПУЛЬС)
echo ========================================================
echo Запуск локального сервера графиков отгрузок РЦ Черная Грязь
echo ========================================================
echo.

set "PY_CMD=python"

:: Проверяем локальное виртуальное окружение
if exist .venv\Scripts\python.exe (
    .venv\Scripts\python.exe --version >nul 2>&1
    if not errorlevel 1 (
        echo Использование виртуального окружения .venv...
        set "PY_CMD=.venv\Scripts\python.exe"
    ) else (
        echo [Инфо] Виртуальное окружение .venv заблокировано политикой, используем системный Python...
    )
)

echo Проверка необходимых библиотек Python...
%PY_CMD% -c "import multipart" 2>nul
if errorlevel 1 (
    echo [Установка] Установка python-multipart...
    %PY_CMD% -m pip install python-multipart
)

echo Запуск сервера...
%PY_CMD% server.py
pause
