@echo off
chcp 65001 > nul
title Веб-сервер графиков отгрузок (ФК ПУЛЬС)
echo ========================================================
echo Запуск локального сервера графиков отгрузок РЦ Черная Грязь
echo ========================================================
echo.

if exist .venv\Scripts\activate.bat (
    echo Активация виртуального окружения .venv...
    call .venv\Scripts\activate.bat
) else (
    if exist ..\.venv\Scripts\activate.bat (
        echo Активация виртуального окружения ..\.venv...
        call ..\.venv\Scripts\activate.bat
    )
)

echo Запуск сервера...
python server.py
pause
