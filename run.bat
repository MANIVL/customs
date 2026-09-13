@echo off
chcp 65001 >nul
set PYTHONUTF8=1
set PYTHONIOENCODING=utf-8
cd /d "%~dp0"

:: ── Настройка Yandex GPT ──
:: Установите API-ключ через переменную окружения:
::   set YANDEX_GPT_API_KEY=ваш_ключ
:: Или вставьте ключ ниже вместо YOUR_YANDEX_API_KEY_HERE
if not defined YANDEX_GPT_API_KEY set "YANDEX_GPT_API_KEY=YOUR_YANDEX_API_KEY_HERE"
set "YANDEX_MODEL_URN=gpt://b1g9m7q1q3s2t4u5v6w/yandexgpt-lite"

set "PY="
where python >nul 2>nul && set "PY=python"
if not defined PY (
  if exist "python-runtime\python.exe" set "PY=python-runtime\python.exe"
)

if not defined PY (
  echo.
  echo [ERROR] Не найден Python. Установите Python 3.11+ или положите папку python-runtime рядом с этим файлом.
  echo.
  pause
  exit /b 1
)

:: ── Установить зависимости если нужно ──
if not exist "venv" (
  echo Создаю виртуальное окружение...
  "%PY%" -m venv venv
)
call venv\Scripts\activate.bat

if exist "requirements.txt" (
  pip install -q -r requirements.txt 2>nul
)

echo Запуск ТаможенФормат (веб)...
if defined YANDEX_GPT_API_KEY (
  if defined YANDEX_MODEL_URN (
    echo ИИ-анализ файлов: Yandex GPT подключён
  ) else (
    echo ИИ-анализ файлов: нужен YANDEX_MODEL_URN в run.bat
  )
) else (
  echo ИИ-анализ файлов: не настроен (будет использоваться стандартный режим)
)
"%PY%" main.py

if errorlevel 1 (
  echo.
  echo Приложение завершилось с ошибкой. См. сообщение выше.
  pause
)
