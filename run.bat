@echo off
chcp 65001 >nul
set PYTHONUTF8=1
set PYTHONIOENCODING=utf-8
cd /d "%~dp0"

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

echo Запуск ТаможенФормат (веб)...
"%PY%" main.py

if errorlevel 1 (
  echo.
  echo Приложение завершилось с ошибкой. См. сообщение выше.
  pause
)
