@echo off
chcp 65001 >nul
cd /d "%~dp0"
title Демо панели УООЗД
if not exist "start_demo.py" (
  echo Архив не распакован. Правой кнопкой по архиву - "Извлечь все", затем запустите этот файл из распакованной папки.
  pause
  exit /b
)
set "PY="
py -3 --version >nul 2>nul && set "PY=py -3"
if not defined PY (
  python --version >nul 2>nul && set "PY=python"
)
if not defined PY (
  echo.
  echo Python не найден.
  echo Установите Python с python.org и при установке отметьте "Add Python to PATH".
  echo Если на рабочем компьютере программы ставит только ИТ - попросите их.
  pause
  exit /b
)
echo Python: %PY%
%PY% -c "import openpyxl, docx, pypdf, docxtpl, xlrd" 2>nul
if errorlevel 1 (
  echo Устанавливаю библиотеки - нужен интернет, 1-2 минуты...
  %PY% -m pip install --user openpyxl python-docx pypdf "docxtpl[subdoc]" xlrd
  if errorlevel 1 (
    echo.
    echo Не удалось установить библиотеки. Скопируйте текст выше и пришлите.
    pause
    exit /b
  )
)
rem pywin32 нужен только для прямого COM-подключения к 1С; без него демо работает, а раздел 1С покажет подсказку
%PY% -c "import pythoncom" 2>nul
if errorlevel 1 (
  echo Ставлю pywin32 для прямого подключения к 1С ^(необязательно^)...
  %PY% -m pip install --user pywin32 2>nul
)
%PY% -X utf8 start_demo.py
echo.
echo Панель остановлена. Если выше есть ошибка - скопируйте её текст.
pause
