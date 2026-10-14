#!/bin/bash
cd "$(dirname "$0")"
if [ ! -f "демо.py" ]; then
  echo "Архив не распакован. Распакуйте его двойным щелчком и запустите этот файл из распакованной папки."
  read -p "Нажмите Enter, чтобы закрыть..."; exit 1
fi
if ! command -v python3 >/dev/null 2>&1; then
  echo "Python не найден. Установите Python с python.org (или согласитесь на установку инструментов разработчика) и запустите снова."
  read -p "Нажмите Enter, чтобы закрыть..."; exit 1
fi
if ! python3 -c "import openpyxl, docx, pypdf, docxtpl, xlrd" >/dev/null 2>&1; then
  echo "Устанавливаю библиотеки (один раз)..."
  python3 -m pip install --user openpyxl python-docx pypdf "docxtpl[subdoc]" xlrd 2>/dev/null || python3 -m pip install --user --break-system-packages openpyxl python-docx pypdf "docxtpl[subdoc]" xlrd
fi
python3 демо.py
read -p "Нажмите Enter, чтобы закрыть..."
