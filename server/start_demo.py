# Запуск демо (ASCII-имя файла, чтобы .bat не зависел от кодировки консоли)
import os, runpy, sys
here = os.path.dirname(os.path.abspath(__file__))
os.chdir(here)
sys.path.insert(0, here)
runpy.run_path(os.path.join(here, 'демо.py'), run_name='__main__')
