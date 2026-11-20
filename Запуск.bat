@echo off
chcp 65001 >nul
cd /d "%~dp0"
python -c "import streamlit, docx, openpyxl" 2>nul || (
  echo Устанавливаю библиотеки, это один раз...
  python -m pip install -r requirements.txt
)
echo Запуск приложения. Чтобы остановить - закройте это окно.
python -m streamlit run app.py
pause
