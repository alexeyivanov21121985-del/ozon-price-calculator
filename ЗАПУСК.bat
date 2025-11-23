@echo off
chcp 65001 >nul
echo Zapusk Ozon Calculator...
"C:\Users\sovaj\AppData\Local\Programs\Python\Python312\python.exe" -m streamlit run "C:\ozon_price_gui\app.py"
pause
