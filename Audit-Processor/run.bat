@echo off
cd /d "%~dp0"
python -m streamlit run audit_app.py
pause
