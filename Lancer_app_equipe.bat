@echo off
cd /d "%~dp0"
python -m streamlit run App_dealflow.py --server.address 0.0.0.0 --server.port 8501
pause
