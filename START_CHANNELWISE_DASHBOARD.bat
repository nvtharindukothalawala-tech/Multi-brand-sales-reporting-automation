@echo off
setlocal
cd /d "%~dp0"
if not exist ".venv\Scripts\python.exe" (
    echo Python environment not found. Open STREAMLIT_SETUP.md for setup instructions.
    pause
    exit /b 1
)
".venv\Scripts\python.exe" -m streamlit run streamlit_app.py --server.address 127.0.0.1
if errorlevel 1 (
    echo Dashboard failed to launch. Install Streamlit as described in STREAMLIT_SETUP.md.
    pause
)
