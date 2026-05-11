@echo off
title SAP Migration Validator
echo Starting SAP Migration Validator...
echo.

:: Check if streamlit is installed
streamlit --version >nul 2>&1
if %errorlevel% neq 0 (
    echo Streamlit not found. Installing...
    pip install streamlit openpyxl pandas
    echo.
)

:: Get the directory where this .bat file lives
cd /d "%~dp0"

:: Open browser and run app
streamlit run app.py --server.headless false --browser.gatherUsageStats false

pause
