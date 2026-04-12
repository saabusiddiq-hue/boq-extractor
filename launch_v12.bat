@echo off
echo ==========================================
echo    BOQ Extractor Pro v12.0 - Launching...
echo    Interactive Preview Edition
echo ==========================================
echo.

REM Check if Python is installed
python --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python is not installed or not in PATH
    echo Please install Python 3.8+ from https://python.org
    pause
    exit /b 1
)

REM Check if requirements are installed
echo Checking dependencies...
pip show streamlit >nul 2>&1
if errorlevel 1 (
    echo Installing dependencies... This may take a few minutes.
    pip install -r requirements_v12.txt
    if errorlevel 1 (
        echo ERROR: Failed to install dependencies
        pause
        exit /b 1
    )
)

echo.
echo Starting BOQ Extractor Pro v12.0...
echo Features: Interactive Preview + Material Anchors
echo.
streamlit run boq_extractor_app_v12.py

pause
