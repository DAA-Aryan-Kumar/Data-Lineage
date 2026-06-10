@echo off
rem Launches the Data Lineage Builder GUI.
rem Requires Python with pandas + openpyxl on PATH (anaconda works fine).
cd /d "%~dp0"
where python >nul 2>nul
if errorlevel 1 (
    echo Python was not found on PATH. Install Python with pandas + openpyxl.
    pause
    exit /b 1
)
python lineage_ui.py
if errorlevel 1 pause
