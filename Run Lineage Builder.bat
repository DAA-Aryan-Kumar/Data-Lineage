@echo off
rem Launches the Data Lineage Builder GUI quietly, then closes this window.
rem Prefers Anaconda's windowless pythonw.exe (which has pandas + openpyxl) so
rem no console window flashes or lingers.
setlocal
cd /d "%~dp0"

set "UI=lineage_ui.py"

rem Fail loudly if the UI script is missing, instead of letting pythonw die
rem silently on a non-existent file.
if not exist "%UI%" (
    echo ERROR: Could not find "%UI%" in this folder:
    echo   %~dp0
    echo The launcher must sit in the same folder as the script.
    pause
    exit /b 1
)

rem Anaconda / Miniconda base env (the one with pandas + openpyxl).
for %%P in (
    "%USERPROFILE%\anaconda3\pythonw.exe"
    "%USERPROFILE%\miniconda3\pythonw.exe"
    "%LOCALAPPDATA%\anaconda3\pythonw.exe"
    "%LOCALAPPDATA%\miniconda3\pythonw.exe"
    "%ProgramData%\anaconda3\pythonw.exe"
) do (
    if exist "%%~P" (
        start "" "%%~P" "%UI%"
        exit /b
    )
)

rem Otherwise use a python on PATH, but only if it has the libraries
rem (avoids a silent pythonw failure on an interpreter that lacks them).
where python >nul 2>nul
if %errorlevel%==0 (
    python -c "import pandas, openpyxl" >nul 2>nul
    if not errorlevel 1 (
        start "" pythonw "%UI%"
        exit /b
    )
)

echo Could not find a Python with pandas + openpyxl ^(e.g. Anaconda^).
echo Install them with:  pip install pandas openpyxl
pause
