@echo off
rem Double-click: reads every PDF in data\ and writes result\yields.xlsx and result\yields.csv
cd /d "%~dp0"
where /q python
if %errorlevel%==0 (
    python extract_yields.py
) else (
    py extract_yields.py
)
echo.
pause
