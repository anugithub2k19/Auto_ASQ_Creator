@echo off
cd /d "%~dp0"
echo Starting ASQ Ticket Migration Tool...
echo.
.venv\Scripts\python.exe web/app.py
echo.
echo App exited with code %errorlevel%
pause