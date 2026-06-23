@echo off
cd /d "%~dp0"
echo Checking openpyxl...
pip install openpyxl 2>nul || python -m pip install openpyxl 2>nul || py -3 -m pip install openpyxl 2>nul
echo Starting...
py -3 server.py || python server.py
pause
