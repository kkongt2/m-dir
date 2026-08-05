@echo off
setlocal
cd /d "%~dp0"

python -m pip install -r requirements-build.txt
python -m PyInstaller --onefile --windowed --name MultiPaneExplorer multipane_explorer.py

if errorlevel 1 exit /b %errorlevel%
echo Built dist\MultiPaneExplorer.exe
