@echo off
setlocal
cd /d "%~dp0"

python -m pip install -r requirements-build.txt
python -m PyInstaller --clean --noconfirm MultiPaneExplorer.spec

if errorlevel 1 exit /b %errorlevel%
echo Built dist\MultiPaneExplorer.exe
