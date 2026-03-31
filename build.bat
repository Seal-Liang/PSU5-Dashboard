@echo off
chcp 65001 >nul
echo Installing PyInstaller...
python -m pip install pyinstaller

echo:
echo Building application (this might take a few minutes)...
python -m PyInstaller --noconfirm --name "WeeklyDashboard" --add-data "static;static" --hidden-import openpyxl --hidden-import pandas --onefile server.py

echo:
echo Build completed!
echo You can find 'WeeklyDashboard.exe' in the 'dist' folder.
pause
