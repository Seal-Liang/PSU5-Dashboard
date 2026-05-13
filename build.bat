@echo off
echo Installing dependencies...
python -m pip install -r requirements.txt
python -m pip install pyinstaller

echo:
echo Building application (this might take a few minutes)...
python -m PyInstaller --noconfirm --noconsole --name "WeeklyDashboard" --add-data "static;static" --hidden-import openpyxl --hidden-import pandas --hidden-import flaskwebgui --onefile server.py

echo:
echo Creating standalone release folder...
if not exist "WeeklyDashboard_Release" mkdir "WeeklyDashboard_Release"
copy /y "dist\WeeklyDashboard.exe" "WeeklyDashboard_Release\" >nul
if exist "data.json" (
    copy /y "data.json" "WeeklyDashboard_Release\" >nul
    echo [OK] Copied data.json into release folder.
) else (
    echo [WARNING] data.json not found!
    echo Please manually place data.json into the Release folder if you want to transfer data.
)

echo:
echo Build completed successfully!
echo The ready-to-share application is in the 'WeeklyDashboard_Release' folder.
echo You can simply compress this folder into a ZIP and send it to your colleagues!
pause
