@echo off
setlocal
title Build VFlap - Synchro Voltage Measurement System
cd /d "%~dp0"

if exist "dist\VFlap\2708 Voltase Accuracy STALL, AUTOPILOT, FLAP.xlsx" (
    del "dist\VFlap\2708 Voltase Accuracy STALL, AUTOPILOT, FLAP.xlsx"
)

".venv\Scripts\pyinstaller.exe" --clean --noconfirm "VFlap.spec"
if errorlevel 1 (
    echo.
    echo Build gagal. Periksa pesan error di atas.
    pause
    exit /b 1
)

echo Menyinkronkan file aplikasi VFlap ke folder utama...
copy /y "dist\VFlap\VFlap.exe" "VFlap.exe" >nul
if exist "_internal" rmdir /s /q "_internal"
robocopy "dist\VFlap\_internal" "_internal" /e /njh /njs /ndl /nc /ns >nul

echo.
echo Build berhasil! VFlap.exe siap dijalankan di folder synchro_app.
pause
