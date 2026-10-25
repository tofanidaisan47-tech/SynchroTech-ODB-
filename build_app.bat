@echo off
setlocal enabledelayedexpansion
title Build SyncroTech ODB - GMF AeroAsia Workstation
cd /d "%~dp0"

echo =====================================================================
echo  BUILDING SYNCROTECH ODB (GMF AEROASIA DIAGNOSTIC WORKSTATION)
echo  Combined Executable: API 8810A Data Logger + Synchro Voltage Monitor
echo =====================================================================
echo.

python -m PyInstaller --clean --noconfirm "SyncroTech_ODB.spec"

if errorlevel 1 (
    echo.
    echo [ERROR] Build failed! Check the error log above.
    pause
    exit /b 1
)

echo.
echo [1/4] Ensuring password_config.json is present in dist directory...
copy /Y "password_config.json" "dist\SyncroTech_ODB\password_config.json" >nul

echo [2/4] Copying Excel Master & Template workbooks for easy operator access...
if exist "datalogger_repo\MASTER_rev3.xlsx" (
    copy /Y "datalogger_repo\MASTER_rev3.xlsx" "dist\SyncroTech_ODB\MASTER_rev3.xlsx" >nul
)
if exist "synchro_app\2708 Voltase Accuracy STALL, AUTOPILOT, FLAP.xlsx" (
    copy /Y "synchro_app\2708 Voltase Accuracy STALL, AUTOPILOT, FLAP.xlsx" "dist\SyncroTech_ODB\2708 Voltase Accuracy STALL, AUTOPILOT, FLAP.xlsx" >nul
)
if exist "synchro_app\synchro_test_template.xlsx" (
    copy /Y "synchro_app\synchro_test_template.xlsx" "dist\SyncroTech_ODB\synchro_test_template.xlsx" >nul
)

echo [3/4] Creating Quick Launcher and Operator Guide...
(
    echo @echo off
    echo cd /d "%%~dp0"
    echo start "" "SyncroTech ODB.exe"
) > "dist\SyncroTech_ODB\Buka_SyncroTech_ODB.bat"

(
    echo =====================================================================
    echo  PANDUAN SINGKAT - SYNCROTECH ODB (GMF AEROASIA DIAGNOSTIC WORKSTATION)
    echo =====================================================================
    echo.
    echo 1. CARA MENJALANKAN:
    echo    - Cukup double-click file "SyncroTech ODB.exe" atau "Buka_SyncroTech_ODB.bat".
    echo    - Masukkan PIN Operator (Default: 000).
    echo    - Tekan AUTHENTICATE atau Enter.
    echo.
    echo 2. SUBSISTEM YANG TERSEDIA:
    echo    [01] API 8810A Data Logger
    echo         Inspeksi flap rigging dan sudut protractor digital.
    echo         Menggunakan koneksi Serial RS-232 atau Ethernet TCP/IP.
    echo    [02] Synchro Voltage Monitor
    echo         Pengukuran voltase akurasi synchro motor (Agilent 34401A).
    echo         Mendukung GPIB / PyVISA dan mode simulasi.
    echo    [1+2] Launch Both Instruments
    echo         Membuka kedua instrumen sekaligus dalam proses terisolasi.
    echo.
    echo 3. FILE EXCEL MASTER / TEMPLATE:
    echo    - MASTER_rev3.xlsx : Master sheet untuk API 8810A Data Logger.
    echo    - 2708 Voltase Accuracy STALL, AUTOPILOT, FLAP.xlsx : Master sheet Synchro.
    echo    File-file ini sudah diletakkan langsung di folder ini agar mudah
    echo    diedit, disalin, atau diganti oleh operator teknisi GMF.
    echo.
    echo 4. KEAMANAN & PASSWORD:
    echo    - Konfigurasi password tersimpan di file "password_config.json".
    echo.
    echo Catatan untuk Laptop Perusahaan:
    echo - Aplikasi ini bersifat PORTABLE (tidak perlu install Python di laptop perusahaan).
    echo - Cukup copy seluruh folder "SyncroTech_ODB" atau extract file ZIP ke laptop.
    echo - Jika menggunakan komunikasi hardware GPIB, pastikan kabel/driver VISA (NI-VISA/Keysight) terhubung.
    echo =====================================================================
) > "dist\SyncroTech_ODB\PANDUAN_PENGGUNAAN.txt"

echo [4/4] Creating distribution ZIP package for direct transfer...
powershell -Command "if (Test-Path 'SyncroTech_ODB_Release.zip') { Remove-Item 'SyncroTech_ODB_Release.zip' -Force }; Compress-Archive -Path 'dist\SyncroTech_ODB' -DestinationPath 'SyncroTech_ODB_Release.zip' -Force"

echo.
echo =====================================================================
echo  BUILD DAN PACKAGING BERHASIL!
echo.
echo  Folder Aplikasi Siap Pakai:
echo  dist\SyncroTech_ODB\SyncroTech ODB.exe
echo.
echo  File ZIP Siap Kirim ke Laptop Perusahaan:
echo  SyncroTech_ODB_Release.zip
echo =====================================================================
echo.
pause
