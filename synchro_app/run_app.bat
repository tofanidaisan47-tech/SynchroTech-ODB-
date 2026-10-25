@echo off
title VFlap - Synchro Voltage Measurement System
cd /d "%~dp0"
echo Menjalankan VFlap - Synchro Voltage Measurement System...
if exist ".venv\Scripts\python.exe" (
    ".venv\Scripts\python.exe" main.py
) else (
    python main.py
)
if %ERRORLEVEL% NEQ 0 (
    echo.
    echo Terjadi kesalahan saat menjalankan aplikasi.
    pause
)
