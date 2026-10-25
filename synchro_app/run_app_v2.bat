@echo off
title Synchro Measurement System - GMF AeroAsia (Versi Revisi)
cd /d "%~dp0"
if exist ".venv\Scripts\python.exe" (
    ".venv\Scripts\python.exe" main.py
) else (
    python main.py
)
pause
