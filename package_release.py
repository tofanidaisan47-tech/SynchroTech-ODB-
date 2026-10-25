import os
import shutil
import zipfile

dist = os.path.abspath('dist/SyncroTech_ODB')
print('Dist dir:', dist)

# 1. Config
shutil.copy2('password_config.json', dist)

# 2. Excel templates
for f in [
    'datalogger_repo/MASTER_rev3.xlsx',
    'synchro_app/2708 Voltase Accuracy STALL, AUTOPILOT, FLAP.xlsx',
    'synchro_app/synchro_test_template.xlsx',
]:
    if os.path.exists(f):
        shutil.copy2(f, dist)
        print('Copied:', f)

# 3. Batch launcher
bat_content = "@echo off\r\ncd /d \"%~dp0\"\r\nstart \"\" \"SyncroTech ODB.exe\"\r\n"
with open(os.path.join(dist, 'Buka_SyncroTech_ODB.bat'), 'w', encoding='ascii') as fp:
    fp.write(bat_content)

# 4. User guide
guide = """=====================================================================
 PANDUAN PENGGUNAAN - SYNCROTECH ODB (GMF AEROASIA DIAGNOSTIC WORKSTATION)
=====================================================================

1. CARA MENJALANKAN DI LAPTOP PERUSAHAAN:
   - Aplikasi ini bersifat 100% PORTABLE (TIDAK MEMERLUKAN INSTALASI PYTHON).
   - Seluruh library, driver komunikasi, dan GUI sudah dibundle di dalam paket ini.
   - Cukup double-click file 'SyncroTech ODB.exe' atau 'Buka_SyncroTech_ODB.bat'.
   - Masukkan PIN Operator (Default: 000), lalu tekan AUTHENTICATE.

2. FITUR & SUBSISTEM TERSEDIA:
   [01] API 8810A Data Logger
        - Pengukuran & inspeksi sudut Flap Rigging Digital Protractor (API 8810A).
        - Koneksi via Serial RS-232 atau Ethernet TCP/IP.
        - Logging otomatis ke file Excel dengan deteksi kestabilan real-time.
   [02] Synchro Voltage Monitor
        - Pengukuran presisi tegangan Synchro Motor Flap (Agilent 34401A).
        - Mendukung deteksi port GPIB / PyVISA otomatis & mode simulasi.
        - Deteksi auto-save 1.6 detik untuk setiap titik sudut (SP Degrees).
   [1+2] Launch Both Instruments
        - Menjalankan kedua instrumen secara bersamaan dalam proses terisolasi.

3. TEMPLATE FILE EXCEL:
   - MASTER_rev3.xlsx : Master sheet untuk API 8810A Data Logger.
   - 2708 Voltase Accuracy STALL, AUTOPILOT, FLAP.xlsx : Master sheet Synchro.
   File-file ini sengaja diletakkan langsung di folder utama ini agar
   operator dapat langsung membuka, mengedit, atau membuat salinannya.

4. PENGATURAN PASSWORD:
   - Password terenkripsi disimpan di file 'password_config.json'.
   - Password default: 000

=====================================================================
 Tim Pengembang: Daisan Tofani, Odayakana Upupha Konnikof, Bramantyo Al Fajri
 PT GMF AeroAsia Tbk - Avionics & Component Services
=====================================================================
"""
with open(os.path.join(dist, 'PANDUAN_PENGGUNAAN.txt'), 'w', encoding='utf-8') as fp:
    fp.write(guide)

print('Creating ZIP archive...')
zip_path = 'SyncroTech_ODB_Release.zip'
if os.path.exists(zip_path):
    os.remove(zip_path)

with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zf:
    for root, dirs, files in os.walk(dist):
        for file in files:
            full_path = os.path.join(root, file)
            rel_path = os.path.relpath(full_path, os.path.dirname(dist))
            zf.write(full_path, rel_path)

print('ZIP archive created successfully:', os.path.abspath(zip_path))
print('Size (MB):', round(os.path.getsize(zip_path) / (1024 * 1024), 2))
