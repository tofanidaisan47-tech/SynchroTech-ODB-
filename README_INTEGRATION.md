# SyncroTech ODB — Aircraft Diagnostic Workstation (GMF AeroAsia)

Integrated Windows Desktop Application yang menggabungkan:
1. **API 8810A Data Logger** (Flap Sensor Digital Protractor Rigging Inspection)
2. **Synchro Voltage Measurement System** (Agilent 34401A Multimeter for Flap Position Sensor)

---

## 🎨 Desain Antarmuka (Hallmark Anti-AI-Slop Framework)

Sesuai panduan desain autentik dari **Nutlope/hallmark**:
- **Bebas Klise AI**: Tidak menggunakan tata letak generik bertumpuk tengah (*centered hero cards*), emoji kartun, maupun efek gradien ungu-biru template.
- **Aerospace Cockpit Console**: Mengadopsi tata letak asimetris konsol diagnostik teknis dengan panel telemetri, spesifikasi perangkat keras, dan status bus terisolasi.
- **Tipografi Presisi**: Kombinasi `Segoe UI` untuk hirarki teks dan `Consolas` monospace untuk penanda telemetri, identitas stasiun, dan parameter instrumen.
- **Tema Warna Industrial Dark**: Menggunakan palet Slate & Obsidian dengan aksen *Cobalt/Cyan* untuk Digital Protractor dan *Industrial Amber* untuk Synchro Voltage Monitor.

---

## 🔐 Kredensial & Akses Operator

- **Default PIN / Password**: `000`
- **Keterangan Pengembang**: `Developed by Daisan Tofani, Odayakana Upupha Konnikof, Bramantyo Al Fajri  18 Oktober 2026`
- **Konfigurasi Keamanan**: Tersimpan pada `password_config.json` menggunakan algoritma hashing **PBKDF2-SHA256 (100.000 iterasi + salt)**.
- Password dapat dikustomisasi kapan saja melalui file konfigurasi.

---

## 🚀 Cara Menjalankan Aplikasi

### Opsi 1: Menjalankan Versi Executable (.exe) — Siap Pakai
Buka folder:
`dist\SyncroTech_ODB\`
Lalu klik dua kali pada:
👉 **`SyncroTech ODB.exe`**

### Opsi 2: Menjalankan via Batch Script (Mode Pengembang)
Klik dua kali pada file:
👉 **`run_app.bat`**

### Opsi 3: Menjalankan via Terminal
```powershell
python launcher.py
```

---

## 🛠️ Navigasi Subsystem

1. **API 8810A Data Logger**:
   - Protokol: RS-232 / TCP Ethernet
   - Target Workbook: `MASTER_rev3.xlsx`
   - Fitur: Live stability detection, circular angle math, continuous auto-save.
2. **Synchro Voltage Monitor**:
   - Protokol: GPIB / PyVISA Direct SCPI
   - Target Sheet: `2708 Accuracy Sheet.xlsx`
   - Fitur: AC voltage acquisition, test point stepping, VFD visualization.
3. **Dual Subsystem Inspection [1 + 2]**:
   - Membuka kedua instrumen sekaligus dalam proses terpisah secara simultan tanpa konflik resource atau tema.
4. **Lock Session & Exit**:
   - Tombol logout dan keluar yang secara otomatis menghentikan (*graceful shutdown*) seluruh proses instrumen yang aktif.

---

## 📦 Cara Membangun Ulang Executable (.exe)

1. Jalankan **`build_app.bat`**.
2. Script otomatis memanggil PyInstaller dengan spesifikasi `SyncroTech_ODB.spec`.
3. Output executable terbaru tersedia di `dist\SyncroTech_ODB\SyncroTech ODB.exe`.
