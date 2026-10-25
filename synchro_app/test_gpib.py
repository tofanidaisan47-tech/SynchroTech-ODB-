"""
Script Verifikasi Koneksi GPIB & Agilent 34401A
GMF AeroAsia - Synchro Voltage Measurement System
"""

import sys
import time

def main():
    print("=" * 65)
    print("   TES KONEKSI HARDWARE GPIB & MULTIMETER AGILENT 34401A")
    print("=" * 65)
    print()

    # 1. Cek PyVISA & VISA Backend
    try:
        import pyvisa
        rm = pyvisa.ResourceManager()
        print("[✓] Pustaka PyVISA dan VISA Engine terdeteksi.")
    except Exception as e:
        print(f"[X] Gagal memuat PyVISA atau VISA Engine: {e}")
        input("\nTekan Enter untuk keluar...")
        return

    # 2. Pindai Port / Resources
    print("\n[+] Memindai port yang terhubung ke laptop...")
    try:
        resources = rm.list_resources()
    except Exception as e:
        print(f"[X] Error saat scanning resources: {e}")
        resources = ()

    gpib_resources = [r for r in resources if "GPIB" in r]
    serial_resources = [r for r in resources if "ASRL" in r]

    print(f"    Total Port Terdeteksi: {len(resources)}")
    if serial_resources:
        print(f"    Port Serial (COM)    : {', '.join(serial_resources)}")
    if gpib_resources:
        print(f"    Port GPIB Ditemukan  : {', '.join(gpib_resources)}")
    else:
        print("    Port GPIB Ditemukan  : (Tidak ada)")

    # 3. Analisis Hasil Pindai
    if not gpib_resources:
        print("\n" + "-" * 65)
        print("[!] PERINGATAN: Perangkat GPIB belum terdeteksi.")
        print("-" * 65)
        print("Penyebab umum & Solusi:")
        print("1. Driver NI-488.2 belum selesai dipasang:")
        print("   -> Buka folder Downloads dan jalankan installer:")
        print("      C:\\Users\\USER\\Downloads\\ni-488.2_26.5_online.exe")
        print("2. Lampu READY di adapter USB-GPIB belum hijau mantap.")
        print("3. Multimeter belum dinyalakan atau kabel GPIB belum kencang.")
        print("4. Multimeter belum diset ke mode GPIB di menu front panel:")
        print("   -> Tekan Shift + Menu > I/O MENU > 1: HP-IB (Enter) > 2: ADDR 22 (Enter).")
        print("-" * 65)
        input("\nTekan Enter untuk keluar...")
        return

    # 4. Tes Komunikasi ke Alat Agilent 34401A
    target_resource = gpib_resources[0]
    print(f"\n[+] Mencoba berkomunikasi dengan alat di {target_resource}...")

    try:
        inst = rm.open_resource(target_resource)
        inst.timeout = 5000  # 5 detik

        # Identifikasi (*IDN?)
        idn = inst.query("*IDN?").strip()
        print(f"[✓] Berhasil terhubung ke: {idn}")

        # Konfigurasi AC Voltage
        print("[+] Mengatur instrumen ke mode AC VOLTAGE (CONF:VOLT:AC)...")
        inst.write("*RST")
        time.sleep(0.3)
        inst.write("CONF:VOLT:AC")
        time.sleep(0.3)

        # Ambil 1 sample pembacaan
        reading = inst.query("READ?").strip()
        print(f"[✓] Hasil Pembacaan Multimeter Saat Ini: {float(reading):.4f} VAC")
        print()
        print("=" * 65)
        print("   STATUS: KONEKSI GPIB & AGILENT 34401A SUKSES 100%!")
        print("   Silakan buka aplikasi utama lewat 'run_app.bat'")
        print("=" * 65)

        inst.close()

    except Exception as e:
        print(f"[X] Gagal membaca dari {target_resource}: {e}")
        print("Pastikan format SCPI pada multimeter aktif (Shift + Menu > I/O MENU > 3: SCPI).")

    print()
    input("Tekan Enter untuk keluar...")

if __name__ == "__main__":
    main()
