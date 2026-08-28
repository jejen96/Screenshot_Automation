"""
main.py
-------
Tanggung jawab: Entry point aplikasi. Mengorkestrasi semua modul lain.

Perubahan dari versi sebelumnya:
- Tidak ada input folder penyimpanan — screenshot tidak disimpan ke file.
- Loop monitoring memanggil capture_screenshot() → PIL Image di memori.
- PIL Image diteruskan langsung ke excel_manager.add_entry(image).
- Tidak ada file PNG permanen yang dibuat.

Alur kerja:
1. Tampilkan menu di terminal.
2. Terima input: interval (detik) dan durasi (menit).
3. Konfirmasi.
4. Loop monitoring:
   a. Ambil screenshot → PIL Image di RAM
   b. Kirim PIL Image ke Excel report → thumbnail masuk Excel
   c. Tampilkan status di terminal
   d. Tunggu interval berikutnya
5. Finalisasi Excel saat selesai.

Hubungan dengan file lain:
- Mengimpor screenshot.py, excel_report.py, utils.py, config.py, logger.py
- File ini yang dijalankan: python main.py
"""

import time
import signal
import sys
from datetime import datetime

from config import APP_TITLE, APP_VERSION
from logger import get_logger, log_exception
from screenshot import capture_screenshot
from excel_report import ExcelReportManager
from utils import (
    validate_positive_integer,
    seconds_to_display,
    get_time_str,
    calculate_end_time,
    calculate_total_screenshots,
    print_separator,
    print_header,
    print_status,
)

logger = get_logger()


# =============================================================================
# GLOBAL STATE
# =============================================================================

_monitoring_active = False


def signal_handler(sig, frame):
    """Handler Ctrl+C — hentikan monitoring dengan bersih."""
    global _monitoring_active
    if _monitoring_active:
        print("\n\nMonitoring dihentikan oleh pengguna.")
        _monitoring_active = False
    else:
        print("\n\nTerima kasih!")
        sys.exit(0)


signal.signal(signal.SIGINT, signal_handler)


# =============================================================================
# INPUT DARI TERMINAL
# =============================================================================

def get_user_input() -> tuple:
    """
    Menampilkan menu di terminal dan menerima input pengguna.

    Input yang diminta:
    1. Interval screenshot (detik).
    2. Durasi monitoring (menit).
    3. Konfirmasi mulai.

    Returns:
        tuple(int, int):
            - interval_seconds: Interval antar screenshot.
            - duration_seconds: Total durasi monitoring.
    """
    print_header(APP_TITLE, APP_VERSION)

    # Input interval
    while True:
        raw = input("Interval screenshot (detik): ").strip()
        ok, msg, interval_seconds = validate_positive_integer(raw, "Interval")
        if ok:
            break
        print(f"  ❌ {msg}")

    # Input durasi
    while True:
        raw = input("Durasi monitoring (menit)  : ").strip()
        ok, msg, duration_minutes = validate_positive_integer(raw, "Durasi")
        if ok:
            duration_seconds = duration_minutes * 60
            break
        print(f"  ❌ {msg}")

    # Ringkasan & konfirmasi
    print()
    print_separator()
    print(f"  Interval   : {interval_seconds} detik")
    print(f"  Durasi     : {duration_minutes} menit")
    print(f"  Output     : reports/Report_<tanggal>.xlsx")
    print_separator()

    confirm = input("\nMulai monitoring? (Y/N): ").strip().upper()
    if confirm != "Y":
        print("Monitoring dibatalkan.")
        sys.exit(0)

    return interval_seconds, duration_seconds


# =============================================================================
# MONITORING LOOP
# =============================================================================

def run_monitoring(interval_seconds: int, duration_seconds: int) -> None:
    """
    Loop monitoring utama.

    Setiap interval:
    1. Ambil screenshot → PIL Image di RAM.
    2. Kirim PIL Image ke ExcelReportManager.
    3. Thumbnail langsung masuk Excel.
    4. Tampilkan status di terminal.
    5. Tunggu interval berikutnya.

    Args:
        interval_seconds : Interval antar screenshot (detik).
        duration_seconds : Total durasi monitoring (detik).
    """
    global _monitoring_active
    _monitoring_active = True

    logger.info("=" * 60)
    logger.info("MONITORING STARTED")
    logger.info(f"Interval: {interval_seconds}s | Duration: {duration_seconds}s")
    logger.info("=" * 60)

    # Inisialisasi Excel report
    excel_manager = ExcelReportManager()

    # Hitung waktu selesai dan total target screenshot
    end_time = calculate_end_time(duration_seconds)
    total_screenshots = calculate_total_screenshots(duration_seconds, interval_seconds)
    screenshot_count = 0

    print()
    print("Monitoring dimulai... (Tekan Ctrl+C untuk berhenti)")
    print()

    while _monitoring_active:
        current_time = datetime.now()

        # Cek durasi
        if current_time >= end_time:
            break

        # ── AMBIL SCREENSHOT ──────────────────────────────────────────────
        # capture_screenshot() mengembalikan PIL Image di RAM
        # Tidak ada file PNG yang dibuat
        image = capture_screenshot()

        if image:
            # ── MASUKKAN KE EXCEL ─────────────────────────────────────────
            # add_entry() menerima PIL Image langsung
            # Thumbnail dibuat di dalam excel_report.py
            excel_manager.add_entry(image)
            screenshot_count += 1

            # Bebaskan memori setelah gambar diproses
            image.close()

        # Hitung sisa waktu
        remaining_seconds = max(
            0, int((end_time - datetime.now()).total_seconds())
        )

        # Hitung waktu screenshot berikutnya
        next_time = datetime.fromtimestamp(
            datetime.now().timestamp() + interval_seconds
        )

        # Tampilkan status di terminal
        print_status(
            status="RUNNING",
            screenshot_count=screenshot_count,
            total_screenshots=total_screenshots,
            next_screenshot_at=get_time_str(next_time),
            remaining_time=seconds_to_display(remaining_seconds),
        )

        # Tunggu interval, cek Ctrl+C setiap detik
        for _ in range(interval_seconds):
            if not _monitoring_active:
                break
            time.sleep(1)

    # ── SELESAI ───────────────────────────────────────────────────────────
    excel_manager.finalize()

    logger.info("=" * 60)
    logger.info("MONITORING COMPLETED")
    logger.info(f"Total screenshots: {screenshot_count}")
    logger.info(f"Excel report: {excel_manager.filepath}")
    logger.info("=" * 60)

    print()
    print_separator()
    print("  MONITORING SELESAI")
    print(f"  Total Screenshot : {screenshot_count}")
    print(f"  Excel Report     : {excel_manager.filepath.name}")
    print(f"  Lokasi           : {excel_manager.filepath.parent}")
    print_separator()


# =============================================================================
# ENTRY POINT
# =============================================================================

def main() -> None:
    """Entry point. Dipanggil saat: python main.py"""
    try:
        interval_seconds, duration_seconds = get_user_input()
        run_monitoring(interval_seconds, duration_seconds)
    except KeyboardInterrupt:
        print("\n\nMonitoring dihentikan.")
    except Exception as e:
        log_exception(logger, e, "Saat menjalankan monitoring")
        print(f"\n❌ Error: {e}")
        print("Lihat file log untuk detail.")


if __name__ == "__main__":
    main()
