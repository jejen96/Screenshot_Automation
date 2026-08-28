"""
config.py
---------
Tanggung jawab: Menyimpan semua konstanta dan path project.

Mengapa file ini dibuat terpisah?
- Single Source of Truth: satu tempat untuk semua konfigurasi.
- Jika ada perubahan, cukup ubah di sini — modul lain tidak perlu disentuh.

Perubahan dari versi sebelumnya:
- SCREENSHOTS_DIR dihapus — program tidak lagi menyimpan file PNG.
- THUMBNAIL_WIDTH diubah ke 200px sesuai spesifikasi baru.
- EXCEL_COLUMNS disesuaikan: hanya No, Tanggal, Waktu, Screenshot.

Hubungan dengan file lain:
- Diimpor oleh semua modul lain.
- Tidak mengimpor modul aplikasi lain (hindari circular import).
"""

from pathlib import Path
from datetime import datetime


# =============================================================================
# PATH FOLDER
# =============================================================================

BASE_DIR: Path = Path(__file__).parent.resolve()

# Program hanya membutuhkan dua folder:
# - reports/ : untuk file Excel
# - logs/    : untuk file log aktivitas
REPORTS_DIR: Path = BASE_DIR / "reports"
LOGS_DIR: Path = BASE_DIR / "logs"

# Buat folder otomatis saat pertama kali di-import
REPORTS_DIR.mkdir(exist_ok=True)
LOGS_DIR.mkdir(exist_ok=True)


# =============================================================================
# KONFIGURASI EXCEL
# =============================================================================

EXCEL_COLUMNS: list = ["No", "Tanggal", "Waktu", "Screenshot"]

# Ukuran thumbnail di Excel (pixel)
# 480px = sekitar 1/3 layar 1366px — cukup besar untuk dibaca teks di dalamnya
# Tinggi dihitung proporsional dari rasio layar:
# Layar 1366x768 → rasio 16:9 → tinggi = 480 × (768/1366) ≈ 270px
THUMBNAIL_WIDTH: int = 480
THUMBNAIL_HEIGHT_RATIO: float = 768 / 1366   # Rasio tinggi/lebar layar umum

# Tinggi baris Excel dalam points untuk menampung thumbnail
# 270px ÷ 1.33pt/px ≈ 203pt — tambah padding → 210pt
EXCEL_ROW_HEIGHT: int = 210


def get_excel_filename() -> str:
    """
    Generate nama file Excel berdasarkan timestamp saat program dijalankan.

    Format : Report_YYYY-MM-DD_HH-MM-SS.xlsx
    Contoh : Report_2026-07-07_14-30-00.xlsx

    Setiap kali python main.py dijalankan → file Excel baru dibuat.
    Semua file tersimpan di folder reports/ sehingga Anda bisa
    melihat riwayat setiap sesi monitoring.

    Returns:
        str: Nama file Excel unik per sesi.
    """
    timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    return f"Report_{timestamp}.xlsx"


# =============================================================================
# KONFIGURASI LOGGING
# =============================================================================

LOG_FORMAT: str = "%(asctime)s | %(levelname)-8s | %(message)s"
LOG_DATE_FORMAT: str = "%Y-%m-%d %H:%M:%S"


def get_log_filename() -> str:
    """
    Generate nama file log untuk hari ini.

    Format : activity_YYYY-MM-DD.log
    Contoh : activity_2026-07-07.log

    Returns:
        str: Nama file log.
    """
    date = datetime.now().strftime("%Y-%m-%d")
    return f"activity_{date}.log"


# =============================================================================
# INFORMASI APLIKASI
# =============================================================================

APP_TITLE: str = "Screen Activity Logger"
APP_VERSION: str = "1.0.0"
