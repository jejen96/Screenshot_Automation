"""
utils.py
--------
Tanggung jawab: Fungsi-fungsi utilitas yang dipakai oleh berbagai modul.

Prinsip DRY (Don't Repeat Yourself):
- Tulis sekali di sini, pakai dari mana saja.
- Tidak ada duplikasi logika di seluruh project.

Ciri fungsi yang cocok di utils.py:
- Fungsinya kecil dan sederhana.
- Tidak spesifik milik satu modul.
- Tidak mengubah state global.

Hubungan dengan file lain:
- Diimpor oleh screenshot.py, excel_report.py, main.py.
- Tidak mengimpor modul aplikasi lain.
"""

from datetime import datetime, timedelta
from pathlib import Path
from typing import Optional


# =============================================================================
# VALIDASI INPUT TERMINAL
# =============================================================================

def validate_positive_integer(value: str, field_name: str) -> tuple:
    """
    Memvalidasi bahwa input adalah angka bulat positif.

    Args:
        value     : String input dari terminal.
        field_name: Nama field untuk pesan error.

    Returns:
        tuple(bool, str, int):
            - bool: True jika valid.
            - str : Pesan error (kosong jika valid).
            - int : Nilai integer (0 jika tidak valid).

    Contoh:
        ok, msg, val = validate_positive_integer("5", "Interval")
        ok, msg, val = validate_positive_integer("abc", "Interval")
    """
    # Kosong?
    if not value or not value.strip():
        return False, f"{field_name} tidak boleh kosong.", 0

    value = value.strip()

    # Bukan angka?
    if not value.isdigit():
        return False, f"{field_name} harus berupa angka bulat.", 0

    val = int(value)

    # Nol atau negatif?
    if val <= 0:
        return False, f"{field_name} harus lebih besar dari 0.", 0

    return True, "", val


def validate_folder_path(path_str: str) -> tuple:
    """
    Memvalidasi path folder yang dimasukkan pengguna.

    Validasi:
    1. Tidak kosong.
    2. Folder harus ada di filesystem.
    3. Harus berupa directory (bukan file).
    4. Harus bisa ditulis (write permission).

    Args:
        path_str: String path dari input terminal.

    Returns:
        tuple(bool, str, Optional[Path]):
            - bool           : True jika valid.
            - str            : Pesan error.
            - Optional[Path] : Path object jika valid.
    """
    # Kosong → pakai default
    if not path_str or not path_str.strip():
        return False, "Path tidak boleh kosong.", None

    folder = Path(path_str.strip())

    # Tidak ada?
    if not folder.exists():
        return False, f"Folder tidak ditemukan: {folder}", None

    # Bukan folder?
    if not folder.is_dir():
        return False, "Path yang dimasukkan bukan folder.", None

    # Tidak bisa ditulis?
    test_file = folder / ".write_test"
    try:
        test_file.touch()
        test_file.unlink()
    except (PermissionError, OSError):
        return False, "Tidak punya izin menulis ke folder ini.", None

    return True, "", folder


# =============================================================================
# FORMAT WAKTU
# =============================================================================

def seconds_to_display(total_seconds: int) -> str:
    """
    Mengubah jumlah detik ke format HH:MM:SS untuk ditampilkan di terminal.

    Args:
        total_seconds: Jumlah detik.

    Returns:
        str: Format HH:MM:SS

    Contoh:
        seconds_to_display(3661) → "01:01:01"
        seconds_to_display(125)  → "00:02:05"
        seconds_to_display(0)    → "00:00:00"
    """
    if total_seconds < 0:
        total_seconds = 0
    hours = total_seconds // 3600
    minutes = (total_seconds % 3600) // 60
    seconds = total_seconds % 60
    return f"{hours:02d}:{minutes:02d}:{seconds:02d}"


def get_date_str(dt: Optional[datetime] = None) -> str:
    """
    Mengembalikan tanggal dalam format YYYY-MM-DD.

    Args:
        dt: Datetime. Jika None, gunakan waktu sekarang.

    Returns:
        str: Tanggal format YYYY-MM-DD.
    """
    if dt is None:
        dt = datetime.now()
    return dt.strftime("%Y-%m-%d")


def get_time_str(dt: Optional[datetime] = None) -> str:
    """
    Mengembalikan waktu dalam format HH:MM:SS.

    Args:
        dt: Datetime. Jika None, gunakan waktu sekarang.

    Returns:
        str: Waktu format HH:MM:SS.
    """
    if dt is None:
        dt = datetime.now()
    return dt.strftime("%H:%M:%S")


def calculate_end_time(duration_seconds: int) -> datetime:
    """
    Menghitung waktu berakhirnya monitoring.

    Args:
        duration_seconds: Durasi dalam detik.

    Returns:
        datetime: Waktu selesai monitoring.
    """
    return datetime.now() + timedelta(seconds=duration_seconds)


def calculate_total_screenshots(duration_seconds: int, interval_seconds: int) -> int:
    """
    Menghitung total screenshot yang akan diambil.

    Args:
        duration_seconds : Total durasi monitoring dalam detik.
        interval_seconds : Interval antar screenshot dalam detik.

    Returns:
        int: Perkiraan jumlah screenshot.
    """
    if interval_seconds <= 0:
        return 0
    return duration_seconds // interval_seconds


# =============================================================================
# DISPLAY HELPERS
# =============================================================================

def print_separator(char: str = "=", width: int = 50) -> None:
    """
    Mencetak garis pemisah di terminal.

    Args:
        char : Karakter untuk garis.
        width: Lebar garis.
    """
    print(char * width)


def print_header(title: str, version: str) -> None:
    """
    Mencetak header aplikasi di terminal.

    Args:
        title  : Judul aplikasi.
        version: Versi aplikasi.
    """
    print_separator()
    print(f"  {title} v{version}")
    print_separator()
    print()


def print_status(
    status: str,
    screenshot_count: int,
    total_screenshots: int,
    next_screenshot_at: str,
    remaining_time: str,
) -> None:
    """
    Mencetak status monitoring real-time di terminal.

    Menggunakan \\r dan end="" agar output ditimpa di baris yang sama
    sehingga terminal tidak scroll terus-menerus.

    Args:
        status           : Status saat ini (RUNNING / STOPPED / dll).
        screenshot_count : Jumlah screenshot yang sudah diambil.
        total_screenshots: Total target screenshot.
        next_screenshot_at: Waktu screenshot berikutnya (HH:MM:SS).
        remaining_time   : Sisa waktu (HH:MM:SS).
    """
    print_separator()
    print(f"  STATUS       : {status}")
    print(f"  Screenshot   : {screenshot_count} / {total_screenshots}")
    print(f"  Next at      : {next_screenshot_at}")
    print(f"  Remaining    : {remaining_time}")
    print_separator()
