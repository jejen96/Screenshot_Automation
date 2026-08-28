"""
screenshot.py
-------------
Tanggung jawab: Mengambil screenshot layar dan mengembalikannya
sebagai objek PIL Image di memori — TIDAK menyimpan ke file.

Perubahan dari versi sebelumnya:
- Tidak ada folder screenshots/.
- Tidak ada file PNG permanen.
- Screenshot hidup di RAM sebagai PIL Image.
- excel_report.py yang bertanggung jawab menyisipkan ke Excel.

Mengapa PIL Image dan bukan bytes?
- PIL Image lebih fleksibel — bisa di-resize, dikonversi, dianalisis.
- Konversi ke bytes (untuk Excel) dilakukan di excel_report.py.
- Pemisahan yang bersih: screenshot.py hanya tahu cara "mengambil gambar",
  excel_report.py tahu cara "memasukkan gambar ke Excel".

Library:
- pyautogui : mengambil screenshot seluruh layar.
- Pillow    : tipe data PIL Image untuk menyimpan gambar di RAM.

Hubungan dengan file lain:
- Dipanggil oleh main.py setiap interval.
- Mengembalikan PIL Image ke main.py.
- main.py meneruskan PIL Image ke excel_report.py.
"""

from typing import Optional

import pyautogui
from PIL import Image

from logger import get_logger

logger = get_logger()


def capture_screenshot() -> Optional[Image.Image]:
    """
    Mengambil screenshot seluruh layar dan mengembalikannya sebagai PIL Image.

    Tidak ada file yang dibuat. Screenshot sepenuhnya ada di memori (RAM).

    Cara kerja:
    1. pyautogui.screenshot() mengambil screenshot layar.
    2. Hasilnya adalah PIL Image — langsung dikembalikan.
    3. Tidak ada operasi file sama sekali.

    Returns:
        Optional[Image.Image]:
            - PIL Image jika berhasil.
            - None jika gagal (error dicatat ke log).

    Contoh penggunaan:
        img = capture_screenshot()
        if img:
            # Lakukan sesuatu dengan img
            # Misalnya: excel_report.add_entry(img)
    """
    try:
        # pyautogui.screenshot() mengembalikan PIL Image secara langsung
        # Tidak perlu simpan ke file — gambar ada di RAM
        img: Image.Image = pyautogui.screenshot()

        logger.info(  # type: ignore[attr-defined]
            f"Screenshot captured | Size: {img.width}x{img.height}px"
        )
        return img

    except Exception as e:
        logger.error(f"Screenshot failed: {e}")
        return None
