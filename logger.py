"""
logger.py
---------
Tanggung jawab: Sistem logging terpusat untuk seluruh aplikasi.

Mengapa logging penting?
- Debugging: saat error terjadi, log menjelaskan apa yang terjadi.
- Audit: melacak aktivitas program untuk analisis.
- Monitoring: tahu berapa screenshot berhasil/gagal.

Hubungan dengan file lain:
- Diimpor oleh semua modul yang perlu logging.
- Menggunakan konstanta dari config.py.
"""

import logging
import sys
from typing import Optional

from config import LOGS_DIR, LOG_FORMAT, LOG_DATE_FORMAT, get_log_filename


# =============================================================================
# CUSTOM LOG LEVEL: SUCCESS
# =============================================================================
# Python default: DEBUG, INFO, WARNING, ERROR, CRITICAL
# Kita tambahkan SUCCESS (level 25, antara INFO dan WARNING)

SUCCESS_LEVEL = 25
logging.addLevelName(SUCCESS_LEVEL, "SUCCESS")


def success(self: logging.Logger, message: str, *args, **kwargs) -> None:
    """Method untuk logger.success('...')"""
    if self.isEnabledFor(SUCCESS_LEVEL):
        self._log(SUCCESS_LEVEL, message, args, **kwargs)


# Inject method success ke class Logger
logging.Logger.success = success  # type: ignore[attr-defined]


# =============================================================================
# SETUP LOGGER (Singleton)
# =============================================================================

_logger_instance: Optional[logging.Logger] = None


def get_logger(name: str = "screen_activity_logger") -> logging.Logger:
    """
    Mengembalikan logger instance yang sudah dikonfigurasi.

    Singleton Pattern:
    - Logger hanya dibuat SATU KALI saat pertama dipanggil.
    - Pemanggilan berikutnya mengembalikan instance yang sama.
    - Mencegah duplikasi handler.

    Args:
        name: Nama logger.

    Returns:
        logging.Logger: Logger siap pakai.

    Contoh penggunaan:
        from logger import get_logger
        logger = get_logger()

        logger.debug("Detail teknis")
        logger.info("Informasi umum")
        logger.success("Operasi berhasil")
        logger.warning("Peringatan")
        logger.error("Error terjadi")
    """
    global _logger_instance

    if _logger_instance is not None:
        return _logger_instance

    logger = logging.getLogger(name)
    logger.setLevel(logging.DEBUG)

    if logger.handlers:
        return logger

    formatter = logging.Formatter(fmt=LOG_FORMAT, datefmt=LOG_DATE_FORMAT)

    # Handler 1: File — simpan log ke file di folder logs/
    log_file_path = LOGS_DIR / get_log_filename()
    file_handler = logging.FileHandler(log_file_path, encoding="utf-8", mode="a")
    file_handler.setLevel(logging.DEBUG)
    file_handler.setFormatter(formatter)

    # Handler 2: Console — tampilkan log di terminal
    console_handler = logging.StreamHandler(sys.stdout)
    console_handler.setLevel(logging.INFO)  # Console hanya tampilkan INFO ke atas
    console_handler.setFormatter(formatter)

    logger.addHandler(file_handler)
    logger.addHandler(console_handler)

    _logger_instance = logger

    logger.info("=" * 60)
    logger.info(f"Logger initialized | Log file: {log_file_path.name}")
    logger.info("=" * 60)

    return logger


# =============================================================================
# HELPER FUNCTIONS
# =============================================================================

def log_exception(logger: logging.Logger, exc: Exception, context: str = "") -> None:
    """
    Helper untuk log exception dengan format konsisten.

    Args:
        logger : Logger instance.
        exc    : Exception yang terjadi.
        context: Konteks di mana error terjadi (opsional).
    """
    logger.error("-" * 60)
    if context:
        logger.error(f"ERROR CONTEXT: {context}")
    logger.error(f"ERROR TYPE: {type(exc).__name__}")
    logger.error(f"ERROR MESSAGE: {str(exc)}")
    logger.exception("TRACEBACK:")
    logger.error("-" * 60)
