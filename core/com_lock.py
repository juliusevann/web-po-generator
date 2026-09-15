"""
Lock lintas-proses sederhana berbasis file, untuk memastikan hanya SATU
proses yang menyentuh Excel/Word (COM automation) di satu waktu.

KENAPA INI DIPERLUKAN: Microsoft TIDAK mendukung otomatisasi Office
(Excel/Word via COM automation) untuk dipakai bersamaan oleh banyak proses
sekaligus di server - ini bukan keterbatasan kode kita, ini batasan resmi
dari Office sendiri. Kalau dipaksa jalan bersamaan, muncul error semacam
"RPC server is unavailable" (Word) atau "Server execution failed" (Excel),
karena proses COM server (WINWORD.EXE / EXCEL.EXE) gagal start atau crash
saat direbut beberapa proses sekaligus.

Solusinya: serialize BAGIAN yang menyentuh Excel/Word lewat lock file ini.
Job lain tetap bisa jalan bersamaan untuk bagian lain (baca file, panggil
API, dst), tapi begitu masuk ke bagian buka Excel/Word, cuma satu proses
yang boleh jalan di satu waktu - proses lain menunggu sebentar sampai
lock ini dilepas.

Konsekuensi: kalau 2 job diproses "bersamaan", bagian Excel/Word-nya tetap
akan antre sebentar satu-satu di titik ini - tapi ini pengorbanan yang
diperlukan supaya tidak crash, karena Office memang tidak didesain untuk
dipakai begini.
"""

import os
import time

from config import Config

_LOCK_PATH = os.path.join(Config.OUTPUT_FOLDER, ".office_com.lock")
_STALE_SECONDS = 300  # anggap lock basi kalau proses pemegangnya crash tanpa sempat release


def acquire(timeout=180, poll_interval=0.3):
    """Tunggu sampai dapat lock (atau timeout). Panggil sebelum menyentuh Excel/Word."""
    os.makedirs(Config.OUTPUT_FOLDER, exist_ok=True)
    start = time.time()
    while True:
        try:
            fd = os.open(_LOCK_PATH, os.O_CREAT | os.O_EXCL | os.O_WRONLY)
            os.write(fd, str(os.getpid()).encode())
            os.close(fd)
            return
        except FileExistsError:
            # Cek apakah lock ini "basi" - pemegangnya kemungkinan crash
            # tanpa sempat release, jadi lock-nya nggak akan pernah dilepas
            # normal. Kalau lebih tua dari _STALE_SECONDS, ambil paksa.
            try:
                age = time.time() - os.path.getmtime(_LOCK_PATH)
                if age > _STALE_SECONDS:
                    os.remove(_LOCK_PATH)
                    continue
            except OSError:
                continue
            if time.time() - start > timeout:
                raise TimeoutError(
                    "Gagal mendapatkan lock Excel/Word automation (timeout) - "
                    "kemungkinan ada proses lain yang macet lama memegang lock ini."
                )
            time.sleep(poll_interval)


def release():
    """Panggil setelah selesai menyentuh Excel/Word, idealnya di blok finally."""
    try:
        os.remove(_LOCK_PATH)
    except OSError:
        pass