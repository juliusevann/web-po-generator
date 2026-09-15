import os
from dotenv import load_dotenv

load_dotenv()

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

class Config:
    # --- Koneksi SQL Server (SSMS) ---
    # Dipakai langsung oleh excel_generator.py (query size-pack untuk template TH/TK/VH/VL)
    DB_DRIVER = os.getenv("DB_DRIVER", "ODBC Driver 17 for SQL Server")
    DB_SERVER = os.getenv("DB_SERVER", "")
    DB_DATABASE = os.getenv("DB_DATABASE", "")
    DB_USERNAME = os.getenv("DB_USERNAME", "")
    DB_PASSWORD = os.getenv("DB_PASSWORD", "")

    # --- po_api (service Flask terpisah yang expose /data, /vh_vl, /pdf) ---
    PO_API_BASE_URL = os.getenv("PO_API_BASE_URL", "http://10.1.2.11:5000")

    # --- Folder-folder penting (WAJIB disesuaikan di .env untuk mesin Windows kamu) ---
    # Folder berisi template Excel (PO EB Bags.xlsx, PO TH TEMPLATE.xlsx, dst)
    EXCEL_TEMPLATE_FOLDER = os.getenv(
        "EXCEL_TEMPLATE_FOLDER", os.path.join(BASE_DIR, "po_template")
    )
    # Folder berisi template Word untuk PDF (PO TEMPLATE CONVERT.docx)
    PDF_TEMPLATE_FOLDER = os.getenv(
        "PDF_TEMPLATE_FOLDER", os.path.join(BASE_DIR, "pdf_template")
    )
    # Folder network drive foto sample produk (dipakai insert_images di excel_generator.py)
    # Contoh asli: P:\SAMPLE PHOTO
    IMAGE_FOLDER = os.getenv("IMAGE_FOLDER", "P:\\SAMPLE PHOTO")

    # Folder kerja sementara (upload file excel & hasil generate sebelum di-zip)
    # Dipakai bersama oleh PO dan Barcode - aman karena nama file upload
    # sudah dibikin unik (prefix timestamp) di app.py.
    UPLOAD_FOLDER = os.getenv("UPLOAD_FOLDER", os.path.join(BASE_DIR, "uploads"))
    OUTPUT_FOLDER = os.getenv("OUTPUT_FOLDER", os.path.join(BASE_DIR, "outputs"))

    # --- Flask ---
    SECRET_KEY = os.getenv("SECRET_KEY", "ubah-dengan-kunci-rahasia-sendiri")
    PORT = int(os.getenv("PORT", "8000"))
    HOST = os.getenv("HOST", "0.0.0.0")

    # --- Login sederhana (Basic Auth) ---
    # WAJIB diisi kalau web app ini diakses lewat internet (misal via ngrok),
    # supaya tidak sembarang orang yang punya link bisa masuk & pakai.
    APP_USERNAME = os.getenv("APP_USERNAME", "")
    APP_PASSWORD = os.getenv("APP_PASSWORD", "")

    # Berapa lama (jam) file output/zip disimpan sebelum otomatis dibersihkan
    RETENTION_HOURS = int(os.getenv("RETENTION_HOURS", "24"))

    # Berapa banyak job (upload) boleh diproses BERSAMAAN untuk PO. Tiap job
    # yang jalan = 1 proses worker terpisah yang buka instance Excel + Word
    # sendiri-sendiri. Makin besar angka ini, makin berat RAM/CPU server -
    # mulai dari angka kecil (2) dan naikkan kalau server-nya kuat.
    MAX_CONCURRENT_JOBS = int(os.getenv("MAX_CONCURRENT_JOBS", "2"))

    # ============================================================
    # --- BARCODE: koneksi & folder khusus (sengaja dipisah dari yang
    # di atas, karena barcode_generator.py connect ke database YANG
    # BEDA (Y2_PROD) dari yang dipakai excel_generator.py punya PO) ---
    # ============================================================
    BARCODE_DB_DRIVER = os.getenv("BARCODE_DB_DRIVER", "SQL Server")
    BARCODE_DB_SERVER = os.getenv("BARCODE_DB_SERVER", "")
    BARCODE_DB_DATABASE = os.getenv("BARCODE_DB_DATABASE", "")
    BARCODE_DB_USERNAME = os.getenv("BARCODE_DB_USERNAME", "")
    BARCODE_DB_PASSWORD = os.getenv("BARCODE_DB_PASSWORD", "")

    # Folder foto khusus barcode label (beda dari IMAGE_FOLDER punya PO)
    BARCODE_PHOTO_FOLDER = os.getenv(
        "BARCODE_PHOTO_FOLDER", "P:\\SAMPLE PHOTO\\barcode label"
    )
    # File Excel master material (kode item+warna -> material & deskripsi)
    BARCODE_MATERIAL_FILE = os.getenv(
        "BARCODE_MATERIAL_FILE", "P:\\SAMPLE PHOTO\\MASTER MATERIAL DETAIL-new.xls"
    )
    # Folder berisi file .xlsx deskripsi produk EB (kolom "Develop Code" &
    # "Product Description Web")
    BARCODE_DESCRIPTION_FOLDER = os.getenv(
        "BARCODE_DESCRIPTION_FOLDER", os.path.join(BASE_DIR, "description")
    )

    # Output barcode dipisah total dari OUTPUT_FOLDER punya PO (folder
    # sendiri), supaya proses cleanup & scan job dua fitur ini tidak
    # saling bentrok/butuh saling tahu satu sama lain.
    BARCODE_OUTPUT_FOLDER = os.getenv(
        "BARCODE_OUTPUT_FOLDER", os.path.join(BASE_DIR, "outputs_barcode")
    )

    # Sama seperti MAX_CONCURRENT_JOBS tapi khusus barcode - dipisah
    # supaya bisa diatur independen (barcode tidak buka Excel/Word
    # otomasi seperti PO, jadi biasanya lebih ringan per job).
    BARCODE_MAX_CONCURRENT_JOBS = int(os.getenv("BARCODE_MAX_CONCURRENT_JOBS", "2"))
