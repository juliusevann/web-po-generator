import os
from dotenv import load_dotenv

load_dotenv()

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

class Config:
    DB_DRIVER = os.getenv("DB_DRIVER", "ODBC Driver 18 for SQL Server")
    DB_SERVER = os.getenv("DB_SERVER", "")
    DB_DATABASE = os.getenv("DB_DATABASE", "")
    DB_USERNAME = os.getenv("DB_USERNAME", "")
    DB_PASSWORD = os.getenv("DB_PASSWORD", "")

    PO_API_BASE_URL = os.getenv("PO_API_BASE_URL", "http://10.1.2.11:5000")

    EXCEL_TEMPLATE_FOLDER = os.getenv(
        "EXCEL_TEMPLATE_FOLDER", os.path.join(BASE_DIR, "po_template")
    )
    
    PDF_TEMPLATE_FOLDER = os.getenv(
        "PDF_TEMPLATE_FOLDER", os.path.join(BASE_DIR, "pdf_template")
    )

    IMAGE_FOLDER = os.getenv("IMAGE_FOLDER", "P:\\SAMPLE PHOTO")

    # Folder kerja sementara (upload file excel & hasil generate sebelum di-zip)
    # Dipakai bersama oleh PO dan Barcode - aman karena nama file upload
    # sudah dibikin unik (prefix timestamp) di app.py.
    UPLOAD_FOLDER = os.getenv("UPLOAD_FOLDER", os.path.join(BASE_DIR, "uploads"))
    OUTPUT_FOLDER = os.getenv("OUTPUT_FOLDER", os.path.join(BASE_DIR, "outputs"))

    SECRET_KEY = os.getenv("SECRET_KEY", "ubah-dengan-kunci-rahasia-sendiri")
    PORT = int(os.getenv("PORT", "8000"))
    HOST = os.getenv("HOST", "0.0.0.0")

    APP_USERNAME = os.getenv("APP_USERNAME", "")
    APP_PASSWORD = os.getenv("APP_PASSWORD", "")


    RETENTION_HOURS = int(os.getenv("RETENTION_HOURS", "24"))


    MAX_CONCURRENT_JOBS = int(os.getenv("MAX_CONCURRENT_JOBS", "2"))


    BARCODE_DB_DRIVER = os.getenv("BARCODE_DB_DRIVER", "SQL Server")
    BARCODE_DB_SERVER = os.getenv("BARCODE_DB_SERVER", "")
    BARCODE_DB_DATABASE = os.getenv("BARCODE_DB_DATABASE", "")
    BARCODE_DB_USERNAME = os.getenv("BARCODE_DB_USERNAME", "")
    BARCODE_DB_PASSWORD = os.getenv("BARCODE_DB_PASSWORD", "")

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


    BARCODE_OUTPUT_FOLDER = os.getenv(
        "BARCODE_OUTPUT_FOLDER", os.path.join(BASE_DIR, "outputs_barcode")
    )


    BARCODE_MAX_CONCURRENT_JOBS = int(os.getenv("BARCODE_MAX_CONCURRENT_JOBS", "2"))
