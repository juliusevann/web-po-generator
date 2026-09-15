import os
import time

import requests
import win32com.client as win32
from docxtpl import DocxTemplate

from config import Config

TEMPLATE_NAME = "PO TEMPLATE CONVERT.docx"


def get_api_data(gl_souche, gl_numero):
    location = "CWH" if gl_souche == "JKT" else "WMKR"
    url = f"{Config.PO_API_BASE_URL}/pdf?gl_souche={location}&gl_numero={gl_numero}"
    try:
        response = requests.get(url, timeout=10)
        if response.status_code == 200:
            return response.json()
    except Exception:
        pass
    return None


def format_number(value):
    value = float(value)
    return f"{value:,.2f}"


def parse_ref(ref_raw):
    """
    Parsing format REF: BRAND_GENDER_KATEGORI_PRODUK_RONA_LOKASI
    contoh: EB_LDS_SHS_GABIN1_NA_JKT
    -> jenis_combo="EB_LDS_SHS", produk="GABIN1", ro_na="NA", lokasi="JKT"

    Di sini "jenis" TIDAK dipakai buat memilih file template, karena hanya
    ada satu template (PO TEMPLATE CONVERT.docx) untuk semua kategori.
    """
    parts = [p.strip() for p in str(ref_raw).strip().split("_")]
    if len(parts) < 6:
        raise ValueError(
            f"Format referensi salah (butuh 6 segmen dipisah '_', dapat {len(parts)}): '{ref_raw}'"
        )
    jenis_combo = "_".join(parts[0:3]).upper()
    produk = parts[3]
    ro_na = parts[4].upper()
    lokasi = parts[5].upper()
    return jenis_combo, produk, ro_na, lokasi


def fill_docx_template(template_path, output_path, api_data, ref_info):
    jenis, produk, ro_na, lokasi = ref_info

    subtotal = sum(float(item["nett"]) for item in api_data["item"])
    subqty = sum(float(item["qty"]) for item in api_data["item"])

    addressFrom = """
    JL. MARSEKAL SURYADHARMA NO. 01A 
    KEL. NEGLASARI, KEC. NEGLASARI 
    KOTA TANGERANG, PROP. BANTEN
    Phone : (021) 2252 9090 / (021) 2252 9393
    """ if lokasi == "JKT" else """
    KOMPLEKS RUKO LATIMOJONG
    JL. NICO BLOK H NO. 19 MAKASSAR
    01.813.768.7-046.000
    Phone : 0411-3651053 /
    """

    addressTo = """
    CENTRAL WAREHOUSE JAKARTA
    JL. MARSEKAL SURYADHARMA NO. 01A
    KEL NEGLASARI, KEC NEGLASARI, TANGERANG BANTEN
    BANTEN
    """ if lokasi == "JKT" else """
    Warehouse Makasar 
    KOMPLEKS RUKO LATIMOJONG
    JL. NICO BLOK H NO. 19 MAKASSAR
    SOUTH SULAWESI
    """

    items = []
    for item in api_data["item"]:
        items.append({
            "reference": item.get("ref", "-"),
            "desc": item.get("desc", ""),
            "qty": format_number(item.get("qty", 0)),
            "nett": format_number(item.get("nett", 0)),
            "upper": item.get("upper", "-"),
            "tax": format_number(item.get("tax", 0)),
        })

    context = {
        "kode_po": api_data["PO_NO"],
        "date": api_data["date"],
        "ext_ref": api_data["ext_ref"],
        "delivery_date": api_data["delivery_date"],
        "items": items,
        "subtotal": format_number(subtotal),
        "sub_qty": format_number(subqty),
        "total_qty": format_number(subqty),
        "first_tax": format_number(subtotal),
        "address_from": addressFrom,
        "address_to": addressTo,
    }

    doc = DocxTemplate(template_path)
    doc.render(context)
    doc.save(output_path)


def convert_to_pdf(docx_path, pdf_path, delete_docx=True):

    from core import com_lock
    com_lock.acquire()
    word = None
    try:
        word = win32.Dispatch("Word.Application")
        doc = word.Documents.Open(os.path.abspath(docx_path))
        try:
            doc.SaveAs(os.path.abspath(pdf_path), FileFormat=17)
        finally:
            try:
                doc.Close()
            except Exception as close_err:
                print(f"[PDF] Peringatan: gagal Close() dokumen (diabaikan): {close_err}", flush=True)
    finally:
        if word is not None:
            try:
                word.Quit()
            except Exception as quit_err:
                print(f"[PDF] Peringatan: gagal Quit() Word (diabaikan): {quit_err}", flush=True)
        com_lock.release()

    if delete_docx:
        time.sleep(1)
        os.remove(docx_path)


def generate_po_pdf(po_code, ref, pdf_dir):
    """
    Entry point yang dipanggil job_manager.py.
    Return (ok: bool, path: str|None, err: str|None).
    """
    try:
        jenis_combo, produk, ro_na, lokasi = parse_ref(ref)
    except ValueError as e:
        return False, None, str(e)

    ref_info = (jenis_combo, produk, ro_na, lokasi)
    
    api_data = get_api_data(lokasi, po_code)
    if not api_data or api_data.get("message") != "Sukses":
        return False, None, f"Gagal ambil data dari API untuk PO {po_code}"

    template_path = os.path.join(Config.PDF_TEMPLATE_FOLDER, TEMPLATE_NAME)
    if not os.path.exists(template_path):
        return False, None, f"Template tidak ditemukan: {template_path}"

    os.makedirs(pdf_dir, exist_ok=True)
    file_base = f"PO_{po_code}_{produk}_{ro_na}_{lokasi}"
    output_docx = os.path.join(pdf_dir, file_base + ".docx")
    output_pdf = os.path.join(pdf_dir, file_base + ".pdf")

    try:
        fill_docx_template(template_path, output_docx, api_data, ref_info)
        convert_to_pdf(output_docx, output_pdf, delete_docx=True)
    except Exception as e:
        msg = f"Gagal generate PDF untuk PO {po_code}: {e}"
        return False, None, msg

    return True, output_pdf, None