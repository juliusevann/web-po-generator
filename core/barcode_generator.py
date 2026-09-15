import datetime as _datetime
import glob
import os
import re
import shutil

import pyodbc
import xlrd
import xlwt

from config import Config

# ============================================================
# CONCERN: import cv2, qrcode, PIL (Image), numpy DIHAPUS dari versi asli.
# Di script asli, semua kode yang benar-benar memakai library itu (sketch
# processing pakai cv2, bikin QR code pakai qrcode) sudah di-comment-out
# semua - jadi library itu sebenarnya tidak dipakai di jalur aktif mana pun.
# Menghapusnya mengurangi dependency yang wajib diinstall di server web.
# Kalau ternyata kode yang di-comment itu masih mau dipakai suatu saat,
# tinggal import lagi.
# ============================================================


# ============================================================
# ADAPTER FORMAT FILE: xlrd (.xls, engine lama) HANYA BISA baca .xls,
# TIDAK BISA baca .xlsx sama sekali (dihapus dukungannya sejak xlrd 2.0).
# Supaya barcode_generator bisa terima DUA-DUANYA (.xls dan .xlsx, karena
# contoh file dari mentor ternyata .xlsx), dibuat adapter tipis di sini -
# jadi seluruh kode grouping & pemrosesan di bawah TETAP sama, tidak perlu
# tahu file aslinya .xls atau .xlsx.
# ============================================================
class _XlsxSheetAdapter:
    """Meniru interface xlrd Sheet (cell_value(row, col), .nrows) tapi baca .xlsx via openpyxl."""

    def __init__(self, ws):
        self._ws = ws
        self.nrows = ws.max_row

    def cell_value(self, row, col):
        # xlrd: index 0-based. openpyxl: index 1-based. +1 di sini menjembataninya.
        val = self._ws.cell(row=row + 1, column=col + 1).value
        return '' if val is None else val


class _XlsxWorkbookAdapter:
    """Meniru xlrd Book (cuma butuh .datemode, tidak dipakai untuk .xlsx
    karena openpyxl sudah otomatis balikin objek datetime, bukan serial number)."""
    datemode = 0


def _open_sheet(input_path):
    """Buka file .xls ATAU .xlsx, return (workbook_like, sheet_like) dengan
    interface seragam (nrows, cell_value) supaya sisa kode tidak perlu tahu
    bedanya."""
    ext = os.path.splitext(input_path)[1].lower()

    if ext == '.xls':
        wb = xlrd.open_workbook(input_path)
        return wb, wb.sheet_by_index(0)

    if ext == '.xlsx':
        import openpyxl
        wb_raw = openpyxl.load_workbook(input_path, data_only=True)
        ws_raw = wb_raw.worksheets[0]
        return _XlsxWorkbookAdapter(), _XlsxSheetAdapter(ws_raw)

    raise ValueError(f"Format file tidak didukung: '{ext}' (harus .xls atau .xlsx)")


# ============================================================
# CACHE (lazy-loaded, supaya tidak baca file Excel master berkali-kali
# tiap job baru - sama pola seperti _photo_index_cache di excel_generator.py)
# ============================================================
_material_map_cache = None
_eb_description_map_cache = None


def normalize_key(s: str) -> str:
    """Normalisasi string untuk matching: uppercase, underscore/spasi jadi 1 spasi."""
    s = str(s).strip().upper()
    s = re.sub(r'[_\s]+', ' ', s)
    return s


def _load_material_map():
    global _material_map_cache
    if _material_map_cache is not None:
        return _material_map_cache

    material_map = {}
    material_file = Config.BARCODE_MATERIAL_FILE
    if os.path.exists(material_file):
        wb_material = xlrd.open_workbook(material_file)
        ws_material = wb_material.sheet_by_index(0)
        for row in range(1, ws_material.nrows):
            key = normalize_key(ws_material.cell_value(row, 2))
            material = str(ws_material.cell_value(row, 3)).strip()
            description = str(ws_material.cell_value(row, 5)).strip()
            material_map[key] = {"material": material, "description": description}
        print(f"[BARCODE] Master material loaded: {len(material_map)} entries", flush=True)
    else:
        print(f"[BARCODE][WARNING] MASTER MATERIAL DETAIL tidak ditemukan: {material_file}", flush=True)

    _material_map_cache = material_map
    return material_map


def find_material_info(item_code: str, colour: str, material_map: dict) -> dict:
    norm_item = normalize_key(item_code)
    norm_colour = normalize_key(colour)
    query_key = f"{norm_item} {norm_colour}".strip()

    if query_key in material_map:
        return material_map[query_key]

    for key, info in material_map.items():
        if norm_item in key and norm_colour in key:
            return info

    for key, info in material_map.items():
        if norm_item in key:
            return info

    return {"material": "Material not found", "description": ""}


def format_batchno(date_value, workbook):
    month_letters = "ABCDEFGHIJKL"
    try:
        # Kasus file .xlsx (openpyxl): tanggal sudah berupa objek datetime asli
        if isinstance(date_value, (_datetime.datetime, _datetime.date)):
            month = date_value.month
            year = date_value.year % 100
            return f"{month_letters[month - 1]}{year:02d}"

        # Kasus file .xls (xlrd): tanggal berupa serial number
        if isinstance(date_value, (float, int)):
            dt = xlrd.xldate_as_tuple(date_value, getattr(workbook, "datemode", 0))
            month = dt[1]
            year = dt[0] % 100
            return f"{month_letters[month - 1]}{year:02d}"

        if isinstance(date_value, str) and date_value.strip():
            month_str, _, year_str = date_value.strip().split('/')
            month = int(month_str)
            year = int(year_str) % 100
            return f"{month_letters[month - 1]}{year:02d}"
    except Exception:
        pass
    return ""


def get_db_connection():
    conn_str = (
        f"DRIVER={{{Config.BARCODE_DB_DRIVER}}};"
        f"SERVER={Config.BARCODE_DB_SERVER};"
        f"DATABASE={Config.BARCODE_DB_DATABASE};"
        f"UID={Config.BARCODE_DB_USERNAME};"
        f"PWD={Config.BARCODE_DB_PASSWORD};"
        f"TrustServerCertificate=yes;"
    )
    return pyodbc.connect(conn_str)


def get_article_info(conn, item_code, barcode):
    sql = """
    select
        GA_FAMILLENIV3 as Segment,
        CE.YX_LIBELLE as UpperStyle,
        GA_CHARLIBRE2 as Dimension,
        GA_CODEARTICLE as article,
        GA_CODEBARRE as barcode,
        GA_LIBELLE as deskripsi
    from ARTICLEDIM_MODE
        LEFT OUTER JOIN CHOIXEXT CE
            ON GA_LIBREART1 = CE.YX_CODE
            AND CE.YX_TYPE = 'LA1'
    where GA_CODEARTICLE = ?
      and GA_CODEBARRE = ?
    """
    cur = conn.cursor()
    cur.execute(sql, item_code, barcode)
    row = cur.fetchone()

    if not row:
        return {"upper_style": "", "dimension": "", "description": ""}

    return {
        "upper_style": row.UpperStyle or "",
        "dimension": row.Dimension or "",
        "description": row.deskripsi or "",
    }


def parse_dimension(dimension, desc):
    length = ''
    width = ''
    height = ''

    if not dimension:
        return length, width, height

    pattern = r'(\d+(?:\.\d+)?)\((L|W|H)\)'

    for value, dim_type in re.findall(pattern, dimension.upper()):
        value = f'{value}' if desc.upper() == 'VB' else f'{value}cm'
        if dim_type == 'L':
            length = value
        elif dim_type == 'W':
            width = value
        elif dim_type == 'H':
            height = value

    return length, width, height


def extract_first_sentence(html_text):
    if not html_text or str(html_text).strip().lower() == 'nan':
        return ""

    clean = re.sub(r'<[^>]+>', '', str(html_text))
    clean = re.sub(r'\s+', ' ', clean).strip()

    match = re.match(r'^(.*?\.)', clean)
    if match:
        return match.group(1).strip()
    return clean


def _load_eb_description_map_from_folder(folder):
    import pandas as pd

    desc_map = {}
    if not os.path.exists(folder):
        print(f"[BARCODE][WARNING] Folder description tidak ditemukan: {folder}", flush=True)
        return desc_map

    xlsx_files = [f for f in os.listdir(folder) if f.upper().endswith('.XLSX')]
    if not xlsx_files:
        print(f"[BARCODE][WARNING] Tidak ada file .xlsx di folder description: {folder}", flush=True)
        return desc_map

    for fname in xlsx_files:
        fpath = os.path.join(folder, fname)
        try:
            df = pd.read_excel(fpath, dtype=str)
            if 'Develop Code' not in df.columns or 'Product Description Web' not in df.columns:
                print(f"[BARCODE][WARNING] Kolom 'Develop Code' / 'Product Description Web' "
                      f"tidak ditemukan di: {fname}", flush=True)
                continue

            for _, row in df.iterrows():
                code = str(row['Develop Code']).strip()
                raw_html = str(row['Product Description Web']).strip()
                cleaned = extract_first_sentence(raw_html)
                if code and code.lower() != 'nan' and cleaned:
                    desc_map[code] = cleaned

            print(f"[BARCODE] Loaded description dari: {fname} ({len(desc_map)} entries)", flush=True)
        except Exception as e:
            print(f"[BARCODE][WARNING] Gagal membaca {fname}: {e}", flush=True)

    print(f"[BARCODE] Total EB description map: {len(desc_map)} entries", flush=True)
    return desc_map


def _load_eb_description_map():
    global _eb_description_map_cache
    if _eb_description_map_cache is not None:
        return _eb_description_map_cache
    _eb_description_map_cache = _load_eb_description_map_from_folder(Config.BARCODE_DESCRIPTION_FOLDER)
    return _eb_description_map_cache


def normalize_name(name: str) -> str:
    name = os.path.splitext(os.path.basename(name))[0]
    name = name.upper()
    name = re.sub(r'[^A-Z0-9]+', '_', name)
    return name.strip('_')


def capitalize(name: str) -> str:
    name = re.sub(r'\s*\+\s*', ' + ', name)
    name = name.title()
    return name


def get_description(text, max_length=99):
    if text is None:
        return ""
    text = str(text).strip()
    if not text:
        return ""
    return text.split('.', 1)[0][:max_length]


def score_photo(filename, item_code, colour):
    name = normalize_name(filename)
    item = normalize_name(item_code)
    color = normalize_name(colour) if colour else ""

    if name.startswith(f"{item}_BLACK"):
        return 1000

    score = 0
    if color and name.startswith(f"{item}_{color}"):
        score += 300
    if item in name and color and color in name:
        score += 200
    if name == item:
        score += 100
    if name.startswith(item):
        score += 50

    return score


def find_best_photo(item_code, colour):
    photo_folder = Config.BARCODE_PHOTO_FOLDER
    best_photo = None
    best_score = -1

    for ext in ('*.JPG', '*.JPEG', '*.PNG'):
        for photo in glob.glob(os.path.join(photo_folder, ext)):
            s = score_photo(photo, item_code, colour)
            if s > best_score:
                best_score = s
                best_photo = photo

    return (best_photo, best_score) if best_score > 0 else (None, best_score)


def is_first3_non_digit(code: str) -> bool:
    """True jika 3 digit/karakter pertama item_code BUKAN angka."""
    if not code:
        return False
    return not code[:3].isdigit()


# ============================================================
# GROUPING
# ============================================================
def _group_rows(ws_input):
    """Mengelompokkan baris input jadi grup (po - item - store)."""
    grouped_data = {}
    for row in range(1, ws_input.nrows):
        po = str(ws_input.cell_value(row, 1)).replace('.0', '').strip()
        item = str(ws_input.cell_value(row, 8)).replace('.0', '').strip()
        store_raw = str(ws_input.cell_value(row, 15)).strip().upper()
        store = "JKT" if store_raw == "CWH" else "MKS"
        key = f"{po} - {item} - {store}"
        grouped_data.setdefault(key, []).append(row)
    return grouped_data


def count_groups(input_path):
    """
    Dipanggil job_manager SAAT FILE DI-UPLOAD, cuma untuk tahu berapa
    total grup (buat progress bar) - cepat, tidak menyentuh DB/foto.
    """
    _, ws_input = _open_sheet(input_path)
    return len(_group_rows(ws_input))


# ============================================================
# PROSES 1 GRUP (persis logic dari script asli, dipindah ke fungsi)
# ============================================================
def _process_group(conn, wb_input, ws_input, group_key, rows, material_map, output_dir):
    wb_output = xlwt.Workbook()
    ws_output = wb_output.add_sheet('tblbarcode')

    headers = [
        'itemno', 'grpdesc', 'uitemno', 'remark6', 'remark8', 'artcol',
        'price', 'price3', 'batchno', 'picture', 'qty', 'LINK', 'Material',
        'Length', 'Width', 'Height', 'uitemname'
    ]
    for c, h in enumerate(headers):
        ws_output.write(0, c, h)

    po, item, store = [x.strip() for x in group_key.split(' - ')]
    first_row = rows[0]
    first_barcode = str(int(ws_input.cell_value(first_row, 5)))
    first_item_code = str(ws_input.cell_value(first_row, 8)).replace('.0', '')
    first_description = get_article_info(conn, first_item_code, first_barcode).get("description", "").strip()
    new_name = f"{po} - {item} - {first_description} - {store}"

    output_subfolder = os.path.join(output_dir, new_name)
    os.makedirs(output_subfolder, exist_ok=True)

    output_row = 1
    photo_done = False
    link = ""

    for row in rows:
        barcode = str(int(ws_input.cell_value(row, 5)))
        desc = str(ws_input.cell_value(row, 6))
        item_code = str(ws_input.cell_value(row, 8)).replace('.0', '')

        is_eb_desc = "EB" in desc.upper()
        item_code_first3_non_digit = is_first3_non_digit(item_code)

        article_info = get_article_info(conn, item_code, barcode)
        upper_style = article_info["upper_style"]
        dimension = article_info["dimension"]
        description = article_info["description"]
        length, width, height = parse_dimension(dimension, desc)
        uitemname = description if desc.upper() == 'VB' else f"{description} {upper_style}".strip()
        uitemname = capitalize(uitemname)

        if is_eb_desc and item_code_first3_non_digit:
            uitemname = f"000EB000 {upper_style}".strip()
        elif not is_eb_desc and item_code_first3_non_digit:
            uitemname = item_code

        colour = str(ws_input.cell_value(row, 19)).replace('_', ' ')
        colour = capitalize(colour)
        size = ws_input.cell_value(row, 20)
        qty = int(float(ws_input.cell_value(row, 11)))
        delivery_date = ws_input.cell_value(row, 12)
        batchno = format_batchno(delivery_date, wb_input)

        names_desc = description if item_code and item_code[0].isdigit() else item_code

        if not photo_done:
            photo, score = find_best_photo(item_code, colour)
            if photo:
                print(f"[BARCODE] SELECTED PHOTO [{score}] :: {photo}", flush=True)
                ext = os.path.splitext(photo)[1]
                if ext != '.png':
                    print(f"[BARCODE][WARNING] Format foto bukan PNG untuk item_code: {item_code} "
                          f"(format ditemukan: {ext}) -> {os.path.basename(photo)}", flush=True)
                shutil.copy(photo, os.path.join(output_subfolder, f"{item_code}{ext}"))
                photo_done = True
            else:
                print(f"[BARCODE][WARNING] Foto tidak ditemukan untuk {item_code}", flush=True)

        if "EB" in desc.upper():
            link = f"https://everbestgroup.com/products/{names_desc.lower()}"
        elif "TR" in desc.upper():
            link = f"https://tracceshoes.com/products/{names_desc.lower()}"
        else:
            link = f"https://evbshoes.com/products/{names_desc.lower()}"

        artcol = f"{item_code} {colour}".strip()

        material_info = find_material_info(item_code, colour, material_map)
        material = material_info.get("material", "")
        material = re.sub(r'\s*\+\s*', ' + ', material)
        material = material.replace(' + ', ', ')
        material = material.title()

        if desc.upper() == 'VB':
            grpdesc = upper_style
        else:
            grpdesc = material_info["description"]
            if not grpdesc:
                print(f"[BARCODE][WARNING] EB description kosong untuk item_code: {item_code} "
                      f"(tidak ditemukan di file description)", flush=True)

        uitemno_value = item_code
        if is_eb_desc and item_code_first3_non_digit:
            uitemno_value = "000EB000"

        if desc.upper() == 'VB':
            colour_out = colour.upper()
            artcol_out = artcol.upper()
            material_out = material.upper()
            uitemname_out = uitemname.upper()
        else:
            colour_out = colour
            artcol_out = artcol
            material_out = material
            uitemname_out = uitemname

        for _ in range(qty):
            data = [
                barcode, grpdesc, uitemno_value, colour_out, size, artcol_out,
                '', '', batchno, f"{item_code}.PNG", qty, link, material_out,
                length, width, height, uitemname_out
            ]
            for c, v in enumerate(data):
                ws_output.write(output_row, c, v)
            output_row += 1

    wb_output.save(os.path.join(output_subfolder, f"{new_name}.xls"))
    return output_subfolder


def process_barcode_excel(input_path, output_dir, on_group_done=None):
    """
    Proses 1 file excel upload jadi banyak folder barcode (1 folder per grup).

    on_group_done(result_dict) dipanggil setiap 1 grup selesai diproses,
    supaya job_manager bisa update progress secara real-time.

    Return: list of dict, masing-masing:
        {"group": str, "ok": bool, "error": str|None, "folder": str|None}
    """
    material_map = _load_material_map()
    _load_eb_description_map()  # dimuat & di-cache; tersedia untuk dipakai kalau nanti diperlukan

    os.makedirs(output_dir, exist_ok=True)

    conn = get_db_connection()
    try:
        wb_input, ws_input = _open_sheet(input_path)
        grouped_data = _group_rows(ws_input)

        results = []
        for group_key, rows in grouped_data.items():
            print(f"\n[BARCODE][PROCESS] {group_key}", flush=True)
            try:
                output_subfolder = _process_group(
                    conn, wb_input, ws_input, group_key, rows, material_map, output_dir
                )
                result = {"group": group_key, "ok": True, "error": None, "folder": output_subfolder}
            except Exception as e:
                result = {"group": group_key, "ok": False, "error": str(e), "folder": None}

            results.append(result)
            if on_group_done:
                on_group_done(result)

        return results
    finally:
        conn.close()