import hashlib
import os
import re
import shutil
import tempfile
import threading
import time
import warnings

import pandas as pd
import pyodbc
import requests
import xlwings as xw
from openpyxl.cell.cell import MergedCell

from config import Config


template_mapping = {
    "HBL": "PO EB Bags.xlsx",
    "HBM": "PO EB Bags.xlsx",
    "EBSL": "PO EB Bags.xlsx",
    "EBSM": "PO EB Bags.xlsx",
    "EBL": "PO EB Ladies Shoes.xlsx",
    "EBM": "PO EB Mens Shoes.xlsx",
    "TRH": "PO TH TEMPLATE.xlsx",
    "TRK": "PO TK TEMPLATE.xlsx",
    "TRL": "PO TL_TM TEMPLATE.xlsx",
    "TRM": "PO TL_TM TEMPLATE.xlsx",
    "VBH": "PO VH TEMPLATE.xlsx",
    "VBL": "PO VL TEMPLATE.xlsx",
    "EBL_RO" : "RO Template Shoes.xlsx",
    "EBM_RO" : "RO Template Shoes MEN.xlsx",
    "VBL_RO" : "TEMPLATE RO VBL.xlsx",
    "EBSLS": "PO Belts.xlsx",
    "EBSMS": "PO Belts.xlsx",
}

JENIS_COMBO_TO_CODE = {
    # EB - Ladies (LDS)
    "EB_LDS_SHS": "EBL",
    "EB_LDS_ACC": "EBSL",
    "EB_LDS_BAG": "HBL",
    "EB_LDS_BGC": "EBSL",
    "EB_LDS_FWC": "EBSL",
    "EB_LDS_SLG": "EBSL",
    "EB_LDS_WAL": "EBSL",
    "EB_LDS_WRS": "EBSL",
    "EB_LDS_SHC": "EBSL",  

    # EB - Men
    "EB_MEN_SHS": "EBM",
    "EB_MEN_ACC": "EBSM",
    "EB_MEN_BAG": "HBM",
    "EB_MEN_BLT": "EBSM",
    "EB_MEN_SCK": "EBSM",
    "EB_MEN_SLG": "EBSM",
    "EB_MEN_WAL": "EBSM",
    "EB_MEN_SHC": "EBSM",
    "EB_MEN_WRS": "EBSM",  

    # TR
    "TR_CLD_SHS": "TRK",
    "TR_LDS_BAG": "TRH",
    "TR_LDS_SHS": "TRL",
    "TR_MEN_BAG": "TRH",
    "TR_MEN_SHS": "TRM",
    "TR_CLD_BAG": "TRH",   

    # VB
    "VB_LDS_BAG": "VBH",
    "VB_LDS_SHS": "VBL",

}


def parse_ref(ref_raw):
    """
    Parsing format REF: BRAND_GENDER_KATEGORI_PRODUK_RONA_LOKASI
    contoh: EB_LDS_SHS_GABIN1_NA_JKT
    -> jenis_combo="EB_LDS_SHS", produk="GABIN1", ro_na="NA", lokasi="JKT"
    """
    parts = str(ref_raw).strip().split("_")
    if len(parts) < 6:
        raise ValueError(
            f"Format REF tidak lengkap (butuh 6 segmen dipisah '_', dapat {len(parts)}): '{ref_raw}'"
        )
    jenis_combo = "_".join(parts[0:3]).strip().upper()
    produk = parts[3].strip()
    ro_na = parts[4].strip().upper()
    lokasi = parts[5].strip().upper()
    return jenis_combo, produk, ro_na, lokasi


def resolve_jenis(jenis_combo):
    """Translate kombinasi (misal EB_LDS_SHS) -> kode template (misal EBL)."""
    jenis = JENIS_COMBO_TO_CODE.get(jenis_combo, "UNKNOWN")
    if jenis == "UNKNOWN":
        raise ValueError(f"Kombinasi jenis '{jenis_combo}' tidak dikenal")
    template_name = template_mapping.get(jenis, "UNKNOWN")
    if template_name == "UNKNOWN":
        raise ValueError(f"Jenis '{jenis}' tidak dikenal di template_mapping")
    return jenis


def get_db_connection():

    conn = pyodbc.connect(
        "DRIVER={ODBC Driver 18 for SQL Server};"
        "SERVER=10.1.2.11\\RPTSVR;"
        "DATABASE=CENTRIC;"
        "UID=sa;"
        "PWD=Eshoes09;" 
        "TrustServerCertificate=yes;"
    )
    return conn

def get_grouped_po_data(location):
    conn = get_db_connection()
    query = """
        WITH RankedPO AS (
            SELECT 
                OptionID,
                SizePackID,
                SizePackName,
                DCID,
                NumofSizePack,
                Timestamp,
                DENSE_RANK() OVER (
                    PARTITION BY OptionID
                    ORDER BY CAST(Timestamp AS DATE) DESC
                ) AS DateRank
            FROM [Centric].[dbo].[PO]
            WHERE DCID = ?
              AND Timestamp IS NOT NULL
        )
        SELECT 
            OptionID,
            SizePackID,
            SizePackName,
            DCID,
            NumofSizePack,
            MAX(Timestamp) AS LatestDate
        FROM RankedPO
        WHERE DateRank = 1
        GROUP BY OptionID, SizePackID, SizePackName, DCID, NumofSizePack
    """
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        df = pd.read_sql(query, conn, params=[location])
    conn.close()
    return df

def format_option_id(article_name, color_name):
    name = article_name.strip().replace(" ", "_").upper()
    # color = color_name.strip().replace(" ", "_").upper()
    return f"{name}_{color_name}"

def extract_size_from_name(size_pack_name):
    # Contoh: "EBL EBM - 12C" => "12C"
    return size_pack_name.split('-')[-1].strip()

def extract_tracee_size_from_name(size_pack_name):
    # Contoh: "EBL EBM - 12C" => "12"
    part = size_pack_name.split('-')[-1].strip()
    match = re.match(r'(\d+)', part)
    if match:
        return match.group(1)
    return part

def format_price(value):
    try:
        return f"$ {float(value):,.2f}"
    except (ValueError, TypeError):
        return "$ 0.00"

def normalize(text: str) -> str:
    return re.sub(r'[^A-Z0-9]', '', text.upper())


_photo_index_cache = None

def build_photo_index():
    global _photo_index_cache
    if _photo_index_cache is not None:
        return _photo_index_cache

    index = {}
    for root, dirs, files in os.walk(Config.IMAGE_FOLDER):
        for f in files:
            if not f.lower().endswith(('.jpg', '.jpeg', '.png')):
                continue
            file_key = os.path.splitext(f)[0]


            if "_" not in file_key:
                continue
            raw_article, raw_color = file_key.split("_", 1)
            key = (normalize(raw_article), normalize(raw_color))
            if key not in index:
                index[key] = os.path.join(root, f)

    _photo_index_cache = index
    print(f"[IMG] Index foto dibangun: {len(index)} foto ditemukan di '{Config.IMAGE_FOLDER}' (termasuk subfolder)")
    return index

def _is_probable_article_code(value):
    """Return True for product/style codes commonly used in SAMPLE PHOTO filenames.

    Example: 271VB003, 263EB435, 271TR012.
    We intentionally require both digits and letters so normal names such as
    AVENIA are not treated as codes.
    """
    if value is None:
        return False
    text = str(value).strip().upper()
    if not text or len(text) < 5 or len(text) > 30:
        return False
    compact = normalize(text)
    return (
        len(compact) >= 5
        and any(ch.isdigit() for ch in compact)
        and any(ch.isalpha() for ch in compact)
        and bool(re.match(r"^\d+[A-Z]+\d+[A-Z0-9]*$", compact))
    )


def _collect_article_codes(value, result=None):
    """Recursively collect likely article/style codes from API response."""
    if result is None:
        result = []

    if isinstance(value, dict):
        # Prefer fields that are semantically likely to contain article/style codes.
        preferred_keys = (
            "style_code", "styleCode", "article_code", "articleCode",
            "option_id", "optionId", "optionID", "article_id", "articleId",
            "article_name_vl", "article_name_vh", "article_name",
        )
        for key in preferred_keys:
            if key in value:
                item = value.get(key)
                if _is_probable_article_code(item):
                    item = str(item).strip()
                    if item not in result:
                        result.append(item)

        for item in value.values():
            _collect_article_codes(item, result)

    elif isinstance(value, (list, tuple)):
        for item in value:
            _collect_article_codes(item, result)

    elif isinstance(value, str) and _is_probable_article_code(value):
        item = value.strip()
        if item not in result:
            result.append(item)

    return result


def resolve_article_candidates(api_data):
    """Build all useful identities for locating SAMPLE PHOTO files.

    The API may return the commercial article name (e.g. ``AVENIA``), while
    SAMPLE PHOTO filenames may use the internal article/style code
    (e.g. ``271VB003_CHALKBEIGE.png``).  Therefore photo lookup must NOT rely
    only on ``article.name``.

    Candidate order:
      1. article.name
      2. explicit style/article/option code fields
      3. article_name_vl / article_name_vh
      4. other code-like values found recursively in the API response
    """
    article = api_data.get("article", {}) or {}
    candidates = []

    def add(value):
        if value is None:
            return
        text = str(value).strip()
        if text and text not in candidates:
            candidates.append(text)

    # Keep the human-readable article name first for backward compatibility.
    add(article.get("name", ""))

    # Explicit code fields.
    for key in (
        "style_code", "styleCode", "article_code", "articleCode",
        "option_id", "optionId", "optionID", "article_id", "articleId",
        "article_name_vl", "article_name_vh", "article_name",
    ):
        add(api_data.get(key))
        add(article.get(key))

    # Finally inspect the whole response. This catches cases where the API
    # nests the code deeper, while AVENIA remains only the display name.
    for code in _collect_article_codes(api_data):
        add(code)

    return candidates


def insert_images(ws, colors, name_candidates, positions, color_label_row=37):
    """
    Insert sample photos and write the corresponding color label.

    `positions` are Excel anchors such as A33/E33/N33/W33.
    `color_label_row` lets each template control where the color label belongs.
    """
    if isinstance(name_candidates, str):
        name_candidates = [name_candidates]

    article_keys = [normalize(n) for n in name_candidates if n]
    valid_colors = [
        c for c in colors
        if str(c.get("color", "")).strip() not in ("", "-")
    ]

    photo_index = build_photo_index()

    for idx, color in enumerate(valid_colors):
        if idx >= len(positions):
            break

        raw_color = str(color.get("color", "")).strip()
        color_key = normalize(raw_color)

        print(f"[IMG] Article candidates: {article_keys} | Color: {color_key}")

        photo_file = None
        for article_key in article_keys:
            photo_file = photo_index.get((article_key, color_key))
            if photo_file:
                break

        position = positions[idx]

        m = re.match(r'^([A-Za-z]+)', position.strip())
        col = m.group(1).upper() if m else None
        color_cell = f"{col}{color_label_row}" if col else position

        try:
            safe_set_cell(ws.range(color_cell), raw_color)
        except Exception as e:
            print(f"[IMG] Error write color '{raw_color}' → {color_cell}: {e}")

        if photo_file:
            try:
                cell = ws.range(position)
                ws.pictures.add(
                    photo_file,
                    top=cell.top,
                    left=cell.left,
                    height=180,
                    width=235
                )
                print(f"[IMG] OK → {os.path.basename(photo_file)} @ {position}")
            except Exception as e:
                print(f"[IMG] Error insert image '{raw_color}': {e}")
        else:
            print(
                f"[IMG] NOT FOUND → Article candidates={article_keys}, "
                f"Color={color_key}"
            )


def detect_vbl_layout(ws):
    """
    Detect the structure of the current PO VL TEMPLATE.xlsx instead of relying
    on the old hard-coded layout.

    Current template facts:
      - Confirmed PO sheet
      - size headers: T22:Z22
      - quantity detail rows: 23..27
      - four picture blocks: A33:D42, E33:M42, N33:V42, W33:AC42
      - color labels: A43, E43, N43, W43
    """
    # Size columns are identified from the actual row-22 values.
    size_columns = {}
    for col in range(1, ws.used_range.last_cell.column + 1):
        try:
            value = ws.range((22, col)).value
        except Exception:
            continue
        size_key = normalize_size(value)
        if size_key and size_key.isdigit():
            n = int(size_key)
            if 30 <= n <= 60:
                size_columns[n] = ws.range((22, col)).get_address().split("$")[1]

    # Detect large merged picture blocks below the data table.
    picture_areas = []
    try:
        for area in ws.merged_cells.areas:
            first = area[0, 0]
            last = area[area.rows.count - 1, area.columns.count - 1]
            min_row, min_col = first.row, first.column
            max_row, max_col = last.row, last.column
            if (
                min_row >= 30
                and max_row <= 44
                and (max_row - min_row) >= 5
                and (max_col - min_col) >= 2
            ):
                picture_areas.append(
                    (min_col, min_row, first.address)
                )
    except Exception:
        picture_areas = []

    picture_positions = [
        addr for _, _, addr in sorted(picture_areas)
    ]

    # Fallback to the known current template layout if Excel's merged-area
    # enumeration is unavailable.
    if len(picture_positions) < 4:
        picture_positions = ["A33", "E33", "N33", "W33"]

    return {
        "size_columns": size_columns,
        "picture_positions": picture_positions[:4],
        "color_label_row": 43,
        "detail_start_row": 23,
        "max_colors": 4,
    }


def safe_set_cell(cell, value):
    try:
        if not isinstance(cell, MergedCell):
            cell.value = value
    except Exception as e:
        print(f"Error setting cell {cell.address if hasattr(cell, 'address') else cell}: {e}")



def get_po_data_from_api(gl_souche, gl_numero):
    location = "CWH" if gl_souche == "JKT" else "WMKR"
    url = f"{Config.PO_API_BASE_URL}/vh_vl?gl_souche={location}&gl_numero={gl_numero}"
    print(f"[API] GET {url}", flush=True)
    t0 = time.time()
    try:
        # timeout=(connect, read) -> tidak akan menggantung selamanya
        response = requests.get(url, timeout=(5, 30))
        print(f"[API] status={response.status_code} dalam {time.time() - t0:.1f}s", flush=True)
        if response.status_code == 200:
            return response.json()
        print(f"[ERROR] Gagal panggil API untuk PO {gl_numero} | Status code: {response.status_code}", flush=True)
    except requests.exceptions.Timeout:
        print(f"[ERROR] API TIMEOUT untuk PO {gl_numero} setelah {time.time() - t0:.1f}s | URL={url}", flush=True)
    except requests.exceptions.RequestException as e:
        print(f"[ERROR] Request exception untuk PO {gl_numero}: {e}", flush=True)

    return None

def read_header_value(ws, row, col):
    cell = ws.range((row, col))
    if cell.merge_cells:
        return cell.merge_area[0, 0].value
    return cell.value

def normalize_size(val):
    if val is None:
        return None
    try:
        return str(int(float(val))).strip()
    except:
        return str(val).strip()

def build_size_qty_map(color_obj):
    result = {}
    for s in color_obj.get("sets", []):
        size = normalize_size(s.get("size"))
        qty = s.get("qty")
        if size:
            result[size] = qty
    return result


# ==========================================================================
# TEMPLATE BARU: PO TH (TRH) & PO VH (VBH)  -- layout kedua template IDENTIK
#
#   B2:C2   PO NO  (anchor B2)          B4:C4  DATE (anchor B4)
#   B5:C5   TO     (anchor B5, default template ": EDE")
#
#   Tabel item = 4 baris warna (11-14):
#       A11:A12 DESCRIPTION (merged) | B11 DEVELOP CODE (=> rumus A28)
#       C11 STYLE NAME (=> rumus A29) | D11:D14 COLOUR (=> rumus label foto row 38)
#       E11:E14 QTY | F11:F14 COST (angka) | G11:G14 AMOUNT = rumus F*E
#       A13 "Material :" , A14 "Lining :"  (teks kolom A, overflow ke B/C)
#       E15 / G15 = rumus SUM  -> JANGAN ditimpa
#   B17:C18 Delivery Date (anchor B17)
#   Foto: A30:B37, C30:D37, E30:F37, G30:H37
#   Label warna baris 38 = RUMUS (=D11..D14) -> JANGAN ditulis manual
# ==========================================================================
THVH_ITEM_ROWS = [11, 12, 13, 14]
THVH_PICTURE_BLOCKS = ["A30:B37", "C30:D37", "E30:F37", "G30:H37"]
_EMPTY_TOKENS = {"", "-", "none", "null", "nan"}


def _clean(value):
    """String rapi, atau '' kalau kosong / '-' / None."""
    if value is None:
        return ""
    text = str(value).strip()
    return "" if text.lower() in _EMPTY_TOKENS else text


def _to_float(value):
    """'US$ 1,234.50' / '$ 12.5' / 12 / None -> float (default 0.0)."""
    if value is None:
        return 0.0
    if isinstance(value, (int, float)):
        return float(value)
    text = re.sub(r"[^0-9.\-]", "", str(value).replace(",", ""))
    try:
        return float(text) if text not in ("", "-", ".") else 0.0
    except ValueError:
        return 0.0


def _to_int(value):
    return int(round(_to_float(value)))


def _colon(value):
    """Pastikan teks diawali ': ' (konsisten dengan template)."""
    text = _clean(value)
    if not text:
        return ":"
    return text if text.startswith(":") else f": {text}"


def _join_materials(materials, keys):
    out = []
    for m in materials or []:
        label = _clean(m.get("label")).lower()
        value = _clean(m.get("value"))
        if value and any(k in label for k in keys) and value not in out:
            out.append(value)
    return ", ".join(out)


def build_color_rows(colors):
    """Normalisasi warna API -> list dict(color, qty, price). Warna kosong/'-' dibuang."""
    rows = []
    for c in colors or []:
        name = _clean(c.get("color"))
        if not name:
            continue
        qty = 0
        for s in c.get("sets", []) or []:
            qty += _to_int(s.get("qty"))
        rows.append({"color": name, "qty": qty, "price": _to_float(c.get("price"))})
    return rows


def resolve_codes(api_data):
    """
    Return (develop_code, style_name).
      - nama artikel diawali angka (mis. 271TR012) = kode internal
            -> DEVELOP CODE = kode itu, STYLE NAME = style_code
      - nama artikel bukan kode (mis. AVENIA)
            -> STYLE NAME = nama itu, DEVELOP CODE = develop/supplier code
    """
    article = api_data.get("article", {}) or {}
    name = _clean(article.get("name"))
    style_code = _clean(api_data.get("style_code"))
    explicit_dev = _clean(api_data.get("develop_code") or article.get("develop_code"))
    supplier = _clean(api_data.get("supplier_code"))

    if name[:1].isdigit():
        develop = explicit_dev or name
        style = style_code or _clean(api_data.get("article_name_vl")) or name
    else:
        style = name or style_code
        develop = explicit_dev or supplier or (style_code if style_code != style else "")
    return develop, style


def find_photo(photo_index, article_candidates, color_name):
    color_key = normalize(str(color_name or ""))
    for cand in article_candidates or []:
        path = photo_index.get((normalize(str(cand)), color_key))
        if path:
            return path
    return None


def _put(ws, addr, value):
    """Tulis ke 1 sel; error tidak menghentikan proses."""
    try:
        ws.range(addr).value = value
        return True
    except Exception as e:
        print(f"[TH/VH] gagal tulis {addr}: {e}", flush=True)
        return False


def _clear(ws, addr):
    try:
        ws.range(addr).clear_contents()
    except Exception as e:
        print(f"[TH/VH] gagal clear {addr}: {e}", flush=True)


def _prepare_local_photo(src, max_px=1200, timeout=20):
    """
    Salin foto dari share jaringan ke folder temp lokal (dengan batas waktu),
    lalu perkecil kalau terlalu besar. Return path lokal, atau None kalau gagal/timeout.
    Excel jauh lebih cepat & stabil memasukkan gambar dari disk lokal.
    """
    cache_dir = os.path.join(tempfile.gettempdir(), "po_photo_cache")
    os.makedirs(cache_dir, exist_ok=True)
    key = hashlib.md5(src.encode("utf-8", "ignore")).hexdigest()[:12]
    ext = os.path.splitext(src)[1].lower() or ".png"
    raw = os.path.join(cache_dir, f"{key}_raw{ext}")
    small = os.path.join(cache_dir, f"{key}_small{ext}")

    if os.path.exists(small):  # sudah pernah diproses
        return small

    t0 = time.time()
    state = {}

    def _copy():
        try:
            part = raw + ".part"
            shutil.copyfile(src, part)
            os.replace(part, raw)
            state["ok"] = True
        except Exception as e:
            state["err"] = e

    th = threading.Thread(target=_copy, daemon=True)
    th.start()
    th.join(timeout)
    if th.is_alive():
        print(f"[IMG] TIMEOUT salin foto >{timeout}s: {src}", flush=True)
        return None
    if not state.get("ok"):
        print(f"[IMG] gagal salin foto: {state.get('err')} | {src}", flush=True)
        return None
    print(f"[IMG] salin foto {time.time() - t0:.1f}s | {os.path.getsize(raw) / 1024:.0f} KB", flush=True)

    try:  # perkecil bila PIL tersedia
        from PIL import Image
        im = Image.open(raw)
        im.load()
        if max(im.size) > max_px:
            im.thumbnail((max_px, max_px))
        if ext in (".jpg", ".jpeg") and im.mode not in ("RGB", "L"):
            im = im.convert("RGB")
        im.save(small)
        im.close()
        return small
    except Exception as e:
        print(f"[IMG] resize dilewati ({e}); pakai file asli lokal", flush=True)
        return raw


def _insert_fit_picture(ws, path, block_addr, padding=4):
    """
    Taruh gambar di tengah blok merged, proporsional.
    Ukuran dihitung di Python (PIL) lalu dimasukkan dalam SATU panggilan Shapes.AddPicture
    dengan width/height eksplisit (sama seperti cara lama yang stabil), tanpa baca/ubah
    ukuran bolak-balik lewat COM.
    """
    print(f"[IMG] -> blok {block_addr}", flush=True)
    area = ws.range(block_addr)
    left, top, area_w, area_h = area.left, area.top, area.width, area.height
    max_w = max(area_w - 2 * padding, 10)
    max_h = max(area_h - 2 * padding, 10)

    img_w = img_h = None
    try:
        from PIL import Image
        with Image.open(path) as im:
            img_w, img_h = im.size
    except Exception as e:
        print(f"[IMG] ukuran gambar tidak terbaca ({e}); pakai rasio 4:3", flush=True)
    if not img_w or not img_h:
        img_w, img_h = 4, 3

    ratio = min(max_w / img_w, max_h / img_h)
    w, h = img_w * ratio, img_h * ratio
    x = left + (area_w - w) / 2
    y = top + (area_h - h) / 2

    t0 = time.time()
    # LinkToFile=False, SaveWithDocument=True, Left, Top, Width, Height
    ws.api.Shapes.AddPicture(path, False, True, x, y, w, h)
    print(f"[IMG] AddPicture selesai {time.time() - t0:.1f}s ({w:.0f}x{h:.0f}pt)", flush=True)
    return True


def _replace_placeholders(ws, final_type):
    try:  # native Excel, tanpa loop per sel
        ws.api.UsedRange.Replace(What="{_alejandro_}", Replacement=final_type, LookAt=2)
    except Exception as e:
        print(f"[TH/VH] replace placeholder sel gagal: {e}", flush=True)
    try:
        shapes = list(ws.shapes)
    except Exception:
        shapes = []
    for shape in shapes:
        try:
            text = getattr(shape, "text", None)
            if text and "{_alejandro_}" in text:
                shape.text = text.replace("{_alejandro_}", final_type)
        except Exception:
            continue


def _verify_th_vh_layout(ws, template_file=None):
    """
    Pastikan file template yang dibuka adalah TEMPLATE TH/VH TERBARU.
    Kalau masih versi lama -> berhenti dengan pesan jelas (bukan hasil salah diam-diam).
    """
    problems = []
    try:
        if _clean(ws.range("B8").value).upper() != "DEVELOP CODE":
            problems.append("B8 bukan 'DEVELOP CODE'")
        if _clean(ws.range("C8").value).upper() != "STYLE NAME":
            problems.append("C8 bukan 'STYLE NAME'")
        if not _clean(ws.range("A13").value).lower().startswith("material"):
            problems.append("A13 bukan 'Material :'")
        if not _clean(ws.range("A14").value).lower().startswith("lining"):
            problems.append("A14 bukan 'Lining :'")
        if str(ws.range("A38").formula).replace(" ", "").upper() != "=D11":
            problems.append("A38 bukan rumus =D11")
        if "$A$30:$B$37" not in str(ws.range("A30").merge_area.address):
            problems.append("blok foto A30:B37 tidak ditemukan")
    except Exception as e:
        problems.append(f"gagal membaca template: {e}")

    if template_file:
        try:
            ts = time.strftime("%Y-%m-%d %H:%M:%S", time.localtime(os.path.getmtime(template_file)))
            print(f"[TEMPLATE] {template_file} | terakhir diubah {ts}", flush=True)
        except Exception:
            pass

    if problems:
        raise ValueError(
            "Template yang dibuka BUKAN versi terbaru (" + "; ".join(problems) + "). "
            "Timpa file di folder po_template dengan template TH/VH terbaru."
        )


def fill_th_vh_template(ws, api_data, datas, prefix, template_file=None):
    """
    Isi template PO TH (prefix='TH') atau PO VH (prefix='VH').
    Error TIDAK ditelan di sini -> naik ke generate_po_excel.
    """
    po = datas["po_code"]
    loc_raw = datas["lokasi"]
    tag = f"[{prefix}]"

    _verify_th_vh_layout(ws, template_file)

    article = api_data.get("article", {}) or {}
    materials = api_data.get("materials", []) or []
    article_candidates = resolve_article_candidates(api_data)

    # ---------------- 1. HEADER ----------------
    _put(ws, "B2", f": {po}")

    try:
        ws.range("B4").number_format = "@"  # cegah Excel ubah jadi tanggal
    except Exception:
        pass
    _put(ws, "B4", _colon(api_data.get("date")))

    to_value = _clean(
        api_data.get("to") or api_data.get("customer") or api_data.get("customer_code")
    )
    if to_value:  # kosong -> biarkan default template ": EDE"
        _put(ws, "B5", f": {to_value}")

    # ---------------- 2. DESCRIPTION / DEVELOP / STYLE ----------------
    develop_code, style_name = resolve_codes(api_data)
    description = (
        _clean(api_data.get("description_name"))
        or _clean(api_data.get("article_name_vl"))
        or _clean(api_data.get("article_name_vh"))
        or style_name
    )
    _put(ws, "A11", description)
    _put(ws, "B11", develop_code)
    _put(ws, "C11", style_name)

    # ---------------- 3. WARNA / QTY / COST (baris 11-14) ----------------
    color_rows = build_color_rows(article.get("color", []))
    if len(color_rows) > len(THVH_ITEM_ROWS):
        print(
            f"{tag} WARNING PO {po}: {len(color_rows)} warna, template hanya "
            f"{len(THVH_ITEM_ROWS)} baris -> warna ke-{len(THVH_ITEM_ROWS) + 1} dst TIDAK tampil.",
            flush=True,
        )
    shown = color_rows[: len(THVH_ITEM_ROWS)]

    for row, item in zip(THVH_ITEM_ROWS, shown):
        _put(ws, f"D{row}", item["color"])
        _put(ws, f"E{row}", item["qty"])
        _put(ws, f"F{row}", item["price"])  # G = rumus F*E, tidak disentuh
        print(
            f"{tag} row {row} | {item['color']} | qty={item['qty']} | cost={item['price']}",
            flush=True,
        )

    for row in THVH_ITEM_ROWS[len(shown):]:  # baris tak terpakai
        _clear(ws, f"D{row}")
        _clear(ws, f"E{row}")
        _put(ws, f"F{row}", 0)

    # E15 & G15 = rumus SUM template -> tidak ditulis

    # ---------------- 4. MATERIAL / LINING (kolom A baris 13-14) ----------------
    upper = _join_materials(materials, ("upper", "material"))
    lining = _join_materials(materials, ("lining",))
    _put(ws, "A13", f"Material : {upper}" if upper else "Material :")
    _put(ws, "A14", f"Lining : {lining}" if lining else "Lining :")

    # ---------------- 5. DELIVERY DATE (anchor B17) ----------------
    _put(ws, "B17", api_data.get("delivery_date", ""))

    # ---------------- 6. PLACEHOLDER ----------------
    _replace_placeholders(ws, f"{prefix}-{loc_raw}")

    # ---------------- 7. FOTO (label warna baris 38 = rumus) ----------------
    print(f"{tag} mulai proses foto ({len(shown)} warna)", flush=True)
    photo_index = build_photo_index()
    for idx, item in enumerate(shown):
        photo = find_photo(photo_index, article_candidates, item["color"])
        if not photo:
            print(
                f"{tag} foto TIDAK ketemu | kandidat={article_candidates} | warna={item['color']}",
                flush=True,
            )
            continue
        local_photo = _prepare_local_photo(photo)
        if not local_photo:
            print(f"{tag} foto dilewati (tidak bisa disalin) | warna={item['color']}", flush=True)
            continue
        try:
            _insert_fit_picture(ws, local_photo, THVH_PICTURE_BLOCKS[idx])
            print(f"{tag} foto OK -> {os.path.basename(photo)} @ {THVH_PICTURE_BLOCKS[idx]}", flush=True)
        except Exception as e:
            print(f"{tag} gagal insert foto {item['color']}: {e}", flush=True)

    print(f"{tag} SUCCESS | PO={po}", flush=True)
    return True


def fill_template_with_res(po_code, jenis_template, api_data, output_path,datas):
    print(f"[INFO] Proses input {jenis_template} - {po_code}...")

    is_repeat_order = datas['ro_na'] == 'RO'

    JENIS_WITH_RO_TEMPLATE = ["EBL", "EBM", "VBL"]

    if is_repeat_order and jenis_template in JENIS_WITH_RO_TEMPLATE:
        template = jenis_template + '_RO'
        jenis_template = jenis_template + '_RO'
    else:
        template = jenis_template

    template_filename = template_mapping.get(template, "")
    # Backward-compatible fallback: if the new EBM RO template has not yet
    # been deployed, use the existing legacy file instead of failing the PO.
    if template == "EBM_RO" and template_filename == "RO Template Shoes MEN.xlsx":
        preferred = os.path.join(Config.EXCEL_TEMPLATE_FOLDER, template_filename)
        if not os.path.exists(preferred):
            legacy = "RO Template Shoes.xlsx"
            legacy_path = os.path.join(Config.EXCEL_TEMPLATE_FOLDER, legacy)
            if os.path.exists(legacy_path):
                template_filename = legacy
    if not template_filename:
        print(f"[ERROR] Kategori '{template}' belum punya nama file template "
              f"di template_mapping (masih kosong) - PO {po_code} di-skip.\n")
        return False

    template_file = os.path.join(Config.EXCEL_TEMPLATE_FOLDER, template_filename)
    if not os.path.exists(template_file):
        print(f"[ERROR] Template tidak ditemukan untuk jenis: {template_file}\n")
        return False

    app = None
    wb = None
    try:
        print(f"[XL] membuka Excel + template: {template_file}", flush=True)
        app = xw.App(visible=False)
        app.display_alerts = False
        app.screen_updating = False
        wb = app.books.open(template_file)
        ws = wb.sheets[0]
        ro_na_raw = datas['ro_na']
        loc_raw = datas['lokasi']
        produk = datas['produk']
        location = "JAKARTA" if loc_raw == "JKT" else "MAKASSAR"
        ro_na = "NEW ARRIVAL" if ro_na_raw == "NA" else "REPEAT ORDER"
        po = datas['po_code']

        #EB SHOES WOI
        if jenis_template in ["EBL", "EBM"] and ro_na != 'REPEAT ORDER':
            try:
                print("[INFO] Menulis format EB SHOES")
                safe_set_cell(ws.range('B2'), f": {po}")
                safe_set_cell(ws.range('B4'), api_data.get("date", ""))
                safe_set_cell(ws.range('E4'), ro_na)
                safe_set_cell(ws.range('E5'), location)
                safe_set_cell(ws.range('E6'), api_data.get("seasons", "")) 
                safe_set_cell(ws.range('B19'), api_data.get("delivery_date", ""))
                safe_set_cell(ws.range('A11'), api_data.get("description_name", ""))
                article = api_data.get("article", {})
                safe_set_cell(ws.range('C11'), api_data.get("style_code", "") if (article.get("name", "") or "")[:1].isdigit() else article.get("name", ""))
                safe_set_cell(ws.range('B11'), article.get("name", "") if (article.get("name", "") or "")[:1].isdigit() else '')
                materials = api_data.get("materials", [])
                material_row = 13
                for material in materials:
                    label = material.get("label", "").strip()
                    value = material.get("value", "").strip()
                    safe_set_cell(ws.range(f"A{material_row}"), f"{label} : {value}") 
                    material_row += 1
                    if material_row > 17:
                        break

                api_range = (api_data.get("range") or "").strip().upper()
                if "MEN" in api_range:
                    sizes = [39, 40, 41, 42, 43, 44, 45,46]
                    columns = ['G', 'H', 'I', 'J', 'K','L','M','N']
                    for col, size in zip(columns, sizes):
                        safe_set_cell(ws.range(f"{col}21"), size) 
                        
                size_columns = {}
                for col in ['G', 'H', 'I', 'J', 'K','L','M']:
                    size_value = ws.range(f"{col}21").value
                    if size_value:
                        size_columns[int(size_value)] = col

                colors = article.get("color", []) 
                desc_row_start = 11
                desc_counter = 0
                for color_info in colors:
                    color_name = color_info.get("color", "")
                    if color_name and color_name != "-":
                        safe_set_cell(ws.range(f"D{desc_row_start + desc_counter}"), color_name) 
                        desc_counter += 1

                for row in range(11, 18):
                    cell_color = ws.range(f"D{row}").value
                    if not cell_color:
                        continue
                    for color_info in colors:
                        color_name = color_info.get("color", "")
                        price = color_info.get("price", "")
                        if cell_color.strip().upper() == color_name.strip().upper():
                            try:
                                price_value = float(price)
                                formatted_price = f"$ {price_value:.2f}"
                            except ValueError:
                                formatted_price = "$ 0.00"
                            safe_set_cell(ws.range(f"AA{row}"), formatted_price)  
                            break

                start_row = 22
                for idx, color_info in enumerate(colors):
                    row = start_row + idx
                    color_name = color_info.get("color", "")
                    sets = color_info.get("sets", [])
                    # print(f'SET :: {sets}')
                    if color_name and color_name != "-":
                        for set_info in sets:
                            size = set_info.get("size")
                            qty = set_info.get("qty", 0)
                            try:
                                size = int(size)
                            except (ValueError, TypeError):
                                continue
                            col = size_columns.get(size)
                            # print(f'COL:ROW :: {col}:{row}')
                            if col:
                                safe_set_cell(ws.range(f"{col}{row}"), qty)  

                location_db = "CWH" if loc_raw == "JKT" else "WMKR"
                if ro_na_raw == "NA" or ro_na_raw == "RO":
                    df_sets = get_grouped_po_data(location_db)
                    size_columns = {}
                    
                    for col in range(5, 24):  # G=7, Y=25
                        cell_letter = ws.range((9, col)).value
                        if not cell_letter: 
                            cell_letter = ws.range((9, col)).merge_area[0, 0].value
                        if cell_letter:
                            col_letter = ws.range((10, col)).get_address().split('$')[1]
                            size_columns[cell_letter.strip().upper()] = col_letter

                    # Loop baris data warna
                    size_suffix_list = []
                    for row in range(11, 16):   
                        color_name = ws.range(f"D{row}").value
                        if color_name:
                            color_name = "_".join(str(color_name).split())
                        else:
                            continue

                        article_name = article.get("name", "")
                        if article_name and article_name[0].isdigit():
                            article_name = api_data.get("style_code", "")
                        # print(f'ARTICLE NAME :: {api_data.get("style_code")} - COLOR :: {color_name}')
                        option_id = format_option_id(article_name, color_name)
                        # print(f'Option ID :: {option_id}')
                        sets_for_option = df_sets[df_sets['OptionID'] == option_id]

                        if sets_for_option.empty:
                            print(f"[INFO] Tidak ada data Sets untuk OptionID: {option_id}") 
                            for size_code, col_letter in size_columns.items():
                                ws.range(f"{col_letter}{row}").value = "0"
                            continue
 
                        size_data = {}
                        for _, row_set in sets_for_option.iterrows():
                            size_pack_name = row_set['SizePackName']
                            num_of_pack = row_set['NumofSizePack']
                            size_code = extract_size_from_name(size_pack_name).upper() 

                            if len(size_code) >= 2:
                                suffix = size_code[-1]
                                if suffix not in size_suffix_list:
                                    size_suffix_list.append(suffix)

                            size_data[size_code] = num_of_pack
                
                        for size_code, col_letter in size_columns.items():
                            value = size_data.get(size_code, "0")
                            if not value or pd.isna(value):
                                value = "0"
                            ws.range(f"{col_letter}{row}").value = value

                type = "EM" if api_data.get("range", "") == "MEN" else "EL"
                final_type = f'{type}-{loc_raw}'
                final_sets_str = "/".join(sorted(size_suffix_list)[:5])
                for cell in ws.used_range:
                    if cell.value and isinstance(cell.value, str) and "{_sets_}" in cell.value:
                        cell.value = cell.value.replace("{_sets_}", final_sets_str)
                    if cell.value and isinstance(cell.value, str) and "{_alejandro_}" in cell.value:
                        cell.value = cell.value.replace("{_alejandro_}", final_type)

                for shape in ws.shapes:
                    try:
                        if hasattr(shape, "text") and shape.text and "{_sets_}" in shape.text:
                            shape.text = shape.text.replace("{_sets_}", final_sets_str)
                        if hasattr(shape, "text") and shape.text and "{_alejandro_}" in shape.text:
                            shape.text = shape.text.replace("{_alejandro_}", final_type)
                    except:
                        continue   
                                
                insert_images(ws, colors, resolve_article_candidates(api_data), ['A32', 'C32', 'E32', 'G32', 'N32', 'T32'])

            except Exception as e:
                print(f"Error writing EBL/EBM format: {e}")

        elif jenis_template in ["EBSLS", "EBSMS"]:
            try:
                print("[INFO] Menulis format EB BELTS")
                safe_set_cell(ws.range('B2'), f": {po}")
                safe_set_cell(ws.range('E4'), ro_na + ' - ' + location)
                safe_set_cell(ws.range('E5'), api_data.get("seasons", ""))
                safe_set_cell(ws.range('B18'), api_data.get("delivery_date", ""))
                safe_set_cell(ws.range('B4'), api_data.get("date", ""))
                safe_set_cell(ws.range('A10'), api_data.get("description_name", ""))

                article = api_data.get("article", {})
                safe_set_cell(ws.range('C10'), api_data.get("style_code", "") if (article.get("name", "") or "")[:1].isdigit() else article.get("name", ""))
                safe_set_cell(ws.range('B10'), article.get("name", "") if (article.get("name", "") or "")[:1].isdigit() else '')

                materials = api_data.get("materials", [])
                material_row = 12
                for material in materials:
                    label = material.get("label", "").strip()
                    value = material.get("value", "").strip()
                    safe_set_cell(ws.range(f"A{material_row}"), f"{label} : {value}") 
                    material_row += 1
                    if material_row > 15:
                        break

                api_range = (api_data.get("range") or "").strip().upper()
                if "MEN" in api_range:
                    sizes = [44, 46, 48, 50]
                    columns = ['G', 'H', 'I', 'J']
                    for col, size in zip(columns, sizes):
                        safe_set_cell(ws.range(f"{col}7"), size) 
                        
                size_columns = {}
                for col in ['G', 'H', 'I', 'J']:
                    size_value = ws.range(f"{col}7").value
                    if size_value:
                        size_columns[int(size_value)] = col

                colors = article.get("color", []) 
                desc_row_start = 10
                desc_counter = 0
                for color_info in colors:
                    color_name = color_info.get("color", "")
                    if color_name and color_name != "-":
                        safe_set_cell(ws.range(f"D{desc_row_start + desc_counter}"), color_name) 
                        desc_counter += 1

                for row in range(10, 14):
                    cell_color = ws.range(f"E{row}").value
                    if not cell_color:
                        continue
                    for color_info in colors:
                        color_name = color_info.get("color", "")
                        price = color_info.get("price", "")
                        if cell_color.strip().upper() == color_name.strip().upper():
                            try:
                                price_value = float(price)
                                formatted_price = f"$ {price_value:.2f}"
                            except ValueError:
                                formatted_price = "$ 0.00"
                            safe_set_cell(ws.range(f"L{row}"), formatted_price)  
                            break

                for row in range(10, 14):
                    cell_color = ws.range(f"D{row}").value
                    if not cell_color:
                        continue

                    for color_info in colors:
                        color_name = color_info.get("color", "")

                        if cell_color.strip().upper() == color_name.strip().upper():

                            sets = color_info.get("sets", [])

                            for s in sets:
                                size = s.get("size")
                                qty = s.get("qty")

                                if size in size_columns:
                                    col = size_columns[size]
                                    safe_set_cell(ws.range(f"{col}{row}"), qty)

                            break

                mark = 'ES'
                final_type = f'{mark}-{loc_raw}'
                for cell in ws.used_range:
                    if cell.value and isinstance(cell.value, str) and "{_alejandro_}" in cell.value:
                        cell.value = cell.value.replace("{_alejandro_}", final_type)

                for shape in ws.shapes:
                    try:
                        if hasattr(shape, "text") and shape.text and "{_alejandro_}" in shape.text:
                            shape.text = shape.text.replace("{_alejandro_}", final_type)
                    except:
                        continue   
                                
                insert_images(ws, colors, resolve_article_candidates(api_data), ['A30', 'C30', 'E30', 'H30', 'O30'])

            except Exception as e:
                print(f"Error writing EBL/EBM format: {e}")
    
        elif jenis_template in ['EBL_RO'] and ro_na == 'REPEAT ORDER':
            try:
                print("[INFO] Menulis format template RO EB SHOES")
                
                safe_set_cell(ws.range('B2'), f": {po}")
                safe_set_cell(ws.range('B4'), api_data.get("date", ""))
                safe_set_cell(ws.range('E5'), location)
                safe_set_cell(ws.range('E6'), api_data.get("seasons", ""))

                safe_set_cell(ws.range('B19'), api_data.get("delivery_date", ""))

                # range_val = (api_data.get("range", "") or "").strip().upper()
                # if range_val == "MEN":
                #     safe_set_cell(ws.range('A54'), "2. MEN SHOES BOX USING WHITE EYELET")
                # elif range_val == "LADIES":
                #     safe_set_cell(ws.range('A54'), "2. LADIES SHOES BOX USING BLACK EYELET")
                safe_set_cell(ws.range('A11'), api_data.get("description_name", ""))

                article = api_data.get("article", {})
                safe_set_cell(ws.range('C11'), api_data.get("style_code", "") if (article.get("name", "") or "")[:1].isdigit() else article.get("name", ""))
                safe_set_cell(ws.range('B11'), article.get("name", "") if (article.get("name", "") or "")[:1].isdigit() else '')

                materials = api_data.get("materials", [])
                material_row = 13
                for material in materials:
                    label = material.get("label", "").strip()
                    value = material.get("value", "").strip()
                    if label and value:
                        safe_set_cell(ws.range(f"A{material_row}"), f"{label} : {value}")
                        material_row += 1
                    if material_row > 17:
                        break

                api_range = (api_data.get("range") or "").strip().upper()
                # if "MEN" in api_range:
                #     sizes = [39, 40, 41, 42, 43, 44, 45]
                #     for col, size in zip(['G', 'H', 'I', 'J', 'K','L','M'], sizes):
                #         safe_set_cell(ws.range(f"{col}20"), size)

                size_columns_bottom = {}
                for col in ['G', 'H', 'I', 'J', 'K','L']:
                    val = ws.range(f"{col}20").value
                    if val:
                        size_columns_bottom[int(val)] = col

                size_columns_top = {}
                api_range = (api_data.get("range") or "").strip().upper()

                if api_range == 'MEN':
                    sizes = [39, 40, 41, 42, 43, 44, 45]
                    for col, size in zip(['E','F','G', 'H', 'I', 'J', 'K'], sizes):
                        safe_set_cell(ws.range(f"{col}9"), size)
                        size_columns_top[size] = col
                    
                else:
                    # fallback (non RO / non MEN)
                    for col in ['E','F','G', 'H', 'I', 'J', 'K']:
                        val = ws.range(f"{col}9").value
                        if val:
                            try:
                                size_columns_top[int(val)] = col
                            except:
                                pass

                colors = article.get("color", [])
                desc_row = 11

                for color in colors:
                    color_name = color.get("color", "")
                    if color_name and color_name != "-":
                        safe_set_cell(ws.range(f"D{desc_row}"), color_name)

                        try:
                            price = float(color.get("price", 0))
                            safe_set_cell(ws.range(f"N{desc_row}"), f"$ {price:.2f}")
                        except:
                            safe_set_cell(ws.range(f"N{desc_row}"), "$ 0.00")

                        desc_row += 1
                        if desc_row > 17:
                            break

                start_row_top = 11      # G11–M
                start_row_bottom = 21   # J22–P

                for idx, color in enumerate(colors):
                    row_top = start_row_top + idx
                    row_bottom = start_row_bottom + idx

                    sets = color.get("sets", [])

                    for set_info in sets:
                        try:
                            size = int(set_info.get("size"))
                            qty = set_info.get("qty")
                        except (ValueError, TypeError):
                            continue

                        col_top = size_columns_top.get(size)
                        if col_top:
                            safe_set_cell(ws.range(f"{col_top}{row_top}"), qty)

                        col_bottom = size_columns_bottom.get(size)
                        if col_bottom:
                            safe_set_cell(ws.range(f"{col_bottom}{row_bottom}"), qty)

                type_code = "EM" if api_data.get("range") == "MEN" else "EL"
                final_type = f"{type_code}-{loc_raw}"

                for cell in ws.used_range:
                    if isinstance(cell.value, str) and "{_alejandro_}" in cell.value:
                        cell.value = cell.value.replace("{_alejandro_}", final_type)

                for shape in ws.shapes:
                    try:
                        if hasattr(shape, "text") and shape.text and "{_alejandro_}" in shape.text:
                            shape.text = shape.text.replace("{_alejandro_}", final_type)
                    except:
                        pass

                insert_images(
                    ws,
                    colors,
                    article.get("name", ""),
                    ['A32', 'C32', 'E32', 'G32', 'N32', 'P32']
                )

            except Exception as e:
                print(f"Error writing EBL/EBM RO format: {e}")

        
        elif jenis_template in ['EBM_RO'] and ro_na == 'REPEAT ORDER':
            try:
                print("[INFO] Menulis format template RO EB SHOES MENS")
                
                safe_set_cell(ws.range('B2'), f": {po}")
                safe_set_cell(ws.range('B4'), api_data.get("date", ""))
                safe_set_cell(ws.range('E5'), location)
                safe_set_cell(ws.range('E6'), api_data.get("seasons", ""))

                safe_set_cell(ws.range('B19'), api_data.get("delivery_date", ""))

                safe_set_cell(ws.range('A11'), api_data.get("description_name", ""))

                article = api_data.get("article", {})
                safe_set_cell(ws.range('C11'), api_data.get("style_code", "") if (article.get("name", "") or "")[:1].isdigit() else article.get("name", ""))
                safe_set_cell(ws.range('B11'), article.get("name", "") if (article.get("name", "") or "")[:1].isdigit() else '')

                materials = api_data.get("materials", [])
                material_row = 13
                for material in materials:
                    label = material.get("label", "").strip()
                    value = material.get("value", "").strip()
                    if label and value:
                        safe_set_cell(ws.range(f"A{material_row}"), f"{label} : {value}")
                        material_row += 1
                    if material_row > 17:
                        break

                api_range = (api_data.get("range") or "").strip().upper()
                if "MEN" in api_range:
                    sizes = [39, 40, 41, 42, 43, 44, 45, 46]
                    for col, size in zip(['G', 'H', 'I', 'J', 'K','L','M','N'], sizes):
                        safe_set_cell(ws.range(f"{col}20"), size)

                size_columns_bottom = {}
                for col in ['G', 'H', 'I', 'J', 'K','L','M','N']:
                    val = ws.range(f"{col}20").value
                    if val:
                        size_columns_bottom[int(val)] = col

                size_columns_top = {}
                api_range = (api_data.get("range") or "").strip().upper()

                sizes = [39, 40, 41, 42, 43, 44, 45, 46]
                for col, size in zip(['E','F','G', 'H', 'I', 'J', 'K', 'L'], sizes):
                    safe_set_cell(ws.range(f"{col}9"), size)
                    size_columns_top[size] = col

                colors = article.get("color", [])
                desc_row = 11

                for color in colors:
                    color_name = color.get("color", "")
                    if color_name and color_name != "-":
                        safe_set_cell(ws.range(f"D{desc_row}"), color_name)

                        try:
                            price = float(color.get("price", 0))
                            safe_set_cell(ws.range(f"O{desc_row}"), f"$ {price:.2f}")
                        except:
                            safe_set_cell(ws.range(f"O{desc_row}"), "$ 0.00")

                        desc_row += 1
                        if desc_row > 17:
                            break

                start_row_top = 11      # G11–M
                start_row_bottom = 21   # J22–P

                for idx, color in enumerate(colors):
                    row_top = start_row_top + idx
                    row_bottom = start_row_bottom + idx

                    sets = color.get("sets", [])

                    for set_info in sets:
                        try:
                            size = int(set_info.get("size"))
                            qty = set_info.get("qty")
                        except (ValueError, TypeError):
                            continue

                        col_top = size_columns_top.get(size)
                        if col_top:
                            safe_set_cell(ws.range(f"{col_top}{row_top}"), qty)

                        col_bottom = size_columns_bottom.get(size)
                        if col_bottom:
                            safe_set_cell(ws.range(f"{col_bottom}{row_bottom}"), qty)

                final_type = f"EM-{loc_raw}"

                for cell in ws.used_range:
                    if isinstance(cell.value, str) and "{_alejandro_}" in cell.value:
                        cell.value = cell.value.replace("{_alejandro_}", final_type)

                for shape in ws.shapes:
                    try:
                        if hasattr(shape, "text") and shape.text and "{_alejandro_}" in shape.text:
                            shape.text = shape.text.replace("{_alejandro_}", final_type)
                    except:
                        pass

                insert_images(
                    ws,
                    colors,
                    article.get("name", ""),
                    ['A32', 'C32', 'E32', 'G32', 'N32', 'P32'],
                )

            except Exception as e:
                print(f"Error writing EBL/EBM RO format: {e}")
        
        #PO EB BAGS
        elif jenis_template in ["HBL", "HBM", "EBSM", "EBSL"]: 
            try:
                print("[INFO] Menulis format template EB BAGS")
                safe_set_cell(ws.range("B2"), f": {po}")
                safe_set_cell(ws.range("E4"), f"{ro_na} - {location}")
                safe_set_cell(ws.range("B4"), api_data.get("date", ""))
                safe_set_cell(ws.range("A10"), api_data.get("description_name", ""))
                safe_set_cell(ws.range('C10'), api_data.get("style_code", ""))
                safe_set_cell(ws.range("B18"), api_data.get("delivery_date", ""))
                safe_set_cell(ws.range("E5"), api_data.get("seasons", ""))
                article = api_data.get("article", {})
                safe_set_cell(ws.range('B10'), article.get("name", ""))

                materials = api_data.get("materials", [])
                for mat in materials:
                    label = mat.get("label", "").lower()
                    value = mat.get("value", "")
                    if "upper" in label:
                        safe_set_cell(ws.range('A12'), f"Material : {value}")
                    elif "lining" in label:
                        safe_set_cell(ws.range('A13'), f"Lining : {value}") 

                colors = api_data.get("article", {}).get("color", [])
                start_row = 10
                for idx, color_info in enumerate(colors):
                    row = start_row + idx
                    safe_set_cell(ws.range(f'D{row}'), color_info.get("color", ""))
                    safe_set_cell(ws.range(f'F{row}'), color_info.get("price", "")) 
                    sets = color_info.get("sets", [])
                    if color_info.get("color", "") and color_info["color"] != "-":
                        for set_info in sets:
                            qty = set_info.get("qty", 0)
                            safe_set_cell(ws.range(f"E{row}"), qty) 
                
                mark = 'ES' if jenis_template == 'EBSL' or jenis_template == 'EBSM' else "EH"
                final_type = f'{mark}-{loc_raw}'
                for cell in ws.used_range:
                    if cell.value and isinstance(cell.value, str) and "{_alejandro_}" in cell.value:
                        cell.value = cell.value.replace("{_alejandro_}", final_type)

                for shape in ws.shapes:
                    try:
                        if hasattr(shape, "text") and shape.text and "{_alejandro_}" in shape.text:
                            shape.text = shape.text.replace("{_alejandro_}", final_type)
                    except:
                        continue   

                insert_images(ws, colors, resolve_article_candidates(api_data), ['A31', 'C31', 'E31', 'G31'])
            except Exception as e:
                print(f"Error writing HBL/HBM/EBSM format: {e}")

        # PO EVB VH  (template baru -> lihat fill_th_vh_template)
        elif jenis_template == "VBH":
            print("[INFO] Menulis format template EVB VH")
            fill_th_vh_template(ws, api_data, datas, "VH", template_file)

        # PO EVB VL
        elif jenis_template == "VBL" and ro_na != 'REPEAT ORDER':
            print("[INFO] Menulis format template EVB SHOES")
            try:
                safe_set_cell(ws.range('B2'), f": {po}")
                safe_set_cell(ws.range('B4'), api_data.get('date', ''))
                safe_set_cell(ws.range('L4'), api_data.get('seasons', ''))
                safe_set_cell(ws.range('L5'), location)
                safe_set_cell(ws.range('A11'), api_data.get("description_name", ""))
                article = api_data.get("article", {})
                safe_set_cell(ws.range('C11'), api_data.get("style_code", "") if (article.get("name", "") or "")[:1].isdigit() else article.get("name", ""))
                safe_set_cell(ws.range('B11'), article.get("name", "") if (article.get("name", "") or "")[:1].isdigit() else '')
            
                vbl_layout = detect_vbl_layout(ws)
                size_col = vbl_layout["size_columns"]
                print("[VBL] Detected size columns:", size_col)
                print("[VBL] Detected picture positions:", vbl_layout["picture_positions"])

                materials = api_data.get("materials", [])
                material_row = 13
                for material in materials:
                    label = material.get("label", "").strip()
                    value = material.get("value", "").strip()
                    safe_set_cell(ws.range(f"A{material_row}"), f"{label} : {value}") 
                    material_row += 1
                    if material_row > 17:
                        break
                safe_set_cell(ws.range('B19'), api_data.get('delivery_date', ''))
                for idx, color in enumerate(api_data.get("article", {}).get("color", [])):
                    row_main = 11 + idx
                    row_detail = 23 + idx

                    safe_set_cell(ws.range(f"D{row_main}"), color['color'])
                    safe_set_cell(ws.range(f"AB{row_main}"), format_price(color['price']))
                    for s in color['sets']:
                        size = int(s['size'])
                        qty = int(s['qty'])

                        col_letter = size_col.get(size)
                        if col_letter:
                            safe_set_cell(ws.range(f"{col_letter}{row_detail}"), qty)
                            
                location_db = "CWH" if loc_raw == "JKT" else "WMKR"
                if ro_na_raw == "NA":
                    df_sets = get_grouped_po_data(location_db)
                    size_columns = {}
                    
                    for col in range(6, 25):  # F=6, X=24
                        cell_letter = ws.range((9, col)).value
                        if not cell_letter: 
                            cell_letter = ws.range((9, col)).merge_area[0, 0].value
                        if cell_letter:
                            col_letter = ws.range((10, col)).get_address().split('$')[1]
                            size_columns[cell_letter.strip().upper()] = col_letter

                    # Loop baris data warna
                    size_suffix_list = []
                    for row in range(11, 16):   
                        color_name = ws.range(f"D{row}").value
                        if color_name:
                            color_name = "_".join(str(color_name).split())
                        else:
                            continue

                        article_name = article.get("name", "")
                        if article_name and article_name[0].isdigit():
                            article_name = api_data.get("style_code", "")
                        option_id = format_option_id(article_name, color_name)
                        sets_for_option = df_sets[df_sets['OptionID'] == option_id]

                        if sets_for_option.empty:
                            print(f"[INFO] Tidak ada data Sets untuk OptionID: {option_id}") 
                            continue
 
                        size_data = {}
                        for _, row_set in sets_for_option.iterrows():
                            size_pack_name = row_set['SizePackName']
                            num_of_pack = row_set['NumofSizePack']
                            size_code = extract_size_from_name(size_pack_name).upper() 

                            if len(size_code) >= 2:
                                suffix = size_code[-1]
                                if suffix not in size_suffix_list:
                                    size_suffix_list.append(suffix)

                            size_data[size_code] = num_of_pack
                
                        for size_code, col_letter in size_columns.items():
                            value = size_data.get(size_code, "0")
                            if not value or pd.isna(value):
                                value = "0"
                            ws.range(f"{col_letter}{row}").value = value
                 
                final_type = f'VL-{loc_raw}'
                final_sets_str = "/".join(sorted(size_suffix_list)[:5])
                for row in ws.used_range:
                    for cell in row:
                        if cell.value and isinstance(cell.value, str):
                            if "{_sets_}" in cell.value:
                                cell.value = cell.value.replace("{_sets_}", final_sets_str)
                            if "{_alejandro_}" in cell.value:
                                cell.value = cell.value.replace("{_alejandro_}", final_type)

                for shape in ws.shapes:
                    try:
                        if hasattr(shape, "text") and shape.text and "{_sets_}" in shape.text:
                            shape.text = shape.text.replace("{_sets_}", final_sets_str)
                        if hasattr(shape, "text") and shape.text and "{_alejandro_}" in shape.text:
                            shape.text = shape.text.replace("{_alejandro_}", final_type)
                    except:
                        continue     
                
                insert_images(ws, api_data.get('article', {}).get('color', []), resolve_article_candidates(api_data), ['A33', 'E33', 'N33', 'W33'])
            except Exception as e:
                print(f"Error writing VBL format: {e}")

        # RO EVB VL
        elif jenis_template == "VBL_RO" and ro_na == 'REPEAT ORDER':
            print("[INFO] Menulis format template RO EVB SHOES")
            try:
                article = api_data.get("article", {})
                safe_set_cell(ws.range('B2'), f": {po}")
                safe_set_cell(ws.range('B4'), api_data.get('date', ''))
                safe_set_cell(ws.range('G4'), api_data.get('seasons', ''))
                safe_set_cell(ws.range('G5'), location)
                safe_set_cell(ws.range('A11'), api_data.get("description_name", ""))
                safe_set_cell(ws.range('C11'), api_data.get("style_code", "") if (article.get("name", "") or "")[:1].isdigit() else article.get("name", ""))
                safe_set_cell(ws.range('B11'), article.get("name", "") if (article.get("name", "") or "")[:1].isdigit() else '')

                materials = api_data.get("materials", [])
                material_row = 13
                for material in materials:
                    label = material.get("label", "").strip()
                    value = material.get("value", "").strip()
                    safe_set_cell(ws.range(f"A{material_row}"), f"{label} : {value}")
                    material_row += 1
                    if material_row > 17:
                        break
                safe_set_cell(ws.range('B19'), api_data.get('delivery_date', ''))

                # Header ukuran (35..41) ada di BARIS 22, kolom I(9) s.d. O(15)
                size_col_map = {}
                for col in range(9, 16):        # I=9 ... O=15
                    header_val = ws.range((22, col)).value
                    size_key = normalize_size(header_val)
                    if size_key:
                        size_col_map[size_key] = col

                print("SIZE → COLUMN MAP (row 22, I:O):", size_col_map)

                for idx, color in enumerate(article.get("color", [])):
                    row_main = 11 + idx      # nama warna ditulis di D11..D14
                    row_detail = 23 + idx    # qty per size ditulis di I23:O23 dst (G11:M11 otomatis ikut lewat formula)

                    color_name = color.get("color", "").strip().upper()
                    if not color_name or color_name == "-":
                        continue

                    safe_set_cell(ws.range(f"D{row_main}"), color_name)
                    safe_set_cell(ws.range(f"P{row_main}"), format_price(color.get('price', '')))

                    size_qty_map = build_size_qty_map(color)

                    # kosongkan dulu placeholder di baris detail
                    for col in range(9, 16):
                        ws.range((row_detail, col)).value = "-"

                    for size_key, qty in size_qty_map.items():
                        col = size_col_map.get(size_key)
                        if col:
                            ws.range((row_detail, col)).value = qty

                final_type = f'VL-{loc_raw}'
                for row in ws.used_range:
                    for cell in row:
                        if cell.value and isinstance(cell.value, str):
                            if "{_alejandro_}" in cell.value:
                                cell.value = cell.value.replace("{_alejandro_}", final_type)

                for shape in ws.shapes:
                    try:
                        if hasattr(shape, "text") and shape.text and "{_alejandro_}" in shape.text:
                            shape.text = shape.text.replace("{_alejandro_}", final_type)
                    except:
                        continue

                insert_images(ws, article.get('color', []), resolve_article_candidates(api_data), ['A33', 'D33', 'L33', 'R33'])
            except Exception as e:
                print(f"Error writing RO VBL format: {e}")
                
        # PO TRACCE TK
        elif jenis_template == "TRK":
            print("[INFO] Menulis format template TRACCE KIDS")
            size_columns = {
                26: 'F', 27: 'G', 28: 'H', 29: 'I', 30: 'J', 31: 'K',
                32: 'L', 33: 'M', 34: 'N', 35: 'O', 36: 'P', 37: 'Q'
            }
            try:
                safe_set_cell(ws.range('B2'), f": {po}")
                safe_set_cell(ws.range('E3'), ro_na)
                safe_set_cell(ws.range('E4'), location)
                safe_set_cell(ws.range('B4'), api_data.get('date', ''))
                safe_set_cell(ws.range('C11'), api_data.get("article", {}).get("name", ""))
                safe_set_cell(ws.range('A11'), api_data.get("article_name_vh", ''))
                safe_set_cell(ws.range('A12'), ", ".join([f"Material : {m['value']}" for m in api_data['materials'] if m['label'] == 'Upper']))
                safe_set_cell(ws.range('B11'), api_data.get('supplier_code', ''))
                safe_set_cell(ws.range('B17'), api_data.get('delivery_date', ''))
                safe_set_cell(ws.range('E5'), api_data.get('seasons', ''))
                for idx, color in enumerate(api_data.get("article", {}).get("color", [])):
                    total_qty = sum(s['qty'] for s in color['sets'])
                    price_formatted = format_price(color['price'])
                    safe_set_cell(ws.range(f'D{11 + idx}'), color['color'])
                    safe_set_cell(ws.range(f'R{11 + idx}'), total_qty)
                    safe_set_cell(ws.range(f'S{11 + idx}'), price_formatted)
                    for set_info in color['sets']:
                        try:
                            size = int(set_info.get('size'))
                            qty = set_info.get('qty', 0)
                            col_letter = size_columns.get(size)
                            if col_letter:
                                safe_set_cell(ws.range(f'{col_letter}{11 + idx}'), qty)
                        except:
                            continue

                final_type = f'TK-{loc_raw}'

                location_db = "CWH" if loc_raw == "JKT" else "WMKR"
                size_suffix_list = []
                if ro_na_raw == "NA" or ro_na_raw == "RO":
                    df_sets = get_grouped_po_data(location_db)
                    article_name = api_data.get("article", {}).get("name", "")
                    if article_name and article_name[0].isdigit():
                        article_name = api_data.get("style_code", "")
                    for color in api_data.get("article", {}).get("color", []):
                        color_name = color.get("color", "")
                        if not color_name or color_name == "-":
                            continue
                        color_key = "_".join(str(color_name).split())
                        option_id = format_option_id(article_name, color_key)
                        sets_for_option = df_sets[df_sets['OptionID'] == option_id]
                        if sets_for_option.empty:
                            print(f"[INFO] Tidak ada data Sets untuk OptionID: {option_id}")
                            continue
                        for _, row_set in sets_for_option.iterrows():
                            size_pack_name = row_set['SizePackName']
                            size_code = extract_size_from_name(size_pack_name).upper()
                            if len(size_code) >= 2:
                                suffix = size_code[-1]
                                if suffix not in size_suffix_list:
                                    size_suffix_list.append(suffix)
                final_sets_str = "/".join(sorted(size_suffix_list)[:5])
                for cell in ws.used_range:
                    if cell.value and isinstance(cell.value, str) and "{_alejandro_}" in cell.value:
                        cell.value = cell.value.replace("{_alejandro_}", final_type)

                for shape in ws.shapes:
                    try:
                        if hasattr(shape, "text") and shape.text and "{_alejandro_}" in shape.text:
                            shape.text = shape.text.replace("{_alejandro_}", final_type)
                    except:
                        continue   

                insert_images(ws, api_data.get("article", {}).get("color", []), resolve_article_candidates(api_data), ['A29', 'B29', 'D29', 'F29'])
            except Exception as e:
                print(f"Error writing TK format: {e}")

        # PO TRACCE TL/TM
        elif jenis_template in ["TRM", "TRL"]:
            print("[INFO] Menulis format template TRACCE SHOES")
            try:
                safe_set_cell(ws.range('B2'), f": {po}")
                safe_set_cell(ws.range('E4'), ro_na)
                safe_set_cell(ws.range('E5'), location)
                safe_set_cell(ws.range('B4'), api_data.get('date', ''))
                safe_set_cell(ws.range('C11'), api_data.get("article", {}).get("name", ""))
                safe_set_cell(ws.range('F19'), api_data.get("article", {}).get("name", ""))
                safe_set_cell(ws.range('A13'), ", ".join([f"Material : {m['value']}" for m in api_data['materials'] if m['label'] == 'Upper']))
                safe_set_cell(ws.range('A14'), ", ".join([f"Heel Height : {m['value']}" for m in api_data['materials'] if m['label'] == 'Heel Height']))
                safe_set_cell(ws.range('B11'), api_data.get('supplier_code', ''))
                safe_set_cell(ws.range('B16'), api_data.get('delivery_date', ''))
                safe_set_cell(ws.range('F11'), api_data.get('range', ''))
                safe_set_cell(ws.range('A11'), api_data.get('article_name_vl', ''))
                safe_set_cell(ws.range('E6'), api_data.get('seasons', ''))
                
                api_range = (api_data.get("range") or "").strip().upper()
                size_columns = {}

                if "MEN" in api_range:
                    sizes = [39, 40, 41, 42, 43, 44, 45]
                    columns = ['K', 'L', 'M', 'N', 'O', 'P', 'Q']
                    for col, size in zip(columns, sizes):
                        safe_set_cell(ws.range(f"{col}18"), size) 
                        size_columns[size] = col
                else:
                    for col in ['K', 'L', 'M', 'N', 'O', 'P', 'Q']:
                        try:
                            size = int(ws.range(f"{col}18").value)
                            size_columns[size] = col
                        except:
                            continue

                for idx, color in enumerate(api_data.get("article", {}).get("color", [])): 
                    price_formatted = format_price(color['price']) 
                    safe_set_cell(ws.range(f'D{11 + idx}'), color['color'])
                    safe_set_cell(ws.range(f'Q{11 + idx}'), price_formatted)
                    safe_set_cell(ws.range(f'I{19 + idx}'),  color['color'])
                    for set_info in color['sets']:
                        try:
                            size = int(set_info.get('size'))
                            qty = set_info.get('qty', 0)
                            col_letter = size_columns.get(size)
                            if col_letter:
                                safe_set_cell(ws.range(f'{col_letter}{19 + idx}'), qty)
                        except:
                            continue

                location_db = "CWH" if loc_raw == "JKT" else "WMKR"
                size_suffix_list = []
                if ro_na_raw == "NA":
                    df_sets = get_grouped_po_data(location_db)
                    size_column = {}
                    
                    for col in range(7, 15):  # G=7, N=14
                        cell_letter = ws.range((9, col)).value
                        if not cell_letter: 
                            cell_letter = ws.range((9, col)).merge_area[0, 0].value
                        if cell_letter:
                            col_letter = ws.range((10, col)).get_address().split('$')[1]
                            size_column[cell_letter.strip().upper()] = col_letter

                    # Loop baris data warna
                    
                    for row in range(11, 16):   
                        color_name = ws.range(f"D{row}").value
                        if not color_name:
                            continue

                        option_id = format_option_id(api_data.get("article", {}).get("name", ""), color_name)
                        sets_for_option = df_sets[df_sets['OptionID'] == option_id]

                        if sets_for_option.empty:
                            print(f"[INFO] Tidak ada data Sets untuk OptionID: {option_id}") 
                            for size_code, col_letter in size_column.items():
                                ws.range(f"{col_letter}{row}").value = "-"
                            continue
 
                        size_data = {}
                        for _, row_set in sets_for_option.iterrows():
                            size_pack_name = row_set['SizePackName']
                            num_of_pack = row_set['NumofSizePack']
                            size_code = extract_tracee_size_from_name(size_pack_name).upper() 
                            print('Size code :: ',size_code)

                            if len(size_code) >= 2:
                                suffix = size_code[-1]
                                if suffix not in size_suffix_list:
                                    size_suffix_list.append(suffix)

                            size_data[size_code] = num_of_pack
                
                        for size_code, col_letter in size_column.items():
                            value = size_data.get(size_code, "-")
                            if not value or pd.isna(value):
                                value = "-"
                            ws.range(f"{col_letter}{row}").value = value

                type = "TM" if api_data.get("range", "") == "MEN" else "TL"
                final_type = f'{type}-{loc_raw}'
                final_sets_str = "/".join(sorted(size_suffix_list)[:5])
                for row in ws.used_range:
                    for cell in row:
                        if cell.value and isinstance(cell.value, str):
                            if "{_sets_}" in cell.value:
                                cell.value = cell.value.replace("{_sets_}", final_sets_str)
                            if "{_alejandro_}" in cell.value:
                                cell.value = cell.value.replace("{_alejandro_}", final_type)

                for shape in ws.shapes:
                    try:
                        if hasattr(shape, "text") and shape.text and "{_sets_}" in shape.text:
                            shape.text = shape.text.replace("{_sets_}", final_sets_str)
                        if hasattr(shape, "text") and shape.text and "{_alejandro_}" in shape.text:
                            shape.text = shape.text.replace("{_alejandro_}", final_type)
                    except:
                        continue     
                
                insert_images(ws, api_data.get('article', {}).get('color', []), resolve_article_candidates(api_data), ['A28', 'C28', 'F28', 'L28'])

            except Exception as e:
                print(f"Error writing TL format: {e}")

        # PO TRACCE TH  (template baru -> lihat fill_th_vh_template)
        elif jenis_template == "TRH":
            print("[INFO] Menulis format template TRACCE TH - NEW TEMPLATE", flush=True)
            fill_th_vh_template(ws, api_data, datas, "TH", template_file)

        # ------------------------------------------------------------------
        # SIMPAN HASIL  (nama file harus sama dengan yang dicari generate_po_excel)
        # ------------------------------------------------------------------
        os.makedirs(output_path, exist_ok=True)
        output_file = os.path.join(
            output_path, f"PO {po_code} - {produk} - {ro_na_raw} - {loc_raw}.xlsx"
        )
        print(f"[XL] menyimpan: {output_file}", flush=True)
        wb.save(output_file)
        return True

    finally:
        # Jangan biarkan EXCEL.EXE tertinggal jika satu PO gagal di tengah proses.
        try:
            if wb is not None:
                wb.close()
        except Exception:
            pass
        try:
            if app is not None:
                app.quit()
        except Exception:
            pass



def generate_po_excel(po_code, ref, excel_dir):
    """
    Public entry point used by job_manager.py.

    Contract:
        (excel_ok: bool, excel_path: str|None, excel_err: str|None)
    """
    try:
        jenis_combo, produk, ro_na, lokasi = parse_ref(ref)
        jenis = resolve_jenis(jenis_combo)
    except ValueError as e:
        return False, None, str(e)

    datas = {
        "po_code": po_code,
        "jenis": jenis,
        "produk": produk,
        "ro_na": ro_na,
        "lokasi": lokasi,
    }

    try:
        api_data = get_po_data_from_api(lokasi, po_code)
    except Exception as e:
        import traceback
        traceback.print_exc()
        return False, None, f"Gagal mengambil data API untuk PO {po_code}: {e}"

    if api_data is None:
        return False, None, f"Gagal ambil data dari API untuk PO {po_code}"

    os.makedirs(excel_dir, exist_ok=True)

    try:
        ok = fill_template_with_res(
            po_code,
            jenis,
            api_data,
            excel_dir,
            datas,
        )
    except Exception as e:
        import traceback
        traceback.print_exc()
        return False, None, f"Gagal generate Excel PO {po_code}: {e}"

    if not ok:
        return False, None, f"Gagal generate Excel PO {po_code}"

    output_filename = f"PO {po_code} - {produk} - {ro_na} - {lokasi}.xlsx"
    output_file = os.path.join(excel_dir, output_filename)

    if not os.path.exists(output_file):
        return (
            False,
            None,
            f"Generator mengembalikan sukses tetapi file tidak ditemukan: "
            f"{output_file}",
        )

    return True, output_file, None