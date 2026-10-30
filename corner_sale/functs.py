# from typing import Optional
# from decimal import Decimal
# from copy import copy
# import openpyxl,re,pyodbc
# from class_promo import PromoRow,PromoTable

# SHEET_MAP = {
#     ('LDS', 'SHS'): ('LADIES',           'EBL'),
#     ('LDS', 'BAG'): ('HANDBAG',          'EB LADIES BAG'),
#     ('LDS', 'BGC'): ('HANDBAG',          'EB LADIES BAG'),
#     ('MEN', 'SHS'): ('MENS',             'EBM'),
#     ('MEN', 'BAG'): ('HANDBAG',          'EB MENS BAG'),
#     ('LDS', 'SLG'): ('SUNDRIES',         'EBSL'),
#     ('LDS', 'WAL'): ('SUNDRIES',         'EBSL'),
#     ('LDS', 'ACC'): ('SUNDRIES',         'EBSL'),
#     ('LDS', 'WRS'): ('SUNDRIES',         'EBSL'),
#     ('LDS', 'FWC'): ('SUNDRIES',         'EBSL'),
#     ('MEN', 'SLG'): ('SUNDRIES',         'EBSM'),
#     ('MEN', 'WAL'): ('SUNDRIES',         'EBSM'),
#     ('MEN', 'BLT'): ('SUNDRIES',         'EBSM'),
#     ('MEN', 'ACC'): ('SUNDRIES',         'EBSM'),
#     ('MEN', 'WRS'): ('SUNDRIES',         'EBSM'),
#     ('LDS', 'SHC'): ('SHOECARE & SOCKS', 'SHOECARE'),
#     ('MEN', 'SHC'): ('SHOECARE & SOCKS', 'SHOECARE'),
#     ('MEN', 'SCK'): ('SHOECARE & SOCKS', 'SOCKS'),
# }
 
# def get_sheet_and_suffix(dept: str, segment: str):
#     """Return (sheet_name, suffix) atau (None, None) jika tidak ada di mapping."""
#     return SHEET_MAP.get((dept.strip().upper(), segment.strip().upper()), (None, None))

# def sort_rows(rows: list) -> list:
#     return sorted(rows, key=lambda r: (r.article.upper(), r.colour.upper()))
 
# def is_promo_header_row(row_vals: list) -> bool:
#     texts = [str(v).strip() if v else '' for v in row_vals]
#     return any('article' in t.lower() for t in texts) and \
#            any('colour' in t.lower() for t in texts)

# def detect_column_groups(row_vals: list) -> list:
#     groups = []
#     for i, val in enumerate(row_vals):
#         if val and str(val).strip().lower() == 'article':
#             try:
#                 c_colour = next(
#                     j for j in range(i + 1, min(i + 6, len(row_vals)))
#                     if row_vals[j] and 'colour' in str(row_vals[j]).lower()
#                 )
#                 c_awal = next(
#                     j for j in range(c_colour + 1, min(c_colour + 4, len(row_vals)))
#                     if row_vals[j] and 'harga' in str(row_vals[j]).lower() and 'awal' in str(row_vals[j]).lower()
#                 )
#                 c_diskon = next(
#                     j for j in range(c_awal + 1, min(c_awal + 4, len(row_vals)))
#                     if row_vals[j] and 'harga' in str(row_vals[j]).lower() and 'diskon' in str(row_vals[j]).lower()
#                 )
#                 groups.append((i, c_colour, c_awal, c_diskon))
#             except StopIteration:
#                 pass
#     return groups

# def extract_numeric(val) -> Optional[int]:
#     if val is None:
#         return None
#     if isinstance(val, (int, float)):
#         return int(val)
#     if isinstance(val, str):
#         # If it's a formula, skip
#         if val.startswith('='):
#             return None
#         # Try stripping non-numeric chars
#         cleaned = re.sub(r'[^\d]', '', val)
#         return int(cleaned) if cleaned else None
#     return None

# def make_discount_key(promo_name: str) -> str:
#     m = re.search(r'\bSP(\d+K?)\b', promo_name, re.IGNORECASE)
#     if m:
#         return 'SP' + m.group(1).upper()

#     m = re.search(r'\bD(\d+)\b', promo_name, re.IGNORECASE)
#     if m:
#         return 'D' + m.group(1)

#     m = re.search(r'(\d+)\s*%', promo_name)
#     if m:
#         return 'D' + m.group(1)

#     return promo_name.split()[0].upper()

# def is_promo_name_cell(val) -> bool:
#     if not val or not isinstance(val, str):
#         return False
#     v = val.strip()
#     if not v:
#         return False
#     patterns = [
#         r'DISC[\.\s]*\d+',
#         r'D\d+\s+(SOCKS|SHOECARE)',
#         r'SP\d+K?\s',
#         r'SP\d+K?$',
#     ]
#     return any(re.search(p, v, re.IGNORECASE) for p in patterns)

# # ---------------------------------------------------------------------------
# # FIX: kategori tabel (Mens Bag vs Ladies Bag, dst) harus ikut membedakan
# # tabel promo, bukan cuma persentase diskonnya.
# #
# # Sebelumnya, discount_key hanya berupa 'D40', 'D50', dst -- diambil murni
# # dari angka persen di nama promo. Padahal dalam SATU sheet (contoh: HANDBAG)
# # bisa ada beberapa tabel dengan persen yang SAMA tapi kategori BERBEDA,
# # misalnya "DISC. 40% EB LADIES BAG" dan "DISC.40% EB MENS BAG" -- keduanya
# # menghasilkan discount_key 'D40'. Karena LADIES BAG di-scan lebih dulu
# # (posisinya di atas dalam sheet), dia yang dapat key 'D40' yang bersih,
# # sedangkan MENS BAG kebagian key 'D40_2' hasil dari mekanisme
# # anti-tabrakan. Akibatnya, saat artikel MENS BAG dipindah discount-nya,
# # proses hapus-dari-tabel-lama dan taruh-ke-tabel-baru mencari/menaruh
# # berdasarkan key polos ('D40'/'D50') yang justru menunjuk ke tabel
# # LADIES BAG -- artikel pun nyasar ke tabel LADIES.
# #
# # Fix: setiap tabel diberi 'category_suffix' (mis. 'EB MENS BAG',
# # 'EB LADIES BAG', 'EBL', 'EBSM', dst) yang diekstrak dari teks promo_name.
# # Key yang dipakai untuk MENYIMPAN tabel di dalam dict digabung dari
# # discount_key + category_suffix, sehingga dua tabel dengan persen sama
# # tapi kategori beda tidak akan pernah dianggap tabel yang sama.
# # ---------------------------------------------------------------------------
# CATEGORY_SUFFIXES = sorted(
#     {suffix for (_dept, _seg), (_sheet, suffix) in SHEET_MAP.items()},
#     key=len, reverse=True,
# )

# def extract_category_suffix(promo_name: str) -> Optional[str]:
#     """
#     Cari suffix kategori (mis. 'EB MENS BAG', 'EBL', 'EBSM', 'SHOECARE', ...)
#     yang tercantum dalam teks promo_name. Dicocokkan dari yang paling
#     panjang/spesifik dulu supaya tidak salah tangkap suffix yang lebih pendek.
#     Return None jika promo_name tidak mengandung suffix kategori apa pun
#     (mis. tabel SP yang dipakai bersama utk semua kategori dalam 1 sheet).
#     """
#     name_up = promo_name.upper()
#     for suf in CATEGORY_SUFFIXES:
#         if suf in name_up:
#             return suf
#     return None


# def _table_store_key(discount_key: str, category_suffix: Optional[str]) -> str:
#     """Key gabungan yang dipakai sebagai key dict promo_groups[sheet][...]."""
#     return f"{discount_key}::{category_suffix}" if category_suffix else discount_key


# def resolve_table_key(promo_groups: dict, sheet_name: str,
#                       discount_key: str, category_suffix: Optional[str]) -> str:
#     """
#     Tentukan key dict yang tepat untuk mengakses/menyimpan tabel promo di
#     promo_groups[sheet_name], dengan mempertimbangkan category_suffix supaya
#     tabel dengan discount_key sama tapi kategori berbeda (mis. Mens Bag vs
#     Ladies Bag) tidak saling tertukar. Dipakai baik saat mencari tabel
#     SUMBER (utk dihapus) maupun tabel TUJUAN (utk ditambahkan/dibuat).
#     """
#     return _table_store_key(discount_key, category_suffix)


# def scan_excel(filepath: str) -> dict:
#     wb = openpyxl.load_workbook(filepath, data_only=True)
#     result = {}

#     for sheet_name in wb.sheetnames:
#         ws = wb[sheet_name]
#         sheet_result = {}

#         all_rows = []
#         for row in ws.iter_rows(values_only=True):
#             all_rows.append(list(row))

#         # Walk rows to find promo sections
#         i = 0
#         while i < len(all_rows):
#             row = all_rows[i]
#             non_none = [v for v in row if v is not None]

#             # Detect promo section header
#             first_val = next((v for v in row if v is not None), None)
#             if first_val and is_promo_name_cell(str(first_val).strip()):
#                 promo_name = str(first_val).strip()
#                 # Clean newlines in name
#                 promo_name = promo_name.replace('\n', ' ').strip()
#                 discount_key    = make_discount_key(promo_name)
#                 category_suffix = extract_category_suffix(promo_name)

#                 # Find the column-header row (next row with 'Article')
#                 header_row_idx = None
#                 for j in range(i + 1, min(i + 4, len(all_rows))):
#                     if is_promo_header_row(all_rows[j]):
#                         header_row_idx = j
#                         break

#                 if header_row_idx is None:
#                     i += 1
#                     continue

#                 col_groups = detect_column_groups(all_rows[header_row_idx])
#                 if not col_groups:
#                     i += 1
#                     continue

#                 # FIX: key penyimpanan digabung dengan category_suffix supaya
#                 # tabel dengan persen sama tapi kategori beda (mis. "DISC.40%
#                 # EB MENS BAG" vs "DISC. 40% EB LADIES BAG") tidak dianggap
#                 # tabel yang sama / saling menimpa. Fallback "_2", "_3", ...
#                 # tetap dipertahankan sebagai jaring pengaman untuk kasus
#                 # duplikat sungguhan (persen SAMA + kategori SAMA).
#                 unique_key = _table_store_key(discount_key, category_suffix)
#                 dupe_suffix = 2
#                 while unique_key in sheet_result:
#                     unique_key = f"{_table_store_key(discount_key, category_suffix)}_{dupe_suffix}"
#                     dupe_suffix += 1

#                 table = PromoTable(
#                     sheet=sheet_name,
#                     promo_name=promo_name,
#                     discount_key=discount_key,          # tetap 'D40'/'D50'/'SP490K' yg bersih
#                     category_suffix=category_suffix,    # 'EB MENS BAG' / 'EB LADIES BAG' / None
#                     rows=[]
#                 )

#                 # Collect data rows until next promo header or end
#                 data_start = header_row_idx + 1
#                 k = data_start
#                 while k < len(all_rows):
#                     data_row = all_rows[k]
#                     non_none_dr = [v for v in data_row if v is not None]

#                     # Stop if next promo name found
#                     first_dr_val = next((v for v in data_row if v is not None), None)
#                     if first_dr_val and is_promo_name_cell(str(first_dr_val).strip()):
#                         break

#                     # Stop if another column-header row found (different section)
#                     if is_promo_header_row(data_row) and k != header_row_idx:
#                         # Could be a new section starting horizontally in different col group
#                         # Check if a NEW promo header was in a non-B column
#                         new_groups = detect_column_groups(data_row)
#                         if new_groups:
#                             break

#                     # Extract data from each column group
#                     row_had_data = False
#                     for (ci_art, ci_col, ci_awal, ci_dis) in col_groups:
#                         art_val = data_row[ci_art] if ci_art < len(data_row) else None
#                         col_val = data_row[ci_col] if ci_col < len(data_row) else None
#                         awal_val = data_row[ci_awal] if ci_awal < len(data_row) else None
#                         dis_val = data_row[ci_dis] if ci_dis < len(data_row) else None

#                         if art_val and isinstance(art_val, str) and art_val.strip():
#                             article = art_val.strip()
#                             colour = str(col_val).strip() if col_val else ''
#                             harga_awal = extract_numeric(awal_val)
#                             harga_diskon = extract_numeric(dis_val)

#                             # Skip if article looks like another promo header
#                             if is_promo_name_cell(article):
#                                 continue

#                             table.rows.append(PromoRow(
#                                 article=article,
#                                 colour=colour,
#                                 harga_awal=harga_awal,
#                                 harga_diskon=harga_diskon,
#                             ))
#                             row_had_data = True

#                     k += 1

#                 if table.rows:
#                     sheet_result[unique_key] = table
#                     i = k  # jump past the consumed rows
#                     continue

#             i += 1

#         if sheet_result:
#             result[sheet_name] = sheet_result

#     return result

# # ---------------------------------------------------------------------------
# # Output formatter
# # ---------------------------------------------------------------------------

# def format_harga(val) -> str:
#     if val is None:
#         return 'None'
#     return str(int(val))

# def write_output(data: dict, output_path: str):
#     lines = []
#     lines.append('promo_groups = {')

#     sheet_names = list(data.keys())
#     for s_idx, sheet_name in enumerate(sheet_names):
#         tables = data[sheet_name]
#         is_last_sheet = (s_idx == len(sheet_names) - 1)

#         lines.append(f'   "{sheet_name}": {{')

#         table_keys = list(tables.keys())
#         for t_idx, key in enumerate(table_keys):
#             table = tables[key]
#             is_last_table = (t_idx == len(table_keys) - 1)

#             sorted_rows = sort_rows(table.rows)
#             count = len(sorted_rows)

#             lines.append(f'        "{key}": PromoTable(')
#             lines.append(f"            sheet='{table.sheet}',")
#             lines.append(f"            promo_name='{table.promo_name}',")
#             lines.append(f"            discount_key='{table.discount_key}',")
#             lines.append(f"            count={count},")
#             lines.append(f'            rows=[')

#             for row in sorted_rows:
#                 article = row.article.replace("'", "\\'")
#                 colour = row.colour.replace("'", "\\'")
#                 ha = format_harga(row.harga_awal)
#                 hd = format_harga(row.harga_diskon)
#                 mf        = f"'{row.moved_from}'" if row.moved_from else 'None'
#                 mt        = f"'{row.moved_to}'"   if row.moved_to   else 'None'
#                 lines.append(
#                     f"                PromoRow(article='{article}', "
#                     f"colour='{colour}', "
#                     f"harga_awal={ha}, "
#                     f"harga_diskon={hd}, "
#                     f"moved_from={mf}, moved_to={mt}),"
#                 )

#             lines.append(f'            ]')
#             lines.append(f'        ){","  if not is_last_table else ""}')

#         lines.append(f'    }}{","  if not is_last_sheet else ""}')

#     lines.append('}')

#     with open(output_path, 'w', encoding='utf-8') as f:
#         f.write('\n'.join(lines))
#         f.write('\n')

#     print(f"Output written to: {output_path}")

# # ---------------------------------------------------------------------------
# # Summary printer
# # ---------------------------------------------------------------------------

# def print_summary(data: dict):
#     print("\n=== SCAN SUMMARY ===")
#     total_tables = 0
#     total_rows = 0
#     for sheet_name, tables in data.items():
#         print(f"\nSheet: {sheet_name}")
#         for key, table in tables.items():
#             print(f"  [{key}] {table.promo_name} -> {len(table.rows)} rows")
#             total_tables += 1
#             total_rows += len(table.rows)
#     print(f"\nTotal sheets  : {len(data)}")
#     print(f"Total tables  : {total_tables}")
#     print(f"Total data rows: {total_rows}")

# SERVER   = '149.129.249.131'
# DATABASE = 'Y2_PROD'
# USERNAME = 'sa'
# PASSWORD = 'Eshoes09'

# DB_CONNECTION = (
#     f"DRIVER={{SQL Server}};"
#     f"SERVER={SERVER};"
#     f"DATABASE={DATABASE};"
#     f"UID={USERNAME};"
#     f"PWD={PASSWORD};"
#     f"TrustServerCertificate=yes;"
# )

# SQL_QUERY = """
# SELECT 
#   GA_CODEBARRE as Item_Barcode,
#   ISNULL(RP.GF_PRIXUNITAIRE,0) as Retail_Price,
#   GA_CODEARTICLE as Article,
#   GA_ARTICLE as Item_Identifier, 
#   GA_LIBREART5 as Promotion, 
#   GA_CODEARTICLE as Item_code,
#   GA_CHARLIBRE1 as Vendor_Style_Code, 
#   GA_FAMILLENIV1 as Brand, 
#   GA_FAMILLENIV2 as Department, 
#   GA_FAMILLENIV3 as Segment, 
#   GA2_FAMILLENIV4 as Category, 
#   GDI_LIBELLE as Colour_Desc,
#   GA_DIMMASQUE as Dim2, 
#   GDI_LIBELLE_1 as Size_Desc,
#   YX6.YX_LIBELLE as Material,
#   YX7.YX_LIBELLE as Upper_Style,
#   YX12.YX_LIBELLE as Sub_Upper_Style,
#   YX4.YX_LIBELLE as Toe_Style,
#   YX8.YX_LIBELLE as Heel_Height,
#   YX10.YX_LIBELLE as Heel_Style,
#   YX9.YX_LIBELLE as Wearing_Style,
#   YX11.YX_LIBELLE as Season,
#   0 as Basic_Price,
#   GA2_LIBREARTC as Category_Product,
#   GA_PAHT as Item_Purchase_Price_Tax_Excl,
#   YX15.YX_LIBELLE as Lining_Material,
#   YX14.YX_LIBELLE as Insole_Material,
#   YX16.YX_LIBELLE as Outsole_Material,
#   GA_CODEDOUANIER as Custom_Reference,
#   GA_TYPEARTICLE as Item_Type,
#   GA_PRHT as Item_Cost_Price_Tax_Excl,
#   0.00 as Formula,
#   GA_POIDSBRUT as Gross_Weight_Item,
#   GA_TENUESTOCK as In_Item_Stock,
#   GA_DATEMODIF as Item_Modification_Date,
#   GA_DATECREATION as Item_Creation_Date,
#   GA_CODEDIM2 as Size,
#   GA_EMBALLAGE as Packing_Item,
#   GA_UNITEQTEVTE as Sales_Unit_Quantity
# FROM 
#   ARTICLEDIM_MODE AM
#   OUTER APPLY (
#     SELECT TOP 1
#         RP.GF_PRIXUNITAIRE
#     FROM GCTARFCONMODEART RP
#     WHERE RP.GA_CODEARTICLE = AM.GA_CODEARTICLE
#       AND RP.GF_REGIMEPRIX = 'TTC'
#       AND RP.GFM_TYPETARIF = 'BSP01'
#       AND RP.GFM_PERTARIF IN ('00BASE', 'SGC1')
#       AND RP.GFM_NATURETYPE = 'VTE'
#       AND RP.GF_ARTICLE <> ''
#       AND GETDATE() BETWEEN RP.GF_DATEDEBUT AND RP.GF_DATEFIN
#     ORDER BY
#         RP.GF_DATEDEBUT DESC,
#         RP.GF_DATEMODIF DESC
# ) RP
#   LEFT OUTER JOIN CHOIXCOD CC1 ON GA_FAMILLENIV1 = CC1.CC_CODE AND CC1.CC_TYPE = 'FN1' 
#   LEFT OUTER JOIN CHOIXCOD CC2 ON GA_FAMILLENIV2 = CC2.CC_CODE AND CC2.CC_TYPE = 'FN2'
#   LEFT OUTER JOIN CHOIXCOD CC3 ON GA_FAMILLENIV3 = CC3.CC_CODE AND CC3.CC_TYPE = 'FN3' 
#   LEFT OUTER JOIN CHOIXEXT YX4 ON GA_LIBREART2 = YX4.YX_CODE AND YX4.YX_TYPE = 'LA2' 
#   LEFT OUTER JOIN CHOIXCOD CC5 ON GA2_FAMILLENIV4 = CC5.CC_CODE AND CC5.CC_TYPE = 'FN4' 
#   LEFT OUTER JOIN CHOIXEXT YX6 ON GA_LIBREART7 = YX6.YX_CODE AND YX6.YX_TYPE = 'LA7' 
#   LEFT OUTER JOIN CHOIXEXT YX7 ON GA_LIBREART1 = YX7.YX_CODE AND YX7.YX_TYPE = 'LA1' 
#   LEFT OUTER JOIN CHOIXEXT YX8 ON GA_LIBREART3 = YX8.YX_CODE AND YX8.YX_TYPE = 'LA3' 
#   LEFT OUTER JOIN CHOIXEXT YX9 ON GA_LIBREART4 = YX9.YX_CODE AND YX9.YX_TYPE = 'LA4' 
#   LEFT OUTER JOIN CHOIXEXT YX10 ON GA_LIBREART8 = YX10.YX_CODE AND YX10.YX_TYPE = 'LA8' 
#   LEFT OUTER JOIN CHOIXEXT YX11 ON GA_LIBREART9 = YX11.YX_CODE AND YX11.YX_TYPE = 'LA9' 
#   LEFT OUTER JOIN CHOIXEXT YX12 ON GA2_LIBREARTB = YX12.YX_CODE AND YX12.YX_TYPE = 'LAB' 
#   LEFT OUTER JOIN CHOIXEXT YX13 ON GA2_LIBREARTC = YX13.YX_CODE AND YX13.YX_TYPE = 'LAC' 
#   LEFT OUTER JOIN CHOIXEXT YX14 ON GA2_LIBREARTD = YX14.YX_CODE AND YX14.YX_TYPE = 'LAD' 
#   LEFT OUTER JOIN CHOIXEXT YX15 ON GA2_LIBREARTE = YX15.YX_CODE AND YX15.YX_TYPE = 'LAE' 
#   LEFT OUTER JOIN CHOIXEXT YX16 ON GA2_LIBREARTF = YX16.YX_CODE AND YX16.YX_TYPE = 'LAF' 
# WHERE 
#   GA_TENUESTOCK = 'X' 
#   AND GA_FAMILLENIV1 = ?
#   AND GA_EMBALLAGE <> 'X' 
#   AND GA_TYPEARTICLE IN ('MAR')
#   AND GA_TYPEARTICLE NOT IN ('PRE', 'FI', 'FRA')
#   AND GA_STATUTART = 'DIM'
# """

# def get_db_connection():
#     try:
#         conn = pyodbc.connect(DB_CONNECTION)
#         print(f"  [DB] Koneksi berhasil ke {SERVER} / {DATABASE}")
#         return conn
#     except Exception as e:
#         print(f"  [DB ERROR] Gagal konek ke database: {e}")
#         raise

# def execute_query(brand: str) -> dict:
#     """
#     Eksekusi SQL_QUERY ke database dengan parameter brand.
#     Return format:
#         { article_upper: { colour_upper: { retail_price, promotion, dept, segment, brand, article, colour } } }

#     Eksekusi SQL_QUERY ke database dengan parameter brand.
#     Return format:
#         { article_upper: { colour_upper: { retail_price, promotion, dept, segment, brand, article, colour } } }
#     """
#     print(f"  [DB] Menjalankan query untuk brand='{brand}'...")

#     conn   = get_db_connection()
#     cursor = conn.cursor()

#     try:
#         cursor.execute(SQL_QUERY, brand)
#         columns = [col[0] for col in cursor.description]
#         rows    = cursor.fetchall()
#     finally:
#         cursor.close()
#         conn.close()

#     print(f"  [DB] {len(rows)} baris diterima dari database.")

#     col_idx = {c: i for i, c in enumerate(columns)}

#     data = {}
#     for row in rows:
#         article   = str(row[col_idx['Article']]).strip()     if row[col_idx['Article']]    else ''
#         colour    = str(row[col_idx['Colour_Desc']]).strip() if row[col_idx['Colour_Desc']] else ''
#         retail    = row[col_idx['Retail_Price']]
#         promotion = str(row[col_idx['Promotion']]).strip()   if row[col_idx['Promotion']]  else ''
#         dept      = str(row[col_idx['Department']]).strip()  if row[col_idx['Department']] else ''
#         segment   = str(row[col_idx['Segment']]).strip()     if row[col_idx['Segment']]    else ''
#         brand_val = str(row[col_idx['Brand']]).strip()       if row[col_idx['Brand']]      else ''

#         if not article:
#             continue

#         art_key    = article.upper()
#         col_key    = colour.upper()
#         retail_val = int(retail) if isinstance(retail, (int, float, Decimal)) and retail else 0

#         if art_key not in data:
#             data[art_key] = {}

#         existing = data[art_key].get(col_key)
#         # Simpan jika: belum ada entry, atau harga baru lebih tinggi.
#         # Jangan timpa harga yang sudah ada dengan 0.
#         if existing is None or (retail_val > 0 and retail_val > existing['retail_price']):
#             data[art_key][col_key] = {
#                 'retail_price': retail_val,
#                 'promotion':    promotion,
#                 'dept':         dept,
#                 'segment':      segment,
#                 'brand':        brand_val,
#                 'article':      article,
#                 'colour':       colour,
#             }

#     return data

# def normalize_key(raw: str) -> str:
#     """Normalize input 'Item Group Menjadi' value to canonical discount_key."""
#     raw = raw.strip()
#     # e.g. "SP490" -> "SP490K", "D20" -> "D20"
#     m = re.fullmatch(r'SP(\d+)K?', raw, re.IGNORECASE)
#     if m:
#         return 'SP' + m.group(1) + 'K'
#     m = re.fullmatch(r'D(\d+)', raw, re.IGNORECASE)
#     if m:
#         return 'D' + m.group(1)
#     return raw.upper()
 
# def load_promo_groups_from_txt(txt_path: str) -> dict:
#     with open(txt_path, 'r', encoding='utf-8') as f:
#         raw = f.read()
 
#     result = {}
#     sheet_pattern = re.compile(
#         r'"([^"]+)":\s*\{(.*?)\n    \}',
#         re.DOTALL
#     )
#     table_pattern = re.compile(
#         r'"([^"]+)":\s*PromoTable\((.*?)\n        \)',
#         re.DOTALL
#     )
#     row_pattern = re.compile(
#         r"PromoRow\(article='(.*?)',\s*colour='(.*?)',\s*harga_awal=(\S+?),\s*harga_diskon=(\S+?)"
#         r"(?:,\s*moved_from=(.*?),\s*moved_to=(.*?))?\)"
#     )
 
#     for sm in sheet_pattern.finditer(raw):
#         sheet_name = sm.group(1)
#         sheet_body = sm.group(2)
#         tables = {}
 
#         for tm in table_pattern.finditer(sheet_body):
#             key = tm.group(1)
#             table_body = tm.group(2)
 
#             # Extract fields
#             def _get(field_name):
#                 m = re.search(rf"{field_name}='([^']*)'", table_body)
#                 return m.group(1) if m else ''
 
#             def _get_int(field_name):
#                 m = re.search(rf"{field_name}=(\d+)", table_body)
#                 return int(m.group(1)) if m else 0
 
#             sheet_val    = _get('sheet')
#             promo_name   = _get('promo_name')
#             discount_key = _get('discount_key')
#             count        = _get_int('count')
 
#             rows = []
#             for rm in row_pattern.finditer(table_body):
#                 art  = rm.group(1).replace("\\'", "'")
#                 col  = rm.group(2).replace("\\'", "'")
#                 ha   = None if rm.group(3) == 'None' else int(rm.group(3).rstrip(','))
#                 hd   = None if rm.group(4).rstrip(',') == 'None' else float(rm.group(4).rstrip(','))
#                 mf_raw = rm.group(5).strip().strip("'") if rm.group(5) else 'None'
#                 mt_raw = rm.group(6).strip().strip("'") if rm.group(6) else 'None'
#                 mf = None if mf_raw == 'None' else mf_raw
#                 mt = None if mt_raw == 'None' else mt_raw
#                 rows.append(PromoRow(article=art, colour=col,
#                                      harga_awal=ha, harga_diskon=hd,
#                                      moved_from=mf, moved_to=mt))
 
#             tables[key] = PromoTable(
#                 sheet=sheet_val,
#                 promo_name=promo_name,
#                 discount_key=discount_key,
#                 count=count,
#                 rows=rows,
#             )
 
#         result[sheet_name] = tables
#     return result
 
# def load_query_result(xlsx_path: str) -> dict:
#     wb = openpyxl.load_workbook(xlsx_path, data_only=True)
#     ws = wb.active
 
#     headers = [str(c.value).strip() if c.value else '' for c in next(ws.iter_rows(min_row=1, max_row=1))]
#     idx = {h: i for i, h in enumerate(headers)}
 
#     data = {}  
#     for row in ws.iter_rows(min_row=2, values_only=True):
#         article   = str(row[idx['Article']]).strip() if row[idx['Article']] else ''
#         colour    = str(row[idx['Colour_Desc']]).strip() if row[idx['Colour_Desc']] else ''
#         retail    = row[idx['Retail_Price']]
#         promotion = str(row[idx['Promotion']]).strip() if row[idx['Promotion']] else ''
#         dept      = str(row[idx['Department']]).strip() if row[idx['Department']] else ''
#         segment   = str(row[idx['Segment']]).strip() if row[idx['Segment']] else ''
#         brand     = str(row[idx['Brand']]).strip() if row[idx['Brand']] else ''
 
#         if not article:
#             continue
 
#         art_key = article.upper()
#         col_key = colour.upper()
 
#         retail_val = int(retail) if isinstance(retail, (int, float, Decimal)) and retail else 0
 
#         if art_key not in data:
#             data[art_key] = {}
 
#         existing = data[art_key].get(col_key)
#         # Simpan jika: belum ada entry, atau harga baru lebih tinggi.
#         # Jangan timpa harga yang sudah ada dengan 0.
#         if existing is None or (retail_val > 0 and retail_val > existing['retail_price']):
#             data[art_key][col_key] = {
#                 'retail_price': retail_val,
#                 'promotion':    promotion,
#                 'dept':         dept,
#                 'segment':      segment,
#                 'brand':        brand,
#                 'article':      article,
#                 'colour':       colour,
#             }
 
#     return data
 
# def load_input_file(xlsx_path: str) -> list:
#     wb = openpyxl.load_workbook(xlsx_path, data_only=True)
#     ws = wb.active
 
#     rows = []
#     for i, row in enumerate(ws.iter_rows(min_row=2, values_only=True)):
#         if not any(v for v in row):
#             continue
#         brand        = str(row[0]).strip() if row[0] else ''
#         item_code    = str(row[1]).strip() if row[1] else ''
#         colour       = str(row[2]).strip() if row[2] else None
#         if colour in ('None', '', 'none'):
#             colour = None
#         item_group_menjadi = str(row[4]).strip() if row[4] else ''
#         rows.append({
#             'brand':            brand,
#             'item_code':        item_code.upper(),
#             'colour':           colour.upper() if colour else None,
#             'item_group_menjadi': normalize_key(item_group_menjadi),
#         })
#     return rows
  
# def compute_harga_diskon(harga_awal: int, discount_key: str) -> Optional[int]:
#     if harga_awal is None or harga_awal == 0:
#         return None
 
#     m = re.fullmatch(r'D(\d+)', discount_key, re.IGNORECASE)
#     if m:
#         pct = int(m.group(1))
#         return round(harga_awal * (1 - pct / 100))
 
#     m = re.fullmatch(r'SP(\d+)K', discount_key, re.IGNORECASE)
#     if m:
#         return int(m.group(1)) * 1000
 
#     return None

# def recount(promo_groups: dict):
#     """Sort semua rows dan update count di setiap PromoTable."""
#     for sheet_tables in promo_groups.values():
#         for table in sheet_tables.values():
#             table.rows  = sort_rows(table.rows)
#             table.count = len(table.rows)

# # ---------------------------------------------------------------------------
# # Apply instruksi input ke promo_groups
# # ---------------------------------------------------------------------------
# def _build_promo_name(discount_key: str, suffix: str) -> str:
#     """Buat promo_name yang readable dari discount_key dan suffix sheet."""
#     m = re.fullmatch(r'D(\d+)', discount_key, re.IGNORECASE)
#     if m:
#         return f"DISC. {m.group(1)}% {suffix}"
#     m = re.fullmatch(r'SP(\d+)K', discount_key, re.IGNORECASE)
#     if m:
#         return f"SP{m.group(1)}K {suffix}"
#     return f"{discount_key} {suffix}"


# # ---------------------------------------------------------------------------
# # BUG FIX 1: _get_colour_promo_map
# # Sebelumnya, script mengambil satu sample pertama dari q_records untuk
# # menentukan orig_key (promo asal). Ini salah karena 1 artikel bisa punya
# # promo berbeda per warna (contoh: AMARELA LIGHT BLUE=D50, BLACK=D30).
# # Fungsi ini membangun mapping colour -> orig_key secara individual.
# # ---------------------------------------------------------------------------
# def _get_colour_promo_map(q_records: dict) -> dict:
#     """Return { colour_upper: normalize_key(promotion) or None }"""
#     result = {}
#     for col_key, rec in q_records.items():
#         orig_promo = rec.get('promotion', '')
#         if orig_promo not in ('NML', '', 'NULL', None):
#             result[col_key] = normalize_key(orig_promo)
#         else:
#             result[col_key] = None
#     return result


# def _expand_all_colour_row(src_table, art_code: str, q_records: dict,
#                            exclude_colours: set, promo_groups: dict,
#                            sheet_name: str) -> Optional[int]:
#     """
#     Jika tabel asal punya row 'All Colour' untuk artikel ini, hapus row tersebut
#     lalu kembalikan sisa warna (yang tidak ada di exclude_colours) ke tabel asal
#     sebagai row per warna. Return harga dari row All Colour (sebagai fallback harga),
#     atau None jika tidak ada row All Colour yang ditemukan.

#     Dipanggil saat instruksi adalah per warna tertentu saja (bukan All Colour),
#     sehingga warna yang tidak dipindahkan harus tetap muncul di tabel asal.
#     """
#     harga_fallback = None
#     new_rows = []
#     expanded = False

#     for existing_row in src_table.rows:
#         if existing_row.article.upper() == art_code and existing_row.colour.upper() == 'ALL COLOUR':
#             # Simpan harga sebagai fallback
#             harga_fallback = existing_row.harga_awal
#             expanded = True
#             # Jangan masukkan ke new_rows (row All Colour ini dihapus)
#         else:
#             new_rows.append(existing_row)

#     if expanded:
#         src_table.rows = new_rows
#         # Tambahkan sisa warna individual yang tidak masuk instruksi pindah
#         colour_promo_map = _get_colour_promo_map(q_records)
#         for col_key, rec in q_records.items():
#             if col_key in exclude_colours:
#                 continue  # warna ini sedang dipindahkan, skip
#             actual_colour = rec['colour']
#             harga_awal = rec['retail_price'] or harga_fallback or 0
#             # Tentukan discount_key tabel asal dari colour_promo_map
#             orig_key = colour_promo_map.get(col_key)
#             # Cari tabel yang sesuai (orig_key atau tabel saat ini)
#             src_discount_key = src_table.discount_key
#             harga_diskon = compute_harga_diskon(harga_awal, src_discount_key)
#             # Pastikan belum ada
#             already = any(r.article.upper() == art_code and r.colour.upper() == actual_colour.upper()
#                           for r in src_table.rows)
#             if not already:
#                 src_table.rows.append(PromoRow(
#                     article=rec['article'],
#                     colour=actual_colour,
#                     harga_awal=harga_awal,
#                     harga_diskon=harga_diskon,
#                     moved_from=None,
#                     moved_to=None,
#                 ))
#                 print(f"  [EXPAND] {art_code}/{actual_colour} tetap di {src_discount_key} (sisa dari All Colour)")

#     return harga_fallback


# def _remove_from_source(promo_groups: dict, sheet_name: str,
#                         art_code: str, col_key: str, moved_from: Optional[str],
#                         target_key: str, category_suffix: Optional[str] = None,
#                         also_remove_all_colour: bool = False):
#     """
#     Hapus row (art_code/col_key) dari tabel asal jika moved_from != target.
#     Jika also_remove_all_colour=True, juga hapus row 'All Colour' untuk artikel
#     yang sama (dipakai saat semua warna dipindahkan sekaligus).

#     FIX: tabel sumber dicari dengan discount_key + category_suffix (bukan
#     discount_key polos), supaya kalau dalam satu sheet ada beberapa tabel
#     dengan persen sama tapi kategori beda (mis. EB MENS BAG vs EB LADIES
#     BAG), penghapusan tidak salah sasaran ke tabel kategori lain -- yang
#     sebelumnya menyebabkan artikel Mens Bag tidak pernah terhapus dari
#     tabel diskon lamanya.
#     """
#     if not moved_from or moved_from == target_key:
#         return
#     src_tables = promo_groups.get(sheet_name, {})
#     src_key = resolve_table_key(promo_groups, sheet_name, moved_from, category_suffix)
#     if src_key not in src_tables:
#         # Fallback: coba key polos (kompatibilitas utk tabel tanpa kategori, mis. SP)
#         if moved_from in src_tables:
#             src_key = moved_from
#         else:
#             return
#     src_table = src_tables[src_key]
#     kept = []
#     for existing_row in src_table.rows:
#         art_match = existing_row.article.upper() == art_code
#         col_match = existing_row.colour.upper() == col_key
#         all_colour_match = also_remove_all_colour and existing_row.colour.upper() == 'ALL COLOUR'
#         if art_match and (col_match or all_colour_match):
#             print(f"  [MOVE] {art_code}/{existing_row.colour}: {moved_from} -> {target_key}")
#         else:
#             kept.append(existing_row)
#     src_table.rows = kept


# def _find_source_tables(promo_groups: dict, sheet_name: str,
#                         category_suffix: Optional[str], art_code: str,
#                         exclude_key: Optional[str] = None) -> list:
#     """
#     Cari SEMUA tabel promo di sheet_name (dengan kategori yang cocok --
#     category_suffix sama, atau tabel tanpa kategori seperti SP yang dipakai
#     bersama semua kategori) yang SAAT INI benar-benar menyimpan baris untuk
#     art_code.

#     FIX PENTING: sebelumnya kode menentukan tabel asal artikel (moved_from)
#     dengan MENEBAK dari field 'Promotion' di database (GA_LIBREART5). Kalau
#     field itu kosong/'NML'/tidak sinkron dengan kondisi sebenarnya di file
#     Excel -- yang ternyata sangat sering terjadi -- maka penghapusan dari
#     tabel lama gagal total secara diam-diam: artikel tetap nangkring di
#     tabel lama SEKALIGUS ditambahkan ke tabel baru (duplikat).

#     Dengan mencari langsung ke isi hasil scan Excel (ground truth), proses
#     hapus/pindah artikel tidak lagi bergantung pada field database yang bisa
#     saja sudah usang atau tidak pernah di-update.
#     """
#     found = []
#     for key, table in promo_groups.get(sheet_name, {}).items():
#         if key == exclude_key:
#             continue
#         if table.category_suffix not in (category_suffix, None):
#             continue
#         if any(r.article.upper() == art_code for r in table.rows):
#             found.append((key, table))
#     return found


# def apply_input(promo_groups: dict, input_rows: list, query_data: dict):
#     for instr in input_rows:
#         art_code   = instr['item_code']
#         in_colour  = instr['colour']          # None = all colour
#         target_key = instr['item_group_menjadi']

#         q_records = query_data.get(art_code, {})
#         if not q_records:
#             print(f"  [WARN] {art_code}: tidak ditemukan di hasil query SQL, dilewati.")
#             continue

#         # Ambil dept/segment dari sample pertama (sama untuk semua warna 1 artikel)
#         sample    = next(iter(q_records.values()))
#         dept      = sample['dept']
#         segment   = sample['segment']

#         sheet_name, suffix = get_sheet_and_suffix(dept, segment)
#         if sheet_name is None:
#             print(f"  [WARN] {art_code}: dept={dept} segment={segment} tidak ada di SHEET_MAP, dilewati.")
#             continue

#         if sheet_name not in promo_groups:
#             promo_groups[sheet_name] = {}

#         # Cari/buat tabel target dengan key gabungan (discount_key + kategori)
#         # dulu, supaya "D50" milik EB MENS BAG tidak salah nyambung ke tabel
#         # "D50" milik EB LADIES BAG yang kebetulan sudah ada duluan di sheet
#         # yang sama. Kalau tidak ketemu, coba key polos (utk kompatibilitas
#         # dgn tabel yang memang dipakai bersama semua kategori, mis. tabel
#         # SP di sheet HANDBAG yang tidak dipisah Mens/Ladies). Kalau dua-duanya
#         # tidak ada, baru buat tabel baru dengan key gabungan.
#         combo_key = resolve_table_key(promo_groups, sheet_name, target_key, suffix)
#         if combo_key in promo_groups[sheet_name]:
#             target_store_key = combo_key
#         elif target_key in promo_groups[sheet_name]:
#             target_store_key = target_key
#         else:
#             target_store_key = combo_key
#             print(f"  [INFO] Membuat tabel promo baru: sheet={sheet_name} key={target_store_key}")
#             promo_groups[sheet_name][target_store_key] = PromoTable(
#                 sheet=sheet_name,
#                 promo_name=_build_promo_name(target_key, suffix),
#                 discount_key=target_key,
#                 category_suffix=suffix,
#                 rows=[],
#             )

#         target_table = promo_groups[sheet_name][target_store_key]

#         # Tentukan warna yang diproses
#         colours_to_process = list(q_records.keys()) if in_colour is None else [in_colour]

#         # Cek apakah semua warna yang diproses punya harga sama (untuk All Colour collapse)
#         prices_to_check = [q_records[c]['retail_price'] for c in colours_to_process if c in q_records]
#         all_same_price  = len(set(prices_to_check)) == 1 if prices_to_check else False

#         # FIX: cari tabel sumber SEBENARNYA langsung dari isi hasil scan
#         # (ground truth), bukan dari field 'Promotion' di database.
#         source_tables = _find_source_tables(promo_groups, sheet_name, suffix,
#                                             art_code, exclude_key=target_store_key)

#         # ----------------------------------------------------------------
#         # KASUS A: Instruksi All Colour (in_colour is None) + harga sama
#         # Collapse semua warna menjadi satu row "All Colour" di target.
#         # ----------------------------------------------------------------
#         if in_colour is None and all_same_price:
#             harga_awal = prices_to_check[0]

#             # Jika harga masih 0, warn dan skip
#             if not harga_awal:
#                 print(f"  [WARN] {art_code}/All Colour: harga_awal=0, tidak dapat dihitung. Dilewati.")
#                 continue

#             harga_diskon = compute_harga_diskon(harga_awal, target_key)

#             # Hapus artikel dari SEMUA tabel sumber yang benar-benar
#             # ditemukan berisi artikel ini (bisa >1 kalau sebelumnya
#             # sempat kepencar/duplikat), lalu catat dari tabel mana saja
#             # ia berasal (utk moved_from / pewarnaan legend).
#             moved_from_labels = set()
#             for src_key, src_table in source_tables:
#                 before_n = len(src_table.rows)
#                 src_table.rows = [r for r in src_table.rows if r.article.upper() != art_code]
#                 removed_n = before_n - len(src_table.rows)
#                 if removed_n:
#                     moved_from_labels.add(src_table.discount_key)
#                     print(f"  [MOVE] {art_code}: {src_key} -> {target_store_key} ({removed_n} baris dihapus)")

#             if len(moved_from_labels) == 0:
#                 all_colour_moved_from = None
#             elif len(moved_from_labels) == 1:
#                 all_colour_moved_from = next(iter(moved_from_labels))
#             else:
#                 all_colour_moved_from = ','.join(sorted(moved_from_labels))

#             # Bersihkan SEMUA baris artikel ini yang sudah ada di tabel
#             # TARGET sendiri (per-warna maupun All Colour), termasuk kasus di
#             # mana artikel ini sebenarnya sudah berada di tabel target sejak
#             # awal, supaya tidak terjadi duplikasi dengan row "All Colour" baru.
#             before_count = len(target_table.rows)
#             target_table.rows = [
#                 r for r in target_table.rows if r.article.upper() != art_code
#             ]
#             removed_dupe = before_count - len(target_table.rows)
#             if removed_dupe:
#                 print(f"  [DEDUPE] {art_code}: {removed_dupe} row lama di {target_store_key} "
#                       f"dihapus, digabung jadi All Colour")

#             moved_to_val = target_key if all_colour_moved_from else None
#             target_table.rows.append(PromoRow(
#                 article=sample['article'],
#                 colour='All Colour',
#                 harga_awal=harga_awal,
#                 harga_diskon=harga_diskon,
#                 moved_from=all_colour_moved_from,
#                 moved_to=moved_to_val,
#             ))
#             print(f"  [ADD ] {art_code}/All Colour -> {target_store_key} "
#                   f"harga_awal={harga_awal} harga_diskon={harga_diskon}"
#                   + (f" [dari {all_colour_moved_from}]" if all_colour_moved_from else ""))
#             continue

#         # ----------------------------------------------------------------
#         # KASUS B: Instruksi per warna tertentu (in_colour is not None),
#         # atau instruksi All Colour tapi harga berbeda-beda (expand per warna).
#         # ----------------------------------------------------------------
#         for col_key in colours_to_process:
#             rec = q_records.get(col_key)
#             if rec is None:
#                 print(f"  [WARN] {art_code} colour={col_key}: tidak ditemukan di query, dilewati.")
#                 continue

#             harga_awal    = rec['retail_price']
#             actual_colour = rec['colour']

#             # FIX: cari tabel sumber utk warna ini dari ground truth (isi
#             # tabel sebenarnya), bukan dari field 'Promotion' di database.
#             src_match = None
#             for src_key, src_table in source_tables:
#                 if any(r.article.upper() == art_code and
#                        r.colour.upper() in (col_key, 'ALL COLOUR')
#                        for r in src_table.rows):
#                     src_match = (src_key, src_table)
#                     break

#             harga_fallback = None
#             moved_from = None
#             if src_match:
#                 src_key, src_table = src_match
#                 moved_from = src_table.discount_key
#                 has_all_colour = any(
#                     r.article.upper() == art_code and r.colour.upper() == 'ALL COLOUR'
#                     for r in src_table.rows
#                 )
#                 if has_all_colour:
#                     # Warna lain (di luar instruksi ini) yang berasal dari
#                     # tabel sumber yang sama harus tetap "expand" kembali ke
#                     # tabel sumber, bukan ikut terhapus.
#                     colours_moving_from_this_src = {
#                         c for c in colours_to_process
#                         if any(r.article.upper() == art_code and
#                                r.colour.upper() in (c, 'ALL COLOUR')
#                                for r in src_table.rows)
#                     }
#                     harga_fallback = _expand_all_colour_row(
#                         src_table, art_code, q_records,
#                         exclude_colours=colours_moving_from_this_src,
#                         promo_groups=promo_groups,
#                         sheet_name=sheet_name,
#                     )
#                 else:
#                     before_n = len(src_table.rows)
#                     src_table.rows = [
#                         r for r in src_table.rows
#                         if not (r.article.upper() == art_code and r.colour.upper() == col_key)
#                     ]
#                     if len(src_table.rows) != before_n:
#                         print(f"  [MOVE] {art_code}/{actual_colour}: {src_key} -> {target_store_key}")

#             # Gunakan harga fallback dari row All Colour jika query memberikan 0
#             if (harga_awal is None or harga_awal == 0) and harga_fallback:
#                 harga_awal = harga_fallback

#             # Jika masih 0 setelah semua fallback, skip dan warn
#             if not harga_awal:
#                 print(f"  [WARN] {art_code}/{actual_colour}: harga_awal=0, tidak dapat dihitung. Dilewati.")
#                 continue

#             harga_diskon = compute_harga_diskon(harga_awal, target_key)
#             moved_to = target_key if moved_from else None

#             # Tambah ke tabel target
#             already = any(r.article.upper() == art_code and r.colour.upper() == actual_colour.upper()
#                           for r in target_table.rows)
#             if already:
#                 print(f"  [SKIP] {art_code}/{actual_colour}: sudah ada di {target_store_key}")
#                 continue

#             target_table.rows.append(PromoRow(
#                 article=rec['article'],
#                 colour=actual_colour,
#                 harga_awal=harga_awal,
#                 harga_diskon=harga_diskon,
#                 moved_from=moved_from,
#                 moved_to=moved_to,
#             ))
#             print(f"  [ADD ] {art_code}/{actual_colour} -> {target_store_key} "
#                   f"harga_awal={harga_awal} harga_diskon={harga_diskon}"
#                   + (f" [dari {moved_from}]" if moved_from else ""))

# # ---------------------------------------------------------------------------
# # Excel output
# # ---------------------------------------------------------------------------
# import math
# from datetime import date
# from openpyxl import Workbook
# from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
# from openpyxl.utils import get_column_letter

# _XL_COLS_PER_GROUP = 4   # Article | Colour | Harga Awal | Harga Diskon
# _XL_GAP_COLS       = 1
# _XL_GROUPS_PER_ROW = 3
# _XL_TOTAL_COLS     = _XL_GROUPS_PER_ROW * (_XL_COLS_PER_GROUP + _XL_GAP_COLS) - _XL_GAP_COLS  # 14
# _XL_GROUP_WIDTHS   = [20, 16, 14, 14, 2]

# _XC = {
#     'title_bg':  '1F3864', 'title_fg':  'FFFFFF',
#     'meta_bg':   'D6E4F0', 'meta_fg':   '000000',
#     'promo_bg':  '2E75B6', 'promo_fg':  'FFFFFF',
#     'hdr_bg':    'BDD7EE', 'hdr_fg':    '000000',
#     'row_odd':   'FFFFFF', 'row_even':  'F2F2F2',
# }

# # Warna baris berdasarkan promo asal (moved_from)
# MOVE_COLOURS: dict = {
#     'NML':  'FFFF00',   # kuning
#     'D20':  'FF9900',   # oranye
#     'D30':  '00B0F0',   # biru muda
#     'D40':  '92D050',   # hijau
#     'D50':  'FF0000',   # merah
#     'SP':   '7030A0',   # ungu (semua SP*)
# }

# def _xl_fill(hex_code: str) -> PatternFill:
#     return PatternFill('solid', start_color=hex_code, end_color=hex_code)

# def _xl_border() -> Border:
#     s = Side(style='thin')
#     return Border(top=s, bottom=s, left=s, right=s)

# def _xl_write_sheet(wb: Workbook, sheet_name: str, tables: dict, periode: str):
#     ws = wb.create_sheet(title=sheet_name)

#     col_widths = (_XL_GROUP_WIDTHS * _XL_GROUPS_PER_ROW)[:_XL_TOTAL_COLS]
#     for i, w in enumerate(col_widths, 1):
#         ws.column_dimensions[get_column_letter(i)].width = w

#     row = 1

#     # Judul
#     ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=_XL_TOTAL_COLS)
#     c = ws.cell(row=row, column=1, value="PROMLIST CORNER SALE EVERBEST")
#     c.font      = Font(name='Arial', bold=True, size=14, color=_XC['title_fg'])
#     c.fill      = _xl_fill(_XC['title_bg'])
#     c.alignment = Alignment(horizontal='center', vertical='center')
#     ws.row_dimensions[row].height = 22
#     row += 1

#     # BERLAKU DI
#     ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=2)
#     c = ws.cell(row=row, column=1, value="BERLAKU DI")
#     c.font = Font(name='Arial', bold=True, size=10); c.fill = _xl_fill(_XC['meta_bg'])
#     c.alignment = Alignment(horizontal='left', vertical='center')
#     ws.merge_cells(start_row=row, start_column=3, end_row=row, end_column=_XL_TOTAL_COLS)
#     c = ws.cell(row=row, column=3, value=": ALL STORE, COUNTER, BAZAR, WEBSITE, MARKETPLACE")
#     c.font = Font(name='Arial', size=10); c.fill = _xl_fill(_XC['meta_bg'])
#     c.alignment = Alignment(horizontal='left', vertical='center')
#     ws.row_dimensions[row].height = 16
#     row += 1

#     # PERIODE
#     ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=2)
#     c = ws.cell(row=row, column=1, value="PERIODE")
#     c.font = Font(name='Arial', bold=True, size=10); c.fill = _xl_fill(_XC['meta_bg'])
#     c.alignment = Alignment(horizontal='left', vertical='center')
#     ws.merge_cells(start_row=row, start_column=3, end_row=row, end_column=_XL_TOTAL_COLS)
#     c = ws.cell(row=row, column=3, value=f": {periode}")
#     c.font = Font(name='Arial', size=10); c.fill = _xl_fill(_XC['meta_bg'])
#     c.alignment = Alignment(horizontal='left', vertical='center')
#     ws.row_dimensions[row].height = 16
#     row += 1

#     row += 1  # baris kosong

#     hdr_labels = ['Article', 'Colour', 'Harga Awal', 'Harga Diskon']

#     for tbl in tables.values():
#         rows_data = tbl.rows if hasattr(tbl, 'rows') else tbl.get('rows', [])
#         if not rows_data:
#             continue

#         promo_name = tbl.promo_name if hasattr(tbl, 'promo_name') else tbl.get('promo_name', '')

#         # Header nama promo
#         ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=_XL_TOTAL_COLS)
#         c = ws.cell(row=row, column=1, value=promo_name.upper())
#         c.font      = Font(name='Arial', bold=True, size=11, color=_XC['promo_fg'])
#         c.fill      = _xl_fill(_XC['promo_bg'])
#         c.alignment = Alignment(horizontal='center', vertical='center')
#         c.border    = _xl_border()
#         ws.row_dimensions[row].height = 18
#         row += 1

#         # Header kolom
#         for g in range(_XL_GROUPS_PER_ROW):
#             base = g * (_XL_COLS_PER_GROUP + _XL_GAP_COLS) + 1
#             for ci, label in enumerate(hdr_labels):
#                 c = ws.cell(row=row, column=base + ci, value=label)
#                 c.font      = Font(name='Arial', bold=True, size=9, color=_XC['hdr_fg'])
#                 c.fill      = _xl_fill(_XC['hdr_bg'])
#                 c.alignment = Alignment(horizontal='center', vertical='center')
#                 c.border    = _xl_border()
#         ws.row_dimensions[row].height = 15
#         row += 1

#         # Baris data
#         n         = len(rows_data)
#         n_per_col = math.ceil(n / _XL_GROUPS_PER_ROW)

#         for r_idx in range(n_per_col):
#             base_fill = _XC['row_odd'] if r_idx % 2 == 0 else _XC['row_even']
#             for g in range(_XL_GROUPS_PER_ROW):
#                 data_idx = g * n_per_col + r_idx
#                 base     = g * (_XL_COLS_PER_GROUP + _XL_GAP_COLS) + 1
#                 if data_idx < n:
#                     d = rows_data[data_idx]
#                     if hasattr(d, 'article'):   # PromoRow dataclass
#                         art  = d.article
#                         col  = d.colour
#                         ha   = d.harga_awal
#                         hd   = int(d.harga_diskon) if d.harga_diskon is not None else None
#                         mf   = d.moved_from
#                     else:                        # dict
#                         art  = d['article']
#                         col  = d['colour']
#                         ha   = d['harga_awal']
#                         hd   = d['harga_diskon']
#                         mf   = d.get('moved_from')
#                     vals = [art, col, ha, hd]
#                     # Warna baris: hanya artikel yang dipindahkan (moved_from terisi)
#                     if mf:
#                         # Cari warna: cek exact key dulu, lalu prefix 'SP'
#                         row_color = MOVE_COLOURS.get(mf)
#                         if row_color is None:
#                             row_color = MOVE_COLOURS.get(mf[:2]) if len(mf) >= 2 else None
#                         row_color = row_color or base_fill
#                     else:
#                         row_color = base_fill
#                     # Bold hanya pada kolom Colour (ci=1) jika bukan All Colour
#                     is_specific = col.strip().upper() != 'ALL COLOUR'
#                     for ci, v in enumerate(vals):
#                         c = ws.cell(row=row, column=base + ci, value=v)
#                         c.font      = Font(name='Arial', size=9,
#                                           bold=(ci == 1 and is_specific))
#                         c.fill      = _xl_fill(row_color)
#                         c.alignment = Alignment(
#                             horizontal='center' if ci >= 2 else 'left',
#                             vertical='center',
#                         )
#                         c.border = _xl_border()
#                         if ci >= 2 and v is not None:
#                             c.number_format = '#,##0'
#                 else:
#                     for ci in range(_XL_COLS_PER_GROUP):
#                         c = ws.cell(row=row, column=base + ci)
#                         c.fill   = _xl_fill(base_fill)
#                         c.border = _xl_border()
#             ws.row_dimensions[row].height = 14
#             row += 1

#         row += 1  # baris kosong antar tabel

#     ws.freeze_panes = 'A5'

#     # ── Legend warna: baris vertikal di kolom 16 ─────────────────────────
#     leg_col = _XL_TOTAL_COLS + 2   # kolom 16
#     ws.column_dimensions[get_column_letter(leg_col)].width = 24

#     today_str = str(date.today().day) + date.today().strftime(' %B %Y')
#     legend_entries = [
#         ('NML',  f'PER {today_str} dari NML'),
#         ('D20',  f'PER {today_str} dari D20'),
#         ('D30',  f'PER {today_str} dari D30'),
#         ('D40',  f'PER {today_str} dari D40'),
#         ('D50',  f'PER {today_str} dari D50'),
#         ('SP',   f'PER {today_str} dari SP'),
#     ]

#     for leg_row, (key, label) in enumerate(legend_entries, start=1):
#         bg = MOVE_COLOURS[key]
#         c  = ws.cell(row=leg_row, column=leg_col, value=label)
#         c.font      = Font(name='Arial', size=9, bold=True)
#         c.fill      = _xl_fill(bg)
#         c.alignment = Alignment(horizontal='left', vertical='center')
#         c.border    = _xl_border()
#         ws.row_dimensions[leg_row].height = 16

# # ---------------------------------------------------------------------------
# # Copy sheet khusus (SP COUNTER MP) dari template ke workbook output
# # ---------------------------------------------------------------------------
# SPECIAL_SHEETS = ['SP COUNTER MP']  # tambahkan nama sheet lain di sini jika perlu

# def copy_special_sheets(template_path: str, dst_wb: Workbook, sheet_names: list = None) -> list:
#     """
#     Mengecek file template untuk sheet khusus (misal: 'SP COUNTER MP') dan
#     menyalin sheet tersebut apa adanya (data + format) ke workbook output.

#     Parameter
#     ---------
#     template_path : str
#         Path ke file template Excel sumber.
#     dst_wb : openpyxl.Workbook
#         Workbook output yang sedang dibangun (sebelum di-save).
#     sheet_names : list[str], optional
#         Daftar nama sheet yang ingin disalin. Default: SPECIAL_SHEETS.

#     Return
#     ------
#     list[str] : nama-nama sheet yang berhasil disalin.
#     """
#     names = sheet_names or SPECIAL_SHEETS
#     copied = []

#     src_wb = openpyxl.load_workbook(template_path, data_only=False)

#     for name in names:
#         if name not in src_wb.sheetnames:
#             continue

#         src_ws = src_wb[name]
#         dst_ws = dst_wb.create_sheet(title=name[:31])

#         # Column widths & hidden state
#         for col_letter, dim in src_ws.column_dimensions.items():
#             dst_ws.column_dimensions[col_letter].width = dim.width
#             dst_ws.column_dimensions[col_letter].hidden = dim.hidden

#         # Row heights
#         for row_idx, dim in src_ws.row_dimensions.items():
#             dst_ws.row_dimensions[row_idx].height = dim.height

#         # Cell values + full style
#         for row in src_ws.iter_rows():
#             for cell in row:
#                 new_cell = dst_ws.cell(row=cell.row, column=cell.column, value=cell.value)
#                 if cell.has_style:
#                     new_cell.font = copy(cell.font)
#                     new_cell.fill = copy(cell.fill)
#                     new_cell.border = copy(cell.border)
#                     new_cell.alignment = copy(cell.alignment)
#                     new_cell.number_format = cell.number_format
#                     new_cell.protection = copy(cell.protection)

#         # Merged cells
#         for merged_range in src_ws.merged_cells.ranges:
#             dst_ws.merge_cells(str(merged_range))

#         # View properties
#         dst_ws.freeze_panes = src_ws.freeze_panes
#         dst_ws.sheet_view.showGridLines = src_ws.sheet_view.showGridLines

#         copied.append(name)

#     return copied


# def write_promolist_excel(promo_groups: dict, output_path: str, template_path: str = None):
#     """
#     Buat file Excel promolist corner sale dari promo_groups.

#     Menerima promo_groups langsung dari apply_input() / recount() —
#     berisi PromoTable dataclass maupun dict, keduanya didukung.

#     Parameter
#     ---------
#     promo_groups : dict
#         { sheet_name: { discount_key: PromoTable } }
#     output_path  : str
#         Path file .xlsx yang akan dibuat.
#     template_path : str, optional
#         Path ke file template Excel sumber. Jika ada sheet khusus
#         (lihat SPECIAL_SHEETS), sheet tersebut akan disalin apa adanya
#         ke file output.
#     """
#     today = date.today()
#     periode = str(today.day) + today.strftime(' %B %Y') + ' - 31 DEC 2099'

#     wb = Workbook()
#     wb.remove(wb.active)

#     for sheet_name, tables in promo_groups.items():
#         safe_name = re.sub(r'[\\/*?:\[\]]', '-', sheet_name)[:31]
#         _xl_write_sheet(wb, safe_name, tables, periode)

#     if template_path:
#         copied = copy_special_sheets(template_path, wb)
#         if copied:
#             print(f"  [EXCEL] Sheet khusus disalin: {', '.join(copied)}")

#     wb.save(output_path)


from typing import Optional
from decimal import Decimal
from copy import copy
import openpyxl,re,pyodbc
from class_promo import PromoRow,PromoTable

SHEET_MAP = {
    ('LDS', 'SHS'): ('LADIES',           'EBL'),
    ('LDS', 'BAG'): ('HANDBAG',          'EB LADIES BAG'),
    ('LDS', 'BGC'): ('HANDBAG',          'EB LADIES BAG'),
    ('MEN', 'SHS'): ('MENS',             'EBM'),
    ('MEN', 'BAG'): ('HANDBAG',          'EB MENS BAG'),
    ('LDS', 'SLG'): ('SUNDRIES',         'EBSL'),
    ('LDS', 'WAL'): ('SUNDRIES',         'EBSL'),
    ('LDS', 'ACC'): ('SUNDRIES',         'EBSL'),
    ('LDS', 'WRS'): ('SUNDRIES',         'EBSL'),
    ('LDS', 'FWC'): ('SUNDRIES',         'EBSL'),
    ('MEN', 'SLG'): ('SUNDRIES',         'EBSM'),
    ('MEN', 'WAL'): ('SUNDRIES',         'EBSM'),
    ('MEN', 'BLT'): ('SUNDRIES',         'EBSM'),
    ('MEN', 'ACC'): ('SUNDRIES',         'EBSM'),
    ('MEN', 'WRS'): ('SUNDRIES',         'EBSM'),
    ('LDS', 'SHC'): ('SHOECARE & SOCKS', 'SHOECARE'),
    ('MEN', 'SHC'): ('SHOECARE & SOCKS', 'SHOECARE'),
    ('MEN', 'SCK'): ('SHOECARE & SOCKS', 'SOCKS'),
}
 
# Kode brand dari file input -> nama brand di judul "PROMLIST CORNER SALE {BRAND}"
BRAND_MAP = {
    'EB': 'EVERBEST',
    'TR': 'TRACCE',
    'VB': 'EVB',
}
DEFAULT_BRAND = 'EB'

def brand_full_name(brand_code: str) -> str:
    """'EB' -> 'EVERBEST', 'TR' -> 'TRACCE', 'VB' -> 'EVB'. Kode tak dikenal dikembalikan apa adanya."""
    code = (brand_code or DEFAULT_BRAND).strip().upper()
    return BRAND_MAP.get(code, code)

# Sheet yang header tabel promonya ikut kode brand input (SP99K EBL -> SP99K TRL).
# Sheet lain (SUNDRIES, SHOECARE & SOCKS, ...) header-nya dibiarkan apa adanya.
BRAND_HEADER_SHEETS = {'LADIES', 'MENS', 'HANDBAG'}

def apply_brand_to_header(promo_name: str, brand_code: str) -> str:
    """
    Ganti prefix brand 'EB' di header tabel promo dengan kode brand input,
    mis. brand 'TR':
        'SP99K EBL'            -> 'SP99K TRL'
        'SP149K EBM'           -> 'SP149K TRM'
        'SP199K EB LADIES BAG' -> 'SP199K TR LADIES BAG'
    """
    code = (brand_code or DEFAULT_BRAND).strip().upper()
    return re.sub(r'\bEB(L|M)?\b', lambda m: code + (m.group(1) or ''),
                  promo_name, flags=re.IGNORECASE)

def get_sheet_and_suffix(dept: str, segment: str):
    """Return (sheet_name, suffix) atau (None, None) jika tidak ada di mapping."""
    return SHEET_MAP.get((dept.strip().upper(), segment.strip().upper()), (None, None))

def sort_rows(rows: list) -> list:
    return sorted(rows, key=lambda r: (r.article.upper(), r.colour.upper()))
 
def is_promo_header_row(row_vals: list) -> bool:
    texts = [str(v).strip() if v else '' for v in row_vals]
    return any('article' in t.lower() for t in texts) and \
           any('colour' in t.lower() for t in texts)

def detect_column_groups(row_vals: list) -> list:
    groups = []
    for i, val in enumerate(row_vals):
        if val and str(val).strip().lower() == 'article':
            try:
                c_colour = next(
                    j for j in range(i + 1, min(i + 6, len(row_vals)))
                    if row_vals[j] and 'colour' in str(row_vals[j]).lower()
                )
                c_awal = next(
                    j for j in range(c_colour + 1, min(c_colour + 4, len(row_vals)))
                    if row_vals[j] and 'harga' in str(row_vals[j]).lower() and 'awal' in str(row_vals[j]).lower()
                )
                c_diskon = next(
                    j for j in range(c_awal + 1, min(c_awal + 4, len(row_vals)))
                    if row_vals[j] and 'harga' in str(row_vals[j]).lower() and 'diskon' in str(row_vals[j]).lower()
                )
                groups.append((i, c_colour, c_awal, c_diskon))
            except StopIteration:
                pass
    return groups

def extract_numeric(val) -> Optional[int]:
    if val is None:
        return None
    if isinstance(val, (int, float)):
        return int(val)
    if isinstance(val, str):
        # If it's a formula, skip
        if val.startswith('='):
            return None
        # Try stripping non-numeric chars
        cleaned = re.sub(r'[^\d]', '', val)
        return int(cleaned) if cleaned else None
    return None

def make_discount_key(promo_name: str) -> str:
    m = re.search(r'\bSP(\d+K?)\b', promo_name, re.IGNORECASE)
    if m:
        return 'SP' + m.group(1).upper()

    m = re.search(r'\bD(\d+)\b', promo_name, re.IGNORECASE)
    if m:
        return 'D' + m.group(1)

    m = re.search(r'(\d+)\s*%', promo_name)
    if m:
        return 'D' + m.group(1)

    return promo_name.split()[0].upper()

def is_promo_name_cell(val) -> bool:
    if not val or not isinstance(val, str):
        return False
    v = val.strip()
    if not v:
        return False
    patterns = [
        r'DISC[\.\s]*\d+',
        r'D\d+\s+(SOCKS|SHOECARE)',
        r'SP\d+K?\s',
        r'SP\d+K?$',
    ]
    return any(re.search(p, v, re.IGNORECASE) for p in patterns)

# ---------------------------------------------------------------------------
# FIX: kategori tabel (Mens Bag vs Ladies Bag, dst) harus ikut membedakan
# tabel promo, bukan cuma persentase diskonnya.
#
# Sebelumnya, discount_key hanya berupa 'D40', 'D50', dst -- diambil murni
# dari angka persen di nama promo. Padahal dalam SATU sheet (contoh: HANDBAG)
# bisa ada beberapa tabel dengan persen yang SAMA tapi kategori BERBEDA,
# misalnya "DISC. 40% EB LADIES BAG" dan "DISC.40% EB MENS BAG" -- keduanya
# menghasilkan discount_key 'D40'. Karena LADIES BAG di-scan lebih dulu
# (posisinya di atas dalam sheet), dia yang dapat key 'D40' yang bersih,
# sedangkan MENS BAG kebagian key 'D40_2' hasil dari mekanisme
# anti-tabrakan. Akibatnya, saat artikel MENS BAG dipindah discount-nya,
# proses hapus-dari-tabel-lama dan taruh-ke-tabel-baru mencari/menaruh
# berdasarkan key polos ('D40'/'D50') yang justru menunjuk ke tabel
# LADIES BAG -- artikel pun nyasar ke tabel LADIES.
#
# Fix: setiap tabel diberi 'category_suffix' (mis. 'EB MENS BAG',
# 'EB LADIES BAG', 'EBL', 'EBSM', dst) yang diekstrak dari teks promo_name.
# Key yang dipakai untuk MENYIMPAN tabel di dalam dict digabung dari
# discount_key + category_suffix, sehingga dua tabel dengan persen sama
# tapi kategori beda tidak akan pernah dianggap tabel yang sama.
# ---------------------------------------------------------------------------
CATEGORY_SUFFIXES = sorted(
    {suffix for (_dept, _seg), (_sheet, suffix) in SHEET_MAP.items()},
    key=len, reverse=True,
)

def extract_category_suffix(promo_name: str) -> Optional[str]:
    """
    Cari suffix kategori (mis. 'EB MENS BAG', 'EBL', 'EBSM', 'SHOECARE', ...)
    yang tercantum dalam teks promo_name. Dicocokkan dari yang paling
    panjang/spesifik dulu supaya tidak salah tangkap suffix yang lebih pendek.
    Return None jika promo_name tidak mengandung suffix kategori apa pun
    (mis. tabel SP yang dipakai bersama utk semua kategori dalam 1 sheet).
    """
    name_up = promo_name.upper()
    for suf in CATEGORY_SUFFIXES:
        if suf in name_up:
            return suf
    return None


def _table_store_key(discount_key: str, category_suffix: Optional[str]) -> str:
    """Key gabungan yang dipakai sebagai key dict promo_groups[sheet][...]."""
    return f"{discount_key}::{category_suffix}" if category_suffix else discount_key


def resolve_table_key(promo_groups: dict, sheet_name: str,
                      discount_key: str, category_suffix: Optional[str]) -> str:
    """
    Tentukan key dict yang tepat untuk mengakses/menyimpan tabel promo di
    promo_groups[sheet_name], dengan mempertimbangkan category_suffix supaya
    tabel dengan discount_key sama tapi kategori berbeda (mis. Mens Bag vs
    Ladies Bag) tidak saling tertukar. Dipakai baik saat mencari tabel
    SUMBER (utk dihapus) maupun tabel TUJUAN (utk ditambahkan/dibuat).
    """
    return _table_store_key(discount_key, category_suffix)


def scan_excel(filepath: str) -> dict:
    wb = openpyxl.load_workbook(filepath, data_only=True)
    result = {}

    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        sheet_result = {}

        all_rows = []
        for row in ws.iter_rows(values_only=True):
            all_rows.append(list(row))

        # Walk rows to find promo sections
        i = 0
        while i < len(all_rows):
            row = all_rows[i]
            non_none = [v for v in row if v is not None]

            # Detect promo section header
            first_val = next((v for v in row if v is not None), None)
            if first_val and is_promo_name_cell(str(first_val).strip()):
                promo_name = str(first_val).strip()
                # Clean newlines in name
                promo_name = promo_name.replace('\n', ' ').strip()
                discount_key    = make_discount_key(promo_name)
                category_suffix = extract_category_suffix(promo_name)

                # Find the column-header row (next row with 'Article')
                header_row_idx = None
                for j in range(i + 1, min(i + 4, len(all_rows))):
                    if is_promo_header_row(all_rows[j]):
                        header_row_idx = j
                        break

                if header_row_idx is None:
                    i += 1
                    continue

                col_groups = detect_column_groups(all_rows[header_row_idx])
                if not col_groups:
                    i += 1
                    continue

                # FIX: key penyimpanan digabung dengan category_suffix supaya
                # tabel dengan persen sama tapi kategori beda (mis. "DISC.40%
                # EB MENS BAG" vs "DISC. 40% EB LADIES BAG") tidak dianggap
                # tabel yang sama / saling menimpa. Fallback "_2", "_3", ...
                # tetap dipertahankan sebagai jaring pengaman untuk kasus
                # duplikat sungguhan (persen SAMA + kategori SAMA).
                unique_key = _table_store_key(discount_key, category_suffix)
                dupe_suffix = 2
                while unique_key in sheet_result:
                    unique_key = f"{_table_store_key(discount_key, category_suffix)}_{dupe_suffix}"
                    dupe_suffix += 1

                table = PromoTable(
                    sheet=sheet_name,
                    promo_name=promo_name,
                    discount_key=discount_key,          # tetap 'D40'/'D50'/'SP490K' yg bersih
                    category_suffix=category_suffix,    # 'EB MENS BAG' / 'EB LADIES BAG' / None
                    rows=[]
                )

                # Collect data rows until next promo header or end
                data_start = header_row_idx + 1
                k = data_start
                while k < len(all_rows):
                    data_row = all_rows[k]
                    non_none_dr = [v for v in data_row if v is not None]

                    # Stop if next promo name found
                    first_dr_val = next((v for v in data_row if v is not None), None)
                    if first_dr_val and is_promo_name_cell(str(first_dr_val).strip()):
                        break

                    # Stop if another column-header row found (different section)
                    if is_promo_header_row(data_row) and k != header_row_idx:
                        # Could be a new section starting horizontally in different col group
                        # Check if a NEW promo header was in a non-B column
                        new_groups = detect_column_groups(data_row)
                        if new_groups:
                            break

                    # Extract data from each column group
                    row_had_data = False
                    for (ci_art, ci_col, ci_awal, ci_dis) in col_groups:
                        art_val = data_row[ci_art] if ci_art < len(data_row) else None
                        col_val = data_row[ci_col] if ci_col < len(data_row) else None
                        awal_val = data_row[ci_awal] if ci_awal < len(data_row) else None
                        dis_val = data_row[ci_dis] if ci_dis < len(data_row) else None

                        if art_val and isinstance(art_val, str) and art_val.strip():
                            article = art_val.strip()
                            colour = str(col_val).strip() if col_val else ''
                            harga_awal = extract_numeric(awal_val)
                            harga_diskon = extract_numeric(dis_val)

                            # Skip if article looks like another promo header
                            if is_promo_name_cell(article):
                                continue

                            table.rows.append(PromoRow(
                                article=article,
                                colour=colour,
                                harga_awal=harga_awal,
                                harga_diskon=harga_diskon,
                            ))
                            row_had_data = True

                    k += 1

                if table.rows:
                    sheet_result[unique_key] = table
                    i = k  # jump past the consumed rows
                    continue

            i += 1

        if sheet_result:
            result[sheet_name] = sheet_result

    return result

# ---------------------------------------------------------------------------
# Output formatter
# ---------------------------------------------------------------------------

def format_harga(val) -> str:
    if val is None:
        return 'None'
    return str(int(val))

def write_output(data: dict, output_path: str):
    lines = []
    lines.append('promo_groups = {')

    sheet_names = list(data.keys())
    for s_idx, sheet_name in enumerate(sheet_names):
        tables = data[sheet_name]
        is_last_sheet = (s_idx == len(sheet_names) - 1)

        lines.append(f'   "{sheet_name}": {{')

        table_keys = list(tables.keys())
        for t_idx, key in enumerate(table_keys):
            table = tables[key]
            is_last_table = (t_idx == len(table_keys) - 1)

            sorted_rows = sort_rows(table.rows)
            count = len(sorted_rows)

            lines.append(f'        "{key}": PromoTable(')
            lines.append(f"            sheet='{table.sheet}',")
            lines.append(f"            promo_name='{table.promo_name}',")
            lines.append(f"            discount_key='{table.discount_key}',")
            lines.append(f"            count={count},")
            lines.append(f'            rows=[')

            for row in sorted_rows:
                article = row.article.replace("'", "\\'")
                colour = row.colour.replace("'", "\\'")
                ha = format_harga(row.harga_awal)
                hd = format_harga(row.harga_diskon)
                mf        = f"'{row.moved_from}'" if row.moved_from else 'None'
                mt        = f"'{row.moved_to}'"   if row.moved_to   else 'None'
                lines.append(
                    f"                PromoRow(article='{article}', "
                    f"colour='{colour}', "
                    f"harga_awal={ha}, "
                    f"harga_diskon={hd}, "
                    f"moved_from={mf}, moved_to={mt}),"
                )

            lines.append(f'            ]')
            lines.append(f'        ){","  if not is_last_table else ""}')

        lines.append(f'    }}{","  if not is_last_sheet else ""}')

    lines.append('}')

    with open(output_path, 'w', encoding='utf-8') as f:
        f.write('\n'.join(lines))
        f.write('\n')

    print(f"Output written to: {output_path}")

# ---------------------------------------------------------------------------
# Summary printer
# ---------------------------------------------------------------------------

def print_summary(data: dict):
    print("\n=== SCAN SUMMARY ===")
    total_tables = 0
    total_rows = 0
    for sheet_name, tables in data.items():
        print(f"\nSheet: {sheet_name}")
        for key, table in tables.items():
            print(f"  [{key}] {table.promo_name} -> {len(table.rows)} rows")
            total_tables += 1
            total_rows += len(table.rows)
    print(f"\nTotal sheets  : {len(data)}")
    print(f"Total tables  : {total_tables}")
    print(f"Total data rows: {total_rows}")

SERVER   = '149.129.249.131'
DATABASE = 'Y2_PROD'
USERNAME = 'sa'
PASSWORD = 'Eshoes09'

DB_CONNECTION = (
    f"DRIVER={{SQL Server}};"
    f"SERVER={SERVER};"
    f"DATABASE={DATABASE};"
    f"UID={USERNAME};"
    f"PWD={PASSWORD};"
    f"TrustServerCertificate=yes;"
)

# SQL_QUERY = """
# SELECT 
#   GA_CODEBARRE as Item_Barcode,
#   ISNULL(RP.GF_PRIXUNITAIRE,0) as Retail_Price,
#   GA_CODEARTICLE as Article,
#   GA_ARTICLE as Item_Identifier, 
#   GA_LIBREART5 as Promotion, 
#   GA_CODEARTICLE as Item_code,
#   GA_CHARLIBRE1 as Vendor_Style_Code, 
#   GA_FAMILLENIV1 as Brand, 
#   GA_FAMILLENIV2 as Department, 
#   GA_FAMILLENIV3 as Segment, 
#   GA2_FAMILLENIV4 as Category, 
#   GDI_LIBELLE as Colour_Desc,
#   GA_DIMMASQUE as Dim2, 
#   GDI_LIBELLE_1 as Size_Desc,
#   YX6.YX_LIBELLE as Material,
#   YX7.YX_LIBELLE as Upper_Style,
#   YX12.YX_LIBELLE as Sub_Upper_Style,
#   YX4.YX_LIBELLE as Toe_Style,
#   YX8.YX_LIBELLE as Heel_Height,
#   YX10.YX_LIBELLE as Heel_Style,
#   YX9.YX_LIBELLE as Wearing_Style,
#   YX11.YX_LIBELLE as Season,
#   0 as Basic_Price,
#   GA2_LIBREARTC as Category_Product,
#   GA_PAHT as Item_Purchase_Price_Tax_Excl,
#   YX15.YX_LIBELLE as Lining_Material,
#   YX14.YX_LIBELLE as Insole_Material,
#   YX16.YX_LIBELLE as Outsole_Material,
#   GA_CODEDOUANIER as Custom_Reference,
#   GA_TYPEARTICLE as Item_Type,
#   GA_PRHT as Item_Cost_Price_Tax_Excl,
#   0.00 as Formula,
#   GA_POIDSBRUT as Gross_Weight_Item,
#   GA_TENUESTOCK as In_Item_Stock,
#   GA_DATEMODIF as Item_Modification_Date,
#   GA_DATECREATION as Item_Creation_Date,
#   GA_CODEDIM2 as Size,
#   GA_EMBALLAGE as Packing_Item,
#   GA_UNITEQTEVTE as Sales_Unit_Quantity
# FROM 
#   ARTICLEDIM_MODE AM
#   OUTER APPLY (
#     SELECT TOP 1
#         RP.GF_PRIXUNITAIRE
#     FROM GCTARFCONMODEART RP
#     WHERE RP.GA_CODEARTICLE = AM.GA_CODEARTICLE
#       AND RP.GF_REGIMEPRIX = 'TTC'
#       AND RP.GFM_TYPETARIF = 'BSP01'
#       AND RP.GFM_PERTARIF IN ('00BASE', 'SGC1')
#       AND RP.GFM_NATURETYPE = 'VTE'
#       AND RP.GF_ARTICLE <> ''
#       AND GETDATE() BETWEEN RP.GF_DATEDEBUT AND RP.GF_DATEFIN
#     ORDER BY
#         RP.GF_DATEDEBUT DESC,
#         RP.GF_DATEMODIF DESC
# ) RP
#   LEFT OUTER JOIN CHOIXCOD CC1 ON GA_FAMILLENIV1 = CC1.CC_CODE AND CC1.CC_TYPE = 'FN1' 
#   LEFT OUTER JOIN CHOIXCOD CC2 ON GA_FAMILLENIV2 = CC2.CC_CODE AND CC2.CC_TYPE = 'FN2'
#   LEFT OUTER JOIN CHOIXCOD CC3 ON GA_FAMILLENIV3 = CC3.CC_CODE AND CC3.CC_TYPE = 'FN3' 
#   LEFT OUTER JOIN CHOIXEXT YX4 ON GA_LIBREART2 = YX4.YX_CODE AND YX4.YX_TYPE = 'LA2' 
#   LEFT OUTER JOIN CHOIXCOD CC5 ON GA2_FAMILLENIV4 = CC5.CC_CODE AND CC5.CC_TYPE = 'FN4' 
#   LEFT OUTER JOIN CHOIXEXT YX6 ON GA_LIBREART7 = YX6.YX_CODE AND YX6.YX_TYPE = 'LA7' 
#   LEFT OUTER JOIN CHOIXEXT YX7 ON GA_LIBREART1 = YX7.YX_CODE AND YX7.YX_TYPE = 'LA1' 
#   LEFT OUTER JOIN CHOIXEXT YX8 ON GA_LIBREART3 = YX8.YX_CODE AND YX8.YX_TYPE = 'LA3' 
#   LEFT OUTER JOIN CHOIXEXT YX9 ON GA_LIBREART4 = YX9.YX_CODE AND YX9.YX_TYPE = 'LA4' 
#   LEFT OUTER JOIN CHOIXEXT YX10 ON GA_LIBREART8 = YX10.YX_CODE AND YX10.YX_TYPE = 'LA8' 
#   LEFT OUTER JOIN CHOIXEXT YX11 ON GA_LIBREART9 = YX11.YX_CODE AND YX11.YX_TYPE = 'LA9' 
#   LEFT OUTER JOIN CHOIXEXT YX12 ON GA2_LIBREARTB = YX12.YX_CODE AND YX12.YX_TYPE = 'LAB' 
#   LEFT OUTER JOIN CHOIXEXT YX13 ON GA2_LIBREARTC = YX13.YX_CODE AND YX13.YX_TYPE = 'LAC' 
#   LEFT OUTER JOIN CHOIXEXT YX14 ON GA2_LIBREARTD = YX14.YX_CODE AND YX14.YX_TYPE = 'LAD' 
#   LEFT OUTER JOIN CHOIXEXT YX15 ON GA2_LIBREARTE = YX15.YX_CODE AND YX15.YX_TYPE = 'LAE' 
#   LEFT OUTER JOIN CHOIXEXT YX16 ON GA2_LIBREARTF = YX16.YX_CODE AND YX16.YX_TYPE = 'LAF' 
# WHERE 
#   GA_TENUESTOCK = 'X' 
#   AND GA_FAMILLENIV1 = ?
#   AND GA_EMBALLAGE <> 'X' 
#   AND GA_TYPEARTICLE IN ('MAR')
#   AND GA_TYPEARTICLE NOT IN ('PRE', 'FI', 'FRA')
#   AND GA_STATUTART = 'DIM'
# """

SQL_QUERY = """
SELECT
  A.GA_CODEBARRE                 AS Item_Barcode,
  ISNULL(RP.GF_PRIXUNITAIRE, 0)  AS Retail_Price,
  A.GA_CODEARTICLE               AS Article,
  A.GA_ARTICLE                   AS Item_Identifier,
  A.GA_LIBREART5                 AS Promotion,
  A.GA_CODEARTICLE               AS Item_code,
  A.GA_CHARLIBRE1                AS Vendor_Style_Code,
  A.GA_FAMILLENIV1               AS Brand,
  A.GA_FAMILLENIV2               AS Department,
  A.GA_FAMILLENIV3               AS Segment,
  A2.GA2_FAMILLENIV4             AS Category,
  D1.GDI_LIBELLE                 AS Colour_Desc,
  A.GA_DIMMASQUE                 AS Dim2,
  D2.GDI_LIBELLE                 AS Size_Desc,
  YX6.YX_LIBELLE                 AS Material,
  YX7.YX_LIBELLE                 AS Upper_Style,
  YX12.YX_LIBELLE                AS Sub_Upper_Style,
  YX4.YX_LIBELLE                 AS Toe_Style,
  YX8.YX_LIBELLE                 AS Heel_Height,
  YX10.YX_LIBELLE                AS Heel_Style,
  YX9.YX_LIBELLE                 AS Wearing_Style,
  YX11.YX_LIBELLE                AS Season,
  0                              AS Basic_Price,
  A2.GA2_LIBREARTC               AS Category_Product,
  A.GA_PAHT                      AS Item_Purchase_Price_Tax_Excl,
  YX15.YX_LIBELLE                AS Lining_Material,
  YX14.YX_LIBELLE                AS Insole_Material,
  YX16.YX_LIBELLE                AS Outsole_Material,
  A.GA_CODEDOUANIER              AS Custom_Reference,
  A.GA_TYPEARTICLE               AS Item_Type,
  A.GA_PRHT                      AS Item_Cost_Price_Tax_Excl,
  0.00                           AS Formula,
  A.GA_POIDSBRUT                 AS Gross_Weight_Item,
  A.GA_TENUESTOCK                AS In_Item_Stock,
  A.GA_DATEMODIF                 AS Item_Modification_Date,
  A.GA_DATECREATION              AS Item_Creation_Date,
  A.GA_CODEDIM2                  AS Size,
  A.GA_EMBALLAGE                 AS Packing_Item,
  A.GA_UNITEQTEVTE               AS Sales_Unit_Quantity
FROM ARTICLE A
  LEFT JOIN ARTICLECOMPL A2
         ON A2.GA2_ARTICLE = A.GA_ARTICLE
  LEFT JOIN DIMENSION D1                       -- warna
         ON D1.GDI_TYPEDIM    = 'DI1'
        AND D1.GDI_GRILLEDIM  = A.GA_GRILLEDIM1
        AND D1.GDI_CODEDIM    = A.GA_CODEDIM1
  LEFT JOIN DIMENSION D2                       -- ukuran
         ON D2.GDI_TYPEDIM    = 'DI2'
        AND D2.GDI_GRILLEDIM  = A.GA_GRILLEDIM2
        AND D2.GDI_CODEDIM    = A.GA_CODEDIM2
  OUTER APPLY (
      SELECT TOP 1 RP.GF_PRIXUNITAIRE
      FROM GCTARFCONMODEART RP
      WHERE RP.GA_CODEARTICLE = A.GA_CODEARTICLE
        AND RP.GF_REGIMEPRIX  = 'TTC'
        AND RP.GFM_TYPETARIF  = 'BSP01'
        AND RP.GFM_PERTARIF IN ('00BASE', 'SGC1')
        AND RP.GFM_NATURETYPE = 'VTE'
        AND RP.GF_ARTICLE <> ''
        AND GETDATE() BETWEEN RP.GF_DATEDEBUT AND RP.GF_DATEFIN
      ORDER BY RP.GF_DATEDEBUT DESC, RP.GF_DATEMODIF DESC
  ) RP
  LEFT JOIN CHOIXEXT YX4  ON A.GA_LIBREART2   = YX4.YX_CODE  AND YX4.YX_TYPE  = 'LA2'
  LEFT JOIN CHOIXEXT YX6  ON A.GA_LIBREART7   = YX6.YX_CODE  AND YX6.YX_TYPE  = 'LA7'
  LEFT JOIN CHOIXEXT YX7  ON A.GA_LIBREART1   = YX7.YX_CODE  AND YX7.YX_TYPE  = 'LA1'
  LEFT JOIN CHOIXEXT YX8  ON A.GA_LIBREART3   = YX8.YX_CODE  AND YX8.YX_TYPE  = 'LA3'
  LEFT JOIN CHOIXEXT YX9  ON A.GA_LIBREART4   = YX9.YX_CODE  AND YX9.YX_TYPE  = 'LA4'
  LEFT JOIN CHOIXEXT YX10 ON A.GA_LIBREART8   = YX10.YX_CODE AND YX10.YX_TYPE = 'LA8'
  LEFT JOIN CHOIXEXT YX11 ON A.GA_LIBREART9   = YX11.YX_CODE AND YX11.YX_TYPE = 'LA9'
  LEFT JOIN CHOIXEXT YX12 ON A2.GA2_LIBREARTB = YX12.YX_CODE AND YX12.YX_TYPE = 'LAB'
  LEFT JOIN CHOIXEXT YX14 ON A2.GA2_LIBREARTD = YX14.YX_CODE AND YX14.YX_TYPE = 'LAD'
  LEFT JOIN CHOIXEXT YX15 ON A2.GA2_LIBREARTE = YX15.YX_CODE AND YX15.YX_TYPE = 'LAE'
  LEFT JOIN CHOIXEXT YX16 ON A2.GA2_LIBREARTF = YX16.YX_CODE AND YX16.YX_TYPE = 'LAF'
WHERE A.GA_STATUTART   = 'DIM'
  AND A.GA_FAMILLENIV1 = ?
  AND ISNULL(A.GA_TENUESTOCK,  'X')   = 'X'     -- NULL dianggap tenue stock
  AND ISNULL(A.GA_EMBALLAGE,   '-')  <> 'X'     -- NULL dianggap bukan emballage
  AND ISNULL(A.GA_TYPEARTICLE, 'MAR') = 'MAR'   -- NULL dianggap MAR
"""



def get_db_connection():
    try:
        conn = pyodbc.connect(DB_CONNECTION)
        print(f"  [DB] Koneksi berhasil ke {SERVER} / {DATABASE}")
        return conn
    except Exception as e:
        print(f"  [DB ERROR] Gagal konek ke database: {e}")
        raise

def execute_query(brand: str) -> dict:
    """
    Eksekusi SQL_QUERY ke database dengan parameter brand.
    Return format:
        { article_upper: { colour_upper: { retail_price, promotion, dept, segment, brand, article, colour } } }

    Eksekusi SQL_QUERY ke database dengan parameter brand.
    Return format:
        { article_upper: { colour_upper: { retail_price, promotion, dept, segment, brand, article, colour } } }
    """
    print(f"  [DB] Menjalankan query untuk brand='{brand}'...")

    conn   = get_db_connection()
    cursor = conn.cursor()

    try:
        cursor.execute(SQL_QUERY, (brand,))
        columns = [col[0] for col in cursor.description]
        rows    = cursor.fetchall()
    finally:
        cursor.close()
        conn.close()

    print(f"  [DB] {len(rows)} baris diterima dari database.")

    col_idx = {c: i for i, c in enumerate(columns)}

    data = {}
    for row in rows:
        article   = str(row[col_idx['Article']]).strip()     if row[col_idx['Article']]    else ''
        colour    = str(row[col_idx['Colour_Desc']]).strip() if row[col_idx['Colour_Desc']] else ''
        retail    = row[col_idx['Retail_Price']]
        promotion = str(row[col_idx['Promotion']]).strip()   if row[col_idx['Promotion']]  else ''
        dept      = str(row[col_idx['Department']]).strip()  if row[col_idx['Department']] else ''
        segment   = str(row[col_idx['Segment']]).strip()     if row[col_idx['Segment']]    else ''
        brand_val = str(row[col_idx['Brand']]).strip()       if row[col_idx['Brand']]      else ''

        if not article:
            continue

        art_key    = article.upper()
        col_key    = colour.upper()
        retail_val = int(retail) if isinstance(retail, (int, float, Decimal)) and retail else 0

        if art_key not in data:
            data[art_key] = {}

        existing = data[art_key].get(col_key)
        # Simpan jika: belum ada entry, atau harga baru lebih tinggi.
        # Jangan timpa harga yang sudah ada dengan 0.
        if existing is None or (retail_val > 0 and retail_val > existing['retail_price']):
            data[art_key][col_key] = {
                'retail_price': retail_val,
                'promotion':    promotion,
                'dept':         dept,
                'segment':      segment,
                'brand':        brand_val,
                'article':      article,
                'colour':       colour,
            }

    return data

def normalize_key(raw: str) -> str:
    """Normalize input 'Item Group Menjadi' value to canonical discount_key."""
    raw = raw.strip()
    # e.g. "SP490" -> "SP490K", "D20" -> "D20"
    m = re.fullmatch(r'SP(\d+)K?', raw, re.IGNORECASE)
    if m:
        return 'SP' + m.group(1) + 'K'
    m = re.fullmatch(r'D(\d+)', raw, re.IGNORECASE)
    if m:
        return 'D' + m.group(1)
    return raw.upper()
 
def load_promo_groups_from_txt(txt_path: str) -> dict:
    with open(txt_path, 'r', encoding='utf-8') as f:
        raw = f.read()
 
    result = {}
    sheet_pattern = re.compile(
        r'"([^"]+)":\s*\{(.*?)\n    \}',
        re.DOTALL
    )
    table_pattern = re.compile(
        r'"([^"]+)":\s*PromoTable\((.*?)\n        \)',
        re.DOTALL
    )
    row_pattern = re.compile(
        r"PromoRow\(article='(.*?)',\s*colour='(.*?)',\s*harga_awal=(\S+?),\s*harga_diskon=(\S+?)"
        r"(?:,\s*moved_from=(.*?),\s*moved_to=(.*?))?\)"
    )
 
    for sm in sheet_pattern.finditer(raw):
        sheet_name = sm.group(1)
        sheet_body = sm.group(2)
        tables = {}
 
        for tm in table_pattern.finditer(sheet_body):
            key = tm.group(1)
            table_body = tm.group(2)
 
            # Extract fields
            def _get(field_name):
                m = re.search(rf"{field_name}='([^']*)'", table_body)
                return m.group(1) if m else ''
 
            def _get_int(field_name):
                m = re.search(rf"{field_name}=(\d+)", table_body)
                return int(m.group(1)) if m else 0
 
            sheet_val    = _get('sheet')
            promo_name   = _get('promo_name')
            discount_key = _get('discount_key')
            count        = _get_int('count')
 
            rows = []
            for rm in row_pattern.finditer(table_body):
                art  = rm.group(1).replace("\\'", "'")
                col  = rm.group(2).replace("\\'", "'")
                ha   = None if rm.group(3) == 'None' else int(rm.group(3).rstrip(','))
                hd   = None if rm.group(4).rstrip(',') == 'None' else float(rm.group(4).rstrip(','))
                mf_raw = rm.group(5).strip().strip("'") if rm.group(5) else 'None'
                mt_raw = rm.group(6).strip().strip("'") if rm.group(6) else 'None'
                mf = None if mf_raw == 'None' else mf_raw
                mt = None if mt_raw == 'None' else mt_raw
                rows.append(PromoRow(article=art, colour=col,
                                     harga_awal=ha, harga_diskon=hd,
                                     moved_from=mf, moved_to=mt))
 
            tables[key] = PromoTable(
                sheet=sheet_val,
                promo_name=promo_name,
                discount_key=discount_key,
                count=count,
                rows=rows,
            )
 
        result[sheet_name] = tables
    return result
 
def load_query_result(xlsx_path: str) -> dict:
    wb = openpyxl.load_workbook(xlsx_path, data_only=True)
    ws = wb.active
 
    headers = [str(c.value).strip() if c.value else '' for c in next(ws.iter_rows(min_row=1, max_row=1))]
    idx = {h: i for i, h in enumerate(headers)}
 
    data = {}  
    for row in ws.iter_rows(min_row=2, values_only=True):
        article   = str(row[idx['Article']]).strip() if row[idx['Article']] else ''
        colour    = str(row[idx['Colour_Desc']]).strip() if row[idx['Colour_Desc']] else ''
        retail    = row[idx['Retail_Price']]
        promotion = str(row[idx['Promotion']]).strip() if row[idx['Promotion']] else ''
        dept      = str(row[idx['Department']]).strip() if row[idx['Department']] else ''
        segment   = str(row[idx['Segment']]).strip() if row[idx['Segment']] else ''
        brand     = str(row[idx['Brand']]).strip() if row[idx['Brand']] else ''
 
        if not article:
            continue
 
        art_key = article.upper()
        col_key = colour.upper()
 
        retail_val = int(retail) if isinstance(retail, (int, float, Decimal)) and retail else 0
 
        if art_key not in data:
            data[art_key] = {}
 
        existing = data[art_key].get(col_key)
        # Simpan jika: belum ada entry, atau harga baru lebih tinggi.
        # Jangan timpa harga yang sudah ada dengan 0.
        if existing is None or (retail_val > 0 and retail_val > existing['retail_price']):
            data[art_key][col_key] = {
                'retail_price': retail_val,
                'promotion':    promotion,
                'dept':         dept,
                'segment':      segment,
                'brand':        brand,
                'article':      article,
                'colour':       colour,
            }
 
    return data
 
def load_input_file(xlsx_path: str) -> list:
    wb = openpyxl.load_workbook(xlsx_path, data_only=True)
    ws = wb.active
 
    rows = []
    for i, row in enumerate(ws.iter_rows(min_row=2, values_only=True)):
        if not any(v for v in row):
            continue
        brand        = str(row[0]).strip() if row[0] else ''
        item_code    = str(row[1]).strip() if row[1] else ''
        colour       = str(row[2]).strip() if row[2] else None
        if colour in ('None', '', 'none'):
            colour = None
        item_group   = str(row[3]).strip() if row[3] else ''
        item_group_menjadi = str(row[4]).strip() if row[4] else ''
        # Kolom 'Seg' (opsional, kolom ke-6) -- dipakai untuk memilih sheet template
        seg          = str(row[5]).strip().upper() if len(row) > 5 and row[5] else None
        rows.append({
            'brand':            brand,
            'item_code':        item_code.upper(),
            'colour':           colour.upper() if colour else None,
            'item_group':       normalize_key(item_group) if item_group else None,
            'item_group_menjadi': normalize_key(item_group_menjadi),
            'seg':              seg,
        })
    return rows
  
def compute_harga_diskon(harga_awal: int, discount_key: str) -> Optional[int]:
    if harga_awal is None or harga_awal == 0:
        return None
 
    m = re.fullmatch(r'D(\d+)', discount_key, re.IGNORECASE)
    if m:
        pct = int(m.group(1))
        return round(harga_awal * (1 - pct / 100))
 
    m = re.fullmatch(r'SP(\d+)K', discount_key, re.IGNORECASE)
    if m:
        return int(m.group(1)) * 1000
 
    return None

def recount(promo_groups: dict):
    """Sort semua rows dan update count di setiap PromoTable."""
    for sheet_tables in promo_groups.values():
        for table in sheet_tables.values():
            table.rows  = sort_rows(table.rows)
            table.count = len(table.rows)

# ---------------------------------------------------------------------------
# Apply instruksi input ke promo_groups
# ---------------------------------------------------------------------------
def _build_promo_name(discount_key: str, suffix: str) -> str:
    """Buat promo_name yang readable dari discount_key dan suffix sheet."""
    m = re.fullmatch(r'D(\d+)', discount_key, re.IGNORECASE)
    if m:
        return f"DISC. {m.group(1)}% {suffix}"
    m = re.fullmatch(r'SP(\d+)K', discount_key, re.IGNORECASE)
    if m:
        return f"SP{m.group(1)}K {suffix}"
    return f"{discount_key} {suffix}"


# ---------------------------------------------------------------------------
# BUG FIX 1: _get_colour_promo_map
# Sebelumnya, script mengambil satu sample pertama dari q_records untuk
# menentukan orig_key (promo asal). Ini salah karena 1 artikel bisa punya
# promo berbeda per warna (contoh: AMARELA LIGHT BLUE=D50, BLACK=D30).
# Fungsi ini membangun mapping colour -> orig_key secara individual.
# ---------------------------------------------------------------------------
def _get_colour_promo_map(q_records: dict) -> dict:
    """Return { colour_upper: normalize_key(promotion) or None }"""
    result = {}
    for col_key, rec in q_records.items():
        orig_promo = rec.get('promotion', '')
        if orig_promo not in ('NML', '', 'NULL', None):
            result[col_key] = normalize_key(orig_promo)
        else:
            result[col_key] = None
    return result


def _expand_all_colour_row(src_table, art_code: str, q_records: dict,
                           exclude_colours: set, promo_groups: dict,
                           sheet_name: str) -> Optional[int]:
    """
    Jika tabel asal punya row 'All Colour' untuk artikel ini, hapus row tersebut
    lalu kembalikan sisa warna (yang tidak ada di exclude_colours) ke tabel asal
    sebagai row per warna. Return harga dari row All Colour (sebagai fallback harga),
    atau None jika tidak ada row All Colour yang ditemukan.

    Dipanggil saat instruksi adalah per warna tertentu saja (bukan All Colour),
    sehingga warna yang tidak dipindahkan harus tetap muncul di tabel asal.
    """
    harga_fallback = None
    new_rows = []
    expanded = False

    for existing_row in src_table.rows:
        if existing_row.article.upper() == art_code and existing_row.colour.upper() == 'ALL COLOUR':
            # Simpan harga sebagai fallback
            harga_fallback = existing_row.harga_awal
            expanded = True
            # Jangan masukkan ke new_rows (row All Colour ini dihapus)
        else:
            new_rows.append(existing_row)

    if expanded:
        src_table.rows = new_rows
        # Tambahkan sisa warna individual yang tidak masuk instruksi pindah
        colour_promo_map = _get_colour_promo_map(q_records)
        for col_key, rec in q_records.items():
            if col_key in exclude_colours:
                continue  # warna ini sedang dipindahkan, skip
            actual_colour = rec['colour']
            harga_awal = rec['retail_price'] or harga_fallback or 0
            # Tentukan discount_key tabel asal dari colour_promo_map
            orig_key = colour_promo_map.get(col_key)
            # Cari tabel yang sesuai (orig_key atau tabel saat ini)
            src_discount_key = src_table.discount_key
            harga_diskon = compute_harga_diskon(harga_awal, src_discount_key)
            # Pastikan belum ada
            already = any(r.article.upper() == art_code and r.colour.upper() == actual_colour.upper()
                          for r in src_table.rows)
            if not already:
                src_table.rows.append(PromoRow(
                    article=rec['article'],
                    colour=actual_colour,
                    harga_awal=harga_awal,
                    harga_diskon=harga_diskon,
                    moved_from=None,
                    moved_to=None,
                ))
                print(f"  [EXPAND] {art_code}/{actual_colour} tetap di {src_discount_key} (sisa dari All Colour)")

    return harga_fallback


def _remove_from_source(promo_groups: dict, sheet_name: str,
                        art_code: str, col_key: str, moved_from: Optional[str],
                        target_key: str, category_suffix: Optional[str] = None,
                        also_remove_all_colour: bool = False):
    """
    Hapus row (art_code/col_key) dari tabel asal jika moved_from != target.
    Jika also_remove_all_colour=True, juga hapus row 'All Colour' untuk artikel
    yang sama (dipakai saat semua warna dipindahkan sekaligus).

    FIX: tabel sumber dicari dengan discount_key + category_suffix (bukan
    discount_key polos), supaya kalau dalam satu sheet ada beberapa tabel
    dengan persen sama tapi kategori beda (mis. EB MENS BAG vs EB LADIES
    BAG), penghapusan tidak salah sasaran ke tabel kategori lain -- yang
    sebelumnya menyebabkan artikel Mens Bag tidak pernah terhapus dari
    tabel diskon lamanya.
    """
    if not moved_from or moved_from == target_key:
        return
    src_tables = promo_groups.get(sheet_name, {})
    src_key = resolve_table_key(promo_groups, sheet_name, moved_from, category_suffix)
    if src_key not in src_tables:
        # Fallback: coba key polos (kompatibilitas utk tabel tanpa kategori, mis. SP)
        if moved_from in src_tables:
            src_key = moved_from
        else:
            return
    src_table = src_tables[src_key]
    kept = []
    for existing_row in src_table.rows:
        art_match = existing_row.article.upper() == art_code
        col_match = existing_row.colour.upper() == col_key
        all_colour_match = also_remove_all_colour and existing_row.colour.upper() == 'ALL COLOUR'
        if art_match and (col_match or all_colour_match):
            print(f"  [MOVE] {art_code}/{existing_row.colour}: {moved_from} -> {target_key}")
        else:
            kept.append(existing_row)
    src_table.rows = kept


def _remove_to_normal(promo_groups: dict, sheet_name: str, suffix: str,
                      art_code: str, in_colour: Optional[str], q_records: dict):
    """
    Instruksi '-> NML': keluarkan artikel (atau 1 warna) dari semua tabel
    promo di sheet. Kalau warnanya tercakup row 'All Colour', row tersebut
    dipecah dulu supaya warna lain tetap di tabel asalnya.
    """
    source_tables = _find_source_tables(promo_groups, sheet_name, suffix, art_code)
    if not source_tables:
        print(f"  [INFO] {art_code}: -> NML, tidak ditemukan di tabel promo manapun, dilewati.")
        return

    for src_key, src_table in source_tables:
        had = any(r.article.upper() == art_code and
                  (in_colour is None or r.colour.upper() in (in_colour, 'ALL COLOUR'))
                  for r in src_table.rows)
        if not had:
            continue
        if in_colour is None:
            src_table.rows = [r for r in src_table.rows if r.article.upper() != art_code]
        else:
            has_all_colour = any(r.article.upper() == art_code and r.colour.upper() == 'ALL COLOUR'
                                 for r in src_table.rows)
            if has_all_colour:
                _expand_all_colour_row(src_table, art_code, q_records,
                                       exclude_colours={in_colour},
                                       promo_groups=promo_groups, sheet_name=sheet_name)
            src_table.rows = [r for r in src_table.rows
                              if not (r.article.upper() == art_code and r.colour.upper() == in_colour)]
        print(f"  [NML ] {art_code}/{in_colour or 'All Colour'}: dikeluarkan dari {src_key}")


def find_duplicate_articles(promo_groups: dict) -> list:
    """
    Cari artikel+warna yang muncul di lebih dari 1 tabel promo dalam sheet
    yang sama. Return [(sheet, article, colour, [promo_name, ...]), ...].
    """
    dupes = []
    for sheet_name, tables in promo_groups.items():
        where = {}
        for t in tables.values():
            for r in t.rows:
                where.setdefault((r.article.upper(), r.colour.upper()), []).append(t.promo_name)
        for (art, col), names in where.items():
            if len(names) > 1:
                dupes.append((sheet_name, art, col, names))
    return dupes


def print_duplicate_warnings(promo_groups: dict, label: str):
    dupes = find_duplicate_articles(promo_groups)
    if not dupes:
        return
    print(f"  [WARN] {len(dupes)} artikel dobel di {label} (cek manual, script tidak menghapusnya):")
    for sheet_name, art, col, names in dupes:
        print(f"         {sheet_name}: {art} / {col} -> {', '.join(names)}")


def _sheet_uses_shared_tables(promo_groups: dict, sheet_name: str) -> bool:
    """True jika sheet sudah punya tabel dan SEMUA tabelnya tanpa category_suffix
    (header template tidak memakai suffix SHEET_MAP, mis. 'DISC. 30% VBH')."""
    tables = list(promo_groups.get(sheet_name, {}).values())
    return bool(tables) and all(t.category_suffix is None for t in tables)


_PROMO_PREFIX_RE = re.compile(r'^\s*(?:DISC[\.\s]*\d+\s*%|SP\s*\d+\s*K?|D\d+)\s*', re.IGNORECASE)

def _pick_header_suffix(promo_groups: dict, sheet_name: str, art_code: str) -> Optional[str]:
    """
    Ambil akhiran header (bagian setelah 'DISC. 30%' / 'SP90K') dari tabel
    yang sudah ada di sheet, mis. 'VBH', 'SLG', 'WAL'. Prioritas: tabel yang
    saat ini berisi art_code, lalu akhiran yang paling sering di sheet itu.
    """
    tables = list(promo_groups.get(sheet_name, {}).values())

    def _tail(t):
        return _PROMO_PREFIX_RE.sub('', t.promo_name).strip()

    for t in tables:
        if any(r.article.upper() == art_code for r in t.rows) and _tail(t):
            return _tail(t)

    tails = [_tail(t) for t in tables if _tail(t)]
    return max(set(tails), key=tails.count) if tails else None


def _find_source_tables(promo_groups: dict, sheet_name: str,
                        category_suffix: Optional[str], art_code: str,
                        exclude_key: Optional[str] = None) -> list:
    """Cari tabel sumber aktual. Utamakan kategori yang cocok, lalu fallback
    ke tabel lain pada sheet yang sama jika artikel tidak ditemukan di kategori
    tersebut. Fallback ini menangani template yang suffix/header-nya tidak
    konsisten dengan mapping database, tanpa mengubah tabel lain jika match
    kategori yang tepat sudah tersedia.
    """
    tables = promo_groups.get(sheet_name, {})
    exact, fallback = [], []
    for key, table in tables.items():
        if key == exclude_key:
            continue
        if not any(str(r.article).strip().upper() == art_code for r in table.rows):
            continue
        fallback.append((key, table))
        if getattr(table, 'category_suffix', None) in (category_suffix, None):
            exact.append((key, table))
    if exact:
        return exact
    if fallback:
        print(f"  [WARN] {art_code}: sumber tidak ditemukan pada kategori '{category_suffix}', "
              f"fallback mencari seluruh tabel di sheet '{sheet_name}'.")
    return fallback


def _resolve_colour_key(q_records: dict, requested_colour: Optional[str]) -> Optional[str]:
    """Temukan key warna dari hasil query meskipun beda kapital/spasi/tanda baca."""
    if requested_colour is None:
        return None
    requested = str(requested_colour).strip().upper()
    if requested in q_records:
        return requested
    norm = re.sub(r'[^A-Z0-9]', '', requested)
    for key, rec in q_records.items():
        key_norm = re.sub(r'[^A-Z0-9]', '', str(key).upper())
        rec_colour = re.sub(r'[^A-Z0-9]', '', str(rec.get('colour', '')).upper())
        if norm and norm in (key_norm, rec_colour):
            return key
    return None


def apply_input(promo_groups: dict, input_rows: list, query_data: dict):
    for instr in input_rows:
        art_code   = instr['item_code']
        in_colour  = instr['colour']          # None = all colour
        target_key = instr['item_group_menjadi']
        src_group_fallback = instr.get('item_group')   # 'Item Group' dari file input

        q_records = query_data.get(art_code, {})
        if not q_records:
            print(f"  [WARN] {art_code}: tidak ditemukan di hasil query SQL, dilewati.")
            continue

        # Ambil dept/segment dari sample pertama (sama untuk semua warna 1 artikel)
        sample    = next(iter(q_records.values()))
        dept      = sample['dept']
        segment   = sample['segment']

        sheet_name, suffix = get_sheet_and_suffix(dept, segment)

        # Kalau template punya sheet sendiri untuk segmen ini (mis. sheet 'ACC'
        # di template EVB), pakai sheet itu, bukan sheet dari SHEET_MAP.
        # Prioritas: kolom 'Seg' di file input, lalu segment dari database
        # (mis. DB bilang 'SUN' tapi input bilang 'ACC' -> sheet ACC).
        seg_sheet = None
        for seg in (instr.get('seg'), segment):
            if seg:
                seg_sheet = next((s for s in promo_groups if s.strip().upper() == seg.strip().upper()), None)
                if seg_sheet:
                    break
        if seg_sheet:
            sheet_name = seg_sheet

        if sheet_name is None:
            print(f"  [WARN] {art_code}: dept={dept} segment={segment} tidak ada di SHEET_MAP, dilewati.")
            continue

        if sheet_name not in promo_groups:
            promo_groups[sheet_name] = {}

        # Item Group Menjadi = NML -> artikel keluar dari promo: hapus dari
        # tabel asalnya saja, tidak dibuatkan tabel "NML".
        if target_key == 'NML':
            _remove_to_normal(promo_groups, sheet_name, suffix, art_code, in_colour, q_records)
            continue

        # Cari/buat tabel target dengan key gabungan (discount_key + kategori)
        # dulu, supaya "D50" milik EB MENS BAG tidak salah nyambung ke tabel
        # "D50" milik EB LADIES BAG yang kebetulan sudah ada duluan di sheet
        # yang sama. Kalau tidak ketemu, coba key polos (utk kompatibilitas
        # dgn tabel yang memang dipakai bersama semua kategori, mis. tabel
        # SP di sheet HANDBAG yang tidak dipisah Mens/Ladies). Kalau dua-duanya
        # tidak ada, baru buat tabel baru dengan key gabungan.
        combo_key = resolve_table_key(promo_groups, sheet_name, target_key, suffix)
        if combo_key in promo_groups[sheet_name]:
            target_store_key = combo_key
        elif (target_key in promo_groups[sheet_name]
              and getattr(promo_groups[sheet_name][target_key], 'category_suffix', None) in (suffix, None)):
            target_store_key = target_key
        else:
            # Sheet yang tabelnya dipakai bersama (header tanpa kategori SHEET_MAP,
            # mis. 'SP90K VBH'): tabel baru juga dipakai bersama dan akhiran
            # header-nya meminjam dari tabel yang sudah ada di sheet ini.
            shared = _sheet_uses_shared_tables(promo_groups, sheet_name)
            target_store_key = target_key if shared else combo_key
            header_suffix = (_pick_header_suffix(promo_groups, sheet_name, art_code) if shared else None) or suffix or sheet_name
            print(f"  [INFO] Membuat tabel promo baru: sheet={sheet_name} key={target_store_key}")
            promo_groups[sheet_name][target_store_key] = PromoTable(
                sheet=sheet_name,
                promo_name=_build_promo_name(target_key, header_suffix),
                discount_key=target_key,
                category_suffix=None if shared else suffix,
                rows=[],
            )

        target_table = promo_groups[sheet_name][target_store_key]

        # Tentukan warna yang diproses. Input warna dapat berbeda kapital,
        # spasi, atau tanda baca dari Colour_Desc hasil database.
        if in_colour is None:
            colours_to_process = list(q_records.keys())
        else:
            resolved_colour = _resolve_colour_key(q_records, in_colour)
            if resolved_colour is None:
                if len(q_records) == 1:
                    resolved_colour = next(iter(q_records))
                    print(f"  [COLOR-MAP] {art_code}: warna input '{in_colour}' tidak cocok dengan DB; "
                          f"menggunakan satu-satunya warna DB '{q_records[resolved_colour].get('colour', resolved_colour)}'.")
                else:
                    print(f"  [WARN] {art_code}: warna input '{in_colour}' tidak ditemukan. "
                          f"Warna tersedia: {list(q_records.keys())}. Instruksi dilewati agar tidak salah warna.")
                    continue
            colours_to_process = [resolved_colour]

        # Cek apakah semua warna yang diproses punya harga sama (untuk All Colour collapse)
        prices_to_check = [q_records[c]['retail_price'] for c in colours_to_process if c in q_records]
        all_same_price  = len(set(prices_to_check)) == 1 if prices_to_check else False

        # FIX: cari tabel sumber SEBENARNYA langsung dari isi hasil scan
        # (ground truth), bukan dari field 'Promotion' di database.
        source_tables = _find_source_tables(promo_groups, sheet_name, suffix,
                                            art_code, exclude_key=target_store_key)

        # ----------------------------------------------------------------
        # KASUS A: Instruksi All Colour (in_colour is None) + harga sama
        # Collapse semua warna menjadi satu row "All Colour" di target.
        # ----------------------------------------------------------------
        if in_colour is None and all_same_price:
            harga_awal = prices_to_check[0]

            # Jika harga masih 0, warn dan skip
            if not harga_awal:
                print(f"  [WARN] {art_code}/All Colour: harga_awal=0, tidak dapat dihitung. Dilewati.")
                continue

            harga_diskon = compute_harga_diskon(harga_awal, target_key)

            # Hapus artikel dari SEMUA tabel sumber yang benar-benar
            # ditemukan berisi artikel ini (bisa >1 kalau sebelumnya
            # sempat kepencar/duplikat), lalu catat dari tabel mana saja
            # ia berasal (utk moved_from / pewarnaan legend).
            moved_from_labels = set()
            for src_key, src_table in source_tables:
                before_n = len(src_table.rows)
                src_table.rows = [r for r in src_table.rows if r.article.upper() != art_code]
                removed_n = before_n - len(src_table.rows)
                if removed_n:
                    moved_from_labels.add(src_table.discount_key)
                    print(f"  [MOVE] {art_code}: {src_key} -> {target_store_key} ({removed_n} baris dihapus)")

            if len(moved_from_labels) == 0:
                if src_group_fallback and src_group_fallback != target_key:
                    all_colour_moved_from = src_group_fallback
                    print(f"  [INFO] {art_code}: tidak ditemukan di tabel diskon manapun hasil "
                          f"scan, pakai 'Item Group' dari input -> {src_group_fallback}")
                else:
                    all_colour_moved_from = None
            elif len(moved_from_labels) == 1:
                all_colour_moved_from = next(iter(moved_from_labels))
            else:
                all_colour_moved_from = ','.join(sorted(moved_from_labels))

            # Bersihkan SEMUA baris artikel ini yang sudah ada di tabel
            # TARGET sendiri (per-warna maupun All Colour), termasuk kasus di
            # mana artikel ini sebenarnya sudah berada di tabel target sejak
            # awal, supaya tidak terjadi duplikasi dengan row "All Colour" baru.
            before_count = len(target_table.rows)
            target_table.rows = [
                r for r in target_table.rows if r.article.upper() != art_code
            ]
            removed_dupe = before_count - len(target_table.rows)
            if removed_dupe:
                print(f"  [DEDUPE] {art_code}: {removed_dupe} row lama di {target_store_key} "
                      f"dihapus, digabung jadi All Colour")

            # Warna highlight baris mengikuti Item Group ASAL pada file input.
            # Tabel sumber hasil scan bisa berbeda/keliru karena data template
            # lama atau artikel sudah lebih dulu berpindah. Untuk label warna,
            # instruksi input adalah acuan: D40 -> hijau, D30 -> biru, dst.
            if src_group_fallback and src_group_fallback != target_key:
                input_source_key = normalize_key(src_group_fallback)
                if input_source_key:
                    all_colour_moved_from = input_source_key

            moved_to_val = target_key if all_colour_moved_from else None
            target_table.rows.append(PromoRow(
                article=sample['article'],
                colour='All Colour',
                harga_awal=harga_awal,
                harga_diskon=harga_diskon,
                moved_from=all_colour_moved_from,
                moved_to=moved_to_val,
            ))
            print(f"  [ADD ] {art_code}/All Colour -> {target_store_key} "
                  f"harga_awal={harga_awal} harga_diskon={harga_diskon}"
                  + (f" [dari {all_colour_moved_from}]" if all_colour_moved_from else ""))
            continue

        # ----------------------------------------------------------------
        # KASUS B: Instruksi per warna tertentu (in_colour is not None),
        # atau instruksi All Colour tapi harga berbeda-beda (expand per warna).
        # ----------------------------------------------------------------
        for col_key in colours_to_process:
            rec = q_records.get(col_key)
            if rec is None:
                print(f"  [WARN] {art_code} colour={col_key}: tidak ditemukan di query, dilewati.")
                continue

            harga_awal = rec['retail_price']
            # Jika user mengisi warna secara eksplisit, tulis persis warna dari
            # file input (bukan label warna lama/template dari database).
            actual_colour = str(in_colour).strip() if in_colour is not None else rec['colour']

            # FIX: cari tabel sumber utk warna ini dari ground truth (isi
            # tabel sebenarnya), bukan dari field 'Promotion' di database.
            src_match = None
            for src_key, src_table in source_tables:
                if any(r.article.strip().upper() == art_code and
                       (r.colour.strip().upper() in (str(col_key).strip().upper(), 'ALL COLOUR'))
                       for r in src_table.rows):
                    src_match = (src_key, src_table)
                    break

            # Jika warna pada template salah (mis. template masih abu/putih,
            # sedangkan master DB menunjukkan hijau), exact match tidak ketemu.
            # Koreksi otomatis hanya jika tepat satu baris artikel tersedia di
            # tabel sumber; jika lebih dari satu, jangan menebak warna mana.
            source_row_to_replace = None

            if src_match is None:
                # Warna input adalah acuan. Cari baris lama yang warnanya
                # TIDAK dikenal di master DB, sehingga warna lain yang valid
                # untuk artikel yang sama tetap aman.
                valid_db_colours = {
                    re.sub(r'[^A-Z0-9]', '', str(value).strip().upper())
                    for key, value in (
                        [(k, k) for k in q_records.keys()]
                        + [(k, v.get('colour', '')) for k, v in q_records.items()]
                    )
                    if value
                }

                stale_candidates = []
                all_article_candidates = []
                for candidate_key, candidate_table in source_tables:
                    for candidate_row in candidate_table.rows:
                        if (candidate_row.article.strip().upper() == art_code
                                and candidate_row.colour.strip().upper() != 'ALL COLOUR'):
                            item = (candidate_key, candidate_table, candidate_row)
                            all_article_candidates.append(item)
                            old_colour_norm = re.sub(
                                r'[^A-Z0-9]', '', candidate_row.colour.strip().upper()
                            )
                            if old_colour_norm not in valid_db_colours:
                                stale_candidates.append(item)

                # Prioritaskan satu warna lama yang tidak ada di master DB.
                # Fallback ke satu-satunya baris artikel jika memang hanya ada
                # satu baris, seperti kasus warna template yang salah.
                candidates_to_replace = stale_candidates or all_article_candidates
                if len(candidates_to_replace) == 1:
                    candidate_key, candidate_table, candidate_row = candidates_to_replace[0]
                    src_match = (candidate_key, candidate_table)
                    source_row_to_replace = candidate_row
                    print(f"  [COLOR-CORRECTION] {art_code}: warna lama "
                          f"'{candidate_row.colour}' akan diganti menjadi '{actual_colour}' "
                          f"sesuai warna input.")
                elif len(candidates_to_replace) > 1:
                    print(f"  [WARN] {art_code}/{actual_colour}: ditemukan beberapa baris warna "
                          f"lama yang mungkin cocok; tidak menghapus warna lain secara otomatis.")

            harga_fallback = None
            moved_from = None
            if src_match:
                src_key, src_table = src_match
                moved_from = src_table.discount_key
                has_all_colour = any(
                    r.article.upper() == art_code and r.colour.upper() == 'ALL COLOUR'
                    for r in src_table.rows
                )
                if has_all_colour:
                    # Warna lain (di luar instruksi ini) yang berasal dari
                    # tabel sumber yang sama harus tetap "expand" kembali ke
                    # tabel sumber, bukan ikut terhapus.
                    colours_moving_from_this_src = {
                        c for c in colours_to_process
                        if any(r.article.upper() == art_code and
                               r.colour.upper() in (c, 'ALL COLOUR')
                               for r in src_table.rows)
                    }
                    harga_fallback = _expand_all_colour_row(
                        src_table, art_code, q_records,
                        exclude_colours=colours_moving_from_this_src,
                        promo_groups=promo_groups,
                        sheet_name=sheet_name,
                    )
                else:
                    before_n = len(src_table.rows)

                    if source_row_to_replace is not None:
                        # Hapus objek baris lama yang sudah dipastikan sebagai
                        # kandidat koreksi warna. Jangan bandingkan dengan
                        # col_key karena nilainya memang berbeda dari warna input.
                        src_table.rows = [
                            r for r in src_table.rows if r is not source_row_to_replace
                        ]
                        print(f"  [COLOR-REPLACE] {art_code}: "
                              f"'{source_row_to_replace.colour}' -> '{actual_colour}' "
                              f"di tabel sumber {src_key}.")
                    else:
                        src_table.rows = [
                            r for r in src_table.rows
                            if not (r.article.strip().upper() == art_code and
                                    r.colour.strip().upper() == str(col_key).strip().upper())
                        ]

                    if len(src_table.rows) != before_n:
                        print(f"  [MOVE] {art_code}/{actual_colour}: {src_key} -> {target_store_key}")
            else:
                if src_group_fallback and src_group_fallback != target_key:
                    moved_from = src_group_fallback
                    print(f"  [INFO] {art_code}/{actual_colour}: tidak ditemukan di tabel diskon "
                          f"manapun hasil scan, pakai 'Item Group' dari input -> {src_group_fallback}")

            # Gunakan harga fallback dari row All Colour jika query memberikan 0
            if (harga_awal is None or harga_awal == 0) and harga_fallback:
                harga_awal = harga_fallback

            # Jika masih 0 setelah semua fallback, skip dan warn
            if not harga_awal:
                print(f"  [WARN] {art_code}/{actual_colour}: harga_awal=0, tidak dapat dihitung. Dilewati.")
                continue

            harga_diskon = compute_harga_diskon(harga_awal, target_key)

            # Sumber warna highlight wajib mengikuti Item Group pada input,
            # bukan discount_key tabel yang kebetulan ditemukan saat scan.
            # Contoh: D40 -> D50 tetap diberi warna hijau (asal D40),
            # sedangkan D30 -> D50 diberi warna biru (asal D30).
            if src_group_fallback and src_group_fallback != target_key:
                input_source_key = normalize_key(src_group_fallback)
                if input_source_key:
                    moved_from = input_source_key

            moved_to = target_key if moved_from else None

            # Tambah ke tabel target
            article_target_rows = [r for r in target_table.rows
                                   if r.article.strip().upper() == art_code]
            existing_target = next((r for r in article_target_rows
                                    if r.colour.strip().upper() == actual_colour.strip().upper()), None)
            if existing_target:
                existing_target.harga_awal = harga_awal
                existing_target.harga_diskon = harga_diskon
                if moved_from:
                    existing_target.moved_from = moved_from
                    existing_target.moved_to = target_key
                print(f"  [UPDATE] {art_code}/{actual_colour} di {target_store_key}: "
                      f"harga_awal={harga_awal}, harga_diskon={harga_diskon}")
                continue

            # Jika warna input belum ada di tabel tujuan, cari satu baris
            # artikel yang warnanya tidak dikenal di master DB. Ganti hanya
            # baris tersebut; warna lain yang valid tetap dipertahankan.
            if in_colour is not None and article_target_rows:
                valid_db_colours = {
                    re.sub(r'[^A-Z0-9]', '', str(value).strip().upper())
                    for key, value in (
                        [(k, k) for k in q_records.keys()]
                        + [(k, v.get('colour', '')) for k, v in q_records.items()]
                    )
                    if value
                }
                stale_target_rows = [
                    r for r in article_target_rows
                    if re.sub(r'[^A-Z0-9]', '', r.colour.strip().upper())
                    not in valid_db_colours
                ]

                # Jika hanya ada satu baris artikel di target, itu juga boleh
                # dikoreksi meskipun warna lama masih tercatat di DB.
                candidates_to_replace = stale_target_rows or (
                    article_target_rows if len(article_target_rows) == 1 else []
                )

                if len(candidates_to_replace) == 1:
                    stale_target = candidates_to_replace[0]
                    old_colour = stale_target.colour
                    stale_target.colour = actual_colour
                    stale_target.harga_awal = harga_awal
                    stale_target.harga_diskon = harga_diskon
                    if moved_from:
                        stale_target.moved_from = moved_from
                        stale_target.moved_to = target_key
                    print(f"  [COLOR-CORRECTION] {art_code}: '{old_colour}' -> "
                          f"'{actual_colour}' sesuai input di tabel tujuan {target_store_key}.")
                    continue

            target_table.rows.append(PromoRow(
                article=rec['article'],
                colour=actual_colour,
                harga_awal=harga_awal,
                harga_diskon=harga_diskon,
                moved_from=moved_from,
                moved_to=moved_to,
            ))
            print(f"  [ADD ] {art_code}/{actual_colour} -> {target_store_key} "
                  f"harga_awal={harga_awal} harga_diskon={harga_diskon}"
                  + (f" [dari {moved_from}]" if moved_from else ""))

# ---------------------------------------------------------------------------
# Pencocokan harga: Template vs Database (dijalankan SETELAH artikel
# ditempatkan di tabel akhirnya, mis. setelah apply_input())
#
# Untuk SETIAP baris di promo_groups (bukan cuma artikel yang ada di file
# input instruksi), harga_awal dibandingkan dengan retail_price terkini di
# database. Kalau berbeda:
#   - harga_awal & harga_diskon baris tsb DIPERBARUI mengikuti database
#   - baris ditandai (harga_disesuaikan = True) supaya kolom Article-nya
#     diwarnai abu-abu muda saat ditulis ke Excel (lihat _xl_write_sheet()).
# ---------------------------------------------------------------------------

def build_query_data_for_brands(brands: list) -> dict:
    """
    Jalankan execute_query() untuk setiap brand di 'brands' dan gabungkan
    hasilnya menjadi satu query_data besar yang mencakup SEMUA artikel brand
    tersebut di database -- bukan cuma artikel yang ada di file input.
    Dipakai sebagai basis pencocokan harga template vs database.
    """
    query_data = {}
    for brand in brands:
        result = execute_query(brand)
        for art_key, colours in result.items():
            if art_key not in query_data:
                query_data[art_key] = {}
            query_data[art_key].update(colours)
    return query_data


def _lookup_db_price(row: PromoRow, q_records: dict) -> Optional[int]:
    """
    Ambil harga_awal yang seharusnya berlaku utk baris ini menurut database.
    Return None kalau tidak bisa disimpulkan dgn yakin (data tidak lengkap/
    tidak konsisten), supaya baris tsb dibiarkan apa adanya (aman).
    """
    if row.colour.strip().upper() == 'ALL COLOUR':
        # Row 'All Colour' -> semua warna di DB seharusnya seharga sama.
        # Kalau ternyata beda-beda, jangan tebak -- biarkan saja, jangan diubah.
        db_prices = {r['retail_price'] for r in q_records.values() if r['retail_price']}
        if len(db_prices) != 1:
            return None
        return next(iter(db_prices))

    rec = q_records.get(row.colour.strip().upper())
    if not rec or not rec['retail_price']:
        return None
    return rec['retail_price']


def sync_harga_dengan_database(promo_groups: dict, query_data: dict,
                               toleransi: int = 0) -> list:
    """
    Cek harga_awal SETIAP baris di promo_groups (semua artikel, bukan cuma
    yang ada di file input) terhadap harga terkini di database. Kalau beda
    lebih dari 'toleransi', harga_awal & harga_diskon baris tsb DIPERBARUI
    mengikuti database, dan baris ditandai harga_disesuaikan=True (dipakai
    utk mewarnai kolom Article abu-abu muda di Excel).

    Panggil fungsi ini SETELAH apply_input() + recount(), sebelum
    write_promolist_excel(), supaya artikel sudah berada di tabel akhirnya
    saat dicek.

    Parameter
    ---------
    promo_groups : dict
        Hasil scan_excel() yang sudah diproses apply_input().
    query_data : dict
        Hasil build_query_data_for_brands() -- mencakup semua artikel brand
        terkait, bukan cuma yang ada di file input instruksi.
    toleransi : int
        Selisih harga (Rupiah) yang masih dianggap "cocok". Default 0.

    Return
    ------
    list[dict] : ringkasan setiap baris yang harganya disesuaikan, dengan
    field: sheet, tabel, article, colour, harga_lama, harga_baru, selisih.
    """
    changes = []

    for sheet_name, tables in promo_groups.items():
        for table_key, table in tables.items():
            for row in table.rows:
                q_records = query_data.get(row.article.upper())
                if not q_records:
                    continue  # artikel tidak ada di database, biarkan apa adanya

                db_price = _lookup_db_price(row, q_records)
                if db_price is None:
                    continue

                selisih = (row.harga_awal or 0) - db_price
                if abs(selisih) > toleransi:
                    harga_lama = row.harga_awal
                    row.harga_awal        = db_price
                    row.harga_diskon      = compute_harga_diskon(db_price, table.discount_key)
                    row.harga_disesuaikan = True

                    changes.append({
                        'sheet': sheet_name, 'tabel': table_key,
                        'article': row.article, 'colour': row.colour,
                        'harga_lama': harga_lama, 'harga_baru': db_price,
                        'selisih': selisih,
                    })
                    print(f"  [SYNC] {row.article}/{row.colour} [{sheet_name}/{table_key}]: "
                          f"harga_awal {harga_lama} -> {db_price} (selisih {selisih:+}), "
                          f"harga_diskon dihitung ulang -> {row.harga_diskon}")

    return changes


def print_sync_summary(changes: list):
    """Cetak ringkasan hasil sync_harga_dengan_database() ke console."""
    if not changes:
        print("\n[SYNC] Semua harga di template sudah sesuai dengan database. "
              "Tidak ada yang disesuaikan.")
        return
    print(f"\n=== {len(changes)} BARIS DISESUAIKAN MENGIKUTI HARGA DATABASE ===")
    for c in changes:
        print(f"  [{c['sheet']}/{c['tabel']}] {c['article']}/{c['colour']}: "
              f"{c['harga_lama']} -> {c['harga_baru']} (selisih {c['selisih']:+})")



# ---------------------------------------------------------------------------
# Excel output
# ---------------------------------------------------------------------------
import math
from datetime import date
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

_XL_COLS_PER_GROUP = 4   # Article | Colour | Harga Awal | Harga Diskon
_XL_GAP_COLS       = 1
_XL_GROUPS_PER_ROW = 3
_XL_TOTAL_COLS     = _XL_GROUPS_PER_ROW * (_XL_COLS_PER_GROUP + _XL_GAP_COLS) - _XL_GAP_COLS  # 14
_XL_GROUP_WIDTHS   = [20, 16, 14, 14, 2]

_XC = {
    'title_bg':  '1F3864', 'title_fg':  'FFFFFF',
    'meta_bg':   'D6E4F0', 'meta_fg':   '000000',
    'promo_bg':  '2E75B6', 'promo_fg':  'FFFFFF',
    'hdr_bg':    'BDD7EE', 'hdr_fg':    '000000',
    'row_odd':   'FFFFFF', 'row_even':  'F2F2F2',
}

# Warna baris berdasarkan promo asal (moved_from)
MOVE_COLOURS: dict = {
    'NML':  'FFFF00',   # kuning
    'D20':  'FF9900',   # oranye
    'D30':  '00B0F0',   # biru muda
    'D40':  '92D050',   # hijau
    'D50':  'FF0000',   # merah
    'D60':  '201654',   # biru tua
    'SP':   '7030A0',   # ungu (semua SP*)
}

# Warna abu-abu muda utk kolom Article pada baris yang harganya disesuaikan
# mengikuti database (lihat sync_harga_dengan_database()).
HARGA_DISESUAIKAN_COLOR = 'D9D9D9'

def _xl_fill(hex_code: str) -> PatternFill:
    return PatternFill('solid', start_color=hex_code, end_color=hex_code)

def _xl_border() -> Border:
    s = Side(style='thin')
    return Border(top=s, bottom=s, left=s, right=s)

def _promo_table_order(tbl) -> tuple:
    """
    Urutan tabel promo di sheet: semua DISC dulu (persen kecil -> besar),
    lalu SP (nominal kecil -> besar), sisanya di akhir. Tabel dengan diskon
    sama (mis. Ladies Bag vs Mens Bag) tetap pada urutan aslinya.
    """
    key = tbl.discount_key if hasattr(tbl, 'discount_key') else tbl.get('discount_key', '')
    m = re.fullmatch(r'D(\d+)', key or '', re.IGNORECASE)
    if m:
        return (0, int(m.group(1)))
    m = re.fullmatch(r'SP(\d+)K?', key or '', re.IGNORECASE)
    if m:
        return (1, int(m.group(1)))
    return (2, 0)

def _xl_write_sheet(wb: Workbook, sheet_name: str, tables: dict, periode: str,
                    brand: str = DEFAULT_BRAND):
    ws = wb.create_sheet(title=sheet_name)

    col_widths = (_XL_GROUP_WIDTHS * _XL_GROUPS_PER_ROW)[:_XL_TOTAL_COLS]
    for i, w in enumerate(col_widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    row = 1

    # Judul
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=_XL_TOTAL_COLS)
    c = ws.cell(row=row, column=1, value=f"PROMLIST CORNER SALE {brand_full_name(brand)}")
    c.font      = Font(name='Arial', bold=True, size=14, color=_XC['title_fg'])
    c.fill      = _xl_fill(_XC['title_bg'])
    c.alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[row].height = 22
    row += 1

    # BERLAKU DI
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=2)
    c = ws.cell(row=row, column=1, value="BERLAKU DI")
    c.font = Font(name='Arial', bold=True, size=10); c.fill = _xl_fill(_XC['meta_bg'])
    c.alignment = Alignment(horizontal='left', vertical='center')
    ws.merge_cells(start_row=row, start_column=3, end_row=row, end_column=_XL_TOTAL_COLS)
    c = ws.cell(row=row, column=3, value=": ALL STORE, COUNTER, BAZAR, WEBSITE, MARKETPLACE")
    c.font = Font(name='Arial', size=10); c.fill = _xl_fill(_XC['meta_bg'])
    c.alignment = Alignment(horizontal='left', vertical='center')
    ws.row_dimensions[row].height = 16
    row += 1

    # PERIODE
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=2)
    c = ws.cell(row=row, column=1, value="PERIODE")
    c.font = Font(name='Arial', bold=True, size=10); c.fill = _xl_fill(_XC['meta_bg'])
    c.alignment = Alignment(horizontal='left', vertical='center')
    ws.merge_cells(start_row=row, start_column=3, end_row=row, end_column=_XL_TOTAL_COLS)
    c = ws.cell(row=row, column=3, value=f": {periode}")
    c.font = Font(name='Arial', size=10); c.fill = _xl_fill(_XC['meta_bg'])
    c.alignment = Alignment(horizontal='left', vertical='center')
    ws.row_dimensions[row].height = 16
    row += 1

    row += 1  # baris kosong

    hdr_labels = ['Article', 'Colour', 'Harga Awal', 'Harga Diskon']

    for tbl in sorted(tables.values(), key=_promo_table_order):
        rows_data = tbl.rows if hasattr(tbl, 'rows') else tbl.get('rows', [])
        # Data di setiap tabel selalu urut abjad (Article, lalu Colour),
        # termasuk saat output dibuat tanpa file input (tanpa recount()).
        rows_data = sorted(rows_data, key=lambda d: (
            (d.article if hasattr(d, 'article') else d['article']).upper(),
            (d.colour if hasattr(d, 'colour') else d['colour']).upper(),
        ))
        if not rows_data:
            continue

        promo_name = tbl.promo_name if hasattr(tbl, 'promo_name') else tbl.get('promo_name', '')

        # Header nama promo: hanya sheet LADIES/MENS/HANDBAG yang brand-nya
        # disesuaikan, sheet lain apa adanya dari template / _build_promo_name
        if sheet_name.upper() in BRAND_HEADER_SHEETS:
            promo_name = apply_brand_to_header(promo_name, brand)

        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=_XL_TOTAL_COLS)
        c = ws.cell(row=row, column=1, value=promo_name.upper())
        c.font      = Font(name='Arial', bold=True, size=11, color=_XC['promo_fg'])
        c.fill      = _xl_fill(_XC['promo_bg'])
        c.alignment = Alignment(horizontal='center', vertical='center')
        c.border    = _xl_border()
        ws.row_dimensions[row].height = 18
        row += 1

        # Header kolom
        for g in range(_XL_GROUPS_PER_ROW):
            base = g * (_XL_COLS_PER_GROUP + _XL_GAP_COLS) + 1
            for ci, label in enumerate(hdr_labels):
                c = ws.cell(row=row, column=base + ci, value=label)
                c.font      = Font(name='Arial', bold=True, size=9, color=_XC['hdr_fg'])
                c.fill      = _xl_fill(_XC['hdr_bg'])
                c.alignment = Alignment(horizontal='center', vertical='center')
                c.border    = _xl_border()
        ws.row_dimensions[row].height = 15
        row += 1

        # Baris data
        n         = len(rows_data)
        n_per_col = math.ceil(n / _XL_GROUPS_PER_ROW)

        for r_idx in range(n_per_col):
            base_fill = _XC['row_odd'] if r_idx % 2 == 0 else _XC['row_even']
            for g in range(_XL_GROUPS_PER_ROW):
                data_idx = g * n_per_col + r_idx
                base     = g * (_XL_COLS_PER_GROUP + _XL_GAP_COLS) + 1
                if data_idx < n:
                    d = rows_data[data_idx]
                    if hasattr(d, 'article'):   # PromoRow dataclass
                        art  = d.article
                        col  = d.colour
                        ha   = d.harga_awal
                        hd   = int(d.harga_diskon) if d.harga_diskon is not None else None
                        mf   = d.moved_from
                        harga_disesuaikan = getattr(d, 'harga_disesuaikan', False)
                    else:                        # dict
                        art  = d['article']
                        col  = d['colour']
                        ha   = d['harga_awal']
                        hd   = d['harga_diskon']
                        mf   = d.get('moved_from')
                        harga_disesuaikan = d.get('harga_disesuaikan', False)
                    vals = [art, col, ha, hd]
                    # Warna baris: hanya artikel yang dipindahkan (moved_from terisi)
                    if mf:
                        # Cari warna: cek exact key dulu, lalu prefix 'SP'
                        row_color = MOVE_COLOURS.get(mf)
                        if row_color is None:
                            row_color = MOVE_COLOURS.get(mf[:2]) if len(mf) >= 2 else None
                        row_color = row_color or base_fill
                    else:
                        row_color = base_fill
                    # Bold hanya pada kolom Colour (ci=1) jika bukan All Colour
                    is_specific = col.strip().upper() != 'ALL COLOUR'
                    for ci, v in enumerate(vals):
                        c = ws.cell(row=row, column=base + ci, value=v)
                        c.font      = Font(name='Arial', size=9,
                                          bold=(ci == 1 and is_specific))
                        # Kolom Article (ci==0) diwarnai abu-abu muda kalau
                        # harganya baru saja disesuaikan mengikuti database,
                        # menimpa warna baris normal HANYA utk sel ini.
                        if ci == 0 and harga_disesuaikan:
                            c.fill = _xl_fill(HARGA_DISESUAIKAN_COLOR)
                        else:
                            c.fill = _xl_fill(row_color)
                        c.alignment = Alignment(
                            horizontal='center' if ci >= 2 else 'left',
                            vertical='center',
                        )
                        c.border = _xl_border()
                        if ci >= 2 and v is not None:
                            c.number_format = '#,##0'
                else:
                    for ci in range(_XL_COLS_PER_GROUP):
                        c = ws.cell(row=row, column=base + ci)
                        c.fill   = _xl_fill(base_fill)
                        c.border = _xl_border()
            ws.row_dimensions[row].height = 14
            row += 1

        row += 1  # baris kosong antar tabel

    ws.freeze_panes = 'A5'

    # ── Legend warna: baris vertikal di kolom 16 ─────────────────────────
    leg_col = _XL_TOTAL_COLS + 2   # kolom 16
    ws.column_dimensions[get_column_letter(leg_col)].width = 24

    today_str = str(date.today().day) + date.today().strftime(' %B %Y')
    legend_entries = [
        ('NML',  f'PER {today_str} dari NML'),
        ('D20',  f'PER {today_str} dari D20'),
        ('D30',  f'PER {today_str} dari D30'),
        ('D40',  f'PER {today_str} dari D40'),
        ('D50',  f'PER {today_str} dari D50'),
        ('SP',   f'PER {today_str} dari SP'),
    ]

    for leg_row, (key, label) in enumerate(legend_entries, start=1):
        bg = MOVE_COLOURS[key]
        c  = ws.cell(row=leg_row, column=leg_col, value=label)
        c.font      = Font(name='Arial', size=9, bold=True)
        c.fill      = _xl_fill(bg)
        c.alignment = Alignment(horizontal='left', vertical='center')
        c.border    = _xl_border()
        ws.row_dimensions[leg_row].height = 16

# ---------------------------------------------------------------------------
# Copy sheet khusus (SP COUNTER MP) dari template ke workbook output
# ---------------------------------------------------------------------------
SPECIAL_SHEETS = ['SP COUNTER MP']  # tambahkan nama sheet lain di sini jika perlu

def copy_special_sheets(template_path: str, dst_wb: Workbook, sheet_names: list = None) -> list:
    """
    Mengecek file template untuk sheet khusus (misal: 'SP COUNTER MP') dan
    menyalin sheet tersebut apa adanya (data + format) ke workbook output.

    Parameter
    ---------
    template_path : str
        Path ke file template Excel sumber.
    dst_wb : openpyxl.Workbook
        Workbook output yang sedang dibangun (sebelum di-save).
    sheet_names : list[str], optional
        Daftar nama sheet yang ingin disalin. Default: SPECIAL_SHEETS.

    Return
    ------
    list[str] : nama-nama sheet yang berhasil disalin.
    """
    names = sheet_names or SPECIAL_SHEETS
    copied = []

    src_wb = openpyxl.load_workbook(template_path, data_only=False)

    for name in names:
        if name not in src_wb.sheetnames:
            continue

        src_ws = src_wb[name]
        dst_ws = dst_wb.create_sheet(title=name[:31])

        # Column widths & hidden state
        for col_letter, dim in src_ws.column_dimensions.items():
            dst_ws.column_dimensions[col_letter].width = dim.width
            dst_ws.column_dimensions[col_letter].hidden = dim.hidden

        # Row heights
        for row_idx, dim in src_ws.row_dimensions.items():
            dst_ws.row_dimensions[row_idx].height = dim.height

        # Cell values + full style
        for row in src_ws.iter_rows():
            for cell in row:
                new_cell = dst_ws.cell(row=cell.row, column=cell.column, value=cell.value)
                if cell.has_style:
                    new_cell.font = copy(cell.font)
                    new_cell.fill = copy(cell.fill)
                    new_cell.border = copy(cell.border)
                    new_cell.alignment = copy(cell.alignment)
                    new_cell.number_format = cell.number_format
                    new_cell.protection = copy(cell.protection)

        # Merged cells
        for merged_range in src_ws.merged_cells.ranges:
            dst_ws.merge_cells(str(merged_range))

        # View properties
        dst_ws.freeze_panes = src_ws.freeze_panes
        dst_ws.sheet_view.showGridLines = src_ws.sheet_view.showGridLines

        copied.append(name)

    return copied


def write_promolist_excel(promo_groups: dict, output_path: str, template_path: str = None,
                          brand: str = DEFAULT_BRAND):
    """
    Buat file Excel promolist corner sale dari promo_groups.

    Menerima promo_groups langsung dari apply_input() / recount() —
    berisi PromoTable dataclass maupun dict, keduanya didukung.

    Parameter
    ---------
    promo_groups : dict
        { sheet_name: { discount_key: PromoTable } }
    output_path  : str
        Path file .xlsx yang akan dibuat.
    template_path : str, optional
        Path ke file template Excel sumber. Jika ada sheet khusus
        (lihat SPECIAL_SHEETS), sheet tersebut akan disalin apa adanya
        ke file output.
    brand : str, optional
        Kode brand dari file input ('EB', 'TR', 'VB'). Dipakai untuk judul
        "PROMLIST CORNER SALE {BRAND}" dan header setiap tabel promo.
    """
    today = date.today()
    periode = str(today.day) + today.strftime(' %B %Y') + ' - 31 DEC 2099'

    wb = Workbook()
    wb.remove(wb.active)

    for sheet_name, tables in promo_groups.items():
        safe_name = re.sub(r'[\\/*?:\[\]]', '-', sheet_name)[:31]
        _xl_write_sheet(wb, safe_name, tables, periode, brand=brand)

    if template_path:
        copied = copy_special_sheets(template_path, wb)
        if copied:
            print(f"  [EXCEL] Sheet khusus disalin: {', '.join(copied)}")

    wb.save(output_path)
    print(f"  [EXCEL] Promolist disimpan: {output_path}")

