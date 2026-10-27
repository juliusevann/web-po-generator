import sys
from pathlib import Path
from functs import (
    scan_excel, print_summary,
    load_input_file, execute_query,
    apply_input, recount,
    write_promolist_excel, print_duplicate_warnings,
)

TEMPLATE_DIR = r'D:\EVERBEST\NEW_GENERATOR\webapp\corner_sale\template'
OUTPUT_DIR   = r'D:\EVERBEST\NEW_GENERATOR\webapp\corner_sale\output'
INPUT_DIR    = r'D:\EVERBEST\NEW_GENERATOR\webapp\corner_sale\input'

if __name__ == '__main__':
    template_dir = Path(TEMPLATE_DIR)
    output_dir   = Path(OUTPUT_DIR)
    input_dir    = Path(INPUT_DIR)

    # --- Validasi folder ---
    for d, label in [(template_dir, 'Template'), (input_dir, 'Input')]:
        if not d.exists():
            print(f"[ERROR] Folder {label} tidak ditemukan: {d}")
            sys.exit(1)

    output_dir.mkdir(parents=True, exist_ok=True)

    # =========================================================
    # TAHAP 1 — Scan file Excel template -> promo_groups (memory)
    # =========================================================
    excel_files = [f for f in list(template_dir.glob('*.xlsx')) + list(template_dir.glob('*.xlsm'))
                   if not f.name.startswith('~$')]

    if not excel_files:
        print(f"[ERROR] Tidak ada file .xlsx / .xlsm di folder: {template_dir}")
        sys.exit(1)

    print(f"Template folder : {template_dir}")
    print(f"Output folder   : {output_dir}")
    print(f"Input folder    : {input_dir}")
    print(f"File template   : {len(excel_files)} file ditemukan")

    # { excel_stem: promo_groups_dict }
    scanned = {}

    for excel_path in excel_files:
        print(f"\n{'='*60}")
        print(f"Scanning : {excel_path.name}")

        try:
            data = scan_excel(str(excel_path))
            print_summary(data)
            print_duplicate_warnings(data, 'template')
            scanned[excel_path.stem] = data
        except Exception as e:
            print(f"[ERROR] Gagal memproses {excel_path.name}: {e}")

    # =========================================================
    # TAHAP 2 — Baca input, hit SQL, apply ke promo_groups,
    #           langsung output Excel
    # =========================================================
    input_files = list(input_dir.glob('*.xlsx')) + list(input_dir.glob('*.xlsm'))
    # Lewati file lock Excel (~$...) yang muncul saat file input sedang dibuka
    input_file  = next((f for f in input_files
                        if not f.name.lower().startswith('hasil_query') and not f.name.startswith('~$')), None)

    if not input_file:
        print(f"\n[INFO] Tidak ada file input di: {input_dir}")
        print("[INFO] Menyimpan hasil scan template tanpa perubahan...")

        for stem, promo_groups in scanned.items():
            excel_out = output_dir / (stem + '_promolist.xlsx')
            template_path = template_dir / (stem + '.xlsx')
            if not template_path.exists():
                template_path = template_dir / (stem + '.xlsm')
            print(f"\n{'='*60}")
            print(f"Output : {excel_out.name}")
            try:
                write_promolist_excel(promo_groups, str(excel_out), template_path=str(template_path))
            except Exception as e:
                print(f"[ERROR] {stem}: {e}")

    else:
        print(f"\n{'='*60}")
        print(f"TAHAP 2 — Apply Input")
        print(f"Input file : {input_file.name}")

        # Baca instruksi dari file input
        input_rows = load_input_file(str(input_file))
        print(f"  Instruksi ditemukan : {len(input_rows)} baris")

        # Ambil semua brand unik
        brands = list(dict.fromkeys(r['brand'] for r in input_rows if r['brand']))
        print(f"  Brand ditemukan     : {brands}")
        brand_code = brands[0].upper() if brands else 'EB'
        if len(brands) > 1:
            print(f"  [WARN] Lebih dari 1 brand di input, header memakai brand pertama: {brand_code}")

        # Eksekusi query SQL per brand
        print(f"\n  Mengeksekusi query SQL...")
        query_data = {}
        for brand in brands:
            try:
                result = execute_query(brand)
                for art_key, colours in result.items():
                    if art_key not in query_data:
                        query_data[art_key] = {}
                    query_data[art_key].update(colours)
                print(f"  [OK] Brand '{brand}': {len(result)} artikel diterima")
            except Exception as e:
                print(f"  [ERROR] Gagal query brand '{brand}': {e}")
                sys.exit(1)

        print(f"  Total artikel dari DB : {len(query_data)}")

        # Apply ke setiap promo_groups hasil scan, lalu output Excel
        for stem, promo_groups in scanned.items():
            excel_out = output_dir / (stem + '_promolist.xlsx')
            template_path = template_dir / (stem + '.xlsx')
            if not template_path.exists():
                template_path = template_dir / (stem + '.xlsm')
            print(f"\n{'='*60}")
            print(f"Memproses : {stem}")
            print(f"Output    : {excel_out.name}")

            try:
                apply_input(promo_groups, input_rows, query_data)
                recount(promo_groups)
                print_duplicate_warnings(promo_groups, 'output')
                write_promolist_excel(promo_groups, str(excel_out), template_path=str(template_path),
                                      brand=brand_code)
            except Exception as e:
                print(f"  [ERROR] {stem}: {e}")

    print(f"\n{'='*60}")
    print("Selesai.")