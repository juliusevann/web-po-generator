import argparse
import sys
from pathlib import Path

from functs import (
    scan_excel,
    print_summary,
    load_input_file,
    execute_query,
    apply_input,
    recount,
    write_promolist_excel,
    print_duplicate_warnings,
)


def parse_args():
    parser = argparse.ArgumentParser(
        description="Corner Sale Promolist Generator"
    )
    parser.add_argument("--template-file", required=True)
    parser.add_argument("--input-file", required=True)
    parser.add_argument("--output-dir", required=True)
    return parser.parse_args()


def main():
    args = parse_args()

    template_path = Path(args.template_file).resolve()
    input_path = Path(args.input_file).resolve()
    output_dir = Path(args.output_dir).resolve()

    # Validasi file
    for path, label in [
        (template_path, "Template"),
        (input_path, "Input"),
    ]:
        if not path.is_file():
            print(f"[ERROR] File {label} tidak ditemukan: {path}")
            return 1

        if path.suffix.lower() not in {".xlsx", ".xlsm"}:
            print(f"[ERROR] Format file {label} harus .xlsx atau .xlsm")
            return 1

    output_dir.mkdir(parents=True, exist_ok=True)

    print(f"Template : {template_path.name}")
    print(f"Input    : {input_path.name}")
    print(f"Output   : {output_dir}")

    # Tahap 1: scan template yang diunggah user
    print(f"\n{'=' * 60}")
    print(f"Scanning template: {template_path.name}")

    try:
        promo_groups = scan_excel(str(template_path))
        print_summary(promo_groups)
        print_duplicate_warnings(promo_groups, "template")
    except Exception as e:
        print(f"[ERROR] Gagal scan template: {e}")
        return 1

    # Tahap 2: baca file input
    print(f"\n{'=' * 60}")
    print(f"Membaca input: {input_path.name}")

    try:
        input_rows = load_input_file(str(input_path))
    except Exception as e:
        print(f"[ERROR] Gagal membaca file input: {e}")
        return 1

    print(f"Instruksi ditemukan: {len(input_rows)} baris")

    if not input_rows:
        print("[ERROR] Tidak ada instruksi promo di file input.")
        return 1

    brands = list(dict.fromkeys(
        row["brand"] for row in input_rows if row.get("brand")
    ))

    print(f"Brand ditemukan: {brands}")

    if not brands:
        print("[ERROR] Brand tidak ditemukan di file input.")
        return 1

    brand_code = brands[0].upper()

    if len(brands) > 1:
        print(
            "[WARN] Ada lebih dari satu brand. "
            f"Header menggunakan brand pertama: {brand_code}"
        )

    # Tahap 3: query database
    print("\nMengeksekusi query SQL...")
    query_data = {}

    for brand in brands:
        try:
            result = execute_query(brand)

            for article_key, colours in result.items():
                query_data.setdefault(article_key, {}).update(colours)

            print(
                f"[OK] Brand '{brand}': "
                f"{len(result)} artikel diterima"
            )
        except Exception as e:
            print(f"[ERROR] Gagal query brand '{brand}': {e}")
            return 1

    print(f"Total artikel dari DB: {len(query_data)}")

    # Tahap 4: apply promo dan simpan hasil
    try:
        apply_input(promo_groups, input_rows, query_data)
        recount(promo_groups)
        print_duplicate_warnings(promo_groups, "output")

        output_file = output_dir / (
            f"{template_path.stem}_promolist.xlsx"
        )

        write_promolist_excel(
            promo_groups,
            str(output_file),
            template_path=template_path,
            brand=brand_code,
        )

        print(f"\n[OK] Output berhasil dibuat: {output_file}")
    except Exception as e:
        print(f"[ERROR] Gagal membuat output: {e}")
        return 1

    print("\nSelesai.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
