#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
main.py — CLI entry point for logo_updater.

Usage:
    python main.py kaynak.xlsx
    python main.py dosya1.xlsx dosya2.xlsx dosya3.xlsx
    python main.py --input-dir ./kaynaklar/
"""

import argparse
import glob as glob_mod
import os
import sys

sys.path.insert(0, os.path.dirname(__file__))
import reader
import mapper
import writer
from config import SHEET_CONFIGS


def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        prog="logo_updater",
        description="Tedarikçi Excel dosyalarından Logo_Güncellemeler.xlsx üretir.",
    )
    p.add_argument(
        "files",
        nargs="*",
        metavar="DOSYA.xlsx",
        help="İşlenecek kaynak Excel dosyaları",
    )
    p.add_argument(
        "--input-dir",
        metavar="DİZİN",
        help="Bir dizindeki tüm .xlsx dosyalarını işle",
    )
    return p


def resolve_input_files(args) -> list:
    """Resolve command-line arguments to a list of existing .xlsx file paths."""
    if args.input_dir:
        pattern = os.path.join(args.input_dir, "*.xlsx")
        files = sorted(glob_mod.glob(pattern))
        if not files:
            print(f"HATA: '{args.input_dir}' dizininde .xlsx dosyası bulunamadı.")
            sys.exit(1)
        return files

    if args.files:
        missing = [f for f in args.files if not os.path.isfile(f)]
        if missing:
            for m in missing:
                print(f"HATA: Dosya bulunamadı: {m}")
            sys.exit(1)
        return args.files

    build_parser().print_help()
    sys.exit(0)


def main():
    parser = build_parser()
    args = parser.parse_args()
    input_files = resolve_input_files(args)

    output_filename = writer.make_output_filename()

    all_colored_cells = []
    file_stats = []  # list of (filepath, product_count | None, error_msg | None)

    for filepath in input_files:
        filename = os.path.basename(filepath)
        try:
            cells = reader.read_colored_cells(filepath)
            all_colored_cells.extend(cells)
            # Count unique product codes in this file's cells
            unique_products = len({cc.product_code for cc in cells})
            file_stats.append((filename, unique_products, None))
        except ValueError as e:
            file_stats.append((filename, None, str(e)))
        except Exception as e:
            file_stats.append((filename, None, f"Beklenmeyen hata: {e}"))

    # Map and write even if some files failed
    sheet_data = mapper.map_to_sheets(all_colored_cells)
    writer.write_output(sheet_data, output_filename)

    # Print summary
    print()
    max_name_len = max((len(name) for name, _, _ in file_stats), default=10)
    for filename, count, error in file_stats:
        if error:
            print(f"  ✗ {filename:<{max_name_len}}  → HATA: {error}")
        else:
            print(f"  ✓ {filename:<{max_name_len}}  → {count} ürün işlendi")

    print("  " + "─" * 55)

    total_products = sum(c for _, c, _ in file_stats if c is not None)
    num_sheets = len(SHEET_CONFIGS)
    print(f"  Toplam: {total_products} ürün | {num_sheets} sheet | Çıktı: {output_filename}")
    print()


if __name__ == "__main__":
    main()
