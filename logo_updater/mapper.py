# -*- coding: utf-8 -*-
"""
mapper.py — Convert a list of ColoredCell objects into a dict of
{sheet_name: [output_row, ...]} ready to be written by writer.py.
"""

import sys
import os

sys.path.insert(0, os.path.dirname(__file__))
from config import SHEET_CONFIGS, PRODUCT_CODE_COL, CATEGORY_SEP


def _is_empty(value) -> bool:
    """Return True if value is None or a blank string."""
    if value is None:
        return True
    if isinstance(value, str) and value.strip() == "":
        return True
    return False


def _get_unique_rows_by_product(colored_cells, trigger_cols):
    """
    Return a dict {product_code: row_values} for each unique product code
    that has at least one colored cell in any of the trigger_cols.
    When the same product code appears multiple times (multiple source files),
    last row_values wins — this is acceptable because the data should be identical.
    """
    rows = {}
    trigger_set = set(trigger_cols)
    for cc in colored_cells:
        if cc.col_idx in trigger_set:
            rows[cc.product_code] = cc.row_values
    return rows


def _build_standard_row(cfg, row_values):
    """Build an output row list for a standard (non-multi_row) sheet config."""
    output_row = []
    for col_letter in cfg["output_cols"]:
        src_idx = cfg["source_map"].get(col_letter)
        if src_idx is None:
            # Fixed value
            output_row.append(cfg.get("fixed_values", {}).get(col_letter, ""))
        else:
            value = row_values[src_idx] if src_idx < len(row_values) else None
            output_row.append("" if value is None else value)
    return output_row


def _build_kategori_row(cfg, row_values):
    """Build the output row for the Kategori sheet (column B = concat of D/E/F/G)."""
    product_code = row_values[PRODUCT_CODE_COL] if PRODUCT_CODE_COL < len(row_values) else ""
    concat_parts = []
    for src_idx in cfg["concat_cols"]:
        v = row_values[src_idx] if src_idx < len(row_values) else None
        if not _is_empty(v):
            concat_parts.append(str(v).strip())
    kategori_kodu = CATEGORY_SEP.join(concat_parts) if concat_parts else ""
    return [
        "" if product_code is None else product_code,
        kategori_kodu,
    ]


def _build_barkod_rows(cfg, product_code, row_values):
    """Expand one source row into up to 3 Barkod rows (URUN / İÇ KOLİ / KOLİ)."""
    output_rows = []
    for birim_label, src_idx in cfg["barcode_map"].items():
        barkod_value = row_values[src_idx] if src_idx < len(row_values) else None
        if not _is_empty(barkod_value):
            output_rows.append([
                "" if product_code is None else product_code,
                birim_label,
                barkod_value,
            ])
    return output_rows


def _build_birim_koli_rows(cfg, product_code, row_values):
    """Expand one source row into up to 2 Birim Koli rows (KOLİ / İÇ KOLİ)."""
    output_rows = []
    for birim_label, src_idx in cfg["birim_map"].items():
        adet_value = row_values[src_idx] if src_idx < len(row_values) else None
        if not _is_empty(adet_value):
            output_rows.append([
                "" if product_code is None else product_code,
                birim_label,
                adet_value,
            ])
    return output_rows


def _build_boyut_rows(cfg, product_code, row_values):
    """Expand one source row into up to 3 Boyut Bulk rows (URUN / İÇ KOLİ / KOLİ)."""
    output_rows = []
    fields = ["en", "boy", "yukseklik", "net", "brut"]
    for unit_label, dim_map in cfg["boyut_map"].items():
        dim_values = []
        any_value = False
        for field in fields:
            src_idx = dim_map[field]
            v = row_values[src_idx] if src_idx < len(row_values) else None
            if not _is_empty(v):
                any_value = True
                dim_values.append(v)
            else:
                dim_values.append("")  # empty → SQL formula will write NULL

        if any_value:
            output_rows.append([
                "" if product_code is None else product_code,
                unit_label,
                dim_values[0],  # EN
                dim_values[1],  # BOY
                dim_values[2],  # YUKSEKLIK
                dim_values[3],  # NET_AGIRLIK
                dim_values[4],  # BRUT_AGIRLIK
            ])
    return output_rows


def map_to_sheets(colored_cells: list) -> dict:
    """
    Convert a list of ColoredCell objects into a dict of sheet data.

    Returns:
        {sheet_name: [row_list, row_list, ...]}
        Each row_list has values for the sheet's output_cols only (no SQL column).
        The SQL formula is added by writer.py.
    """
    sheet_data = {cfg["name"]: [] for cfg in SHEET_CONFIGS}
    # Track seen rows per sheet to avoid duplicates
    seen = {cfg["name"]: set() for cfg in SHEET_CONFIGS}

    for cfg in SHEET_CONFIGS:
        sheet_name = cfg["name"]
        is_multi = cfg.get("multi_row", False)

        if is_multi:
            # Gather unique product rows matching any trigger col
            product_rows = _get_unique_rows_by_product(colored_cells, cfg["trigger_cols"])

            for product_code, row_values in product_rows.items():
                if "barcode_map" in cfg:
                    candidate_rows = _build_barkod_rows(cfg, product_code, row_values)
                elif "birim_map" in cfg:
                    candidate_rows = _build_birim_koli_rows(cfg, product_code, row_values)
                elif "boyut_map" in cfg:
                    candidate_rows = _build_boyut_rows(cfg, product_code, row_values)
                else:
                    candidate_rows = []

                for row in candidate_rows:
                    key = tuple(str(v) for v in row)
                    if key not in seen[sheet_name]:
                        seen[sheet_name].add(key)
                        sheet_data[sheet_name].append(row)

        else:
            # Standard single-output-row sheets
            product_rows = _get_unique_rows_by_product(colored_cells, cfg["trigger_cols"])

            for product_code, row_values in product_rows.items():
                if "concat_cols" in cfg:
                    row = _build_kategori_row(cfg, row_values)
                else:
                    row = _build_standard_row(cfg, row_values)

                # Skip if all non-product-code values are empty
                data_values = row[1:]  # col A is always product code
                if all(_is_empty(v) for v in data_values):
                    continue

                key = tuple(str(v) for v in row)
                if key not in seen[sheet_name]:
                    seen[sheet_name].add(key)
                    sheet_data[sheet_name].append(row)

    return sheet_data
