# -*- coding: utf-8 -*-
"""
writer.py — Create the output Logo_Güncellemeler_YYYYMMDD_HHMMSS.xlsx workbook.
"""

import datetime
import sys
import os

import openpyxl
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

sys.path.insert(0, os.path.dirname(__file__))
from config import SHEET_CONFIGS


def make_output_filename() -> str:
    """Return a timestamped output filename."""
    ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    return f"Logo_Güncellemeler_{ts}.xlsx"


def _build_formula(template: str, r: int) -> str:
    """Replace all {r} placeholders with the actual Excel row number."""
    return template.replace("{r}", str(r))


def _auto_size_columns(ws):
    """Set column widths based on the maximum content length, capped at 80."""
    for col_cells in ws.columns:
        max_len = 0
        col_letter = col_cells[0].column_letter
        for cell in col_cells:
            if cell.value is not None:
                try:
                    max_len = max(max_len, len(str(cell.value)))
                except Exception:
                    pass
        ws.column_dimensions[col_letter].width = min(max_len + 4, 80)


def write_output(sheet_data: dict, output_path: str) -> None:
    """
    Create the output workbook at output_path.

    For each sheet defined in SHEET_CONFIGS (in order):
    - Write a bold header row (row 1)
    - Write data rows starting at row 2
    - Append the SQL formula column for each data row
    - Auto-size all columns

    sheet_data: {sheet_name: [row_list, ...]} as returned by mapper.map_to_sheets()
    """
    wb = openpyxl.Workbook()
    bold = Font(bold=True)

    # Remove the default empty sheet
    default_sheet = wb.active
    wb.remove(default_sheet)

    for cfg in SHEET_CONFIGS:
        sheet_name = cfg["name"]
        ws = wb.create_sheet(title=sheet_name)

        # Build full header list: data cols + SQL col
        sql_col_letter = cfg["sql_col"]
        sql_header = "SQL"
        all_headers = list(cfg["headers"]) + [sql_header]

        # Write bold header row
        for col_idx, header in enumerate(all_headers, start=1):
            cell = ws.cell(row=1, column=col_idx, value=header)
            cell.font = bold

        rows = sheet_data.get(sheet_name, [])
        for row_offset, row_values in enumerate(rows):
            excel_row = row_offset + 2  # data starts at row 2

            # Write data values
            for col_idx, value in enumerate(row_values, start=1):
                ws.cell(row=excel_row, column=col_idx, value=value)

            # Write SQL formula in the sql_col
            # sql_col is defined as an Excel column letter relative to the sheet,
            # so we calculate the 1-based column index from the letter
            sql_col_idx = len(cfg["output_cols"]) + 1
            formula = _build_formula(cfg["sql_template"], excel_row)
            ws.cell(row=excel_row, column=sql_col_idx, value=formula)

        _auto_size_columns(ws)

    wb.save(output_path)
