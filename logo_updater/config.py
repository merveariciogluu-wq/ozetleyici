# -*- coding: utf-8 -*-
"""
config.py — Single source of truth for all constants used across logo_updater.
"""

# Source file constants
SOURCE_SHEET_NAME = "Rapor"
YELLOW_RGB = "FFFFFF00"
CATEGORY_SEP = "/"

# Column letter → 0-based index mapping
COL = {
    "A": 0, "B": 1, "C": 2, "D": 3, "E": 4,
    "F": 5, "G": 6, "H": 7, "I": 8, "J": 9,
    "K": 10, "L": 11, "M": 12, "N": 13, "O": 14,
    "P": 15, "Q": 16, "R": 17, "S": 18, "T": 19,
    "U": 20, "V": 21, "W": 22, "X": 23, "Y": 24,
    "Z": 25, "AA": 26, "AB": 27, "AC": 28, "AD": 29,
    "AE": 30, "AF": 31, "AG": 32, "AH": 33, "AI": 34,
    "AJ": 35, "AK": 36, "AL": 37, "AM": 38, "AN": 39,
    "AO": 40, "AP": 41, "AQ": 42, "AR": 43, "AS": 44,
    "AT": 45, "AU": 46, "AV": 47, "AW": 48, "AX": 49,
    "AY": 50, "AZ": 51, "BA": 52, "BB": 53, "BC": 54,
    "BD": 55, "BE": 56, "BF": 57, "BG": 58,
}

PRODUCT_CODE_COL = COL["J"]

# Boyut Bulk LET formula — triple-quoted so ' and " appear without escaping.
# {r} is replaced at write-time via str.replace, not str.format.
_BOYUT_BULK_TEMPLATE = (
    '=LET(A,A{r},B,B{r},'
    'C,IF(C{r}="","NULL",SUBSTITUTE(ROUND(C{r},4),",",".")),'
    'D,IF(D{r}="","NULL",SUBSTITUTE(ROUND(D{r},4),",",".")),'
    'E,IF(E{r}="","NULL",SUBSTITUTE(ROUND(E{r},4),",",".")),'
    'F,IF(F{r}="","NULL",SUBSTITUTE(ROUND(F{r},4),",",".")),'
    'G,IF(G{r}="","NULL",SUBSTITUTE(ROUND(G{r},4),",",".")),'
    '"UPDATE LG_038_UNITSETF SET WIDTH="&C&", LENGTH="&D'
    '&", HEIGHT="&E&", NETWEIGHT="&F&", GROSSWEIGHT="&G'
    '&" WHERE UNITSETREF=(SELECT UNITSETREF FROM LG_038_UNITSETL'
    ' WHERE UNITSETREF=(SELECT UNITSETREF FROM LG_038_ITEMS'
    " WHERE CODE='\""
    "&A&\""
    "') AND CODE='\""
    "&B&\""
    "')\""
    ")"
)

# Sheet configurations driving both mapper.py and writer.py.
# Each dict keys:
#   name         — output sheet name
#   trigger_cols — 0-based source col indices; a colored cell in any of these triggers a row
#   output_cols  — Excel column letters for data output (left-to-right)
#   headers      — header strings (same length as output_cols)
#   source_map   — {output_col_letter: source_col_index | None}; None = fixed_value
#   fixed_values — {output_col_letter: value} for None entries in source_map
#   sql_col      — Excel column letter where the SQL formula is written
#   sql_template — formula string with {r} as row number placeholder
#   multi_row    — True if one source row can expand to multiple output rows
#   concat_cols  — (Kategori only) list of source col indices to join for col B
#   barcode_map  — (Barkod only) {birim_label: source_col_index}
#   birim_map    — (Birim Koli only) {birim_label: source_col_index}
#   boyut_map    — (Boyut Bulk only) {unit_label: {field: source_col_index}}

SHEET_CONFIGS = [
    # 1. Tedarikçi
    # Excel formula: ="EXEC dbo.Sp_Musteri_Tedarikci_INSERT_038 '" & A{r} & "','" & B{r} & "','" & C{r} & "'," & D{r} & ";"
    {
        "name": "Tedarikçi",
        "trigger_cols": [COL["K"]],
        "output_cols": ["A", "B", "C", "D"],
        "headers": ["Ürün Kodu", "Cari Kodu", "Tedarikçi Ürün Kodu", "Kod Tipi"],
        "source_map": {
            "A": PRODUCT_CODE_COL,
            "B": COL["C"],
            "C": COL["K"],
            "D": None,
        },
        "fixed_values": {"D": 1},
        "sql_col": "E",
        "sql_template": (
            '="EXEC dbo.Sp_Musteri_Tedarikci_INSERT_038 \'"'
            ' & A{r} & "\',\'" & B{r} & "\',\'" & C{r} & "\'," & D{r} & ";"'
        ),
    },
    # 2. Kategori
    # Excel formula: ="UPDATE LG_038_ITEMS SET STGRPCODE = '" & B{r} & "' WHERE CODE = '" & A{r} & "';"
    {
        "name": "Kategori",
        "trigger_cols": [COL["D"], COL["E"], COL["F"], COL["G"]],
        "output_cols": ["A", "B"],
        "headers": ["Ürün Kodu", "Kategori Kodu"],
        "source_map": {
            "A": PRODUCT_CODE_COL,
            "B": None,
        },
        "concat_cols": [COL["D"], COL["E"], COL["F"], COL["G"]],
        "sql_col": "C",
        "sql_template": (
            '="UPDATE LG_038_ITEMS SET STGRPCODE = \'" & B{r} & "\' WHERE CODE = \'" & A{r} & "\';"'
        ),
    },
    # 3. Marka
    {
        "name": "Marka",
        "trigger_cols": [COL["H"]],
        "output_cols": ["A", "B"],
        "headers": ["Ürün Kodu", "Marka REF"],
        "source_map": {"A": PRODUCT_CODE_COL, "B": COL["H"]},
        "sql_col": "C",
        "sql_template": (
            '="UPDATE LG_038_ITEMS SET MARKREF = \'" & B{r} & "\' WHERE CODE = \'" & A{r} & "\';"'
        ),
    },
    # 4. Özel Statü
    {
        "name": "Özel Statü",
        "trigger_cols": [COL["Z"]],
        "output_cols": ["A", "B"],
        "headers": ["Ürün Kodu", "Kodu"],
        "source_map": {"A": PRODUCT_CODE_COL, "B": COL["Z"]},
        "sql_col": "C",
        "sql_template": '="update LG_038_ITEMS set SPECODE = \'"&B{r}&"\' where CODE = \'"&A{r}&"\'"',
    },
    # 5. Lisans
    {
        "name": "Lisans",
        "trigger_cols": [COL["I"]],
        "output_cols": ["A", "B"],
        "headers": ["Stok_Kodu", "Kodu"],
        "source_map": {"A": PRODUCT_CODE_COL, "B": COL["I"]},
        "sql_col": "C",
        "sql_template": '="update LG_038_ITEMS set SPECODE2 = \'"&B{r}&"\' where CODE = \'"&A{r}&"\'"',
    },
    # 6. Cinsiyet
    {
        "name": "Cinsiyet",
        "trigger_cols": [COL["V"]],
        "output_cols": ["A", "B"],
        "headers": ["Ürün Kodu", "Kodu"],
        "source_map": {"A": PRODUCT_CODE_COL, "B": COL["V"]},
        "sql_col": "C",
        "sql_template": '="update LG_038_ITEMS set SPECODE3 = \'"&B{r}&"\' where CODE = \'"&A{r}&"\'"',
    },
    # 7. Yaş Grubu
    {
        "name": "Yaş Grubu",
        "trigger_cols": [COL["X"]],
        "output_cols": ["A", "B"],
        "headers": ["Ürün Kodu", "Kodu"],
        "source_map": {"A": PRODUCT_CODE_COL, "B": COL["X"]},
        "sql_col": "C",
        "sql_template": '="update LG_038_ITEMS set SPECODE4 = \'"&B{r}&"\' where CODE = \'"&A{r}&"\'"',
    },
    # 8. Pazarlama
    {
        "name": "Pazarlama",
        "trigger_cols": [COL["AA"]],
        "output_cols": ["A", "B"],
        "headers": ["Ürün Kodu", "Kodu"],
        "source_map": {"A": PRODUCT_CODE_COL, "B": COL["AA"]},
        "sql_col": "C",
        "sql_template": '="update LG_038_ITEMS set SPECODE5 = \'"&B{r}&"\' where CODE = \'"&A{r}&"\'"',
    },
    # 9. Menşei
    {
        "name": "Menşei",
        "trigger_cols": [COL["Y"]],
        "output_cols": ["A", "B"],
        "headers": ["CODE", "ORIGIN"],
        "source_map": {"A": PRODUCT_CODE_COL, "B": COL["Y"]},
        "sql_col": "C",
        "sql_template": (
            '="UPDATE LG_038_ITEMS SET ORIGIN = \'" & B{r}'
            ' & "\', PRODCOUNTRY = \'" & B{r} & "\' WHERE CODE = \'" & A{r} & "\';"'
        ),
    },
    # 10. Tedarikçi Kodu
    {
        "name": "Tedarikçi Kodu",
        "trigger_cols": [COL["K"]],
        "output_cols": ["A", "B"],
        "headers": ["CODE", "PRODUCERCODE"],
        "source_map": {"A": PRODUCT_CODE_COL, "B": COL["K"]},
        "sql_col": "C",
        "sql_template": (
            '="UPDATE LG_038_ITEMS SET PRODUCERCODE = \'" & B{r} & "\' WHERE CODE = \'" & A{r} & "\';"'
        ),
    },
    # 11. Kart Tipi
    {
        "name": "Kart Tipi",
        "trigger_cols": [COL["A"]],
        "output_cols": ["A", "B"],
        "headers": ["Ürün Kodu", "Kart Tipi"],
        "source_map": {"A": PRODUCT_CODE_COL, "B": COL["A"]},
        "sql_col": "C",
        "sql_template": '="update lg_038_ITEMS SET CARDTYPE= "&B{r}& " WHERE CODE=\'" & A{r} & "\'"',
    },
    # 12. Ürün Adı Güncelleme
    {
        "name": "Ürün Adı Güncelleme",
        "trigger_cols": [COL["L"]],
        "output_cols": ["A", "B"],
        "headers": ["CODE", "Ürün Adı"],
        "source_map": {"A": PRODUCT_CODE_COL, "B": COL["L"]},
        "sql_col": "C",
        "sql_template": (
            '="UPDATE LG_038_ITEMS SET NAME = \'" & B{r} & "\' WHERE CODE = \'" & A{r} & "\';"'
        ),
    },
    # 13. GTIP G
    {
        "name": "GTIP G",
        "trigger_cols": [COL["BF"]],
        "output_cols": ["A", "B"],
        "headers": ["CODE", "GTIPCODE"],
        "source_map": {"A": PRODUCT_CODE_COL, "B": COL["BF"]},
        "sql_col": "C",
        "sql_template": (
            '="UPDATE LG_038_ITEMS SET GTIPCODE = \'" & B{r} & "\' WHERE CODE = \'" & A{r} & "\';"'
        ),
    },
    # 14. Barkod (multi-row: URUN / İÇ KOLİ / KOLİ)
    # Excel formula: =CONCATENATE("EXEC Sp_Barkod_Atama_038 '",A{r},"','",C{r},"','",B{r},"',0")
    {
        "name": "Barkod",
        "trigger_cols": [COL["M"], COL["N"], COL["O"]],
        "output_cols": ["A", "B", "C"],
        "headers": ["Agnaris Kodu", "Birim", "ALT BARKOD"],
        "multi_row": True,
        "barcode_map": {
            "URUN":     COL["M"],
            "İÇ KOLİ": COL["N"],
            "KOLİ":    COL["O"],
        },
        "sql_col": "D",
        "sql_template": (
            '=CONCATENATE("EXEC Sp_Barkod_Atama_038 \'",A{r},"\',\'",C{r},"\',\'",B{r},"\',0")'
        ),
    },
    # 15. Birim Koli (multi-row: KOLİ / İÇ KOLİ)
    # Excel formula: ="EXEC dbo.Sp_Birim_Carpan_UPDATE_038 '" & A{r} & "', " & C{r} & ", '" & B{r} & "';"
    {
        "name": "Birim Koli",
        "trigger_cols": [COL["Q"], COL["P"]],
        "output_cols": ["A", "B", "C"],
        "headers": ["Logo Ürün Kodu", "Birim", "Koli İçi Adet"],
        "multi_row": True,
        "birim_map": {
            "KOLİ":    COL["Q"],
            "İÇ KOLİ": COL["P"],
        },
        "sql_col": "D",
        "sql_template": (
            '="EXEC dbo.Sp_Birim_Carpan_UPDATE_038 \'" & A{r} & "\', " & C{r} & ", \'" & B{r} & "\';"'
        ),
    },
    # 16. Boyut Bulk (multi-row: URUN / İÇ KOLİ / KOLİ)
    {
        "name": "Boyut Bulk",
        "trigger_cols": [
            COL["AG"], COL["AH"], COL["AI"], COL["AJ"], COL["AK"],
            COL["AR"], COL["AS"], COL["AT"], COL["AU"], COL["AV"],
            COL["AW"], COL["AX"], COL["AY"], COL["AZ"], COL["BA"],
        ],
        "output_cols": ["A", "B", "C", "D", "E", "F", "G"],
        "headers": ["TRU KODU", "BİRİM", "EN", "BOY", "YUKSEKLIK", "NET_AGIRLIK", "BRUT_AGIRLIK"],
        "multi_row": True,
        "boyut_map": {
            "URUN": {
                "en": COL["AG"], "boy": COL["AH"], "yukseklik": COL["AI"],
                "net": COL["AJ"], "brut": COL["AK"],
            },
            "İÇ KOLİ": {
                "en": COL["AR"], "boy": COL["AS"], "yukseklik": COL["AT"],
                "net": COL["AU"], "brut": COL["AV"],
            },
            "KOLİ": {
                "en": COL["AW"], "boy": COL["AX"], "yukseklik": COL["AY"],
                "net": COL["AZ"], "brut": COL["BA"],
            },
        },
        "sql_col": "H",
        "sql_template": _BOYUT_BULK_TEMPLATE,
    },
]
