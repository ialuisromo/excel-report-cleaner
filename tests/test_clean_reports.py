"""Tests for clean_reports.py.  Run:  python -m pytest"""
import random

import pandas as pd
import pytest

import clean_reports as cr
import generate_sample_data


# ---------- Unit tests: one small function at a time ----------

@pytest.mark.parametrize("raw, expected", [
    ("$1,250.50", 1250.50),      # US style with currency symbol
    ("1.250,50", 1250.50),       # Spanish style: dot thousands, comma decimals
    ("507,70", 507.70),          # decimal comma only
    ("1,250", 1250.0),           # thousands comma only
    ("2675.27 MXN", 2675.27),    # text suffix
    (1250.5, 1250.5),            # already a number
    ("-350.00", -350.0),         # negative (refund)
    ("N/A", None),               # missing markers
    ("-", None),
    ("", None),
    ("abc", None),               # garbage -> empty, not a crash
])
def test_parse_amount(raw, expected):
    assert cr.parse_amount(raw) == expected


@pytest.mark.parametrize("raw, expected", [
    ("2026-03-05", "2026-03-05"),
    ("05/03/2026", "2026-03-05"),   # day first (Mexico): 5 March
    ("2026/03/05", "2026-03-05"),
    (pd.Timestamp("2026-03-05"), "2026-03-05"),
])
def test_parse_date_valid(raw, expected):
    assert cr.parse_date(raw) == pd.Timestamp(expected)


@pytest.mark.parametrize("raw", ["31/02/2026", "yesterday", "", None])
def test_parse_date_invalid_returns_none(raw):
    assert cr.parse_date(raw) is None


def test_normalize_key_ignores_case_accents_and_spaces():
    assert cr.normalize_key("  Categoría ") == "categoria"
    assert cr.normalize_key("fecha_venta") == "fecha venta"
    assert cr.normalize_key("NOMBRE   VENDEDOR") == "nombre vendedor"


def test_clean_text():
    assert cr.clean_text("  ana   LÓPEZ ") == "Ana López"
    assert cr.clean_text("N/A") is None


def test_find_header_row_skips_title_rows():
    raw = pd.DataFrame([
        ["Sales report - West", None, None],
        [None, None, None],
        ["Date", "Seller", "Amount"],
        ["2026-01-01", "Ana", 100],
    ])
    assert cr.find_header_row(raw) == 2


def test_standardize_columns_maps_synonyms():
    df = pd.DataFrame(columns=["FECHA", "Vendedor ", "Categoría", "Item", "Importe", "source_row"])
    result = cr.standardize_columns(df, "test.csv", log=[])
    assert list(result.columns) == ["date", "seller", "category", "product", "amount", "source_row"]


def test_standardize_columns_missing_column_gives_clear_error():
    df = pd.DataFrame(columns=["Fecha", "Vendedor", "source_row"])
    with pytest.raises(ValueError, match="category"):
        cr.standardize_columns(df, "test.csv", log=[])


# ---------- End-to-end tests: the whole program ----------

def run_cleaner(input_dir, output_dir):
    """Run the program like the terminal would and return the 3 sheets."""
    exit_code = cr.main(["--input", str(input_dir), "--output", str(output_dir)])
    assert exit_code == 0
    report = output_dir / cr.OUTPUT_FILE_NAME
    return pd.read_excel(report, sheet_name=None)  # dict: sheet name -> table


def test_small_messy_files_end_to_end(tmp_path):
    input_dir = tmp_path / "in"
    input_dir.mkdir()

    # CSV: ';' separator, Windows encoding, a blank row and a duplicate
    csv_text = (
        "FECHA;VENDEDOR;CATEGORIA;PRODUCTO;IMPORTE\n"
        "05/01/2026;LUCÍA RAMÍREZ;HOGAR;LÁMPARA;1.000,50\n"
        ";;;;\n"
        "05/01/2026;LUCÍA RAMÍREZ;HOGAR;LÁMPARA;1.000,50\n"
        "2026-02-10;  jorge torres ;OFICINA;SILLA;N/A\n"
    )
    (input_dir / "a.csv").write_text(csv_text, encoding="cp1252")

    # Excel: title above the table and a TOTAL row at the bottom
    excel = pd.DataFrame({
        "Date": ["2026/01/20", "TOTAL"],
        "Salesperson": ["Lucía Ramírez", None],
        "Category": ["Hogar", None],
        "Item": ["Lámpara", None],
        "Amount (MXN)": ["2000.00 MXN", "2000.00 MXN"],
    })
    with pd.ExcelWriter(input_dir / "b.xlsx") as writer:
        excel.to_excel(writer, index=False, startrow=2, sheet_name="Sheet1")
        writer.sheets["Sheet1"]["A1"] = "Report"

    sheets = run_cleaner(input_dir, tmp_path / "out")
    clean = sheets["Clean data"]
    log = sheets["Corrections log"]

    assert len(clean) == 3                               # duplicate, blank, TOTAL removed
    assert clean["amount"].sum() == pytest.approx(3000.50)  # TOTAL not double counted
    assert set(clean["seller"]) == {"Lucía Ramírez", "Jorge Torres"}
    for action in ["Removed duplicate row", "Removed empty row", "Removed total row",
                   "Missing value (left empty)"]:
        assert action in set(log["action"]), action


def test_sample_data_totals_match_original(tmp_path, monkeypatch):
    """Cross-check with the generated sample files: no sale lost or duplicated."""
    input_dir = tmp_path / "entrada"
    monkeypatch.setattr(generate_sample_data, "INPUT_DIR", input_dir)
    random.seed(42)
    generate_sample_data.main()

    sheets = run_cleaner(input_dir, tmp_path / "salida")
    clean = sheets["Clean data"]

    assert len(clean) == 30 + 45 + 40 + 35   # real sales per branch
    assert clean["date"].notna().all()       # every date understood

    # West file: cleaned amounts must add up to its own TOTAL row
    west_raw = pd.read_excel(input_dir / "Reporte Oeste (final).xlsx", header=2)
    total_row = west_raw[west_raw["Date"] == "TOTAL"]["Amount (MXN)"].iloc[0]
    west_total = float(total_row.replace(" MXN", ""))
    west_clean = clean[clean["source_file"] == "Reporte Oeste (final).xlsx"]
    assert west_clean["amount"].sum() == pytest.approx(west_total)


def test_dates_are_shown_without_time(tmp_path):
    """Regression test: dates used to appear as '2026-01-01 00:00:00' in Excel."""
    import openpyxl

    input_dir = tmp_path / "in"
    input_dir.mkdir()
    (input_dir / "a.csv").write_text(
        "Fecha,Vendedor,Categoria,Producto,Monto\n05/01/2026,Ana,Hogar,Silla,10\n",
        encoding="utf-8",
    )
    run_cleaner(input_dir, tmp_path / "out")
    sheet = openpyxl.load_workbook(tmp_path / "out" / cr.OUTPUT_FILE_NAME)["Clean data"]
    assert sheet["A2"].number_format == "yyyy-mm-dd"


def test_missing_input_folder_returns_error_code(tmp_path):
    assert cr.main(["--input", str(tmp_path / "nope"), "--output", str(tmp_path)]) == 1


def test_file_with_unknown_columns_is_skipped_not_fatal(tmp_path):
    input_dir = tmp_path / "in"
    input_dir.mkdir()
    (input_dir / "bad.csv").write_text("foo,bar\n1,2\n", encoding="utf-8")
    (input_dir / "good.csv").write_text(
        "Fecha,Vendedor,Categoria,Producto,Monto\n2026-01-01,Ana,Hogar,Silla,10\n",
        encoding="utf-8",
    )
    sheets = run_cleaner(input_dir, tmp_path / "out")
    assert len(sheets["Clean data"]) == 1
