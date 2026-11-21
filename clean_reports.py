"""Clean and consolidate messy sales reports (Excel/CSV) into one report.

Usage:
    python clean_reports.py
    python clean_reports.py --input other_folder --output results
"""
import argparse
import numbers
import re
import sys
import unicodedata
from collections import Counter
from datetime import datetime
from pathlib import Path

import pandas as pd
from openpyxl.styles import Font

BASE_DIR = Path(__file__).parent
DEFAULT_INPUT_DIR = BASE_DIR / "entrada"
DEFAULT_OUTPUT_DIR = BASE_DIR / "salida"
OUTPUT_FILE_NAME = "reporte_consolidado.xlsx"

SUPPORTED_EXTENSIONS = {".csv", ".xlsx"}

# Standard column name -> names it may have in the source files.
# Written normalized: lowercase, no accents, single spaces ("_" counts as a space).
COLUMN_SYNONYMS = {
    "date": ["fecha", "fecha venta", "date", "sale date"],
    "seller": ["vendedor", "nombre vendedor", "salesperson", "seller"],
    "category": ["categoria", "category"],
    "product": ["producto", "item", "product"],
    "amount": ["monto", "total", "importe", "amount", "amount (mxn)"],
}
STANDARD_COLUMNS = list(COLUMN_SYNONYMS)
TEXT_COLUMNS = ["seller", "category", "product"]

# Reverse lookup: "importe" -> "amount", "vendedor" -> "seller", ...
SYNONYM_LOOKUP = {}
for standard_name, synonyms in COLUMN_SYNONYMS.items():
    for synonym in synonyms:
        SYNONYM_LOOKUP[synonym] = standard_name

DATE_FORMATS = ["%Y-%m-%d", "%d/%m/%Y", "%Y/%m/%d", "%d-%m-%Y"]  # day-first (Mexico)
MISSING_VALUES = {"", "-", "n/a", "na", "nan", "none", "null"}
LOG_COLUMNS = ["file", "row", "column", "action", "before", "after"]


# ---------- Small helpers ----------

def normalize_key(text):
    """'  Categoría_X ' -> 'categoria x' (lowercase, no accents, single spaces)."""
    text = unicodedata.normalize("NFKD", str(text))
    text = "".join(char for char in text if not unicodedata.combining(char))
    text = text.replace("_", " ").lower()
    return " ".join(text.split())


def is_missing(value):
    """True for NaN/None and for texts like '', '-', 'N/A'."""
    if isinstance(value, str):
        return normalize_key(value) in MISSING_VALUES
    return pd.isna(value)


def as_log_text(value):
    """Turn any value into readable text for the corrections log."""
    if is_missing(value):
        return ""
    if isinstance(value, datetime):
        return value.strftime("%Y-%m-%d")
    if isinstance(value, float):
        return f"{value:.2f}"
    return str(value)


def add_log(log, file_name, row, column, action, before="", after=""):
    """Record one correction."""
    log.append({
        "file": file_name,
        "row": row,
        "column": column,
        "action": action,
        "before": as_log_text(before),
        "after": as_log_text(after),
    })


# ---------- Step 2: reading ----------

def read_csv_file(path):
    """Read a CSV, auto-detecting the separator and trying common encodings."""
    for encoding in ["utf-8-sig", "cp1252"]:
        try:
            return pd.read_csv(path, sep=None, engine="python", encoding=encoding)
        except UnicodeDecodeError:
            continue  # wrong encoding -> try the next one
    raise ValueError(f"Could not read {path.name}: unknown text encoding")


def find_header_row(raw):
    """Return the position of the first row where at least half the cells
    are filled in. Titles fill one cell; real headers fill (almost) all."""
    filled_per_row = raw.notna().sum(axis=1)   # count non-empty cells per row
    min_filled = raw.shape[1] / 2
    for position in range(len(raw)):
        if filled_per_row.iloc[position] >= min_filled:
            return position
    return 0  # nothing found -> assume the first row


def read_excel_file(path):
    """Read an Excel file, skipping any title rows above the real header.
    Returns the table and the position of the header row."""
    preview = pd.read_excel(path, header=None, nrows=20)  # first 20 rows, no header
    header_row = find_header_row(preview)
    return pd.read_excel(path, header=header_row), header_row


def read_file(path):
    """Read one CSV or Excel file. Adds 'source_row' = row number in the original file."""
    extension = path.suffix.lower()          # ".CSV" -> ".csv"
    if extension == ".csv":
        df, header_row = read_csv_file(path), 0
    elif extension == ".xlsx":
        df, header_row = read_excel_file(path)
    else:
        raise ValueError(f"Unsupported file type: {path.name}")
    # +1 because Excel counts from 1, +1 more to skip the header line itself
    df["source_row"] = df.index + header_row + 2
    return df


# ---------- Step 3: standard column names ----------

def standardize_columns(df, file_name, log):
    """Rename columns to the standard names. Raise a clear error if one is missing."""
    renames = {}
    for column in df.columns:
        standard_name = SYNONYM_LOOKUP.get(normalize_key(column))
        if standard_name and standard_name not in renames.values():
            renames[column] = standard_name

    missing = [name for name in STANDARD_COLUMNS if name not in renames.values()]
    if missing:
        raise ValueError(
            f"could not find column(s) {missing}. Columns in file: {list(df.columns)}"
        )

    for original, standard_name in renames.items():
        if original != standard_name:
            add_log(log, file_name, "", original, "Renamed column", original, standard_name)

    df = df.rename(columns=renames)
    return df[STANDARD_COLUMNS + ["source_row"]]  # drop any extra columns


# ---------- Step 4: cleaning values ----------

def remove_empty_and_total_rows(df, file_name, log):
    """Drop rows that are completely empty or are 'TOTAL' / 'Subtotal' lines."""
    keep = []
    for _, row in df.iterrows():
        values = [row[column] for column in STANDARD_COLUMNS]
        text_values = [row[column] for column in ["date"] + TEXT_COLUMNS]
        if all(is_missing(value) for value in values):
            add_log(log, file_name, row["source_row"], "-", "Removed empty row")
            keep.append(False)
        elif any(isinstance(v, str) and "total" in normalize_key(v) for v in text_values):
            add_log(log, file_name, row["source_row"], "-", "Removed total row",
                    " | ".join(as_log_text(v) for v in values))
            keep.append(False)
        else:
            keep.append(True)
    return df[keep]


def clean_text(value):
    """'  ana   LÓPEZ ' -> 'Ana López'."""
    if is_missing(value):
        return None
    return " ".join(str(value).split()).title()


def parse_amount(value):
    """'$1,250.50', '1.250,50', '1250.50 MXN' or 1250.5 -> 1250.5 (None if empty)."""
    if is_missing(value):
        return None
    if isinstance(value, numbers.Number):
        return float(value)

    text = re.sub(r"[^\d,.\-]", "", str(value))  # keep only digits , . and -
    if text in ("", "-"):
        return None

    if "," in text and "." in text:
        # Both present: the one that appears LAST is the decimal separator
        if text.rfind(",") > text.rfind("."):
            text = text.replace(".", "").replace(",", ".")   # 1.250,50
        else:
            text = text.replace(",", "")                      # 1,250.50
    elif "," in text:
        decimals = text.split(",")[-1]
        if len(decimals) == 3:
            text = text.replace(",", "")    # 1,250   -> thousands
        else:
            text = text.replace(",", ".")   # 1250,5  -> decimal comma
    elif text.count(".") > 1:
        text = text.replace(".", "")        # 1.250.000 -> thousands

    try:
        return float(text)
    except ValueError:
        return None


def parse_date(value):
    """Real dates stay as they are; texts are tried against DATE_FORMATS."""
    if is_missing(value):
        return None
    if isinstance(value, datetime):          # pandas Timestamp is a datetime too
        return pd.Timestamp(value)
    text = str(value).strip()
    for date_format in DATE_FORMATS:
        try:
            return pd.Timestamp(datetime.strptime(text, date_format))
        except ValueError:
            continue  # not this format -> try the next one
    return None


def clean_column(df, column, cleaner, action, file_name, log):
    """Apply `cleaner` to every value of a column and log what changed."""
    new_values = []
    for source_row, value in zip(df["source_row"], df[column]):
        new_value = cleaner(value)
        if is_missing(new_value):
            note = ("Missing value (left empty)" if is_missing(value)
                    else "Could not interpret value (left empty)")
            add_log(log, file_name, source_row, column, note, value)
        elif new_value != value:
            add_log(log, file_name, source_row, column, action, value, new_value)
        new_values.append(new_value)
    df[column] = new_values


def clean_file(path, log):
    """Read and clean one file. Returns a table with the standard columns."""
    df = read_file(path)
    df = standardize_columns(df, path.name, log)
    df = remove_empty_and_total_rows(df, path.name, log)
    for column in TEXT_COLUMNS:
        clean_column(df, column, clean_text, "Fixed spaces/capitalization", path.name, log)
    clean_column(df, "date", parse_date, "Converted text to date", path.name, log)
    clean_column(df, "amount", parse_amount, "Converted text to number", path.name, log)
    df["source_file"] = path.name
    return df[STANDARD_COLUMNS + ["source_file", "source_row"]]


# ---------- Step 5: consolidating ----------

def unify_text_variants(df, log):
    """'Electronica' and 'Electrónica' -> the most common spelling of the two."""
    for column in TEXT_COLUMNS:
        spellings = {}  # normalized key -> Counter of the real spellings
        for value in df[column].dropna():
            spellings.setdefault(normalize_key(value), Counter())[value] += 1
        best = {key: counter.most_common(1)[0][0] for key, counter in spellings.items()}

        new_values = []
        for file_name, row, value in zip(df["source_file"], df["source_row"], df[column]):
            if is_missing(value):
                new_values.append(value)
                continue
            new_value = best[normalize_key(value)]
            if new_value != value:
                add_log(log, file_name, row, column, "Unified spelling", value, new_value)
            new_values.append(new_value)
        df[column] = new_values


def remove_duplicates(df, log):
    """Drop rows identical in all business columns (keeps the first one)."""
    is_duplicate = df.duplicated(subset=STANDARD_COLUMNS, keep="first")
    for file_name, row in zip(df.loc[is_duplicate, "source_file"],
                              df.loc[is_duplicate, "source_row"]):
        add_log(log, file_name, row, "-", "Removed duplicate row")
    return df[~is_duplicate].reset_index(drop=True)


# ---------- Step 6: summaries ----------

def build_summaries(df, log_df):
    """Return three summary tables: by month x category, by seller, corrections."""
    valid = df.dropna(subset=["date", "amount"])
    valid = valid.assign(month=valid["date"].dt.strftime("%Y-%m"))

    by_month = pd.pivot_table(
        valid, index="month", columns="category", values="amount",
        aggfunc="sum", fill_value=0, margins=True, margins_name="Total",
    ).reset_index()
    by_month.columns.name = None

    by_seller = (
        valid.groupby("seller")["amount"]
        .agg(sales="count", total="sum", average="mean")
        .sort_values("total", ascending=False)
        .reset_index()
    )

    corrections = (
        log_df["action"].value_counts()
        .rename_axis("correction").reset_index(name="times")
    )
    return by_month, by_seller, corrections


# ---------- Step 7: writing the Excel report ----------

def format_sheet(sheet, freeze_header=True):
    """Number/date formats, column widths and frozen header row."""
    for row in sheet.iter_rows():
        for cell in row:
            if isinstance(cell.value, datetime):
                cell.number_format = "yyyy-mm-dd"      # hide the 00:00:00 time
            elif isinstance(cell.value, float):
                cell.number_format = "#,##0.00"
    for column_cells in sheet.columns:
        longest = max(len(as_log_text(cell.value)) for cell in column_cells
                      if cell.value is not None)
        sheet.column_dimensions[column_cells[0].column_letter].width = min(longest + 2, 45)
    if freeze_header:
        sheet.freeze_panes = "A2"


def write_report(clean_df, log_df, output_path):
    """Write the 3-sheet Excel report."""
    by_month, by_seller, corrections = build_summaries(clean_df, log_df)

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        clean_df.to_excel(writer, sheet_name="Clean data", index=False)

        # Summary sheet: three tables one below the other, each with a title
        start = 1
        titles_and_tables = [
            ("Sales by month and category", by_month),
            ("Sales by seller", by_seller),
            ("Corrections made", corrections),
        ]
        for title, table in titles_and_tables:
            table.to_excel(writer, sheet_name="Summary", startrow=start, index=False)
            title_cell = writer.sheets["Summary"].cell(row=start, column=1, value=title)
            title_cell.font = Font(bold=True, size=13)
            start += len(table) + 4

        log_df.to_excel(writer, sheet_name="Corrections log", index=False)

        format_sheet(writer.sheets["Clean data"])
        format_sheet(writer.sheets["Summary"], freeze_header=False)
        format_sheet(writer.sheets["Corrections log"])


# ---------- Step 8: command line ----------

def parse_args(argv=None):
    """argv=None reads the real command line; tests can pass a list instead."""
    parser = argparse.ArgumentParser(description="Clean and consolidate messy sales reports.")
    parser.add_argument("--input", type=Path, default=DEFAULT_INPUT_DIR,
                        help="folder with the Excel/CSV files (default: entrada/)")
    parser.add_argument("--output", type=Path, default=DEFAULT_OUTPUT_DIR,
                        help="folder for the report (default: salida/)")
    return parser.parse_args(argv)


def main(argv=None):
    args = parse_args(argv)

    if not args.input.is_dir():
        print(f"ERROR: input folder not found: {args.input}")
        return 1
    files = sorted(p for p in args.input.iterdir() if p.suffix.lower() in SUPPORTED_EXTENSIONS)
    if not files:
        print(f"ERROR: no .csv or .xlsx files in {args.input}")
        return 1

    log = []
    tables = []
    print(f"Reading {len(files)} file(s) from {args.input}")
    for path in files:
        try:
            tables.append(clean_file(path, log))
            print(f"  OK       {path.name}")
        except Exception as error:  # one bad file must not stop the others
            print(f"  SKIPPED  {path.name}: {error}")
    if not tables:
        print("ERROR: no file could be processed.")
        return 1

    data = pd.concat(tables, ignore_index=True)
    unify_text_variants(data, log)
    data = remove_duplicates(data, log)
    data = data.sort_values(["date", "source_file", "source_row"]).reset_index(drop=True)

    log_df = pd.DataFrame(log, columns=LOG_COLUMNS)
    args.output.mkdir(parents=True, exist_ok=True)
    output_path = args.output / OUTPUT_FILE_NAME
    try:
        write_report(data, log_df, output_path)
    except PermissionError:
        print(f"ERROR: cannot write {output_path}. Is it open in Excel? Close it and retry.")
        return 1

    missing_amounts = data["amount"].isna().sum()
    print(f"\nClean rows: {len(data)}  |  Corrections logged: {len(log_df)}")
    if missing_amounts:
        print(f"Rows without amount: {missing_amounts} (kept, but not counted in totals)")
    print(f"Report saved to: {output_path}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
