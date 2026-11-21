"""Clean and consolidate messy sales reports (Excel/CSV) into one report.

Usage:
    python clean_reports.py
    python clean_reports.py --input other_folder --output results
    python clean_reports.py --map "coffee_name=product" --map "money=amount"
    python clean_reports.py --date-order mdy --currency USD
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
    "date": ["fecha", "fecha venta", "date", "sale date", "order date", "invoice date",
             "transaction date"],
    "seller": ["vendedor", "nombre vendedor", "salesperson", "seller", "sales rep",
               "sales person", "employee", "agent"],
    "category": ["categoria", "category", "product category"],
    "product": ["producto", "item", "product", "product name", "product line", "description"],
    "amount": ["monto", "total", "importe", "amount", "amount (mxn)", "sales", "revenue",
               "total amount"],
}
STANDARD_COLUMNS = list(COLUMN_SYNONYMS)
REQUIRED_COLUMNS = ["date", "amount"]          # the others are optional
TEXT_COLUMNS = ["seller", "category", "product"]
OUTPUT_COLUMNS = STANDARD_COLUMNS + ["currency", "source_file", "source_row"]

# Reverse lookup: "importe" -> "amount", "vendedor" -> "seller", ...
SYNONYM_LOOKUP = {}
for standard_name, synonyms in COLUMN_SYNONYMS.items():
    for synonym in synonyms:
        SYNONYM_LOOKUP[synonym] = standard_name

# Date formats. ISO-like formats (year first) are never ambiguous.
ISO_DATE_FORMATS = ["%Y-%m-%d", "%Y/%m/%d", "%Y-%m-%d %H:%M:%S"]
DATE_FORMATS = {
    "dmy": ISO_DATE_FORMATS + ["%d/%m/%Y", "%d-%m-%Y"],   # 05/03/2026 = 5 March
    "mdy": ISO_DATE_FORMATS + ["%m/%d/%Y", "%m-%d-%Y"],   # 05/03/2026 = May 3
}
SHORT_DATE = re.compile(r"\s*(\d{1,2})[/-](\d{1,2})[/-]\d{4}")  # 05/03/2026 or 5-3-2026

# Currency markers. A bare "$" is NOT here on purpose: it is used by USD, MXN, COP...
CURRENCY_CODES = {"USD", "MXN", "EUR", "GBP", "CAD", "COP", "ARS", "CLP", "PEN", "BRL", "UAH"}
CURRENCY_SYMBOLS = {"US$": "USD", "€": "EUR", "£": "GBP", "₴": "UAH"}

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

def make_row_key(df, columns):
    """One normalized text per row joining ALL its original values.
    Two rows are duplicates only if this whole text is identical."""
    keys = []
    for _, row in df[columns].iterrows():
        keys.append(" | ".join(normalize_key(value) for value in row))
    return keys


def standardize_columns(df, file_name, log, extra_mappings=None):
    """Rename columns to the standard names.
    Required columns missing -> clear error. Optional ones missing -> empty column."""
    extra_mappings = extra_mappings or {}   # from --map, checked before the synonyms
    original_columns = [column for column in df.columns if column != "source_row"]

    renames = {}
    for column in original_columns:
        key = normalize_key(column)
        standard_name = extra_mappings.get(key) or SYNONYM_LOOKUP.get(key)
        if standard_name and standard_name not in renames.values():
            renames[column] = standard_name

    missing = [name for name in REQUIRED_COLUMNS if name not in renames.values()]
    if missing:
        raise ValueError(
            f"could not find required column(s) {missing}. "
            f"Columns in file: {original_columns}. "
            f'Tip: use --map "your column=standard name"'
        )

    for original, standard_name in renames.items():
        if original != standard_name:
            add_log(log, file_name, "", original, "Renamed column", original, standard_name)

    df["row_key"] = make_row_key(df, original_columns)
    df = df[list(renames) + ["source_row", "row_key"]].rename(columns=renames)

    for name in STANDARD_COLUMNS:
        if name not in df.columns:
            df[name] = None
            add_log(log, file_name, "", name, "Optional column not found (left empty)")

    return df[STANDARD_COLUMNS + ["source_row", "row_key"]]


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
    """Fix spaces always; fix capitalization only if the text is ALL UPPER or all lower.
    '  ana   lópez ' -> 'Ana López', 'JORGE' -> 'Jorge',
    'Americano with Milk' and 'Ana de la Cruz' stay as they are."""
    if is_missing(value):
        return None
    text = " ".join(str(value).split())
    if text.isupper() or text.islower():
        text = text.title()
    return text


def detect_currency(value, default=None):
    """'1250.50 MXN' -> 'MXN', '€5' -> 'EUR'. Numbers or a bare '$' -> default."""
    if isinstance(value, str):
        for word in re.findall(r"[A-Za-z]{3}", value):
            if word.upper() in CURRENCY_CODES:
                return word.upper()
        for symbol, code in CURRENCY_SYMBOLS.items():
            if symbol in value:
                return code
    return default


def parse_amount(value):
    """'$1,250.50', '1.250,50', '1250.50 MXN', '(350.00)' or 1250.5 -> number."""
    if is_missing(value):
        return None
    if isinstance(value, numbers.Number):
        return float(value)

    raw = str(value).strip()
    in_parentheses = raw.startswith("(") and raw.endswith(")")  # accounting negative
    text = re.sub(r"[^\d,.\-]", "", raw)  # keep only digits , . and -
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
        number = float(text)
    except ValueError:
        return None
    return -abs(number) if in_parentheses else number


def detect_date_order(values):
    """Look at dates like 05/03/2026 and decide the order for the whole file:
    'dmy' if some first number > 12, 'mdy' if some second number > 12,
    'unclear' if it cannot be decided, None if there are no such dates."""
    found = day_first = month_first = False
    for value in values:
        match = SHORT_DATE.match(str(value)) if isinstance(value, str) else None
        if match:
            found = True
            first, second = int(match.group(1)), int(match.group(2))
            day_first = day_first or first > 12
            month_first = month_first or second > 12
    if not found:
        return None
    if day_first and not month_first:
        return "dmy"
    if month_first and not day_first:
        return "mdy"
    return "unclear"


def parse_date(value, date_order="dmy"):
    """Real dates stay as they are; texts are tried against the formats of date_order."""
    if is_missing(value):
        return None
    if isinstance(value, datetime):          # pandas Timestamp is a datetime too
        return pd.Timestamp(value)
    text = str(value).strip()
    for date_format in DATE_FORMATS[date_order]:
        try:
            return pd.Timestamp(datetime.strptime(text, date_format))
        except ValueError:
            continue  # not this format -> try the next one
    return None


def choose_date_order(values, setting, file_name, log):
    """Use --date-order if given; otherwise detect it from this file's dates."""
    if setting != "auto":
        return setting
    detected = detect_date_order(values)
    if detected == "unclear":
        add_log(log, file_name, "", "date", "Date order unclear, assumed day/month")
        return "dmy"
    if detected == "mdy":
        add_log(log, file_name, "", "date", "Detected month/day date order")
    return detected or "dmy"


def clean_column(df, column, cleaner, action, file_name, log):
    """Apply `cleaner` to every value of a column and log what changed."""
    new_values = []
    for source_row, value in zip(df["source_row"], df[column]):
        new_value = cleaner(value)
        if is_missing(new_value):
            note = ("Missing value (left empty)" if is_missing(value)
                    else "Could not interpret value (left empty)")
            add_log(log, file_name, source_row, column, note, value)
        elif as_log_text(new_value) != as_log_text(value):
            # only log visible changes: text '2024-03-01' -> date 2024-03-01 is not one
            add_log(log, file_name, source_row, column, action, value, new_value)
        new_values.append(new_value)
    df[column] = new_values


def clean_file(path, log, options):
    """Read and clean one file. Returns a table with the output columns + row_key."""
    df = read_file(path)
    df = standardize_columns(df, path.name, log, options.mappings)
    df = remove_empty_and_total_rows(df, path.name, log)

    for column in TEXT_COLUMNS:
        if df[column].notna().any():   # skip optional columns the file does not have
            clean_column(df, column, clean_text, "Fixed spaces/capitalization", path.name, log)

    date_order = choose_date_order(df["date"], options.date_order, path.name, log)
    original_dates = df["date"].dropna().head(3).tolist()
    clean_column(df, "date", lambda value: parse_date(value, date_order),
                 "Converted text to date", path.name, log)
    if len(df) and df["date"].isna().all():
        raise ValueError(f"no value in the date column could be read as a date "
                         f"(examples: {original_dates})")

    df["currency"] = [detect_currency(value, options.currency) for value in df["amount"]]
    original_amounts = df["amount"].dropna().head(3).tolist()
    clean_column(df, "amount", parse_amount, "Converted text to number", path.name, log)
    if len(df) and df["amount"].isna().all():
        raise ValueError(f"no value in the amount column could be read as a number "
                         f"(examples: {original_amounts})")

    df["source_file"] = path.name
    return df[OUTPUT_COLUMNS + ["row_key"]]


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
    """Drop rows whose ORIGINAL values are all identical (keeps the first one).
    Comparing only date+product+amount would delete real sales, e.g. two
    lattes sold the same day at the same price."""
    is_duplicate = df.duplicated(subset=["row_key"], keep="first")
    for file_name, row in zip(df.loc[is_duplicate, "source_file"],
                              df.loc[is_duplicate, "source_row"]):
        add_log(log, file_name, row, "-", "Removed duplicate row")
    return df[~is_duplicate].reset_index(drop=True)


# ---------- Step 6: summaries ----------

def month_table(part):
    """Month x category with totals; just totals per month if there is no category."""
    if part["category"].notna().any():
        table = pd.pivot_table(
            part.fillna({"category": "(none)"}), index="month", columns="category",
            values="amount", aggfunc="sum", fill_value=0, margins=True, margins_name="Total",
        ).reset_index()
        table.columns.name = None
        return table
    table = part.groupby("month")["amount"].agg(sales="count", total="sum").reset_index()
    total_row = pd.DataFrame([{"month": "Total", "sales": table["sales"].sum(),
                               "total": table["total"].sum()}])
    return pd.concat([table, total_row], ignore_index=True)


def dimension_table(part, column):
    """Sales count, total and average per seller / product."""
    return (
        part.groupby(column)["amount"]
        .agg(sales="count", total="sum", average="mean")
        .sort_values("total", ascending=False)
        .reset_index()
    )


def build_summaries(df, log_df):
    """Return a list of (title, table). Split by currency if there is more than one."""
    valid = df.dropna(subset=["date", "amount"])
    valid = valid.assign(month=valid["date"].dt.strftime("%Y-%m"))

    if valid["currency"].nunique() > 1:
        valid = valid.fillna({"currency": "unspecified"})
        groups = [(f" ({currency})", part) for currency, part in valid.groupby("currency")]
    else:
        groups = [("", valid)]

    tables = []
    for suffix, part in groups:
        has_category = part["category"].notna().any()
        title = "Sales by month and category" if has_category else "Sales by month"
        tables.append((title + suffix, month_table(part)))
        for column in ["seller", "product"]:
            if part[column].notna().any():
                tables.append((f"Sales by {column}{suffix}", dimension_table(part, column)))

    corrections = (
        log_df["action"].value_counts()
        .rename_axis("correction").reset_index(name="times")
    )
    tables.append(("Corrections made", corrections))
    return tables


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
    tables = build_summaries(clean_df, log_df)

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        clean_df.to_excel(writer, sheet_name="Clean data", index=False)

        # Summary sheet: tables one below the other, each with a bold title
        start = 1
        for title, table in tables:
            table.to_excel(writer, sheet_name="Summary", startrow=start, index=False)
            title_cell = writer.sheets["Summary"].cell(row=start, column=1, value=title)
            title_cell.font = Font(bold=True, size=13)
            start += len(table) + 4

        log_df.to_excel(writer, sheet_name="Corrections log", index=False)

        format_sheet(writer.sheets["Clean data"])
        format_sheet(writer.sheets["Summary"], freeze_header=False)
        format_sheet(writer.sheets["Corrections log"])


# ---------- Step 8: command line ----------

def parse_mappings(pairs, parser):
    """['coffee_name=product'] -> {'coffee name': 'product'}."""
    mappings = {}
    for pair in pairs:
        source, _, target = pair.partition("=")
        target = target.strip().lower()
        if not source.strip() or target not in STANDARD_COLUMNS:
            parser.error(f'invalid --map "{pair}". Use "column name=one of {STANDARD_COLUMNS}"')
        mappings[normalize_key(source)] = target
    return mappings


def parse_args(argv=None):
    """argv=None reads the real command line; tests can pass a list instead."""
    parser = argparse.ArgumentParser(description="Clean and consolidate messy sales reports.")
    parser.add_argument("--input", type=Path, default=DEFAULT_INPUT_DIR,
                        help="folder with the Excel/CSV files (default: entrada/)")
    parser.add_argument("--output", type=Path, default=DEFAULT_OUTPUT_DIR,
                        help="folder for the report (default: salida/)")
    parser.add_argument("--map", action="append", default=[], metavar="COLUMN=STANDARD",
                        help='use a column as a standard one, e.g. --map "coffee_name=product" '
                             "(repeatable)")
    parser.add_argument("--date-order", choices=["auto", "dmy", "mdy"], default="auto",
                        help="dmy = 05/03 is 5 March, mdy = May 3 (default: detect per file)")
    parser.add_argument("--currency", default=None,
                        help="currency for amounts that do not show one, e.g. MXN")
    args = parser.parse_args(argv)
    args.mappings = parse_mappings(args.map, parser)
    args.currency = args.currency.upper() if args.currency else None
    return args


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
            tables.append(clean_file(path, log, args))
            print(f"  OK       {path.name}")
        except Exception as error:  # one bad file must not stop the others
            print(f"  SKIPPED  {path.name}: {error}")
    if not tables:
        print("ERROR: no file could be processed.")
        return 1

    data = pd.concat(tables, ignore_index=True)
    unify_text_variants(data, log)
    data = remove_duplicates(data, log)
    data = data.drop(columns="row_key")
    data["date"] = pd.to_datetime(data["date"])  # guarantee a real date column
    data = data.sort_values(["date", "source_file", "source_row"]).reset_index(drop=True)

    currencies = sorted(data["currency"].dropna().unique())
    if len(currencies) > 1:
        message = f"Mixed currencies {currencies}: summaries are split by currency"
        add_log(log, "", "", "currency", message)
        print(f"\nWARNING: {message}")

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
