"""Generate realistic, messy sample sales files in the entrada/ folder.

Each file imitates a different branch office that exports its report
in its own (inconsistent) way.

Run:  python generate_sample_data.py
"""
import random
from datetime import date, timedelta
from pathlib import Path

import pandas as pd

random.seed(42)  # fixed seed -> the same "random" data on every run

BASE_DIR = Path(__file__).parent      # folder where this script lives
INPUT_DIR = BASE_DIR / "entrada"

SELLERS = ["Ana López", "Carlos Méndez", "Lucía Ramírez", "Jorge Torres", "Sofía Herrera"]
PRODUCTS = {
    "Electrónica": ["Audífonos", "Teclado", "Monitor", "Mouse"],
    "Hogar": ["Lámpara", "Cafetera", "Licuadora"],
    "Oficina": ["Silla", "Escritorio", "Archivero"],
}


# ---------- Helpers ----------

def make_clean_sales(n_rows, start, end):
    """Return a list of clean sales (one dict per sale) between two dates."""
    total_days = (end - start).days
    sales = []
    for _ in range(n_rows):
        category = random.choice(list(PRODUCTS))
        sales.append({
            "date": start + timedelta(days=random.randint(0, total_days)),
            "seller": random.choice(SELLERS),
            "category": category,
            "product": random.choice(PRODUCTS[category]),
            "amount": round(random.uniform(150, 9500), 2),
        })
    return sales


def add_duplicates(df, count):
    """Append copies of some existing rows (exact duplicates)."""
    copies = df.sample(n=count, random_state=1)
    return pd.concat([df, copies], ignore_index=True)


def add_blank_rows(df, count):
    """Insert completely empty rows at random positions."""
    for _ in range(count):
        position = random.randint(1, len(df) - 1)
        blank = pd.DataFrame([[None] * len(df.columns)], columns=df.columns)
        df = pd.concat([df.iloc[:position], blank, df.iloc[position:]], ignore_index=True)
    return df


def spanish_number(value):
    """Format 1250.5 as '1.250,50' (thousands with '.', decimals with ',')."""
    us_style = f"{value:,.2f}"  # '1,250.50'
    return us_style.replace(",", "#").replace(".", ",").replace("#", ".")


# ---------- One function per branch office ----------

def build_north():
    """Excel. Spanish headers, real dates, extra spaces, blank rows, duplicates."""
    sales = make_clean_sales(40, date(2026, 1, 1), date(2026, 3, 31))
    df = pd.DataFrame({
        "Fecha": [s["date"] for s in sales],
        "Vendedor": ["  " + s["seller"] + " " for s in sales],
        "Categoría": [s["category"] for s in sales],
        "Producto": [s["product"] for s in sales],
        "Monto": [s["amount"] for s in sales],
    })
    df = add_duplicates(df, 3)
    df = add_blank_rows(df, 4)
    df.to_excel(INPUT_DIR / "ventas_norte_ene-mar.xlsx", index=False)


def build_south():
    """CSV with ','. Dates dd/mm/yyyy, amounts like '$1,250.50', mixed-case names."""
    sales = make_clean_sales(35, date(2026, 2, 1), date(2026, 4, 30))
    df = pd.DataFrame({
        "fecha_venta": [s["date"].strftime("%d/%m/%Y") for s in sales],
        "vendedor ": [random.choice([s["seller"], s["seller"].lower()]) for s in sales],
        "categoria": [s["category"] for s in sales],
        "producto": [s["product"] for s in sales],
        "total": [f"${s['amount']:,.2f}" for s in sales],
    })
    df = add_duplicates(df, 2)
    df = add_blank_rows(df, 2)
    df.to_csv(INPUT_DIR / "ventas_sur.csv", index=False, encoding="utf-8")


def build_center():
    """CSV with ';' and Windows encoding. UPPERCASE text, two date formats
    mixed, decimal comma ('1.250,50') and some missing amounts."""
    sales = make_clean_sales(45, date(2026, 1, 15), date(2026, 4, 15))

    dates = [s["date"].strftime(random.choice(["%Y-%m-%d", "%d/%m/%Y"])) for s in sales]

    amounts = [spanish_number(s["amount"]) for s in sales]
    for i in random.sample(range(len(amounts)), 3):
        amounts[i] = random.choice(["", "N/A", "-"])  # missing values

    df = pd.DataFrame({
        "FECHA": dates,
        "NOMBRE VENDEDOR": [s["seller"].upper() for s in sales],
        "CATEGORIA": [s["category"].upper().replace("Ó", "O") for s in sales],
        "PRODUCTO": [s["product"].upper() for s in sales],
        "IMPORTE": amounts,
    })
    df.to_csv(INPUT_DIR / "VENTAS_CENTRO.csv", index=False, sep=";", encoding="cp1252")


def build_west():
    """Excel with a title above the table, English headers,
    amounts as text with 'MXN', and a TOTAL row at the bottom."""
    sales = make_clean_sales(30, date(2026, 3, 1), date(2026, 4, 30))
    df = pd.DataFrame({
        "Date": [s["date"].strftime("%Y/%m/%d") for s in sales],
        "Salesperson": [s["seller"] for s in sales],
        "Category": [s["category"] for s in sales],
        "Item": [s["product"] for s in sales],
        "Amount (MXN)": [f"{s['amount']:.2f} MXN" for s in sales],
    })
    total = sum(s["amount"] for s in sales)
    total_row = pd.DataFrame([{"Date": "TOTAL", "Amount (MXN)": f"{total:.2f} MXN"}])
    df = pd.concat([df, total_row], ignore_index=True)

    path = INPUT_DIR / "Reporte Oeste (final).xlsx"
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, startrow=2, sheet_name="Ventas")
        writer.sheets["Ventas"]["A1"] = "Reporte de ventas - Sucursal Oeste"


def main():
    INPUT_DIR.mkdir(exist_ok=True)
    build_north()
    build_south()
    build_center()
    build_west()
    print("Sample files created in:", INPUT_DIR)
    for path in sorted(INPUT_DIR.iterdir()):
        print("  -", path.name)


if __name__ == "__main__":
    main()
