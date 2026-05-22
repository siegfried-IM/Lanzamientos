"""
build_dashboard.py
Regenera dashboard.html (autocontenido) a partir de los 4 Excel del mismo directorio:

Inputs:
    MAESTRO.xlsx            - crosswalk Qlik <-> IQVIA
    QLICK_VTA_INTERNA.xlsx  - venta interna por SKU/mes
    X_MOLECULA_2.xlsx       - IQVIA por molecula (mercado total por molecula)
    X_ATC.xlsx              - IQVIA por ATC-4 (mercado total por ATC)

Output:
    dashboard.html (sobrescribe el actual; usa su HTML/CSS/JS como template y
                    reemplaza solo el bloque `const DATA = {...}`).
"""

import json
import re
import sys
from datetime import datetime
from pathlib import Path

import pandas as pd

HERE = Path(__file__).parent
MAESTRO_FILE = HERE / "MAESTRO.xlsx"
QLICK_FILE = HERE / "QLICK_VTA_INTERNA.xlsx"
XMOL_FILE = HERE / "X_MOLECULA_2.xlsx"
XATC_FILE = HERE / "X_ATC.xlsx"
DASHBOARD_FILE = HERE / "dashboard.html"

try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:
    pass

MES_EN_TO_ES = {
    "Jan": "Ene", "Feb": "Feb", "Mar": "Mar", "Apr": "Abr",
    "May": "May", "Jun": "Jun", "Jul": "Jul", "Aug": "Ago",
    "Sep": "Sep", "Oct": "Oct", "Nov": "Nov", "Dic": "Dic",
}

ALL_MONTHS = [
    "Ene-2025", "Feb-2025", "Mar-2025", "Abr-2025", "May-2025", "Jun-2025",
    "Jul-2025", "Ago-2025", "Sep-2025", "Oct-2025", "Nov-2025", "Dic-2025",
    "Ene-2026", "Feb-2026", "Mar-2026", "Abr-2026",
]

# Renombrados que se aplican al display (familia en iqvia_products, gran_familia
# en venta_interna y lista de familias del dropdown). No modifica los Excel.
FAMILY_DISPLAY = {
    "ALIDIAL": "Alidial L",
}

# Prefijos a quitar del label de cada presentacion para que no se repita el
# nombre del producto (ej: "BREXIL 1 mg comp rec x 30" -> "1 mg comp rec x 30").
PRES_PREFIX_STRIP = {
    "BREXIL": "BREXIL ",
}


def en_col_to_es(col: str):
    m = re.match(r"^([A-Za-z]{3})\s+(\d{4})", str(col))
    if not m:
        return None
    en, year = m.group(1), m.group(2)
    es = MES_EN_TO_ES.get(en)
    return f"{es}-{year}" if es else None


def read_maestro():
    df = pd.read_excel(MAESTRO_FILE, sheet_name="MAESTRO_SIEGFRIED", header=6)
    df = df.drop(columns=[c for c in df.columns if str(c).startswith("Unnamed")])
    df.columns = [c.strip() for c in df.columns]
    return df


def read_qlick():
    df = pd.read_excel(QLICK_FILE)
    new_cols = ["Sociedad", "Gran Familia", "Familia", "Producto", "Presentacion", "Codigo"] + list(df.columns[6:])
    df.columns = new_cols
    df = df.iloc[1:].reset_index(drop=True)
    return df


def read_xmol():
    df = pd.read_excel(XMOL_FILE)
    if "Molecules Short" in df.columns:
        df = df.rename(columns={"Molecules Short": "Molecule"})
    elif "Molecules Long" in df.columns:
        df = df.rename(columns={"Molecules Long": "Molecule"})
    df["Product"] = df["Product"].astype(str).str.strip()
    df["Pack"] = df["Pack"].astype(str).str.strip()
    df = df.drop_duplicates(subset=["Product", "Pack"], keep="first").reset_index(drop=True)
    return df


def read_xatc():
    df = pd.read_excel(XATC_FILE)
    df["Product"] = df["Product"].astype(str).str.strip()
    df["Pack"] = df["Pack"].astype(str).str.strip()
    df = df.drop_duplicates(subset=["Product", "Pack"], keep="first").reset_index(drop=True)
    return df


def units_by_month(row, month_cols_en):
    out = {}
    for col in month_cols_en:
        es = en_col_to_es(col)
        if es is None or es not in ALL_MONTHS:
            continue
        val = row[col]
        out[es] = float(val) if pd.notna(val) else 0.0
    return out


def market_by_month(df, mask, month_cols_en):
    out = {}
    subset = df[mask]
    for col in month_cols_en:
        es = en_col_to_es(col)
        if es is None or es not in ALL_MONTHS:
            continue
        out[es] = float(subset[col].sum())
    return out


def build_iqvia_products(maestro, xmol, xatc):
    mol_months = [c for c in xmol.columns if "Units" in str(c)]
    atc_months = [c for c in xatc.columns if "Units" in str(c)]

    xmol_real = xmol[xmol["Manufacturer"] != "Grand Total"].copy()
    xatc_real = xatc[xatc["Manufacturer"] != "Grand Total"].copy()

    mae = maestro[maestro["NOMBRE_IQUVIA"].notna()].copy()

    products = []
    for iqv_name, grupo in mae.groupby("NOMBRE_IQUVIA", sort=False):
        iqv_name_clean = str(iqv_name).strip()
        producto_clean = iqv_name_clean.replace(" (SIE)", "")
        familia_raw = str(grupo["NOMBRE_QLICK"].iloc[0]).strip()
        familia = FAMILY_DISPLAY.get(familia_raw, familia_raw)
        pres_prefix = PRES_PREFIX_STRIP.get(producto_clean)

        molecula = None
        atc = None
        presentaciones = []

        for _, mrow in grupo.iterrows():
            pack = str(mrow["PRESENTACION_IQUVIA"]).strip()
            label = str(mrow["PRESENTACION_QLICK"]).strip()
            if pres_prefix and label.upper().startswith(pres_prefix.upper()):
                label = label[len(pres_prefix):].strip()

            mol_match = xmol_real[
                (xmol_real["Product"].astype(str).str.strip() == iqv_name_clean)
                & (xmol_real["Pack"].astype(str).str.strip() == pack)
            ]
            atc_match = xatc_real[
                (xatc_real["Product"].astype(str).str.strip() == iqv_name_clean)
                & (xatc_real["Pack"].astype(str).str.strip() == pack)
            ]

            if mol_match.empty:
                print(f"[WARN] No mol match: {iqv_name_clean} / {pack}", file=sys.stderr)
                si_dict = {m: 0.0 for m in ALL_MONTHS}
            else:
                si_dict = units_by_month(mol_match.iloc[0], mol_months)
                if molecula is None:
                    molecula = str(mol_match.iloc[0]["Molecule"]).strip()

            if atc_match.empty:
                print(f"[WARN] No atc match: {iqv_name_clean} / {pack}", file=sys.stderr)
            else:
                if atc is None:
                    atc = str(atc_match.iloc[0]["ATC-4"]).strip()

            presentaciones.append({
                "label": label,
                "iqvia_pack": pack,
                "si": si_dict,
            })

        mercado_molecula = {}
        if molecula:
            mercado_molecula = market_by_month(
                xmol_real,
                xmol_real["Molecule"].astype(str).str.strip() == molecula,
                mol_months,
            )
        mercado_atc = {}
        if atc:
            mercado_atc = market_by_month(
                xatc_real,
                xatc_real["ATC-4"].astype(str).str.strip() == atc,
                atc_months,
            )

        products.append({
            "producto": producto_clean,
            "familia": familia,
            "molecula": molecula or "",
            "atc": atc or "",
            "presentaciones": presentaciones,
            "mercado_molecula": {m: mercado_molecula.get(m, 0.0) for m in ALL_MONTHS},
            "mercado_atc": {m: mercado_atc.get(m, 0.0) for m in ALL_MONTHS},
        })

    return products


def build_venta_interna(ventas):
    out = []
    leaf = ventas[ventas["Codigo"].notna()].copy()
    for _, row in leaf.iterrows():
        data = {}
        for m in ALL_MONTHS:
            if m in ventas.columns:
                v = row[m]
                if pd.isna(v) or v == "-":
                    data[m] = 0.0
                else:
                    try:
                        data[m] = float(v)
                    except (ValueError, TypeError):
                        data[m] = 0.0
            else:
                data[m] = 0.0
        gf_raw = str(row["Gran Familia"]).strip() if pd.notna(row["Gran Familia"]) else ""
        out.append({
            "gran_familia": FAMILY_DISPLAY.get(gf_raw, gf_raw),
            "familia": str(row["Familia"]).strip() if pd.notna(row["Familia"]) else "",
            "producto": str(row["Producto"]).strip() if pd.notna(row["Producto"]) else "",
            "presentacion": str(row["Presentacion"]).strip() if pd.notna(row["Presentacion"]) else "",
            "codigo": str(row["Codigo"]).strip(),
            "data": data,
        })
    return out


def build_families(ventas):
    fams = []
    seen = set()
    for f in ventas["Gran Familia"].dropna().tolist():
        f = str(f).strip()
        if f and f != "Totales":
            f = FAMILY_DISPLAY.get(f, f)
            if f not in seen:
                seen.add(f)
                fams.append(f)
    return fams


def main():
    print("Leyendo Excels...")
    maestro = read_maestro()
    ventas = read_qlick()
    xmol = read_xmol()
    xatc = read_xatc()
    print(f"  MAESTRO: {maestro.shape}")
    print(f"  QLICK:   {ventas.shape}")
    print(f"  X_MOL:   {xmol.shape}")
    print(f"  X_ATC:   {xatc.shape}")

    print("Construyendo iqvia_products...")
    iqvia_products = build_iqvia_products(maestro, xmol, xatc)
    print(f"  {len(iqvia_products)} productos IQVIA")

    print("Construyendo venta_interna...")
    venta_interna = build_venta_interna(ventas)
    print(f"  {len(venta_interna)} filas")

    families = build_families(ventas)
    print(f"  Familias: {families}")

    data = {
        "all_months": ALL_MONTHS,
        "default_months": ALL_MONTHS[-6:],
        "families": families,
        "iqvia_products": iqvia_products,
        "venta_interna": venta_interna,
        "meta": {
            "last_update": datetime.now().strftime("%d/%m/%Y %H:%M"),
            "latest_iqvia": "Abr-2026",
            "latest_qlick": "Abr-2026",
        },
    }

    template = DASHBOARD_FILE.read_text(encoding="utf-8")
    start = template.find("const DATA = ")
    end = template.find(";\n", start)
    if start == -1 or end == -1:
        raise RuntimeError("No encontre 'const DATA' en dashboard.html template")

    new_json = "const DATA = " + json.dumps(data, ensure_ascii=False)
    new_html = template[:start] + new_json + template[end:]

    DASHBOARD_FILE.write_text(new_html, encoding="utf-8")
    print(f"OK -> {DASHBOARD_FILE}")


if __name__ == "__main__":
    main()
