"""
build_dashboard.py
Genera dashboard.html leyendo los Excel del directorio.

Soporta dos fuentes IQVIA con detección automática:
  - LEGACY (prioridad): X_MOLECULA_2.xlsx + X_ATC.xlsx (granularidad por presentación)
  - MAYO (fallback / formato nuevo): MAYO/MOLECULA_MAYO_1.xlsx + MAYO/ATC_MAYO_1.xlsx
                                    (granularidad por producto)

Si la carpeta MAYO está presente Y los archivos legacy también, usa LEGACY como
fuente y al final imprime un reporte de diferencias contra MAYO para el último
mes (sanity check).

Venta interna: si existe MAYO/QLICK_VENTA_INTERNA_MAYO_1.xlsx, mergea sus meses
2026 sobre los valores del QLICK_VTA_INTERNA.xlsx histórico.

Uso:
    py build_dashboard.py
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
TEMPLATE_FILE = HERE / "dashboard.html"
OUTPUT_FILE = HERE / "dashboard.html"

MAYO_DIR = HERE / "MAYO"
MAYO_MOL_FILE = MAYO_DIR / "MOLECULA_MAYO_1.xlsx"
MAYO_ATC_FILE = MAYO_DIR / "ATC_MAYO_1.xlsx"
MAYO_QLICK_FILE = MAYO_DIR / "QLICK_VENTA_INTERNA_MAYO_1.xlsx"

# Familias que aparecen en QLICK pero no son productos a mostrar
# (regla "no agregar productos nuevos").
FAMILIES_TO_IGNORE = {"AIREAL"}

MES_EN_TO_ES = {
    "Jan": "Ene", "Feb": "Feb", "Mar": "Mar", "Apr": "Abr",
    "May": "May", "Jun": "Jun", "Jul": "Jul", "Aug": "Ago",
    "Sep": "Sep", "Oct": "Oct", "Nov": "Nov", "Dec": "Dic",
}
MES_LARGO_EN_TO_ES = {
    "January": "Ene", "February": "Feb", "March": "Mar", "April": "Abr",
    "May": "May", "June": "Jun", "July": "Jul", "August": "Ago",
    "September": "Sep", "October": "Oct", "November": "Nov", "December": "Dic",
}

ALL_MONTHS = [
    "Ene-2025", "Feb-2025", "Mar-2025", "Abr-2025", "May-2025", "Jun-2025",
    "Jul-2025", "Ago-2025", "Sep-2025", "Oct-2025", "Nov-2025", "Dic-2025",
    "Ene-2026", "Feb-2026", "Mar-2026", "Abr-2026",
]

FAMILY_DISPLAY = {
    "ALIDIAL": "Alidial L",
}

PRES_PREFIX_STRIP = {
    "BREXIL": "BREXIL ",
}


# ---------- Helpers de mes ----------

def col_to_es_mes(col):
    """Convierte un nombre de columna a 'Ene-2025'. Acepta:
       - 'Jan 2025\\nUnits' (formato corto, archivos LEGACY)
       - 'January 2026\\nUnits' (formato largo, archivos MAYO)
    """
    s = str(col).replace("\n", " ").strip()
    m = re.match(r"^([A-Za-z]+)\s+(\d{4})", s)
    if not m:
        return None
    name, year = m.group(1), m.group(2)
    es = MES_LARGO_EN_TO_ES.get(name) or MES_EN_TO_ES.get(name[:3])
    return f"{es}-{year}" if es else None


# ---------- Lectura LEGACY ----------

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


def read_xmol_legacy():
    df = pd.read_excel(XMOL_FILE)
    if "Molecules Short" in df.columns:
        df = df.rename(columns={"Molecules Short": "Molecule"})
    elif "Molecules Long" in df.columns:
        df = df.rename(columns={"Molecules Long": "Molecule"})
    df["Product"] = df["Product"].astype(str).str.strip()
    df["Pack"] = df["Pack"].astype(str).str.strip()
    df = df.drop_duplicates(subset=["Product", "Pack"], keep="first").reset_index(drop=True)
    return df


def read_xatc_legacy():
    df = pd.read_excel(XATC_FILE)
    df["Product"] = df["Product"].astype(str).str.strip()
    df["Pack"] = df["Pack"].astype(str).str.strip()
    df = df.drop_duplicates(subset=["Product", "Pack"], keep="first").reset_index(drop=True)
    return df


# ---------- Lectura MAYO ----------

def read_xmol_mayo():
    df = pd.read_excel(MAYO_MOL_FILE)
    df.columns = [str(c).strip() for c in df.columns]
    if "Molecules Short" in df.columns:
        df = df.rename(columns={"Molecules Short": "Molecule"})
    elif "Molecules Long" in df.columns:
        df = df.rename(columns={"Molecules Long": "Molecule"})
    df["Product"] = df["Product"].astype(str).str.strip()
    return df  # sin Pack


def read_xatc_mayo():
    df = pd.read_excel(MAYO_ATC_FILE)
    # Headers vienen con '\n' al final en este archivo
    df.columns = [str(c).replace("\n", " ").strip() for c in df.columns]
    df["Product"] = df["Product"].astype(str).str.strip()
    return df  # sin Pack


def read_qlick_mayo():
    df = pd.read_excel(MAYO_QLICK_FILE)
    new_cols = ["Sociedad", "Gran Familia", "Familia", "Producto", "Presentacion", "Codigo"] + list(df.columns[6:])
    df.columns = new_cols
    df = df.iloc[1:].reset_index(drop=True)
    return df


# ---------- Detección de fuente ----------

def detect_iqvia_source():
    """Prioridad LEGACY. Si no está, prueba MAYO."""
    if XMOL_FILE.exists() and XATC_FILE.exists():
        return "LEGACY"
    if MAYO_MOL_FILE.exists() and MAYO_ATC_FILE.exists():
        return "MAYO"
    raise FileNotFoundError(
        "No encontré ni X_MOLECULA_2/X_ATC.xlsx ni MAYO/MOLECULA_MAYO_1/ATC_MAYO_1.xlsx"
    )


# ---------- Builders ----------

def units_by_month(row, month_cols):
    out = {}
    for col in month_cols:
        es = col_to_es_mes(col)
        if es is None or es not in ALL_MONTHS:
            continue
        val = row[col]
        out[es] = float(val) if pd.notna(val) else 0.0
    return out


def market_by_month(df, mask, month_cols):
    out = {}
    subset = df[mask]
    for col in month_cols:
        es = col_to_es_mes(col)
        if es is None or es not in ALL_MONTHS:
            continue
        out[es] = float(subset[col].sum())
    return out


def build_iqvia_products(maestro, xmol, xatc, source):
    """source: 'LEGACY' o 'MAYO'."""
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

        # En modo MAYO, el `si` es el total del producto (no por pack).
        # Lo calculamos una sola vez y lo asignamos a la primera presentación;
        # las demás van en 0 para preservar la suma correcta del total.
        si_producto = None
        if source == "MAYO":
            mol_match_prod = xmol_real[
                xmol_real["Product"].astype(str).str.strip() == iqv_name_clean
            ]
            if not mol_match_prod.empty:
                si_producto = units_by_month(mol_match_prod.iloc[0], mol_months)
                molecula = str(mol_match_prod.iloc[0]["Molecule"]).strip()
            atc_match_prod = xatc_real[
                xatc_real["Product"].astype(str).str.strip() == iqv_name_clean
            ]
            if not atc_match_prod.empty:
                atc = str(atc_match_prod.iloc[0]["ATC-4"]).strip()

        for idx, (_, mrow) in enumerate(grupo.iterrows()):
            pack = str(mrow["PRESENTACION_IQUVIA"]).strip()
            label = str(mrow["PRESENTACION_QLICK"]).strip()
            if pres_prefix and label.upper().startswith(pres_prefix.upper()):
                label = label[len(pres_prefix):].strip()

            if source == "LEGACY":
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
            else:  # MAYO
                # primera presentación recibe el total; las demás 0
                if idx == 0 and si_producto:
                    si_dict = si_producto
                else:
                    si_dict = {m: 0.0 for m in ALL_MONTHS}

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


def build_venta_interna(ventas, ventas_mayo=None):
    """Si ventas_mayo está, sobreescribe los meses 2026 con sus valores."""
    out = []
    leaf = ventas[ventas["Codigo"].notna()].copy()

    # Mapa código → fila MAYO (para mergear meses 2026)
    mayo_map = {}
    mayo_months_es = []
    if ventas_mayo is not None:
        leaf_mayo = ventas_mayo[ventas_mayo["Codigo"].notna()]
        mayo_map = {str(r["Codigo"]).strip(): r for _, r in leaf_mayo.iterrows()}
        mayo_months_es = [c for c in ventas_mayo.columns if isinstance(c, str) and re.match(r"^[A-ZÁ-Úa-zá-úñÑ]{3}-\d{4}$", c)]

    for _, row in leaf.iterrows():
        gf_raw = str(row["Gran Familia"]).strip() if pd.notna(row["Gran Familia"]) else ""
        if gf_raw in FAMILIES_TO_IGNORE:
            continue
        data = {}
        for m in ALL_MONTHS:
            v = row[m] if m in ventas.columns else None
            if pd.isna(v) or v == "-":
                data[m] = 0.0
            else:
                try:
                    data[m] = float(v)
                except (ValueError, TypeError):
                    data[m] = 0.0
        # Overlay MAYO si corresponde
        codigo = str(row["Codigo"]).strip()
        if codigo in mayo_map and mayo_months_es:
            mrow = mayo_map[codigo]
            for m in mayo_months_es:
                if m in ALL_MONTHS:
                    v = mrow[m]
                    if pd.isna(v) or v == "-":
                        data[m] = 0.0
                    else:
                        try:
                            data[m] = float(v)
                        except (ValueError, TypeError):
                            pass
        out.append({
            "gran_familia": FAMILY_DISPLAY.get(gf_raw, gf_raw),
            "familia": str(row["Familia"]).strip() if pd.notna(row["Familia"]) else "",
            "producto": str(row["Producto"]).strip() if pd.notna(row["Producto"]) else "",
            "presentacion": str(row["Presentacion"]).strip() if pd.notna(row["Presentacion"]) else "",
            "codigo": codigo,
            "data": data,
        })
    return out


def build_families(ventas):
    fams = []
    seen = set()
    for f in ventas["Gran Familia"].dropna().tolist():
        f = str(f).strip()
        if not f or f == "Totales" or f in FAMILIES_TO_IGNORE:
            continue
        f = FAMILY_DISPLAY.get(f, f)
        if f not in seen:
            seen.add(f)
            fams.append(f)
    return fams


def report_mayo_diffs(maestro, xmol_legacy, xatc_legacy):
    """Compara el último mes común contra los archivos MAYO si están presentes.
    Imprime un reporte si hay diferencias para los productos del dashboard."""
    if not (MAYO_MOL_FILE.exists() and MAYO_ATC_FILE.exists()):
        return
    try:
        xmol_mayo = read_xmol_mayo()
        xatc_mayo = read_xatc_mayo()
    except Exception as e:
        print(f"[WARN] No pude leer archivos MAYO para comparación: {e}", file=sys.stderr)
        return

    # Último mes común
    legacy_months = [col_to_es_mes(c) for c in xmol_legacy.columns if "Units" in str(c)]
    mayo_months = [col_to_es_mes(c) for c in xmol_mayo.columns if "Units" in str(c)]
    common = [m for m in legacy_months if m in mayo_months and m in ALL_MONTHS]
    if not common:
        return
    last_month = common[-1]
    # Buscar columnas originales
    def find_col(df, es):
        for c in df.columns:
            if col_to_es_mes(c) == es:
                return c
        return None

    col_legacy = find_col(xmol_legacy, last_month)
    col_mayo = find_col(xmol_mayo, last_month)
    if not col_legacy or not col_mayo:
        return

    print(f"\n--- Sanity check vs MAYO para {last_month} ---")
    diffs = []
    mae = maestro[maestro["NOMBRE_IQUVIA"].notna()]
    for iqv_name in mae["NOMBRE_IQUVIA"].dropna().unique():
        iqv_clean = str(iqv_name).strip()
        legacy_val = xmol_legacy[
            (xmol_legacy["Manufacturer"] != "Grand Total")
            & (xmol_legacy["Product"].astype(str).str.strip() == iqv_clean)
        ][col_legacy].sum()
        mayo_val = xmol_mayo[
            (xmol_mayo["Manufacturer"] != "Grand Total")
            & (xmol_mayo["Product"].astype(str).str.strip() == iqv_clean)
        ][col_mayo].sum()
        if abs(legacy_val - mayo_val) > 0.5:
            diffs.append((iqv_clean, legacy_val, mayo_val, mayo_val - legacy_val))

    if not diffs:
        print(f"  ✓ Productos Siegfried Pack {last_month}: coinciden con MAYO")
    else:
        print(f"  ⚠ Diferencias en Pack Siegfried {last_month}:")
        for name, ov, nv, d in diffs:
            print(f"    {name:25} LEGACY={ov:.0f}  MAYO={nv:.0f}  DIFF={d:+.0f}")


def main():
    source = detect_iqvia_source()
    print(f"Fuente IQVIA: {source}")
    print("Leyendo Excels...")
    maestro = read_maestro()
    ventas = read_qlick()
    if source == "LEGACY":
        xmol = read_xmol_legacy()
        xatc = read_xatc_legacy()
    else:
        xmol = read_xmol_mayo()
        xatc = read_xatc_mayo()
    ventas_mayo = read_qlick_mayo() if MAYO_QLICK_FILE.exists() else None
    print(f"  MAESTRO: {maestro.shape}")
    print(f"  QLICK:   {ventas.shape}")
    print(f"  X_MOL:   {xmol.shape}")
    print(f"  X_ATC:   {xatc.shape}")
    if ventas_mayo is not None:
        print(f"  QLICK_MAYO (overlay): {ventas_mayo.shape}")

    print("Construyendo iqvia_products...")
    iqvia_products = build_iqvia_products(maestro, xmol, xatc, source)
    print(f"  {len(iqvia_products)} productos IQVIA")

    print("Construyendo venta_interna...")
    venta_interna = build_venta_interna(ventas, ventas_mayo=ventas_mayo)
    print(f"  {len(venta_interna)} filas (filtrando: {FAMILIES_TO_IGNORE})")

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

    template = TEMPLATE_FILE.read_text(encoding="utf-8")
    start = template.find("const DATA = ")
    end = template.find(";\n", start)
    if start == -1 or end == -1:
        raise RuntimeError("No encontré 'const DATA' en el template")

    new_json = "const DATA = " + json.dumps(data, ensure_ascii=False)
    new_html = template[:start] + new_json + template[end:]

    OUTPUT_FILE.write_text(new_html, encoding="utf-8")
    print(f"OK -> {OUTPUT_FILE}")

    # Sanity check vs MAYO si LEGACY estaba activo
    if source == "LEGACY":
        report_mayo_diffs(maestro, xmol, xatc)

    # Validación
    print("\nValidación ACNECLIN PBA:")
    acne = next((p for p in iqvia_products if "ACNECLIN" in p["producto"]), None)
    if acne:
        if source == "LEGACY":
            pres = acne["presentaciones"][0]
            print(f"  Sep-2025 SI: {pres['si']['Sep-2025']} (esperado 645)")
            print(f"  Abr-2026 SI: {pres['si']['Abr-2026']} (esperado 148)")
        print(f"  Mercado ATC Abr-2026: {acne['mercado_atc']['Abr-2026']:.0f} (esperado 127895)")
        print(f"  Mercado MOL Abr-2026: {acne['mercado_molecula']['Abr-2026']:.0f} (esperado 10264)")


if __name__ == "__main__":
    main()
