"""
Uruguay installed generation capacity by source from MIEM / DNE (Ministerio
de Industria, Energia y Mineria - Direccion Nacional de Energia) - free,
public, no key.

Source: MIEM 'Series estadisticas de energia electrica' page
  https://www.gub.uy/ministerio-industria-energia-mineria/datos-y-estadisticas/datos/series-estadisticas-energia-electrica
link 'Potencia instalada por central (.zip)' -> 'Potencia instalada por
central.xlsx' (data: UTE / DNE). Sheets used:
  'Pot Inst Fuente'  installed capacity (MW) by source, one column per year
                     since 1967 (capacity at 31 December; the current year,
                     when present, is the latest half-year)
  'Pot Inst Equipo ' installed capacity (MW) per plant (kept as a raw tab)
Found via discovery_archive/south_america/CAPACITY_PBUE_DISCOVERY2.py.

Granularity: annual (MIEM publishes year-end values, plus a mid-year value
for the current year once released). One row per year dated 1 January of
that year, holding the capacity at the END of that year. 2021 onwards.

Incremental: the whole series is one small file; each run re-reads it and
only rows for years not yet in the workbook, or whose values changed
(MIEM revisions), are updated.

Usage: python3 URUGUAY_MIEM_CAPACITY.py [--out PATH] [--test]
"""

print("STARTING", flush=True)

import argparse
import io
import os
import re
import sys
import unicodedata
import zipfile

import pandas as pd
import requests

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import power_capacity_common as pcc  # noqa: E402

PAGE = ("https://www.gub.uy/ministerio-industria-energia-mineria/datos-y-estadisticas/datos/"
        "series-estadisticas-energia-electrica")
HEADERS = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) "
                         "Chrome/124.0 Safari/537.36"}
DEFAULT_OUT = os.path.join("output", "Data and Chart Outputs", "uruguay_power_capacity.xlsx")


def key(s):
    s = unicodedata.normalize("NFKD", str(s)).encode("ascii", "ignore").decode().upper()
    return " ".join(s.split())


# MIEM source label (accent-stripped, upper case) -> standard fuel. Checked in order; first match wins.
MAPPING = [
    (r"HIDR", "Hydro"),
    (r"EOLIC", "Wind"),
    (r"SOLAR|FOTOVOLT", "Solar"),
    (r"BIOMASA|BIOGAS|BAGAZO|RESIDUO", "Bioenergy"),
    (r"GAS NATURAL", "Gas"),
    (r"FOSIL|TERMIC|GASOIL|FUEL|DIESEL|PETROLEO|MOTOR|TURBINA|CICLO", "Oil"),
]


def fuel_of(label):
    k = key(label)
    for pat, fuel in MAPPING:
        if re.search(pat, k):
            return fuel
    return None


def find_zip_url():
    r = requests.get(PAGE, headers=HEADERS, timeout=(15, 60))
    r.raise_for_status()
    for href, text in re.findall(r'<a[^>]+href=["\']([^"\']+)["\'][^>]*>(.*?)</a>', r.text, re.S | re.I):
        label = key(re.sub(r"<[^>]+>", " ", text) + " " + requests.utils.unquote(href))
        if "POTENCIA INSTALADA POR CENTRAL" in label and ".ZIP" in label:
            return requests.compat.urljoin(r.url, href)
    raise SystemExit(f"'Potencia instalada por central' link not found on {PAGE}")


def read_workbook(url):
    r = requests.get(url, headers=HEADERS, timeout=(15, 120))
    r.raise_for_status()
    z = zipfile.ZipFile(io.BytesIO(r.content))
    names = [n for n in z.namelist() if n.lower().endswith((".xlsx", ".xls"))]
    if not names:
        raise SystemExit(f"no Excel file in {url}: {z.namelist()}")
    print(f"Zip members: {z.namelist()} -> reading {names[0]}", flush=True)
    return pd.ExcelFile(io.BytesIO(z.read(names[0])))


def year_table(xl, sheet_prefix):
    """A MIEM year-column sheet -> DataFrame(index = row label, columns = column header as published)."""
    sheet = next(s for s in xl.sheet_names if key(s).startswith(key(sheet_prefix)))
    raw = pd.read_excel(xl, sheet_name=sheet, header=None)
    hdr = next(i for i in range(min(15, len(raw)))
               if sum(bool(re.fullmatch(r"(19|20)\d\d(\.0)?", str(v).strip())) for v in raw.iloc[i]) >= 5)
    cols = [str(v).strip() for v in raw.iloc[hdr]]
    body = raw.iloc[hdr + 1:].copy()
    body.columns = cols
    lab = body.columns[0]
    body = body[body[lab].notna()]
    body[lab] = body[lab].astype(str).str.strip()
    return sheet, body.set_index(lab)


def year_of(col):
    m = re.match(r"((19|20)\d\d)", str(col))
    return int(m.group(1)) if m else None


def parse(xl):
    sheet, t = year_table(xl, "Pot Inst Fuente")
    print(f"'{sheet}': {t.shape}; columns {list(t.columns)[:3]} .. {list(t.columns)[-6:]}", flush=True)
    print("Row labels:", list(t.index), flush=True)
    year_cols = [c for c in t.columns if year_of(c)]
    num = t[year_cols].apply(lambda s: pd.to_numeric(s, errors="coerce"))
    # Only leaf rows: skip totals / sub-totals; rows mapped to a fuel.
    rows = {}
    for label, vals in num.iterrows():
        k = key(label)
        if vals.isna().all():
            continue  # section header (UTE / PRIVADOS ...)
        if k.startswith("TOTAL") or "TOTAL" in k.split():
            continue
        fuel = fuel_of(label)
        if fuel is None:
            print(f"  unmapped row {label!r} -> Other", flush=True)
            fuel = "Other"
        rows.setdefault(fuel, []).append(label)
        rows[fuel + "_series"] = rows.get(fuel + "_series", 0) + vals.fillna(0)
    out = pd.DataFrame({f"{f}_MW": rows[f + "_series"] for f in pcc.FUELS if f + "_series" in rows})
    out.index = [pd.Timestamp(year=year_of(c), month=1, day=1) for c in out.index]
    mapping = {f: rows[f] for f in pcc.FUELS if f in rows}
    total_rows = [lab for lab in num.index if key(lab).startswith("TOTAL")]
    published_total = num.loc[total_rows[-1]] if total_rows else None
    if published_total is not None:
        published_total.index = out.index
    return sheet, t, out, mapping, published_total


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", default=DEFAULT_OUT)
    ap.add_argument("--test", action="store_true", help="print the parse in full")
    args = ap.parse_args()

    url = find_zip_url()
    print(f"Zip: {url}", flush=True)
    xl = read_workbook(url)
    print(f"Sheets: {xl.sheet_names}", flush=True)
    sheet, table, annual, mapping, published_total = parse(xl)
    annual = pcc.standardise(annual)
    if published_total is not None:
        gap = (annual["Total_MW"] - published_total.reindex(annual.index)).abs()
        print(f"Max |sum of sources - published total| over all years: {gap.max():.2f} MW", flush=True)
    annual = annual[annual.index >= pcc.START]

    old = pcc.load_monthly(args.out)
    if not old.empty:
        old = pcc.standardise(old)
        changed = [d for d in annual.index if d not in old.index or not old.loc[d].round(1).equals(
            annual.loc[d].round(1))]
        print(f"Archive {len(old)} rows; new or revised: {[d.year for d in changed]}", flush=True)
        annual = pd.concat([old.drop(index=[d for d in changed if d in old.index]), annual.loc[changed]]).sort_index()
    if args.test:
        pd.set_option("display.width", 250)
        print(table.to_string()[:6000])
    print(annual.to_string(), flush=True)

    try:
        _, per_plant = year_table(xl, "Pot Inst Equipo")
        per_plant = per_plant[[c for c in per_plant.columns if year_of(c) and year_of(c) >= 2015]]
    except StopIteration:
        per_plant = pd.DataFrame()
    raw_fuente = table[[c for c in table.columns if year_of(c) and year_of(c) >= 2010]]

    val = pcc.validation_lines(annual, "Uruguay")
    print("\n".join(val), flush=True)
    last = annual.index.max()
    notes = pcc.unit_notes("year") + [
        "Annual series: each row is dated 1 January of the year and holds MIEM's installed capacity at 31 December "
        "of that year (for the current year, when MIEM has published it, the latest half-year value).",
        "",
        "COVERAGE",
        f"{annual.index.min():%Y} to {last:%Y} ({len(annual)} years), from 2021. MIEM's national series (UTE, "
        "private and mixed plants; private generators' backup sets excluded per MIEM's notes). Salto Grande counted "
        "at Uruguay's 50% share (MIEM note: 50% of 1,890 MW).",
        "",
        "SOURCE",
        "MIEM / DNE (Direccion Nacional de Energia), 'Series estadisticas de energia electrica' - 'Potencia instalada "
        f"por central' ({url}), sheet '{sheet}'. Page: {PAGE}. Data: UTE / elaboracion propia DNE.",
        "Script: south_america/URUGUAY_MIEM_CAPACITY.py (scheduled by .github/workflows/uruguay_power_capacity.yml).",
        "",
        "MAPPING",
    ] + [f"{f}_MW = MIEM rows {', '.join(repr(x) for x in labels)}" for f, labels in mapping.items()] + [
        "Oil_MW = MIEM's fossil/thermal capacity: UTE's gas-oil and fuel-oil plants (motors, open-cycle turbines and "
        "the Punta del Tigre B combined cycle, which is dual-fuel but runs on gas oil because Uruguay has almost no "
        "natural gas supply). Ember counts Punta del Tigre B as gas, so Ember shows more Gas and less Oil.",
        "No coal, nuclear or geothermal in Uruguay; Other_MW = any MIEM row not matching a fuel above (none now).",
        "",
        "VALIDATION",
    ] + val
    sheets = {"MIEM by source": raw_fuente}
    if not per_plant.empty:
        sheets["MIEM by plant"] = per_plant
    pcc.write(args.out, annual, notes, sheets)
    print(f"Saved {args.out}", flush=True)


if __name__ == "__main__":
    main()
