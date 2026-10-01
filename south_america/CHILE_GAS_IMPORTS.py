"""
Chile natural gas imports (LNG plus pipeline gas from Argentina),
monthly from 2021, from the CNE "Reporte Mensual del Sector Energetico"
(one PDF per month, public):
  https://www.cne.cl/wp-content/uploads/YYYY/MM/RMensual_vYYYYMM.pdf

Section "Importaciones y Exportaciones de Combustibles" has a table of
imports by fuel in thousand tonnes ("Gas Natural 514 15,5% 3,5%") from
customs (Aduana) data. The data month lags the report by about two
months and is stated in the text ("corresponden al mes de Julio de
2026"). The same sentence names the origin countries for gas.

Units: thousand tonnes per month as published; also an approximate
million m3/day at 1,360 m3 of gas per tonne of LNG.
Not reachable from the editing sandbox; runs in GitHub Actions.
"""
import argparse
import io
import os
import re
import sys

import pandas as pd
import pdfplumber
import requests

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))  # repo root, for xlsx_notes
import xlsx_notes

H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0 Safari/537.36"}
T = (10, 120)
START = pd.Period("2021-01", "M")
M3_PER_TONNE = 1360.0
MESES = {"enero": 1, "febrero": 2, "marzo": 3, "abril": 4, "mayo": 5, "junio": 6, "julio": 7, "agosto": 8,
         "septiembre": 9, "setiembre": 9, "octubre": 10, "noviembre": 11, "diciembre": 12}


def out(*a):
    print(*a, flush=True)


def num(s):
    s = s.strip()
    if re.fullmatch(r"\d{1,3}(\.\d{3})+(,\d+)?", s):
        s = s.replace(".", "")
    return float(s.replace(",", "."))


def fetch_report(p):
    """Report for month p, trying its own upload folder then the next."""
    for folder in (p, p + 1):
        u = f"https://www.cne.cl/wp-content/uploads/{folder.year}/{folder.month:02d}/RMensual_v{p.year}{p.month:02d}.pdf"
        try:
            r = requests.get(u, headers=H, timeout=T)
        except requests.RequestException:
            continue
        if r.status_code == 200 and r.content[:4] == b"%PDF":
            return u, r.content
    return None, None


def parse(content):
    with pdfplumber.open(io.BytesIO(content)) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ""
            if not re.search(r"Importaciones", text, re.I) or not re.search(r"Gas Natural", text):
                continue
            flat = re.sub(r"\s+", " ", text)
            dm = re.search(r"corresponden?\s+al\s+mes\s+de\s+([A-Za-zé]+)\s+(?:de|del)?\s*(20\d\d)", flat, re.I)
            m = re.search(r"^\s*Gas Natural\s+([\d.,]+)\s+[-+>\d,.%na/]+", text, re.M)
            if not m:
                continue
            origin = re.search(r"gas natural desde ([^.;]+)", flat, re.I)
            month = None
            if dm and dm.group(1).lower() in MESES:
                month = pd.Period(year=int(dm.group(2)), month=MESES[dm.group(1).lower()], freq="M")
            return month, num(m.group(1)), (origin.group(1).strip() if origin else None), text
    return None, None, None, None


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", default="output/chile_gas_imports.xlsx")
    ap.add_argument("--full", action="store_true")
    args = ap.parse_args()

    have = pd.DataFrame()
    if os.path.exists(args.out) and not args.full:
        have = pd.read_excel(args.out, sheet_name="Gas imports")
        have = have.drop(columns=[c for c in have.columns if str(c).startswith("Unnamed")])
        have["Month"] = have["Month"].astype(str)
    done_reports = set(have["Report"]) if len(have) else set()

    rows = []
    today = pd.Timestamp.today().to_period("M")
    p = START + 2  # report for data month Jan 2021 is published ~Mar 2021
    while p <= today:
        u, c = fetch_report(p)
        if u is None:
            out(f"  {p}: no report found")
        elif u in done_reports and p < today - 2:
            pass
        else:
            month, kt, origin, text = parse(c)
            if kt is None:
                out(f"  {p}: no 'Gas Natural' import line found ({u})")
                for ln in (text or "").splitlines()[:5]:
                    out("      ", ln[:160])
            else:
                month = month or (p - 2)
                rows.append({"Month": str(month), "Imports_kt": kt, "Origins": origin, "Report": u})
                out(f"  report {p}: data {month} gas imports {kt} kt  [{origin}]")
        p += 1

    df = pd.concat([have, pd.DataFrame(rows)], ignore_index=True) if len(have) else pd.DataFrame(rows)
    if df.empty:
        out("nothing parsed")
        sys.exit(1)
    df = df.drop_duplicates("Month", keep="last").sort_values("Month").reset_index(drop=True)
    df = df[df["Month"] >= str(START)]
    days = pd.PeriodIndex(df["Month"], freq="M").days_in_month
    df["Imports_mcm_per_day_approx"] = (df["Imports_kt"] * 1000 * M3_PER_TONNE / 1e6 / days).round(2)
    df = df[["Month", "Imports_kt", "Imports_mcm_per_day_approx", "Origins", "Report"]]
    months = pd.period_range(df["Month"].min(), df["Month"].max(), freq="M").astype(str)
    gaps = sorted(set(months) - set(df["Month"]))
    out(f"\n{len(df)} months {df['Month'].min()}..{df['Month'].max()}; missing: {gaps or 'none'}")
    out(df.tail(8).to_string(index=False))

    notes = [
        "UNITS",
        "Imports_kt: natural gas imports in thousand tonnes per month, as published by CNE from customs data "
        "(LNG cargoes plus pipeline gas from Argentina). Imports_mcm_per_day_approx: million m3 per day at "
        f"{M3_PER_TONNE:.0f} m3 of gas per tonne, a standard LNG conversion; approximate.",
        "",
        "COVERAGE",
        "Monthly from 2021. The data month lags the report by about two months; it is read from the report's own "
        "text. Origins lists the source countries the report names for that month's gas imports.",
        "",
        "SOURCE",
        "Comision Nacional de Energia (CNE), Reporte Mensual del Sector Energetico, section 'Importaciones y "
        "Exportaciones de Combustibles' (Aduana data via COMEX). The Report column holds the PDF used.",
    ]
    os.makedirs(os.path.dirname(os.path.abspath(args.out)), exist_ok=True)
    xlsx_notes.write_workbook(args.out, {"Gas imports": df}, notes, {"UNITS", "COVERAGE", "SOURCE"})
    out(f"Saved {args.out}")


if __name__ == "__main__":
    main()
