"""
Chile natural gas imports (LNG plus pipeline gas from Argentina),
monthly from 2021, from the CNE "Reporte Mensual del Sector Energetico"
(one PDF per month, public):
  https://www.cne.cl/wp-content/uploads/YYYY/MM/RMensual_vYYYYMM.pdf

Section "Importaciones y Exportaciones de Combustibles" has a table of
imports by fuel in thousand tonnes ("Gas Natural 514 15,5% 3,5%": value,
change on the previous month, change on a year earlier) from customs
(Aduana) data. The data month lags the report by about two months and is
stated in the text ("corresponden al mes de Julio de 2026"). The same
sentence names the origin countries for gas.

Report layouts differ:
  - from Oct 2025: full text layer, parsed directly;
  - some older reports (e.g. 2021, mid-2024): the tables have a text layer
    but the prose (section heading, data-month sentence) does not;
  - other older reports (e.g. 2022, 2025): no text layer at all (text drawn
    as outlines), so the imports page is read with OCR (tesseract).
When the data month cannot be read it is taken as report month - 2, and the
series is checked month to month against each report's published monthly
change. A month with no report (e.g. CNE published no May 2026 report) is
derived from the following month's value and its published monthly change.

Units: thousand tonnes per month as published; also an approximate
million m3/day at 1,360 m3 of gas per tonne of LNG.
Not reachable from the editing sandbox; runs in GitHub Actions
(OCR needs the tesseract-ocr and tesseract-ocr-spa packages).
"""
import argparse
import io
import os
import re
import shutil
import subprocess
import sys
import tempfile

import pandas as pd
import pdfplumber
import requests

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))  # repo root, for xlsx_notes
import xlsx_notes

H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0 Safari/537.36"}
T = (10, 120)
START = pd.Period("2021-01", "M")
LAG = 2  # report month - data month, when the report doesn't say
M3_PER_TONNE = 1360.0
MESES = {"enero": 1, "febrero": 2, "marzo": 3, "abril": 4, "mayo": 5, "junio": 6, "julio": 7, "agosto": 8,
         "septiembre": 9, "setiembre": 9, "octubre": 10, "noviembre": 11, "diciembre": 12}
OCR_DPI = 250

PCT = r"(?:[-+−–—>]?\s?\d{1,4}(?:[.,]\d+)?\s?%|n/[ad])"
# imports-table row: "Gas Natural 215 -25,6% -48,1%" (the exports table's own "Gas Natural" row follows on the same line)
SEP = r"[\s|:]+"  # OCR may read table rules as '|'
GAS_ROW = re.compile(r"Gas\s*Natur\w*" + SEP + r"(\d{1,3}(?:\.\d{3})+|\d{1,4}(?:,\d+)?)" + SEP + "(" + PCT + ")" + SEP + "(" + PCT + ")")
MONTH_RE = re.compile(r"corresponden?\s+al\s+mes\s+de\s+([A-Za-zé]+)\s+(?:de|del)?\s*(20\d\d)", re.I)
ORIGIN_RE = re.compile(r"gas\s+natural\s+desde\s+([^.;]+)", re.I)


def out(*a):
    print(*a, flush=True)


def num(s):
    s = s.strip()
    if re.fullmatch(r"\d{1,3}(\.\d{3})+(,\d+)?", s):
        s = s.replace(".", "")
    return float(s.replace(",", "."))


def pct(s):
    """'-25,6%' -> -0.256; 'n/d', '>100%' -> None."""
    s = re.sub(r"\s", "", s).replace("−", "-").replace("–", "-").replace("—", "-")
    if not s.endswith("%") or s.startswith(">"):
        return None
    return round(float(s[:-1].replace(",", ".")) / 100, 4)


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


def ocr_page(content, i, psm):
    """OCR page i (0-based) of the PDF with tesseract (Spanish)."""
    import pypdfium2 as pdfium  # installed with pdfplumber
    pdf = pdfium.PdfDocument(content)
    try:
        img = pdf[i].render(scale=OCR_DPI / 72).to_pil().convert("L")
    finally:
        pdf.close()
    with tempfile.TemporaryDirectory() as td:
        f = os.path.join(td, "p.png")
        img.save(f)
        r = subprocess.run(["tesseract", f, "stdout", "-l", "spa", "--psm", str(psm)],
                           capture_output=True, text=True, timeout=180)
    return r.stdout


def find_row(text):
    """First imports-table gas row in text: (kt, monthly change, annual change)."""
    for m in GAS_ROW.finditer(text or ""):
        if num(m.group(1)) > 0:  # a 0 is the exports table's gas row (Chile exports none)
            return num(m.group(1)), pct(m.group(2)), pct(m.group(3))
    return None


def month_and_origin(text):
    flat = re.sub(r"\s+", " ", re.sub(r"-\s*\n\s*", "", text or ""))
    dm = MONTH_RE.search(flat)
    month = None
    if dm and dm.group(1).lower() in MESES:
        month = pd.Period(year=int(dm.group(2)), month=MESES[dm.group(1).lower()], freq="M")
    origin = ORIGIN_RE.search(flat)
    return month, (origin.group(1).strip() if origin else None)


def parse(content):
    """Gas imports from one report: dict(kt, mom, yoy, month, origin, method, page) or None."""
    with pdfplumber.open(io.BytesIO(content)) as pdf:
        texts = [pg.extract_text() or "" for pg in pdf.pages]
    n = len(texts)

    # 1) text layer: the page holding the imports table ("Combustible [miles de Ton] Mensual Anual")
    for i, t in enumerate(texts):
        if not re.search(r"miles\s+de\s+Ton", t, re.I):
            continue
        row = find_row(t)
        if not row:
            continue
        month, origin = month_and_origin(t)
        method = "text"
        if month is None and shutil.which("tesseract"):
            # older layout: table has a text layer but the prose (data month, origins) does not
            month, origin2 = month_and_origin(ocr_page(content, i, 3))
            origin = origin or origin2
            method = "text+ocr" if month else "text"
        return dict(kt=row[0], mom=row[1], yoy=row[2], month=month, origin=origin, method=method, page=i + 1)

    # 2) no text layer: OCR the pages where the section usually sits (about 2/3 of the way through)
    if not shutil.which("tesseract"):
        out("      (no text-layer table and tesseract not installed)")
        return None
    cands = [i for i in range(n) if len(texts[i]) < 400 and 0.45 * n <= i <= 0.85 * n]
    cands.sort(key=lambda i: abs(i - 0.645 * n))
    for i in cands[:8]:
        t6 = ocr_page(content, i, 6)
        if not re.search(r"miles\s*de\s*Ton|Importaci", t6, re.I):
            continue
        row = find_row(t6)
        t3 = None
        if not row:
            t3 = ocr_page(content, i, 3)
            row = find_row(t3)
        if not row:
            out(f"      OCR page {i + 1} looks like the imports page but no gas row; lines:")
            for ln in t6.splitlines():
                if re.search(r"gas|miles|Combustible", ln, re.I):
                    out("        ", ln[:160])
            continue
        month, origin = month_and_origin(t6)
        if month is None or origin is None:
            m3, o3 = month_and_origin(t3 if t3 is not None else ocr_page(content, i, 3))
            month, origin = month or m3, origin or o3
        return dict(kt=row[0], mom=row[1], yoy=row[2], month=month, origin=origin, method="ocr", page=i + 1)
    return None


def check_and_fill(df):
    """Check each month against the next month's published monthly change; derive missing months from it."""
    kt = dict(zip(df["Month"], df["Imports_kt"]))
    mom = dict(zip(df["Month"], df["MoM_change_published"]))
    yoy = dict(zip(df["Month"], df["YoY_change_published"]))
    rep = dict(zip(df["Month"], df["Report"]))
    flags = []
    for m in sorted(kt):
        p = pd.Period(m, "M")
        prev, ly = str(p - 1), str(p - 12)
        if prev in kt and pd.notna(mom[m]) and mom[m] > -1:
            imp = kt[m] / (1 + mom[m])
            if abs(imp - kt[prev]) > max(3, 0.03 * kt[prev]):
                flags.append(f"{m}: {kt[m]:g} kt with monthly change {mom[m]:+.1%} implies {prev} = {imp:.0f}, "
                             f"series has {kt[prev]:g}")
        if ly in kt and pd.notna(yoy[m]) and yoy[m] > -1:
            imp = kt[m] / (1 + yoy[m])
            if abs(imp - kt[ly]) > max(3, 0.03 * kt[ly]):
                flags.append(f"{m}: {kt[m]:g} kt with annual change {yoy[m]:+.1%} implies {ly} = {imp:.0f}, "
                             f"series has {kt[ly]:g}")
    out("\nconsistency check (published monthly/annual changes vs neighbouring months): "
        + ("all consistent" if not flags else f"{len(flags)} mismatches"))
    for f in flags:
        out("   ", f)

    months = pd.period_range(START, max(pd.Period(m, "M") for m in kt), freq="M")
    add = []
    for p in months:
        m, nxt = str(p), str(p + 1)
        if m in kt:
            continue
        if nxt in kt and pd.notna(mom[nxt]) and mom[nxt] > -1:
            v = round(kt[nxt] / (1 + mom[nxt]), 1)
            add.append({"Month": m, "Imports_kt": v, "MoM_change_published": None, "YoY_change_published": None,
                        "Origins": None, "Method": f"derived: {nxt} value / (1 + its monthly change {mom[nxt]:+.1%})",
                        "Report": rep[nxt]})
            out(f"  {m}: no report; derived {v} kt from {nxt} ({kt[nxt]:g} kt, monthly change {mom[nxt]:+.1%})")
    if add:
        df = pd.concat([df, pd.DataFrame(add)], ignore_index=True)
    return df


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", default="output/Data and Chart Outputs/chile_gas_imports.xlsx")
    ap.add_argument("--full", action="store_true", help="re-parse every report from 2021")
    args = ap.parse_args()

    have = pd.DataFrame()
    if os.path.exists(args.out) and not args.full:
        have = pd.read_excel(args.out, sheet_name="Gas imports")
        have = have.drop(columns=[c for c in have.columns if str(c).startswith("Unnamed")])
        have["Month"] = have["Month"].astype(str)
        if "Method" not in have.columns:  # file from before the old-layout parser: re-parse everything
            have = pd.DataFrame()
        else:
            have = have[~have["Method"].astype(str).str.startswith("derived")]  # re-derive each run
    # report months already parsed (from the Report URL); these aren't downloaded again
    done = {m.group(1) for u in (have["Report"] if len(have) else [])
            for m in [re.search(r"RMensual_v(\d{6})", str(u))] if m}

    rows = []
    today = pd.Timestamp.today().to_period("M")
    p = START + LAG  # report for data month Jan 2021 is published ~Mar 2021
    while p <= today:
        if f"{p.year}{p.month:02d}" in done and p < today - 2:
            p += 1
            continue
        u, c = fetch_report(p)
        if u is None:
            out(f"  {p}: no report found")
        else:
            r = parse(c)
            if r is None:
                out(f"  {p}: no 'Gas Natural' import line found ({u})")
            else:
                month = r["month"]
                if month is None or not (p - 4 <= month <= p - 1):
                    if month is not None:
                        out(f"      stated data month {month} implausible for report {p}; using {p - LAG}")
                    month = p - LAG
                    how = "month assumed"
                else:
                    how = "month stated"
                rows.append({"Month": str(month), "Imports_kt": r["kt"], "MoM_change_published": r["mom"],
                             "YoY_change_published": r["yoy"], "Origins": r["origin"],
                             "Method": f"{r['method']} p{r['page']}; {how}", "Report": u})
                mom = f"{r['mom']:+.1%}" if r["mom"] is not None else "n/a"
                out(f"  report {p}: data {month} ({r['method']} p{r['page']}, {how}) gas imports {r['kt']:g} kt, m/m {mom} [{r['origin']}]")
        p += 1

    df = pd.concat([have, pd.DataFrame(rows)], ignore_index=True) if len(have) else pd.DataFrame(rows)
    if df.empty:
        out("nothing parsed")
        sys.exit(1)
    df = df.drop_duplicates("Month", keep="last").sort_values("Month").reset_index(drop=True)
    df = df[df["Month"] >= str(START)]
    df = check_and_fill(df)
    df = df.sort_values("Month").reset_index(drop=True)
    days = pd.PeriodIndex(df["Month"], freq="M").days_in_month
    df["Imports_mcm_per_day_approx"] = (df["Imports_kt"] * 1000 * M3_PER_TONNE / 1e6 / days).round(2)
    df = df[["Month", "Imports_kt", "Imports_mcm_per_day_approx", "MoM_change_published", "YoY_change_published",
             "Origins", "Method", "Report"]]
    months = pd.period_range(df["Month"].min(), df["Month"].max(), freq="M").astype(str)
    gaps = sorted(set(months) - set(df["Month"]))
    out(f"\n{len(df)} months {df['Month'].min()}..{df['Month'].max()}; missing: {gaps or 'none'}")
    out(df.drop(columns=["Report"]).to_string(index=False))

    notes = [
        "UNITS",
        "Imports_kt: natural gas imports in thousand tonnes per month, as published by CNE from customs data "
        "(LNG cargoes plus pipeline gas from Argentina). Imports_mcm_per_day_approx: million m3 per day at "
        f"{M3_PER_TONNE:.0f} m3 of gas per tonne, a standard LNG conversion; approximate.",
        "MoM_change_published / YoY_change_published: the report's own change on the previous month and on the "
        "same month a year earlier (fractions; blank where the report shows n/d or n/a).",
        "",
        "COVERAGE",
        "Monthly from January 2021. The data month lags the report by about two months; it is read from the "
        "report's own text where the text can be read ('month stated'), otherwise taken as report month - 2. "
        "Each month is checked against the following month's published monthly change (see the run log).",
        "Method: 'text' = PDF text layer; 'text+ocr' = table from the text layer, data month and origins by OCR; "
        "'ocr' = the whole page by OCR (reports from 2021-2025 that have no text layer); p<n> = PDF page. "
        "'derived' = month with no CNE report (e.g. no May 2026 report was published), computed from the "
        "next month's value and its published monthly change.",
        "Origins lists the source countries the report names for that month's gas imports.",
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
