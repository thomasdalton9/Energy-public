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
series is checked against each report's published monthly and annual
changes. Some reports repeat an earlier report's table (same figures, or a
stated data month that is too old); those are skipped. A month left without
a table (e.g. CNE published no May 2026 report) is derived from the
following month's value and its published monthly change, or else from the
same month a year later and its published annual change.

Units: thousand tonnes per month as published; also an approximate
million m3/day at 1,360 m3 of gas per tonne of LNG.

Main series and sheet "Imports by use": CNE's import statistics workbook
(Estadisticas > Hidrocarburo, importaciones-web.xlsx, customs data) has
monthly pipeline gas from Argentina by use and region (energy: II, RM-V,
VIII regions; petrochemical: Magallanes) and LNG by terminal region (II,
V). It is the main source for every month it covers (to Apr 2026 in the
Jun 2026 upload; this fills the months the reports don't show, e.g. Sep-Nov
2021); the monthly report covers later months. Each run logs how the two
sources compare where both have a month. The newest workbook is found via
CNE's media library and only downloaded when it changes.

Sheet "Domestic production": monthly gas production (thousand m3) by ENAP
and CEOP (private operators), Magallanes, from CNE Estadisticas >
Hidrocarburo, file Produccion_combustibles-<month>-<year>.xlsx (Ministerio
de Energia data; latest found: Jun 2024). The newest file is found via CNE's
media library and only downloaded when it changes. Found via
discovery_archive/south_america/CHILE_GAS_DEMAND_DISCOVERY*.py, which also
found no official gas demand split by sector covering 2021 on (CNE's monthly
consumption-by-sector file ends Oct 2018), so none is included.
Not reachable from the editing sandbox; runs in GitHub Actions
(OCR needs the tesseract-ocr and tesseract-ocr-spa packages).
"""
import argparse
import datetime
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
OCR_DPI = 300

# Imports table: header "Combustible [miles de Ton] Mensual Anual", then rows such as
#   "Gas Natural 215 -25,6% -48,1% Gas Natural 0 n/d n/d"   (imports row, then the exports table's row)
# OCR adds arrow icons and stray marks ("4 Gas Natural 444 Y 533% Am 225% 4 Gas Natural 52 y -15% n/d") and
# often drops the decimal comma, so a change is only taken when it reads as one-decimal "-12,6%".
HEADER_RE = re.compile(r"miles\s*de\s*Ton", re.I)
GAS_RE = re.compile(r"Gas\s*Natur\w*\W{0,3}?\s*(\d{1,3}(?:,\d{1,2})?)(?!\d)", re.I)
PCT_RE = re.compile(r"[-+\u2212\u2013\u2014]?\s?\d{1,3},\d\s?%|>\s?100\s?%|n\s?/\s?[ad]|[-+]?\d{1,4}\s?%")
MONTH_RE = re.compile(r"corresponden?\s+al\s+mes\s+de\s+([A-Za-zé]+)\s+(?:de|del)?\s*(20\d\d)", re.I)
ORIGIN_RE = re.compile(r"gas\s+natural\s+(?:tra[i\u00ed]do\s+)?desde\s+([^.;]+)", re.I)


def out(*a):
    print(*a, flush=True)


def num(s):
    s = s.strip()
    if re.fullmatch(r"\d{1,3}(\.\d{3})+(,\d+)?", s):
        s = s.replace(".", "")
    return float(s.replace(",", "."))


def pct(s):
    """'-25,6%' -> -0.256; 'n/d', '>100%' and OCR tokens without the decimal comma ('-126%') -> None."""
    s = re.sub(r"\s", "", s).replace("\u2212", "-").replace("\u2013", "-").replace("\u2014", "-")
    if not re.fullmatch(r"[-+]?\d{1,3},\d%", s):
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
    """Imports-table gas row: (kt, monthly change, annual change). The imports table is the left-hand one,
    so its gas row starts the line; a second 'Gas Natural' further along is the exports table."""
    lines = (text or "").splitlines()
    for h, ln in enumerate(lines):
        if not HEADER_RE.search(ln) or not re.search(r"Combustible|Mensual", ln, re.I):
            continue
        for row in lines[h + 1:h + 9]:
            m = GAS_RE.search(row)
            if not m or m.start() > 8:
                continue
            kt = num(m.group(1))
            if kt <= 0:
                continue
            rest = re.split(r"Gas\s*Natur", row[m.end():], flags=re.I)[0]
            toks = PCT_RE.findall(rest)
            return kt, (pct(toks[0]) if toks else None), (pct(toks[1]) if len(toks) > 1 else None)
    return None


def month_and_origin(text):
    flat = re.sub(r"\s+", " ", re.sub(r"-\s*\n\s*", "", text or ""))
    dm = MONTH_RE.search(flat)
    month = None
    if dm and dm.group(1).lower() in MESES:
        month = pd.Period(year=int(dm.group(2)), month=MESES[dm.group(1).lower()], freq="M")
    origin = ORIGIN_RE.search(flat)
    origin = re.split(r",?\s+El(?:\s|$)", origin.group(1))[0].strip(" ,") if origin else None  # OCR reads '. El' as ', El'
    return month, origin


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
    cands.sort(key=lambda i: abs(i - 0.62 * n))
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
        m, nxt, ny = str(p), str(p + 1), str(p + 12)
        if m in kt:
            continue
        if nxt in kt and pd.notna(mom[nxt]) and mom[nxt] > -1:
            v, how, src = kt[nxt] / (1 + mom[nxt]), f"{nxt} value / (1 + its monthly change {mom[nxt]:+.1%})", nxt
        elif ny in kt and pd.notna(yoy[ny]) and yoy[ny] > -1:
            v, how, src = kt[ny] / (1 + yoy[ny]), f"{ny} value / (1 + its annual change {yoy[ny]:+.1%})", ny
        else:
            out(f"  {m}: no data (no report table for it, and no published change to derive it from)")
            continue
        v = round(v, 1)
        add.append({"Month": m, "Imports_kt": v, "MoM_change_published": None, "YoY_change_published": None,
                    "Origins": None, "Method": f"derived: {how}", "Report": rep[src]})
        out(f"  {m}: no report table; derived {v} kt = {how}")
    if add:
        df = pd.concat([df, pd.DataFrame(add)], ignore_index=True)
    return df


PROD_SHEET = "produccion crudo gas m_energia"
PROD_FALLBACK = "https://www.cne.cl/wp-content/uploads/2024/09/Produccion_combustibles-junio-2024.xlsx"


def production_file_url():
    """Newest monthly 'Produccion_combustibles-<month>-<year>.xlsx' in CNE's media library (Estadisticas >
    Hidrocarburo). The annual-only versions ('Produccion_anual...', 'Produccion_combustibles-2023') are skipped."""
    try:
        r = requests.get("https://www.cne.cl/wp-json/wp/v2/media", headers=H, timeout=T,
                         params={"search": "Produccion_combustibles", "per_page": 100})
        urls = [s.replace("\\/", "/") for s in re.findall(r'"source_url":"([^"]+)"', r.text)] \
            if r.status_code == 200 else []
    except requests.RequestException:
        urls = []
    best = None
    for u in urls:
        m = re.search(r"Produccion_combustibles-([a-z]+)-(20\d\d)\.xlsx$", u, re.I)
        if m and m.group(1).lower() in MESES:
            key = (int(m.group(2)), MESES[m.group(1).lower()])
            if best is None or key > best[0]:
                best = (key, u)
    return best[1] if best else PROD_FALLBACK


def parse_production(content, url):
    """Monthly gas production (thousand m3) by ENAP and CEOP (private operators under special petroleum operation
    contracts), Ministerio de Energia data, sheet 'produccion crudo gas m_energia':
    Mes | ENAP Petroleo (m3) | ENAP Gas (miles de m3) | CEOP Petroleo | CEOP Gas | Total Petroleo | Total Gas."""
    df = pd.read_excel(io.BytesIO(content), sheet_name=PROD_SHEET, header=None)
    rows = []
    for _, r in df.iterrows():
        if not isinstance(r.iloc[0], datetime.datetime):  # header, notes and annual rows are skipped
            continue
        d = pd.Timestamp(r.iloc[0])
        vals = pd.to_numeric(r.iloc[[2, 4, 6]], errors="coerce").tolist()
        if any(pd.isna(v) for v in vals):
            continue
        enap, ceop, total = vals
        if abs(enap + ceop - total) > max(1.0, 0.005 * total):
            out(f"  production {d:%Y-%m}: ENAP {enap:.0f} + CEOP {ceop:.0f} != total {total:.0f}")
        rows.append({"Month": d.strftime("%Y-%m"), "ENAP_thousand_m3": enap, "CEOP_thousand_m3": ceop,
                     "Total_thousand_m3": total, "Source_file": url})
    p = pd.DataFrame(rows)
    if p.empty:
        return p
    dup = p[p.duplicated("Month", keep=False)]
    if len(dup):
        out(f"  production: repeated months in the file (first kept): {sorted(set(dup['Month']))}")
    return p.drop_duplicates("Month", keep="first")


def production(out_path, full):
    """Domestic gas production sheet, monthly from 2021. Re-downloaded only when CNE posts a newer file."""
    have = pd.DataFrame()
    if os.path.exists(out_path) and not full:
        try:
            have = pd.read_excel(out_path, sheet_name="Domestic production")
            have = have.drop(columns=[c for c in have.columns if str(c).startswith("Unnamed")])
            have["Month"] = have["Month"].astype(str)
        except ValueError:  # sheet not there yet
            have = pd.DataFrame()
    url = production_file_url()
    if len(have) and url in set(have["Source_file"]):
        out(f"production: {url} already read; {len(have)} months kept")
        return have
    try:
        r = requests.get(url, headers=H, timeout=T)
        r.raise_for_status()
        p = parse_production(r.content, url)
    except Exception as e:
        out(f"production: could not read {url}: {type(e).__name__} {str(e)[:120]}")
        return have
    p = p[p["Month"] >= str(START)]
    if p.empty:
        out(f"production: no 2021+ months in {url}")
        return have
    if len(have):  # keep months the old file has and the new one doesn't
        p = pd.concat([have[~have["Month"].isin(p["Month"])], p], ignore_index=True)
    p = p.sort_values("Month").reset_index(drop=True)
    days = pd.PeriodIndex(p["Month"], freq="M").days_in_month
    for c in ("ENAP", "CEOP", "Total"):
        p[f"{c}_mcm_per_day"] = (p[f"{c}_thousand_m3"] / 1000 / days).round(3)
    p = p[["Month", "ENAP_mcm_per_day", "CEOP_mcm_per_day", "Total_mcm_per_day", "ENAP_thousand_m3",
           "CEOP_thousand_m3", "Total_thousand_m3", "Source_file"]]
    out(f"production: {len(p)} months {p['Month'].min()}..{p['Month'].max()} from {url}")
    out(p.drop(columns=["Source_file"]).tail(6).to_string(index=False))
    return p


USE_FALLBACK = "https://www.cne.cl/wp-content/uploads/2026/06/importaciones-web.xlsx"
USE_COLS = ["Pipeline_energy_II_region_mcm_per_day", "Pipeline_energy_RM_V_region_mcm_per_day",
            "Pipeline_energy_VIII_region_mcm_per_day", "Pipeline_petrochemical_Magallanes_mcm_per_day",
            "LNG_II_region_Mejillones_mcm_per_day", "LNG_V_region_Quintero_mcm_per_day"]


def imports_file_url():
    """Newest 'importaciones-web*.xlsx' in CNE's media library (one upload folder per update)."""
    try:
        r = requests.get("https://www.cne.cl/wp-json/wp/v2/media", headers=H, timeout=T,
                         params={"search": "importaciones-web", "per_page": 100})
        urls = [s.replace("\\/", "/") for s in re.findall(r'"source_url":"([^"]+)"', r.text)] \
            if r.status_code == 200 else []
    except requests.RequestException:
        urls = []
    best = None
    for u in urls:
        m = re.search(r"/uploads/(20\d\d)/(\d\d)/importaciones-web(-\d+)?\.xlsx$", u)
        if m:
            key = (int(m.group(1)), int(m.group(2)), int((m.group(3) or "-0")[1:]))
            if best is None or key > best[0]:
                best = (key, u)
    return best[1] if best else USE_FALLBACK


def month_rows(df, first_col=1):
    """{YYYY-MM: row} for the monthly rows of a CNE import sheet (column B holds a date; totals are text)."""
    rows = {}
    for _, r in df.iterrows():
        d = r.iloc[first_col]
        if isinstance(d, datetime.datetime):
            rows.setdefault(d.strftime("%Y-%m"), r)
    return rows


def parse_imports_by_use(content, url):
    """CNE 'importaciones-web.xlsx' (Camara de Comercio de Santiago customs data), monthly:
    'GAS NATURAL GASEOSO' (pipeline gas from Argentina, million m3): energy use by region (II, RM-V, VIII; cols C-E)
    and petrochemical use in Magallanes (col J); 'GAS NATURAL LICUADO' (LNG, tonnes) by region (II, V; cols C-D)."""
    x = pd.ExcelFile(io.BytesIO(content))
    g = pd.read_excel(x, sheet_name="GAS NATURAL GASEOSO", header=None)
    lng = pd.read_excel(x, sheet_name="GAS NATURAL LICUADO", header=None)
    # check the layout before reading by position
    head_g = " ".join(str(v) for v in g.iloc[5:9].values.ravel())
    head_l = " ".join(str(v) for v in lng.iloc[4:8].values.ravel())
    for need, txt in (("USO ENERGETIC", head_g), ("PETROQUIMICO", head_g), ("II REGION", head_g),
                      ("RM - V REGION", head_g), ("VIII REGION", head_g), ("V REGION", head_l), ("TONELADAS", head_l)):
        if need not in txt:
            raise ValueError(f"layout changed: '{need}' not found in the sheet headers")
    gr, lr = month_rows(g), month_rows(lng)
    rows = []
    for m in sorted(set(gr) | set(lr)):
        if m < str(START):
            continue
        a = pd.to_numeric(gr[m].iloc[[2, 3, 4, 9]], errors="coerce").tolist() if m in gr else [None] * 4
        b = pd.to_numeric(lr[m].iloc[[2, 3]], errors="coerce").tolist() if m in lr else [None] * 2
        rows.append({"Month": m, "Pipeline_energy_II_region_mcm": a[0], "Pipeline_energy_RM_V_region_mcm": a[1],
                     "Pipeline_energy_VIII_region_mcm": a[2], "Pipeline_petrochemical_Magallanes_mcm": a[3],
                     "LNG_II_region_Mejillones_t": b[0], "LNG_V_region_Quintero_t": b[1]})
    df = pd.DataFrame(rows)
    vals = df.drop(columns="Month").fillna(0)
    # months not yet published are zero placeholders: drop trailing months with nothing in any column
    filled = df.loc[vals.sum(axis=1) > 0, "Month"]
    df = df[df["Month"] <= filled.max()] if len(filled) else df.iloc[0:0]
    df["Source_file"] = url
    return df


def imports_by_use(out_path, full):
    """'Imports by use' sheet: CNE's monthly import workbook, from 2021. Re-downloaded only when CNE posts a new one."""
    have = pd.DataFrame()
    if os.path.exists(out_path) and not full:
        try:
            have = pd.read_excel(out_path, sheet_name="Imports by use")
            have = have.drop(columns=[c for c in have.columns if str(c).startswith("Unnamed")])
            have["Month"] = have["Month"].astype(str)
        except ValueError:
            have = pd.DataFrame()
    url = imports_file_url()
    if len(have) and url in set(have["Source_file"]):
        out(f"imports by use: {url} already read; {len(have)} months kept")
        return have
    try:
        r = requests.get(url, headers=H, timeout=(10, 300))
        r.raise_for_status()
        u = parse_imports_by_use(r.content, url)
    except Exception as e:
        out(f"imports by use: could not read {url}: {type(e).__name__} {str(e)[:160]}")
        return have
    if u.empty:
        out(f"imports by use: no 2021+ months in {url}")
        return have
    if len(have):
        keep = have[~have["Month"].isin(u["Month"])]
        u = pd.concat([keep[[c for c in u.columns if c in keep.columns]], u], ignore_index=True)
    u = u.sort_values("Month").reset_index(drop=True)
    days = pd.PeriodIndex(u["Month"], freq="M").days_in_month
    for c in ("Pipeline_energy_II_region", "Pipeline_energy_RM_V_region", "Pipeline_energy_VIII_region",
              "Pipeline_petrochemical_Magallanes"):
        u[f"{c}_mcm_per_day"] = (u[f"{c}_mcm"] / days).round(3)
    for c in ("LNG_II_region_Mejillones", "LNG_V_region_Quintero"):
        u[f"{c}_mcm_per_day"] = (u[f"{c}_t"] * M3_PER_TONNE / 1e6 / days).round(3)
    u["Total_mcm_per_day_approx"] = u[USE_COLS].sum(axis=1, min_count=1).round(2)
    raw = ["Pipeline_energy_II_region_mcm", "Pipeline_energy_RM_V_region_mcm", "Pipeline_energy_VIII_region_mcm",
           "Pipeline_petrochemical_Magallanes_mcm", "LNG_II_region_Mejillones_t", "LNG_V_region_Quintero_t"]
    u = u[["Month"] + USE_COLS + ["Total_mcm_per_day_approx"] + raw + ["Source_file"]]
    out(f"imports by use: {len(u)} months {u['Month'].min()}..{u['Month'].max()} from {url}")
    out(u[["Month"] + USE_COLS + ["Total_mcm_per_day_approx"]].tail(6).to_string(index=False))
    return u


PIPE_T_PER_MCM = 756.0  # the import workbook's stated density of pipeline gas, tonnes per million m3


def combine(rep, use):
    """Main monthly series: CNE's import workbook where it has the month (customs data, split by use; pipeline
    gas converted to tonnes at the workbook's 756 t per million m3, LNG in tonnes), else the monthly report.
    Returns (combined frame, comparison frame for months in both)."""
    rep = rep.copy()
    days = pd.PeriodIndex(rep["Month"], freq="M").days_in_month
    rep["Report_kt"] = rep["Imports_kt"]
    rep["Report_mcm_per_day"] = (rep["Imports_kt"] * 1000 * M3_PER_TONNE / 1e6 / days).round(2)
    if use is None or use.empty:
        rep["Workbook_kt"] = None
        u = pd.DataFrame(columns=["Month", "Workbook_kt", "Workbook_mcm_per_day"])
    else:
        pipe = use[["Pipeline_energy_II_region_mcm", "Pipeline_energy_RM_V_region_mcm",
                    "Pipeline_energy_VIII_region_mcm", "Pipeline_petrochemical_Magallanes_mcm"]].sum(axis=1)
        lng = use[["LNG_II_region_Mejillones_t", "LNG_V_region_Quintero_t"]].sum(axis=1)
        u = pd.DataFrame({"Month": use["Month"], "Workbook_kt": (pipe * PIPE_T_PER_MCM + lng) / 1000,
                          "Workbook_mcm_per_day": use["Total_mcm_per_day_approx"]})
        u["Workbook_kt"] = u["Workbook_kt"].round(1)
    d = rep.merge(u, on="Month", how="outer").sort_values("Month").reset_index(drop=True)
    wb = d["Workbook_kt"].notna()
    d["Imports_kt"] = d["Workbook_kt"].where(wb, d["Report_kt"])
    d["Imports_mcm_per_day_approx"] = d["Workbook_mcm_per_day"].where(wb, d["Report_mcm_per_day"])
    derived = d["Method"].astype(str).str.startswith("derived")
    d["Source"] = "CNE import workbook (customs)"
    d.loc[~wb, "Source"] = "CNE Reporte Mensual" + derived[~wb].map({True: " (derived)", False: ""})
    both = wb & d["Report_kt"].notna()
    d["Workbook_vs_report_pct"] = (d["Workbook_kt"] / d["Report_kt"] - 1).where(both).round(4)
    cmp = d.loc[both, ["Month", "Workbook_kt", "Report_kt", "Workbook_vs_report_pct", "Method"]]
    if len(cmp):
        a = cmp["Workbook_vs_report_pct"].abs()
        out(f"\nworkbook vs monthly report, {len(cmp)} overlapping months: mean |diff| {a.mean():.1%}, "
            f"median {a.median():.1%}, max {a.max():.1%} ({cmp.loc[a.idxmax(), 'Month']}); "
            f"within 1%: {(a <= 0.01).sum()}, within 3%: {(a <= 0.03).sum()}")
        for _, r in cmp[a > 0.03].iterrows():
            out(f"    {r['Month']}: workbook {r['Workbook_kt']:g} kt vs report {r['Report_kt']:g} kt "
                f"({r['Workbook_vs_report_pct']:+.1%}; report {r['Method']})")
    d = d[d["Imports_kt"].notna() & (d["Month"] >= str(START))]
    d = d[["Month", "Imports_kt", "Imports_mcm_per_day_approx", "Source", "Workbook_kt", "Report_kt",
           "Workbook_vs_report_pct", "MoM_change_published", "YoY_change_published", "Origins", "Method", "Report"]]
    return d.reset_index(drop=True), cmp


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
            if "Report_kt" in have.columns:  # combined sheet: get back the report-only series
                have["Imports_kt"] = have["Report_kt"]
                have = have[have["Report_kt"].notna()]
            have = have[~have["Method"].astype(str).str.startswith("derived")]  # re-derive each run
            have = have[["Month", "Imports_kt", "MoM_change_published", "YoY_change_published", "Origins",
                         "Method", "Report"]]
    # report months already parsed (from the Report URL); these aren't downloaded again
    done = {m.group(1) for u in (have["Report"] if len(have) else [])
            for m in [re.search(r"RMensual_v(\d{6})", str(u))] if m}

    def tag(u):
        m = re.search(r"RMensual_v(\d{6})", str(u))
        return pd.Period(f"{m.group(1)[:4]}-{m.group(1)[4:]}", "M") if m else None

    # last accepted table per report month, to spot reports that repeat an earlier report's table
    accepted = {tag(r["Report"]): (r["Imports_kt"], r["MoM_change_published"]) for _, r in have.iterrows()} \
        if len(have) else {}

    def repeats_previous(p, kt, mom):
        before = [q for q in accepted if q is not None and q < p]
        if not before:
            return False
        pkt, pmom = accepted[max(before)]
        return kt == pkt and (mom is None or pd.isna(pmom) or abs(mom - pmom) < 1e-6)

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
            p += 1
            continue
        r = parse(c)
        if r is None:
            out(f"  {p}: no 'Gas Natural' import line found ({u})")
            p += 1
            continue
        month, mom_s = r["month"], (f"{r['mom']:+.1%}" if r["mom"] is not None else "n/a")
        desc = f"{r['kt']:g} kt, m/m {mom_s}, y/y " + (f"{r['yoy']:+.1%}" if r["yoy"] is not None else "n/a")
        if month is not None and not (p - 4 <= month <= p - 1):
            out(f"  report {p}: SKIPPED, stale: text says data month {month} ({desc})")
        elif repeats_previous(p, r["kt"], r["mom"]):
            out(f"  report {p}: SKIPPED, repeats the previous report's table ({desc}; text says {month})")
        else:
            how = "month stated" if month is not None else "month assumed"
            month = month if month is not None else p - LAG
            accepted[p] = (r["kt"], r["mom"])
            rows.append({"Month": str(month), "Imports_kt": r["kt"], "MoM_change_published": r["mom"],
                         "YoY_change_published": r["yoy"], "Origins": r["origin"],
                         "Method": f"{r['method']} p{r['page']}; {how}", "Report": u})
            out(f"  report {p}: data {month} ({r['method']} p{r['page']}, {how}) {desc} [{r['origin']}]")
        p += 1

    df = pd.concat([have, pd.DataFrame(rows)], ignore_index=True) if len(have) else pd.DataFrame(rows)
    if df.empty:
        out("nothing parsed")
        sys.exit(1)
    df = df.drop_duplicates("Month", keep="last").sort_values("Month").reset_index(drop=True)
    df = df[df["Month"] >= str(START)]
    df = check_and_fill(df)
    df = df.sort_values("Month").reset_index(drop=True)
    rep_gaps = sorted(set(pd.period_range(df["Month"].min(), df["Month"].max(), freq="M").astype(str))
                      - set(df["Month"]))
    use = imports_by_use(args.out, args.full)
    df, cmp = combine(df, use)
    months = pd.period_range(df["Month"].min(), df["Month"].max(), freq="M").astype(str)
    gaps = sorted(set(months) - set(df["Month"]))
    out(f"\n{len(df)} months {df['Month'].min()}..{df['Month'].max()}; missing: {gaps or 'none'}")
    out(df.drop(columns=["Report"]).to_string(index=False))

    wb_last = use["Month"].max() if len(use) else None
    n_cmp = len(cmp)
    mean_diff = f"{cmp['Workbook_vs_report_pct'].abs().mean():.1%}" if n_cmp else "n/a"
    notes = [
        "UNITS",
        "Imports_kt: natural gas imports in thousand tonnes per month (LNG cargoes plus pipeline gas from "
        "Argentina), customs data published by CNE. Imports_mcm_per_day_approx: million m3 per day (monthly "
        f"average); LNG at {M3_PER_TONNE:.0f} m3 of gas per tonne (a standard conversion; approximate), pipeline "
        "gas as published in million m3.",
        "Source column: 'CNE import workbook (customs)' = CNE's monthly import statistics workbook "
        "(importaciones-web.xlsx), used for every month it covers; its pipeline gas (million m3) is converted to "
        f"tonnes at the workbook's own density, {PIPE_T_PER_MCM:.0f} t per million m3. 'CNE Reporte Mensual' = the "
        "monthly report's table, used for months after the workbook ends; '(derived)' marks a month the report "
        "does not show, derived from a neighbouring month's published change.",
        "Workbook_kt / Report_kt: the two sources side by side; Workbook_vs_report_pct = Workbook_kt / Report_kt - "
        "1 (fraction) for months in both. MoM_change_published / YoY_change_published: the report's own change on "
        "the previous month and on the same month a year earlier (fractions; blank where the report shows n/d).",
        "",
        "COVERAGE",
        f"Monthly from January 2021: import workbook to {wb_last or 'n/a'}, then the monthly report "
        f"(to {df['Month'].max()}). {n_cmp} months are in both sources; mean absolute difference {mean_diff} "
        "(both are the same customs data; the report rounds to whole kt and its 2021-2025 values are read by OCR).",
        "Monthly report parsing: the data month lags the report by about two months; it is read from the report's "
        "own text where possible ('month stated'), otherwise taken as report month - 2. Method: 'text' = PDF text "
        "layer; 'text+ocr' = table from the text layer, data month and origins by OCR; 'ocr' = the whole page by "
        "OCR (reports with no text layer); p<n> = PDF page. Report months with no table: "
        + (", ".join(rep_gaps) if rep_gaps else "none") + " (filled from the import workbook).",
        "Months with no value: " + (", ".join(gaps) if gaps else "none") + ".",
        "Origins lists the source countries the report names for that month's gas imports.",
        "",
        "SOURCE",
        "Comision Nacional de Energia (CNE): (1) Estadisticas > Hidrocarburo > 'Importaciones' workbook "
        "(importaciones-web.xlsx; Camara de Comercio de Santiago customs data), https://www.cne.cl/estadisticas/"
        "hidrocarburo/; (2) Reporte Mensual del Sector Energetico, section 'Importaciones y Exportaciones de "
        "Combustibles' (Aduana data via COMEX). The Report column holds the report PDF used.",
    ]
    if len(use):
        notes += [
            "",
            "IMPORTS BY USE",
            "Sheet 'Imports by use' (measured, CNE import workbook): pipeline gas from Argentina for energy use by "
            "region of entry - II region (Antofagasta, north), RM-V (Santiago/Valparaiso, central), VIII (Biobio, "
            "south) - and for petrochemical use in the Magallanes region (methanol, i.e. Methanex's Cabo Negro "
            "plant); LNG by region of the regasification terminal - II region (GNL Mejillones) and V region (GNL "
            "Quintero). *_mcm / *_t = as published (million m3; tonnes); *_mcm_per_day = million m3 per day, LNG at "
            f"{M3_PER_TONNE:.0f} m3/t. The workbook's own pipeline TOTAL column is not used (in some months it "
            "leaves out the VIII region); totals here are the sum of the regions.",
            "This is a split of imports, not of demand: it shows imported gas for petrochemical use, but Methanex "
            "also uses domestic Magallanes gas, which is not in it. Coverage: "
            f"{use['Month'].min()} to {use['Month'].max()}; the workbook is re-read when CNE uploads a new one.",
        ]
    prod = production(args.out, args.full)
    sheets = {"Gas imports": df}
    if len(use):
        sheets["Imports by use"] = use
    if len(prod):
        sheets["Domestic production"] = prod
        notes += [
            "",
            "DOMESTIC PRODUCTION",
            "Sheet 'Domestic production': monthly natural gas production in Chile (all of it in the Magallanes "
            "basin), measured, as published by CNE from Ministerio de Energia data: ENAP (state oil company) and "
            "CEOP (private operators under Contratos Especiales de Operacion Petrolera). Thousand m3 per month as "
            "published; *_mcm_per_day = million m3 per day (monthly average).",
            f"Coverage: {prod['Month'].min()} to {prod['Month'].max()} (2021 on). CNE's latest file stops there; "
            "the sheet extends when CNE posts a newer 'Produccion_combustibles-<month>-<year>.xlsx'. Where the "
            "file repeats a month (Nov 2021 appears twice, identical) the first is kept.",
            "Source: CNE, Estadisticas > Hidrocarburo > 'Produccion de combustibles' "
            "(https://www.cne.cl/estadisticas/hidrocarburo/), sheet 'produccion crudo gas m_energia'. The "
            "Source_file column holds the workbook read.",
            "Not included: gas demand by sector. No official monthly or annual split by sector covering 2021 on "
            "was found (CNE's monthly consumption-by-sector file ends Oct 2018; the CNE monthly report, INE and "
            "SEC publish none; see discovery_archive/south_america/CHILE_GAS_DEMAND_DISCOVERY*.py).",
        ]
    os.makedirs(os.path.dirname(os.path.abspath(args.out)), exist_ok=True)
    xlsx_notes.write_workbook(args.out, sheets, notes,
                              {"UNITS", "COVERAGE", "SOURCE", "IMPORTS BY USE",
                               "DOMESTIC PRODUCTION"})
    out(f"Saved {args.out}")


if __name__ == "__main__":
    main()
