"""
Ecuador natural gas: production and use by sector, monthly from 2021,
from EP Petroecuador's statistical reports (public PDFs):
  https://www.eppetroecuador.ec/?p=3721  ("Cifras Institucionales")
    - Informes Estadisticos Mensuales: current year to date (one PDF)
    - Informes Estadisticos Anuales: one PDF per past year (Jan-Dec)

Ecuador's only commercial gas is the offshore Amistad field (Gulf of
Guayaquil). Petroecuador's dispatch table splits its sales into:
  GAS NATURAL (MMBTU)          - pipeline gas to the Termogas Machala
                                 power plant  -> Power
  GAS NATURAL LICUADO (MMBTU)  - LNG from the Bajo Alto plant, trucked
                                 to industry (mostly Cuenca ceramics) -> Industrial
plus Amistad production in barrels of oil equivalent per day.

The PDF text layer splits numbers with stray spaces ("5 93.016"), so
values are read by x-position: each number fragment is assigned to the
nearest month column of the table header, then the fragments are joined.

Units: MMBtu per day (monthly dispatch / days in month); also approx.
million cubic feet per day at 1,030 Btu/ft3. Production BOE/d is
converted at 5.8 MMBtu/BOE.
Not reachable from the editing sandbox; runs in GitHub Actions.
"""
import argparse
import io
import os
import re
import sys
from urllib.parse import urljoin

import pandas as pd
import pdfplumber
import requests

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))  # repo root, for xlsx_notes
import xlsx_notes

PAGE = "https://www.eppetroecuador.ec/?p=3721"
H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0 Safari/537.36"}
T = (10, 180)
DATA_START = 2021
MONTHS = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto", "septiembre", "octubre",
          "noviembre", "diciembre"]
BTU_PER_FT3 = 1030.0
MMBTU_PER_BOE = 5.8


def out(*a):
    print(*a, flush=True)


def list_reports():
    """(label, url) for the current monthly report and each annual report
    from DATA_START on, from the 'Informes Estadisticos' blocks."""
    r = requests.get(PAGE, headers=H, timeout=T)
    r.raise_for_status()
    html = r.text
    i = html.find("Mensuales")
    j = html.find("Exploraci", i)
    block = html[i:j] if i > 0 and j > i else html
    reports = []
    for href, txt in re.findall(r'<a[^>]+href="([^"]+)"[^>]*>(.*?)</a>', block, re.S | re.I):
        t = re.sub(r"<[^>]+>|\s+", " ", txt).replace("&#8211;", "-").strip()
        years = [int(y) for y in re.findall(r"(20\d\d)", t)]
        if not years or max(years) < DATA_START:
            continue
        reports.append((t, urljoin(PAGE, href.replace("&amp;", "&"))))
    return reports


def es_number(s):
    s = s.replace(" ", "")
    if s in ("", "-", "--"):
        return 0.0
    if "," in s:
        s = s.replace(".", "").replace(",", ".")
    else:
        s = s.replace(".", "")
    try:
        return float(s)
    except ValueError:
        return None


def month_columns(words):
    """[(month number, x centre)] from the table header line."""
    by_top = {}
    for w in words:
        t = w["text"].lower().strip(".")
        if t in MONTHS:
            by_top.setdefault(round(w["top"]), []).append((MONTHS.index(t) + 1, (w["x0"] + w["x1"]) / 2))
    if not by_top:
        return []
    best = max(by_top.values(), key=len)
    return sorted(best, key=lambda mc: mc[1])


def parse_report(content):
    disp, prod = {}, {}
    with pdfplumber.open(io.BytesIO(content)) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ""
            if "GAS" not in text.upper():
                continue
            ym = re.search(r"A[ÑN]O\s+(20\d\d)", text, re.I) or re.search(r"(20\d\d)", text)
            if not ym:
                continue
            year = int(ym.group(1))
            words = page.extract_words(keep_blank_chars=False, x_tolerance=1.5)
            cols = month_columns(words)
            if len(cols) < 1:
                continue
            per_day = re.search(r"barriles\s*/\s*d[ií]a", text, re.I)
            volumes = re.search(r"Cifras en barriles(?!\s*/)", text, re.I) and not re.search(r"US\s*\$", text)
            if volumes and not disp and re.search(r"GAS NATURAL\s*\(MMBTU\)", text):
                g, l = None, None
                for ln_word in [w for w in words if w["text"].upper() == "GAS"]:
                    rest = " ".join(w["text"] for w in words if abs(w["top"] - ln_word["top"]) < 3
                                    and w["x0"] >= ln_word["x0"])[:40].upper()
                    if rest.startswith("GAS NATURAL LICUADO") and l is None:
                        l = _row_at(words, ln_word, cols, year)
                    elif rest.startswith("GAS NATURAL (MMBTU)") and g is None:
                        g = _row_at(words, ln_word, cols, year)
                if g:
                    disp["Power_MMBtu_month"] = g
                if l:
                    disp["Industrial_LNG_MMBtu_month"] = l
            if per_day and not prod and re.search(r"GAS CAMPO AMISTAD", text, re.I):
                for w in words:
                    if w["text"].upper() == "GAS":
                        rest = " ".join(x["text"] for x in words if abs(x["top"] - w["top"]) < 3
                                        and x["x0"] >= w["x0"])[:30].upper()
                        if rest.startswith("GAS CAMPO AMISTAD"):
                            prod["Amistad_production_BOE_per_day"] = _row_at(words, w, cols, year)
                            break
    return disp, prod


def _row_at(words, first_word, cols, year):
    """Row values to the right of a row label that starts at first_word."""
    top = first_word["top"]
    label_end = max(w["x1"] for w in words if abs(w["top"] - top) < 3 and not re.fullmatch(r"[\d.,\-]+", w["text"])
                    and w["x0"] >= first_word["x0"] and w["x0"] < cols[0][1] - 30)
    line = [w for w in words if abs(w["top"] - top) < 3 and w["x0"] > label_end]
    xs = [x for _, x in cols]
    spacing = (xs[-1] - xs[0]) / max(len(xs) - 1, 1) if len(xs) > 1 else 40
    frag = {m: [] for m, _ in cols}
    for w in sorted(line, key=lambda w: w["x0"]):
        if not re.fullmatch(r"[\d.,\-]+", w["text"]):
            continue
        xc = (w["x0"] + w["x1"]) / 2
        m, x = min(cols, key=lambda mc: abs(mc[1] - xc))
        if abs(x - xc) <= spacing * 0.55:
            frag[m].append(w["text"])
    vals = {}
    for m, parts in frag.items():
        if parts:
            v = es_number("".join(parts))
            if v is not None:
                vals[pd.Timestamp(year=year, month=m, day=1)] = v
    return vals


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", default="output/ecuador_gas.xlsx")
    ap.add_argument("--full", action="store_true", help="re-read every annual report, ignoring the archive")
    args = ap.parse_args()

    have = pd.DataFrame()
    if os.path.exists(args.out) and not args.full:
        have = pd.read_excel(args.out, sheet_name="Gas by use")
        have = have.drop(columns=[c for c in have.columns if str(c).startswith("Unnamed")])
        have["Month"] = pd.to_datetime(have["Month"].astype(str))
    complete_years = set()
    if len(have):
        n = have.dropna(subset=["Power_MMBtu_per_day"]).groupby(have["Month"].dt.year)["Month"].nunique()
        complete_years = {int(y) for y, k in n.items() if k == 12}
        out(f"archive: {len(have)} months; complete years {sorted(complete_years)}")

    reports = list_reports()
    out(f"{len(reports)} reports listed: {[t for t, _ in reports]}")
    rows = {}
    for label, url in reports:
        years = [int(y) for y in re.findall(r"(20\d\d)", label)]
        if max(years) in complete_years and not re.search(r"mes|enero\s*-", label, re.I):
            continue
        try:
            r = requests.get(url, headers=H, timeout=T)
            if r.content[:4] != b"%PDF":
                out(f"  {label}: not a PDF")
                continue
            disp, prod = parse_report(r.content)
        except Exception as e:
            out(f"  {label}: ERR {type(e).__name__}: {str(e)[:150]}")
            continue
        n_d = {k: len(v) for k, v in disp.items()}
        n_p = {k: len(v) for k, v in prod.items()}
        out(f"  {label}: dispatch months {n_d}, production months {n_p}")
        for series in (disp, prod):
            for col, vals in series.items():
                for m, v in vals.items():
                    rows.setdefault(m, {})[col] = v

    new = pd.DataFrame.from_dict(rows, orient="index").sort_index()
    if new.empty and have.empty:
        out("nothing parsed")
        sys.exit(1)
    if not new.empty:
        new.index.name = "Month"
        new = new.reset_index()
        days = new["Month"].dt.days_in_month
        for col in ("Power_MMBtu_month", "Industrial_LNG_MMBtu_month"):
            if col not in new.columns:
                new[col] = None
        new["Power_MMBtu_per_day"] = (new["Power_MMBtu_month"] / days).round(0)
        new["Industrial_LNG_MMBtu_per_day"] = (new["Industrial_LNG_MMBtu_month"] / days).round(0)
        new = new.drop(columns=["Power_MMBtu_month", "Industrial_LNG_MMBtu_month"])
    df = pd.concat([have, new], ignore_index=True) if len(have) else new
    df = df.drop_duplicates("Month", keep="last").sort_values("Month").reset_index(drop=True)
    df = df[df["Month"].dt.year >= DATA_START]
    df["Total_sales_MMBtu_per_day"] = df[["Power_MMBtu_per_day", "Industrial_LNG_MMBtu_per_day"]].sum(axis=1, min_count=1)
    df["Total_sales_MMcf_per_day_approx"] = (df["Total_sales_MMBtu_per_day"] / BTU_PER_FT3).round(2)
    if "Amistad_production_BOE_per_day" in df.columns:
        df["Amistad_production_MMBtu_per_day"] = (df["Amistad_production_BOE_per_day"] * MMBTU_PER_BOE).round(0)
    df["Month"] = df["Month"].dt.strftime("%Y-%m")
    out(df.tail(14).to_string(index=False))
    months = pd.period_range(df["Month"].min(), df["Month"].max(), freq="M").astype(str)
    gaps = sorted(set(months) - set(df.dropna(subset=["Power_MMBtu_per_day"])["Month"]))
    out(f"{len(df)} months {df['Month'].min()}..{df['Month'].max()}; missing dispatch: {gaps or 'none'}")

    notes = [
        "UNITS",
        "MMBtu per day = monthly dispatch in MMBtu (as published) / days in month. MMcf/d approx at 1,030 Btu per "
        f"cubic foot. Amistad production is published in barrels of oil equivalent per day; MMBtu at "
        f"{MMBTU_PER_BOE} MMBtu/BOE.",
        "",
        "SECTORS",
        "Power_MMBtu_per_day: Petroecuador's 'GAS NATURAL (MMBTU)' dispatches - pipeline gas from the Amistad field "
        "to the Termogas Machala power plant. Industrial_LNG_MMBtu_per_day: 'GAS NATURAL LICUADO (MMBTU)' - LNG "
        "from the Bajo Alto liquefaction plant, trucked to industrial users (mainly ceramics in Cuenca). These two "
        "are effectively all of Ecuador's commercial natural gas.",
        "",
        "COVERAGE",
        f"Monthly from {DATA_START}. Past years from Petroecuador's annual statistical reports, the current year from "
        "the year-to-date monthly report. Runs are incremental: complete years in this archive aren't re-read.",
        "",
        "SOURCE",
        f"EP Petroecuador, Informes Estadisticos (Cifras Institucionales): {PAGE}",
    ]
    os.makedirs(os.path.dirname(os.path.abspath(args.out)), exist_ok=True)
    xlsx_notes.write_workbook(args.out, {"Gas by use": df}, notes, {"UNITS", "SECTORS", "COVERAGE", "SOURCE"})
    out(f"Saved {args.out}")


if __name__ == "__main__":
    main()
