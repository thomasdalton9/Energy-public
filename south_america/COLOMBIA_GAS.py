"""
Colombia natural gas demand by sector, monthly from 2021, parsed from
the Gestor del Mercado de Gas (Bolsa Mercantil de Colombia) monthly
reports. These are public PDFs; the BI Gas dashboard with the same data
needs a login and a reCAPTCHA.
  https://www.bmcbec.com.co/informes/informes-mensuales

Each report has a national demand-by-sector table: one row for the
Costa (Caribbean coast) region and one for the Interior, with one column
per sector. The script sums the two rows for each sector, then checks
the result against the report's headline "demanda promedio de N GBTUD".
If the regional table can't be read, it falls back to the per-department
table, where each sector row ends in a national total. Power is then
the headline total minus the other sectors. The Source column says
which method was used.

Units: GBTUD (billion BTU per day, a monthly average of daily energy).
1 GBTUD is roughly 0.027 million m3/day at Colombian gas heating values.

Not reachable from the editing sandbox; runs in GitHub Actions.
"""
import argparse
import io
import os
import re
import sys
from urllib.parse import unquote, urljoin

import pandas as pd
import requests

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))  # repo root, for xlsx_notes
import xlsx_notes

BASE = "https://www.bmcbec.com.co"
PAGE = BASE + "/informes/informes-mensuales"
H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0 Safari/537.36"}
T = (10, 120)
DATA_START = "2021-01"
# Re-parse the newest few months on every run in case a report is reissued.
RECHECK_MONTHS = 3

SECTORS = ["Power", "Industrial", "Residential", "Commercial", "Vehicle_CNG", "Refinery",
           "Petrochemical", "Compressors"]
MESES = {"enero": 1, "febrero": 2, "marzo": 3, "abril": 4, "mayo": 5, "junio": 6, "julio": 7,
         "agosto": 8, "septiembre": 9, "setiembre": 9, "octubre": 10, "noviembre": 11, "diciembre": 12}


def out(*a):
    print(*a, flush=True)


def strip_accents(s):
    return (s.lower().replace("á", "a").replace("é", "e").replace("í", "i").replace("ó", "o")
            .replace("ú", "u"))


def sector_of(word, prev):
    w = strip_accents(word)
    if w.startswith("generacion") or w.startswith("termoelec"):
        return "Power"
    if w.startswith("termica") and not strip_accents(prev).startswith("generacion"):
        return "Power"
    if w.startswith("industrial"):
        return "Industrial"
    if w.startswith("residencial"):
        return "Residential"
    if w.startswith("comercial"):
        return "Commercial"
    if w.startswith("gnv"):
        return "Vehicle_CNG"
    if w.startswith("refineri"):
        return "Refinery"
    if w.startswith("petroquimic"):
        return "Petrochemical"
    if w.startswith("compresor"):
        return "Compressors"
    return None


def num(s):
    s = str(s or "").strip().replace("%", "")
    if not s:
        return None
    # "1.234" thousands vs "6,8" decimal: these tables never exceed a few hundred GBTUD per cell
    if re.fullmatch(r"\d{1,3}(\.\d{3})+", s):
        s = s.replace(".", "")
    s = s.replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return None


# ---------------------------------------------------------------- listing

def list_pdfs():
    links = {}
    try:
        r = requests.get(PAGE, headers=H, timeout=T)
        for href in re.findall(r'href="([^"]+\.pdf)"', r.text, re.I):
            links[urljoin(BASE, href)] = 1
        out(f"listing (raw html): {len(links)} pdf links")
    except Exception as e:
        out(f"listing (raw html) failed: {e}")
    if len(links) < 40:
        from playwright.sync_api import sync_playwright
        with sync_playwright() as p:
            b = p.chromium.launch()
            pg = b.new_page(user_agent=H["User-Agent"])
            pg.goto(PAGE, timeout=90000, wait_until="load")
            pg.wait_for_timeout(6000)
            for _ in range(10):
                for a in pg.locator("a[href$='.pdf'], a[href*='.pdf?']").all():
                    try:
                        links[urljoin(BASE, a.get_attribute("href") or "")] = 1
                    except Exception:
                        pass
                nxt = pg.locator("a[rel=next], li.pager__item--next a, a:has-text('Siguiente'), a:has-text('Ver más')")
                if nxt.count() == 0:
                    break
                try:
                    nxt.first.click(timeout=5000)
                    pg.wait_for_timeout(3000)
                except Exception:
                    break
            b.close()
        out(f"listing (with playwright): {len(links)} pdf links")
    pdfs = []
    for u in links:
        name = unquote(u.rsplit("/", 1)[-1])
        if not re.search(r"informe\s*mensual|informe_mensual", name, re.I):
            continue
        m = re.search(r"/files/(20\d\d)-(\d\d)/", u)
        if not m:
            continue
        upload = f"{m.group(1)}-{m.group(2)}"
        if upload < DATA_START:
            continue
        pdfs.append((upload, u))
    return sorted(pdfs)


def guess_month(url, upload):
    """Report month from the file name (month word + year), else the
    month before the upload folder."""
    name = strip_accents(unquote(url.rsplit("/", 1)[-1]))
    mon = next((n for w, n in MESES.items() if w in name), None)
    yr = re.search(r"(20\d\d)", name)
    if mon and yr:
        return f"{yr.group(1)}-{mon:02d}"
    p = pd.Period(upload, "M") - 1
    if mon:
        y = p.year if mon <= p.month else p.year - 1
        return f"{y}-{mon:02d}"
    return str(p)


# ---------------------------------------------------------------- parsing

def report_month(texts):
    for t in texts[:4]:
        m = re.search(r"INFORME\s+MENSUAL\s+([A-ZÁÉÍÓÚa-záéíóú]+)\s+(?:DE\s+)?(20\d\d)", t, re.I)
        if m and strip_accents(m.group(1)) in MESES:
            return f"{m.group(2)}-{MESES[strip_accents(m.group(1))]:02d}"
    return None


def headline_total(texts):
    for t in texts:
        flat = re.sub(r"\s+", " ", t)
        m = re.search(r"demanda promedio de\s*([\d.,]+)\s*GBTUD", flat, re.I)
        if m:
            raw = m.group(1).rstrip(".,")
            # headline is always hundreds-to-thousands of GBTUD: "1,097" / "1.097" are thousands separators
            if re.fullmatch(r"\d{1,2}[.,]\d{3}", raw):
                return float(raw.replace(".", "").replace(",", ""))
            return num(raw)
    return None


def header_sectors(lines):
    """Sector column order, from the header line above the regional table."""
    best = []
    for ln in lines:
        words = ln.split()
        secs = []
        for i, w in enumerate(words):
            s = sector_of(w, words[i - 1] if i else "")
            if s and s not in secs:
                secs.append(s)
        if len(secs) > len(best):
            best = secs
    return best if len(best) >= 6 else []


def regional_rows_from_tables(tables, n):
    """Costa / Interior / TOTAL Nacional rows of the regional table.
    Cells are positional: a blank cell is a sector with no demand in that
    region (2021 reports leave e.g. petrochemical blank for the Interior)."""
    rows = {}
    for tb in tables:
        for row in tb:
            if not row or not row[0]:
                continue
            key = strip_accents(re.sub(r"\s+", " ", str(row[0])).strip())
            key = "total" if key.startswith("total") else key
            if key not in ("costa", "interior", "total"):
                continue
            cells = [c for c in row[1:]]
            if len(cells) == n:
                vals = [num(c) if str(c or "").strip() else 0.0 for c in cells]
            else:
                vals = [v for v in (num(c) for c in cells if str(c or "").strip()) if v is not None]
            if len(vals) == n and None not in vals:
                rows[key] = vals
    return rows


def regional_rows_from_text(lines, n):
    rows = {}
    for ln in lines:
        m = re.match(r"\s*(COSTA|INTERIOR|TOTAL(?:\s+NACIONAL)?)\s+([\d\s.,]+)$", ln, re.I)
        if m:
            vals = [num(x) for x in m.group(2).split()]
            key = "total" if m.group(1).lower().startswith("total") else m.group(1).lower()
            if len(vals) == n and None not in vals:
                rows[key] = vals
    return rows


def parse_regional(pages):
    """Method A: the Costa + Interior table on the '% Segmento' page."""
    for text, tables in pages:
        lines = text.splitlines()
        secs = header_sectors(lines)
        if not secs:
            continue
        rows = regional_rows_from_tables(tables, len(secs))
        for k, v in regional_rows_from_text(lines, len(secs)).items():
            rows.setdefault(k, v)
        if "costa" in rows and "interior" in rows:
            return {s: rows["costa"][i] + rows["interior"][i] for i, s in enumerate(secs)}
        if "total" in rows:
            return dict(zip(secs, rows["total"]))
    return None


def parse_department(pages, total):
    """Method B: per-department table, each sector row ending in its
    national total. The power row's numbers are usually split off its
    label, so power = headline total minus everything else."""
    for text, _ in pages:
        lines = text.splitlines()
        res = {}
        for ln in lines:
            m = re.match(r"\s*([A-Za-zÁÉÍÓÚáéíóú]+)\s+((?:[\d.,]+\s+){5,}[\d.,]+)\s*$", ln)
            if not m:
                continue
            s = sector_of(m.group(1), "")
            if s and s != "Power" and s not in res:
                res[s] = num(m.group(2).split()[-1])
        if len(res) >= 6 and total:
            res["Power"] = round(total - sum(v for v in res.values() if v), 1)
            return res
    return None


def parse_pdf(content):
    import pdfplumber
    pages = []
    with pdfplumber.open(io.BytesIO(content)) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ""
            tables = page.extract_tables() if re.search(r"Segmento|GBTUD|COSTA", text, re.I) else []
            pages.append((text, tables))
    texts = [p[0] for p in pages]
    month = report_month(texts)
    total = headline_total(texts)
    sec = parse_regional(pages)
    src = "regional table"
    if sec and total and abs(sum(sec.values()) - total) > max(0.04 * total, 8):
        out(f"    regional table sums to {sum(sec.values()):.0f} vs headline {total:.0f}; trying department table")
        sec = None
    if not sec:
        sec = parse_department(pages, total)
        src = "department table (power = total - others)"
    return month, total, sec, src, pages


def dump(pages):
    """Diagnostics for a report that didn't parse."""
    for i, (text, tables) in enumerate(pages):
        if re.search(r"Segmento|Residencial", text, re.I):
            out(f"    --- page {i + 1}")
            for ln in text.splitlines()[:40]:
                out("      ", ln[:160])
            for tb in tables[:3]:
                for row in tb[:12]:
                    out("       T", [str(x)[:14] if x else "" for x in row][:12])


# ---------------------------------------------------------------- main

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", default="output/colombia_gas_demand_by_sector.xlsx")
    ap.add_argument("--full", action="store_true", help="re-parse every report, ignoring the archive")
    args = ap.parse_args()

    have = pd.DataFrame()
    if os.path.exists(args.out) and not args.full:
        have = pd.read_excel(args.out, sheet_name="Demand by sector")
        have["Month"] = have["Month"].astype(str).str[:7]
        out(f"archive: {len(have)} months, latest {have['Month'].max() if len(have) else None}")
    done = set(have["Month"]) if len(have) else set()
    recheck = set(sorted(done)[-RECHECK_MONTHS:])

    pdfs = list_pdfs()
    out(f"{len(pdfs)} monthly reports uploaded since {DATA_START}")
    # newest upload wins when a month was re-issued
    by_month = {}
    for upload, u in pdfs:
        by_month[guess_month(u, upload)] = (upload, u)

    new_rows, failed = [], []
    for gm in sorted(by_month):
        upload, u = by_month[gm]
        if gm < DATA_START or (gm in done and gm not in recheck):
            continue
        try:
            c = requests.get(u, headers=H, timeout=T).content
            if c[:4] != b"%PDF":
                out(f"  {gm}: not a PDF ({u})")
                failed.append(gm)
                continue
            month, total, sec, src, pages = parse_pdf(c)
        except Exception as e:
            out(f"  {gm}: ERR {type(e).__name__}: {str(e)[:150]}")
            failed.append(gm)
            continue
        month = month or gm
        if month != gm:
            out(f"  note: file name suggests {gm}, report header says {month}; using header")
        if not sec:
            out(f"  {month}: could not parse sectors (headline total {total}) {u}")
            dump(pages)
            failed.append(month)
            continue
        row = {"Month": month, **{s: sec.get(s) for s in SECTORS}}
        row["Sum_of_sectors"] = round(sum(v for v in sec.values() if v is not None), 1)
        row["Reported_total"] = total
        row["Source"] = src
        row["Report"] = u
        new_rows.append(row)
        out(f"  {month}: sum {row['Sum_of_sectors']:.0f} vs reported {total} GBTUD [{src}]")

    df = pd.concat([have, pd.DataFrame(new_rows)], ignore_index=True) if len(have) else pd.DataFrame(new_rows)
    if df.empty:
        out("nothing parsed")
        sys.exit(1)
    df = df.drop_duplicates("Month", keep="last").sort_values("Month").reset_index(drop=True)
    df = df[["Month"] + SECTORS + ["Sum_of_sectors", "Reported_total", "Source", "Report"]]
    months = pd.period_range(df["Month"].min(), df["Month"].max(), freq="M").astype(str)
    gaps = sorted(set(months) - set(df["Month"]))
    out(f"\n{len(df)} months {df['Month'].min()}..{df['Month'].max()}; missing: {gaps or 'none'}; "
        f"failed this run: {failed or 'none'}")

    notes = [
        "UNITS",
        "GBTUD = billion BTU per day, the month's average daily energy delivered to end users. "
        "1 GBTUD is about 0.027 million m3/day (roughly 1 MMcf/d) at Colombian gas heating values.",
        "",
        "SECTORS",
        "Power (gas-fired thermal plants), Industrial, Residential, Commercial, Vehicle_CNG (GNVC), Refinery, "
        "Petrochemical, Compressors (gas burned by pipeline compressor stations on the national grid, SNT). "
        "Each is Costa (Caribbean coast) + Interior, as published.",
        "",
        "CHECKS",
        "Sum_of_sectors is compared to Reported_total, the report's headline 'demanda promedio' figure. "
        "Source says which table was read. 'department table' means the regional table could not be read, "
        "so Power is Reported_total minus the other sectors.",
        "",
        "SOURCE",
        f"Gestor del Mercado de Gas Natural (Bolsa Mercantil de Colombia), monthly reports: {PAGE}. "
        "Report holds the PDF used for each month. Figures come from SEGAS end-user delivery reports and "
        "may include estimates for distribution-network sectors.",
    ]
    os.makedirs(os.path.dirname(os.path.abspath(args.out)), exist_ok=True)
    xlsx_notes.write_workbook(args.out, {"Demand by sector": df}, notes, {"UNITS", "SECTORS", "CHECKS", "SOURCE"})
    out(f"Saved {args.out}")


if __name__ == "__main__":
    main()
