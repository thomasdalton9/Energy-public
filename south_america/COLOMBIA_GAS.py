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

The same reports open with a supply table ("Suministro por fuente",
section I. OFERTA): average daily supply by field, the national
production total, the SPEC regasification plant at Cartagena (LNG
imports) and the grand total. That goes to the "Supply by source" sheet,
with fields grouped (Cusiana/Cupiagua, Guajira, Canacol blocks, other)
and mcm/d columns at 1,000 BTU/cf. Reports from late 2023 to mid 2024
carry that table as an image: it is read with OCR (tesseract) and kept
only when its rows add up (fields = national production, production +
LNG = total, total = the 'Suministro Prom.' profile figure).

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
           "Petrochemical", "Oil_sector", "Compressors"]
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
    if w.startswith("petroler"):
        return "Oil_sector"   # "Petrolero", split out of Industrial from the Aug 2025 report on
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


# ---------------------------------------------------------------- supply

# Field groups of the "Suministro por fuente" table. The rows changed over
# the years (2021: Cusiana / Cupiagua / Guajira / Clarinete-Pandereta / Nelson;
# 2025: Chuchupa and Ballena apart, Canacol's blocks VIM 5 / VIM 21 /
# Esperanza; 2026: one "Piedemonte llanero" row), so rows are matched by
# keyword and everything else lands in Other_fields (domestic total minus
# the groups).
FIELD_GROUPS = [
    ("Piedemonte_Cusiana_Cupiagua", r"cusiana|cupiagua|pauto|flore[ñn]a|piedemonte"),
    ("Guajira_Chuchupa_Ballena", r"guajira|chuchupa|chucupa|ballena"),
    ("Canacol_VIM5_VIM21_Esperanza", r"\bvim\b|esperanza|clarinete|pandereta|nelson"),
]
SUPPLY_COLS = ([g for g, _ in FIELD_GROUPS] + ["Other_fields", "Domestic_production", "LNG_imports_SPEC",
               "Venezuela_imports", "Total_supply", "Supply_to_SNT", "Production_potential", "LNG_regas_capacity"])
# 1 GBTU = 1e9 BTU; at 1,000 BTU per cubic foot that is 1e6 cf = 0.0283 million m3.
BTU_PER_CF = 1000
MCM_PER_GBTU = 1e9 / BTU_PER_CF * 0.0283168 / 1e6
# a cell: '1.238' / '1,322' (thousands), '4,5' / '0.7' (decimal, small regas volumes in 2021), '-' (zero)
VAL = r"(?:\d{1,3}(?:[.,]\d{3})+|\d+(?:[.,]\d{1,2})?|-)"
# last column is the % of potential ('106%', '102%(1)' with a footnote, '-' when there is no potential)
PCT = r"(?:\s+(?:\d+(?:[.,]\d+)?\s*%\s*(?:\(\d\))?|-))?"
ROW = re.compile(rf"^(?P<label>.*?)\s*(?P<nums>{VAL}(?:\s+{VAL}){{2,3}}){PCT}\s*$")
OCR_DPI = 300


def whole(s):
    """GBTUD cell value: thousands separators dropped, decimal comma read as a point, '-' is zero."""
    if s == "-":
        return 0.0
    if re.fullmatch(r"\d{1,3}(?:[.,]\d{3})+", s):
        return float(re.sub(r"[.,]", "", s))
    return float(s.replace(",", "."))


def supply_rows(text):
    """(label, [potential, to_SNT, to_others, total]) rows of the supply table.
    A row whose label wraps sits on a numbers-only line between its two label
    halves (2025: 'Cupiagua, Cupiagua sur,' / numbers / 'Cusiana y Pauto sur')."""
    lines = [ln.strip() for ln in text.splitlines()]
    rows = []
    for i, ln in enumerate(lines):
        m = ROW.match(ln)
        if not m:
            continue
        vals = [whole(x) for x in m.group("nums").split()]
        label = m.group("label").strip()
        if not re.search(r"[A-Za-z]", label):
            prev = lines[i - 1] if i and not ROW.match(lines[i - 1]) else ""
            nxt = lines[i + 1] if i + 1 < len(lines) and not re.search(r"\d", lines[i + 1]) else ""
            label = f"{prev} {nxt}".strip()
        # drop letters of the rotated 'Región Interior / Costa' side label ("ire Floreña", "t Gibraltar")
        label = re.sub(r"^(?:[a-z]{1,3}\s+|[A-Z]\s+)+(?=[A-ZÁÉÍÓÚ])", "", label)
        if len(vals) == 3:  # potential, SNT, others with the total left blank
            vals.append(vals[1] + vals[2])
        rows.append((label, vals))
    return rows


def supply_table(text):
    """Supply by source (GBTUD) from the text of the 'Suministro por fuente' page, or None."""
    rows = supply_rows(text)
    is_dom = lambda lb: re.search(r"potencial\s+(?:de\s+)?producci", lb, re.I)  # noqa: E731
    k_dom = next((i for i, (lb, _) in enumerate(rows) if is_dom(lb)), None)
    lng = next((v for lb, v in rows if re.search(r"regasific", lb, re.I)), None)
    if k_dom is None or lng is None:
        return None
    dom = rows[k_dom][1]
    fields = [(lb, v) for lb, v in rows[:k_dom] if re.search(r"[A-Za-z]{3}", lb)]
    ven = next((v for lb, v in rows if re.search(r"venezuela", lb, re.I)), None)
    tot = next((v for lb, v in rows[k_dom:]
                if re.fullmatch(r"(?:[a-zA-Z]{1,3}\s+)?total\W*", lb.strip(), re.I)), None)
    res = {"Domestic_production": dom[3], "Production_potential": dom[0],
           # 2021 reports put the regas plant's delivery in 'to SNT' and leave its total blank
           "LNG_imports_SPEC": max(lng[3], lng[1] + lng[2]), "LNG_regas_capacity": lng[0],
           "Venezuela_imports": ven[3] if ven else None}
    grouped = 0.0
    for g, pat in FIELD_GROUPS:
        hit = [v[3] for lb, v in fields if re.search(pat, lb, re.I)]
        res[g] = sum(hit) if hit else None
        grouped += sum(hit)
    res["Other_fields"] = round(dom[3] - grouped, 1)
    res["Fields_sum"] = sum(v[3] for _, v in fields)
    res["Total_supply"] = tot[3] if tot else dom[3] + res["LNG_imports_SPEC"] + (res["Venezuela_imports"] or 0)
    res["Supply_to_SNT"] = tot[1] if tot else None
    res["Field_rows"] = "; ".join(f"{re.sub(r'[*]+', '', lb).strip()} {v[3]:g}" for lb, v in fields)
    return res


def profile_average(pages, month):
    """The month's average supply from the 'Perfil Contratación vs Suministro' table
    ('Suministro Prom.' row, one value per month of the year so far), for cross-checks."""
    k = int(month[5:7]) - 1 if month else None
    for text, _ in pages[:8]:
        m = re.search(r"Suministro\s+Prom\.?\s+((?:[\d.,]+\s*)+)", text)
        if m and k is not None:
            vals = m.group(1).split()
            return whole(vals[k]) if k < len(vals) else None
    return None


def supply_issues(res, profile):
    """Internal consistency of one month's table: sources add up to the total,
    field rows add up to national production, total matches the profile table."""
    issues = []
    srcs = res["Domestic_production"] + res["LNG_imports_SPEC"] + (res["Venezuela_imports"] or 0)
    if abs(srcs - res["Total_supply"]) > 2:
        issues.append(f"domestic+LNG {srcs:g} vs total {res['Total_supply']:g}")
    if abs(res["Fields_sum"] - res["Domestic_production"]) > max(6, 0.01 * res["Domestic_production"]):
        issues.append(f"field rows {res['Fields_sum']:g} vs domestic {res['Domestic_production']:g}")
    if profile is not None and abs(profile - res["Total_supply"]) > 2:
        issues.append(f"total {res['Total_supply']:g} vs profile table {profile:g}")
    return issues


def ocr_page(content, i, psm):
    """OCR page i (0-based) with tesseract (Spanish). From late 2023 to mid 2024 the
    supply table is an image with no text layer."""
    import subprocess
    import tempfile
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
    # table rules and dashes come out as '|', '—', '–'
    return re.sub(r"[|]", " ", r.stdout).replace("—", "-").replace("–", "-")


def parse_supply(pages, content=None, month=None):
    """Supply by source from the 'Suministro por fuente' table: text layer first,
    OCR of the table page when the table is an image. OCR results are kept only
    if every consistency check passes. Returns (result or None, issues)."""
    profile = profile_average(pages, month)
    for text, _ in pages[:8]:
        if re.search(r"Suministro\s+por\s+fuente", text, re.I):
            res = supply_table(text)
            if res:
                res["Method"] = "text"
                # the profile cross-check is for OCR only: from 2025 that table's layout no longer
                # lines up month by month, and text-layer rows are already checked against each other
                return res, supply_issues(res, None)
    page = next((i for i, (t, _) in enumerate(pages[:8])
                 if re.search(r"principales\s+fuentes\s+de\s+suministro", t, re.I)), None)
    if content is None or page is None:
        return None, ["no supply table found"]
    issues = []
    for psm in (6, 4, 11):
        try:
            txt = ocr_page(content, page, psm)
        except Exception as e:
            return None, [f"OCR failed: {type(e).__name__}: {str(e)[:100]}"]
        res = supply_table(txt)
        if not res:
            issues = [f"OCR psm {psm}: no table rows"]
            continue
        issues = supply_issues(res, profile)
        if not issues:
            res["Method"] = f"OCR (psm {psm})"
            return res, []
        issues = [f"OCR psm {psm}: " + "; ".join(issues) + f" | {res['Field_rows']}"]
    return None, issues


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

def load_sheet(path, sheet):
    try:
        d = pd.read_excel(path, sheet_name=sheet)
    except ValueError:  # sheet not there yet (first run with supply)
        return pd.DataFrame()
    d = d.drop(columns=[c for c in d.columns if str(c).startswith("Unnamed")])
    d["Month"] = d["Month"].astype(str).str[:7]
    return d


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", default="output/Data and Chart Outputs/colombia_gas_demand_by_sector.xlsx")
    ap.add_argument("--full", action="store_true", help="re-parse every report, ignoring the archive")
    args = ap.parse_args()

    have, have_sup = pd.DataFrame(), pd.DataFrame()
    if os.path.exists(args.out) and not args.full:
        have = load_sheet(args.out, "Demand by sector")
        have_sup = load_sheet(args.out, "Supply by source")
        out(f"archive: {len(have)} demand months, latest {have['Month'].max() if len(have) else None}; "
            f"{len(have_sup)} supply months")
    done = set(have["Month"]) if len(have) else set()
    done_sup = set(have_sup["Month"]) if len(have_sup) else set()
    recheck = set(sorted(done)[-RECHECK_MONTHS:])

    pdfs = list_pdfs()
    out(f"{len(pdfs)} monthly reports uploaded since {DATA_START}")
    # newest upload wins when a month was re-issued
    by_month = {}
    for upload, u in pdfs:
        by_month[guess_month(u, upload)] = (upload, u)

    new_rows, sup_rows, failed, failed_sup = [], [], [], []
    for gm in sorted(by_month):
        upload, u = by_month[gm]
        # a month already in both sheets is skipped unless it is one of the newest few
        if gm < DATA_START or (gm in done and gm in done_sup and gm not in recheck):
            continue
        try:
            c = requests.get(u, headers=H, timeout=T).content
            if c[:4] != b"%PDF":
                out(f"  {gm}: not a PDF ({u})")
                failed.append(gm)
                continue
            month, total, sec, src, pages = parse_pdf(c)
            month = month or gm
            sup, issues = parse_supply(pages, c, month)
        except Exception as e:
            out(f"  {gm}: ERR {type(e).__name__}: {str(e)[:150]}")
            failed.append(gm)
            continue
        if month != gm:
            out(f"  note: file name suggests {gm}, report header says {month}; using header")
        if sup:
            srow = {"Month": month, **{k: sup.get(k) for k in SUPPLY_COLS}}
            check = sup["Domestic_production"] + sup["LNG_imports_SPEC"] + (sup["Venezuela_imports"] or 0)
            srow["Check_sources_vs_total"] = round(check - sup["Total_supply"], 1)
            srow["Check_fields_vs_domestic"] = round(sup["Fields_sum"] - sup["Domestic_production"], 1)
            srow["Method"] = sup["Method"]
            srow["Field_rows"] = sup["Field_rows"]
            srow["Report"] = u
            sup_rows.append(srow)
            ven = sup["Venezuela_imports"]
            out(f"  {month}: supply {sup['Total_supply']:g} = domestic {sup['Domestic_production']:g} "
                f"+ LNG {sup['LNG_imports_SPEC']:g}" + (f" + Venezuela {ven:g}" if ven is not None else "")
                + f" GBTUD (fields sum {sup['Fields_sum']:g}) [{sup['Method']}] | {sup['Field_rows']}"
                + (f"\n    WARNING: {'; '.join(issues)}" if issues else ""))
        else:
            out(f"  {month}: could not parse the supply table ({'; '.join(issues)}) {u}")
            failed_sup.append(month)
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
    out(f"\n{len(df)} demand months {df['Month'].min()}..{df['Month'].max()}; missing: {gaps or 'none'}; "
        f"failed this run: {failed or 'none'}")

    sheets = {"Demand by sector": df}
    sup = pd.concat([have_sup, pd.DataFrame(sup_rows)], ignore_index=True) if len(have_sup) else pd.DataFrame(sup_rows)
    if len(sup):
        mcmd = ["Domestic_production", "LNG_imports_SPEC", "Venezuela_imports", "Total_supply"]
        sup = sup.drop_duplicates("Month", keep="last").sort_values("Month").reset_index(drop=True)
        for k in mcmd:
            sup[k + "_mcmd"] = (pd.to_numeric(sup[k], errors="coerce") * MCM_PER_GBTU).round(2)
        sup = sup.drop(columns=["Total_demand_SNT"], errors="ignore").merge(
            df[["Month", "Reported_total"]].rename(columns={"Reported_total": "Total_demand_SNT"}), on="Month", how="left")
        sup = sup.reindex(columns=["Month"] + SUPPLY_COLS + [k + "_mcmd" for k in mcmd]
                          + ["Total_demand_SNT", "Check_sources_vs_total", "Check_fields_vs_domestic", "Method",
                             "Field_rows", "Report"])
        sheets["Supply by source"] = sup
        sm = pd.period_range(sup["Month"].min(), sup["Month"].max(), freq="M").astype(str)
        out(f"{len(sup)} supply months {sup['Month'].min()}..{sup['Month'].max()}; "
            f"missing: {sorted(set(sm) - set(sup['Month'])) or 'none'}; failed this run: {failed_sup or 'none'}")
        last = sup.iloc[-1]
        out(f"latest supply month {last['Month']}: " + ", ".join(f"{k} {last[k]}" for k in SUPPLY_COLS))

    notes = [
        "UNITS",
        "GBTUD = billion BTU per day, the month's average daily energy delivered to end users. "
        "1 GBTUD is about 0.027 million m3/day (roughly 1 MMcf/d) at Colombian gas heating values.",
        "",
        "SECTORS",
        "Power (gas-fired thermal plants), Industrial, Residential, Commercial, Vehicle_CNG (GNVC), Refinery, "
        "Petrochemical, Oil_sector ('Petrolero' - oil-industry use, reported separately from Aug 2025; before "
        "that it was inside Industrial, so add the two for a consistent industrial series), Compressors (gas burned by pipeline compressor stations on the national grid, SNT). "
        "Each is Costa (Caribbean coast) + Interior, as published.",
        "",
        "CHECKS",
        "Sum_of_sectors is compared to Reported_total, the report's headline 'demanda promedio' figure. "
        "Source says which table was read. 'department table' means the regional table could not be read, "
        "so Power is Reported_total minus the other sectors.",
        "",
        "SUPPLY",
        "Sheet 'Supply by source': the report's 'Suministro por fuente' table (section I. OFERTA), the month's "
        "average daily gas supplied by each source in GBTUD, whole numbers as published. Each source is gas "
        "delivered into the national transport system (SNT) plus gas delivered to others (dedicated pipelines, "
        "CNG and isolated fields), so Total_supply is above the SNT demand on 'Demand by sector'. "
        "Supply_to_SNT is the SNT part only, comparable with Total_demand_SNT (= Reported_total).",
        "Domestic_production = the table's national production row ('Potencial Producción Nacional' / 'Total "
        "Potencial de Producción'), supply column. Field groups: Piedemonte_Cusiana_Cupiagua = Cusiana, "
        "Cupiagua, Cupiagua Sur, Pauto Sur and Floreña (Llanos foothills, operated by Ecopetrol); "
        "Guajira_Chuchupa_Ballena = Chuchupa and Ballena (La Guajira); Canacol_VIM5_VIM21_Esperanza = the "
        "Sinú-San Jacinto blocks VIM-5, VIM-21 and Esperanza (Clarinete, Pandereta, Nelson and the rest). "
        "Other_fields = Domestic_production minus those groups (Gibraltar, Bonga/Mamey, Bullerengue, Istanbul "
        "and the smaller interior and coast fields). The report's row list changed over time; Field_rows "
        "keeps the rows read each month.",
        "LNG_imports_SPEC = regasified LNG from the SPEC terminal at Cartagena ('Planta Regasificación "
        "Cartagena'); LNG_regas_capacity is the plant capacity the report states. Production_potential is the "
        "declared production potential (Ministerio de Minas y Energía) for the month.",
        "Venezuela_imports: the table has no Venezuela row in this period, so the column is blank (no pipeline "
        "imports reported); it fills only if the report adds such a row.",
        f"mcm/d columns (*_mcmd) = GBTUD x {MCM_PER_GBTU:.4f}: 1 GBTU = 1 million cubic feet at an assumed "
        f"heating value of {BTU_PER_CF:,} BTU per cubic foot = 0.0283 million m3. Colombian gas is often "
        "1,000-1,100 BTU/cf, so the physical volume can be up to ~10% lower.",
        "Checks: Check_sources_vs_total = Domestic + LNG (+ Venezuela) minus the table's Total (0 or +/-1 from "
        "rounding); Check_fields_vs_domestic = sum of the field rows minus Domestic_production. Method: 'text' "
        "= read from the PDF text; 'OCR' = the table is an image in that report (late 2023 to mid 2024) and was "
        "read with OCR, kept only when both checks pass and the total matches the report's 'Suministro Prom.' "
        "profile figure. Months whose table could not be read reliably are left out rather than estimated.",
        "",
        "SOURCE",
        f"Gestor del Mercado de Gas Natural (Bolsa Mercantil de Colombia), monthly reports: {PAGE}. "
        "Report holds the PDF used for each month. Figures come from SEGAS end-user delivery reports and "
        "may include estimates for distribution-network sectors. Supply: SEGAS and Ministerio de Minas y "
        "Energía, as cited in the report.",
    ]
    os.makedirs(os.path.dirname(os.path.abspath(args.out)), exist_ok=True)
    xlsx_notes.write_workbook(args.out, sheets, notes, {"UNITS", "SECTORS", "CHECKS", "SUPPLY", "SOURCE"})
    out(f"Saved {args.out}")


if __name__ == "__main__":
    main()
