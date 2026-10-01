"""
Ecuador installed generation capacity by technology from ARCONEL (Agencia
de Regulacion y Control de Electricidad, formerly ARCERNNR) - free, public,
no key. ARCONEL's site is https://arconel.gob.ec (valid TLS; the old
ARCERNNR host www.controlrecursosyenergia.gob.ec now serves a certificate
for another name and a placeholder page, so it is not used).

Source: the monthly 'Balance Nacional de Energia Electrica' (BNEE) workbook
  https://arconel.gob.ec/balance-nacional-de-energia-electrica/
  -> wp-content/uploads/downloads/<upload year>/<upload month>/BNEE_<mes>_<anio>[_revACH].xls
whose first block is 'Potencia Nominal en Generacion' (MW), national total
and S.N.I.: Hidraulica, Eolica, Fotovoltaica, Biomasa, Biogas (renewable);
MCI, Turbogas, Turbovapor (non-renewable); Importacion (interconnector
capacity, not generation - left out).
The page only links the latest month. Older months keep the same file
name pattern in the folder of the month they were uploaded; ARCONEL also
uploads a PNG of each month's balance to its WordPress media library, so
the media listing (/wp-json/wp/v2/media?search=BNEE) gives every published
month and its upload month, and the script tries the few name variants
ARCONEL uses (plain or '_revACH', .xls) in that folder and its neighbours.
Found via discovery_archive/south_america/CAPACITY_PBUE_DISCOVERY4-8.py.
BNEE workbooks are on arconel.gob.ec from April 2024; earlier months were on
the retired ARCERNNR site.

Monthly: date = 1st of the month the balance refers to; national nominal
capacity (S.N.I. + isolated systems) in MW.

Incremental: months already in the workbook are kept with their source
URL; only months not yet in it are looked up and downloaded.

Usage: python3 ECUADOR_ARCONEL_CAPACITY.py [--out PATH] [--test] [--max-months N]
"""

print("STARTING", flush=True)

import argparse
import io
import os
import re
import sys
import unicodedata
from datetime import datetime

import pandas as pd
import requests

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import power_capacity_common as pcc  # noqa: E402

SITE = "https://arconel.gob.ec"
PAGE = SITE + "/balance-nacional-de-energia-electrica/"
MEDIA = SITE + "/wp-json/wp/v2/media"
HEADERS = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) "
                         "Chrome/124.0 Safari/537.36"}
TIMEOUT = (15, 120)
DEFAULT_OUT = os.path.join("output", "Data and Chart Outputs", "ecuador_power_capacity.xlsx")
MESES = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto", "septiembre", "octubre",
         "noviembre", "diciembre"]
SUFFIXES = ["_revACH", "", "-1", "_rev", "_revACH-1", "_v2"]

# BNEE technology row -> standard fuel
TECH = {"HIDRAULICA": "Hydro", "EOLICA": "Wind", "FOTOVOLTAICA": "Solar", "SOLAR": "Solar", "BIOMASA": "Bioenergy",
        "BIOGAS": "Bioenergy", "MCI": "Oil", "TURBOVAPOR": "Oil", "TURBOGAS": "Gas", "GEOTERMICA": "Other"}
SKIP = {"IMPORTACION", "COLOMBIA", "PERU", "RENOVABLE", "NO RENOVABLE", "NACIONAL (RENOVABLE + NO RENOVABLE)"}

S = requests.Session()
S.headers.update(HEADERS)


def key(s):
    s = unicodedata.normalize("NFKD", str(s)).encode("ascii", "ignore").decode().upper()
    return " ".join(s.split())


def month_index():
    """{Timestamp(month): [upload 'YYYY/MM', ...]} from the PNG of each month in the media library, plus the
    workbook the page links now."""
    months = {}
    for page in range(1, 6):
        r = S.get(MEDIA, params={"search": "BNEE", "per_page": 100, "page": page, "_fields": "date,source_url"},
                  timeout=TIMEOUT)
        if r.status_code != 200:
            break
        items = r.json()
        for m in items:
            name = os.path.basename(m["source_url"]).lower()
            hit = re.match(r"bnee_([a-z]+)_(\d{4})", name)
            if not hit or hit.group(1) not in MESES:
                continue
            month = pd.Timestamp(year=int(hit.group(2)), month=MESES.index(hit.group(1)) + 1, day=1)
            up = datetime.fromisoformat(m["date"][:19])
            months.setdefault(month, [])
            if f"{up:%Y/%m}" not in months[month]:
                months[month].append(f"{up:%Y/%m}")
        if len(items) < 100:
            break
    latest = {}
    r = S.get(PAGE, timeout=TIMEOUT)
    r.raise_for_status()
    for href in re.findall(r'href=["\']([^"\']+BNEE_[^"\']+\.xlsx?)["\']', r.text, re.I):
        hit = re.search(r"BNEE_([a-z]+)_(\d{4})", href, re.I)
        if hit and hit.group(1).lower() in MESES:
            latest[pd.Timestamp(year=int(hit.group(2)), month=MESES.index(hit.group(1).lower()) + 1, day=1)] = href
    return months, latest


def neighbours(ym):
    y, m = map(int, ym.split("/"))
    out = []
    for d in (0, 1, -1, 2):
        mm, yy = m + d, y
        while mm > 12:
            mm, yy = mm - 12, yy + 1
        while mm < 1:
            mm, yy = mm + 12, yy - 1
        out.append(f"{yy}/{mm:02d}")
    return out


def find_url(month, uploads):
    mes = MESES[month.month - 1]
    folders = []
    for up in uploads:
        folders += [f for f in neighbours(up) if f not in folders]
    for folder in folders:
        for suffix in SUFFIXES:
            for name in (f"BNEE_{mes}_{month.year}", f"BNEE_{mes.capitalize()}_{month.year}"):
                for ext in (".xls", ".xlsx"):
                    u = f"{SITE}/wp-content/uploads/downloads/{folder}/{name}{suffix}{ext}"
                    try:
                        if S.head(u, timeout=(10, 30), allow_redirects=True).status_code == 200:
                            return u
                    except requests.RequestException:
                        pass
    return None


def parse_bnee(content):
    """{fuel: MW (national total)} + {technology label: MW} from a BNEE workbook."""
    raw = pd.read_excel(io.BytesIO(content), header=None)
    # Locate the capacity block: the cell 'Potencia Nominal en Generacion ...' and, below it, the 'MW' unit row.
    pos = [(i, j) for i in range(min(15, len(raw))) for j in range(raw.shape[1])
           if key(raw.iat[i, j]).startswith("POTENCIA NOMINAL")]
    if not pos:
        raise ValueError("no 'Potencia Nominal' block")
    i0, j0 = pos[0]
    unit_row = next(i for i in range(i0, i0 + 6) if key(raw.iat[i, j0]) == "MW")
    tech, fuels = {}, {}
    for i in range(unit_row + 1, len(raw)):
        labels = [key(raw.iat[i, j]) for j in range(j0) if isinstance(raw.iat[i, j], str)]
        if not labels:
            if i > unit_row + 3 and pd.isna(raw.iat[i, j0]):
                break
            continue
        lab = labels[-1]
        if lab.startswith("NOTA"):
            break
        val = pd.to_numeric(raw.iat[i, j0], errors="coerce")
        if lab in SKIP or pd.isna(val):
            continue
        fuel = TECH.get(lab)
        if fuel is None:
            print(f"  unmapped BNEE row {lab!r} ({val} MW) -> Other", flush=True)
            fuel = "Other"
        tech[lab] = float(val)
        fuels[fuel] = fuels.get(fuel, 0.0) + float(val)
    national = [pd.to_numeric(raw.iat[i, j0], errors="coerce") for i in range(unit_row + 1, len(raw))
                if any(key(raw.iat[i, j]).startswith("NACIONAL") for j in range(j0) if isinstance(raw.iat[i, j], str))]
    return fuels, tech, (float(national[0]) if national else None)


ANNUAL_PAGE = SITE + "/publicaciones-estadistica-del-sector-electrico-2/"
NUM = r"(\d{1,3}(?:\.\d{3})*,\d+)"
PDF_ROWS = {"HIDRAULICA": "HIDRAULICA", "EOLICA": "EOLICA", "FOTOVOLTAICA": "FOTOVOLTAICA", "BIOMAS": "BIOMASA",
            "BIOMASA": "BIOMASA", "BIOGAS": "BIOGAS", "MCI": "MCI", "TURBOGAS": "TURBOGAS", "TURBOVAPOR": "TURBOVAPOR",
            "RENOVABLE": "RENOVABLE", "NO RENOVABLE": "NO RENOVABLE"}


def annual_pdf_urls():
    """{year: download URL} of ARCONEL's 'Estadistica Anual y Multianual del Sector Electrico' PDFs."""
    r = S.get(ANNUAL_PAGE, timeout=TIMEOUT)
    r.raise_for_status()
    out = {}
    for m in re.finditer(r'href=["\']([^"\']*download\.php\?id=\d+)&(?:amp;)?force=0["\']', r.text):
        before = re.sub(r"<[^>]+>|\s+", " ", r.text[max(0, m.start() - 600):m.start()])
        hits = re.findall(r"Estad\S*stica Anual y Multianual[^+]*?(\d{4})", before, re.I)
        if hits:
            out.setdefault(int(hits[-1]), m.group(1) + "&force=0")
    return out


def parse_annual_pdf(content):
    """December nominal capacity by technology from the 'Balance nacional de energia electrica' table of an
    ARCONEL annual statistics PDF: ({fuel: MW}, {technology: MW})."""
    import pdfplumber
    with pdfplumber.open(io.BytesIO(content)) as pdf:
        for page in pdf.pages[:60]:
            t = page.extract_text() or ""
            k = key(t)
            if not ("POTENCIA NOMINAL" in k and "TURBOVAPOR" in k and "BALANCE NACIONAL DE ENERGIA ELECTRICA" in k):
                continue
            vals = {}
            for line in t.splitlines():
                m = re.match(r"^\s*([A-Za-zÁÉÍÓÚáéíóúñÑ ]+?)\s+" + NUM, line)
                if not m:
                    continue
                lab = PDF_ROWS.get(key(m.group(1)))
                if lab and lab not in vals:
                    vals[lab] = float(m.group(2).replace(".", "").replace(",", "."))
            if "MCI" not in vals or "HIDRAULICA" not in vals:
                continue
            ren = sum(vals.get(x, 0) for x in ("HIDRAULICA", "EOLICA", "FOTOVOLTAICA", "BIOMASA", "BIOGAS"))
            non = sum(vals.get(x, 0) for x in ("MCI", "TURBOGAS", "TURBOVAPOR"))
            for lab, tot in (("RENOVABLE", ren), ("NO RENOVABLE", non)):
                if lab in vals and abs(vals[lab] - tot) > 1:
                    raise ValueError(f"{lab}: rows add to {tot:,.2f}, table total {vals[lab]:,.2f}")
            tech = {x: v for x, v in vals.items() if x not in ("RENOVABLE", "NO RENOVABLE")}
            fuels = {}
            for x, v in tech.items():
                fuels[TECH[x]] = fuels.get(TECH[x], 0.0) + v
            return fuels, tech
    raise ValueError("no 'Balance nacional de energia electrica' capacity table found")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", default=DEFAULT_OUT)
    ap.add_argument("--test", action="store_true")
    ap.add_argument("--max-months", type=int, default=0, help="only fetch this many missing months (test runs)")
    args = ap.parse_args()

    old = pcc.load_monthly(args.out)
    old_src = pcc.load_sheet(args.out, "BNEE by technology")
    months, latest = month_index()
    for m in latest:
        months.setdefault(m, [])
    # Months with no PNG in the media library: try the folders of the following months (ARCONEL uploads each
    # balance two to three months after the month it covers).
    gap_months = set()
    for m in pd.date_range(min(months), max(months), freq="MS"):
        if m not in months:
            months[m] = [f"{m + pd.DateOffset(months=k):%Y/%m}" for k in (2, 3)]
            gap_months.add(m)
    wanted = sorted(m for m in months if m >= pcc.START)
    have = set(old.index) if not old.empty else set()
    todo = [m for m in wanted if m not in have]
    if args.max_months:
        todo = todo[-args.max_months:]
    print(f"BNEE months published: {len(wanted)} ({wanted[0]:%Y-%m} .. {wanted[-1]:%Y-%m}); archive {len(have)}; "
          f"fetching {[f'{m:%Y-%m}' for m in todo]}", flush=True)

    rows, techs, failed = {}, {}, []
    for m in todo:
        url = latest.get(m) or find_url(m, months[m])
        if url is None:
            print(f"  {m:%Y-%m}: workbook not found (upload folders {months[m]})", flush=True)
            if m not in gap_months:
                failed.append(m)
            continue
        try:
            r = S.get(url, timeout=TIMEOUT)
            r.raise_for_status()
            fuels, tech, national = parse_bnee(r.content)
        except Exception as e:  # noqa: BLE001
            print(f"  {m:%Y-%m}: {url} failed: {type(e).__name__}: {str(e)[:150]}", flush=True)
            failed.append(m)
            continue
        total = sum(tech.values())
        print(f"  {m:%Y-%m}: {url} -> {', '.join(f'{k} {v:,.1f}' for k, v in tech.items())}; sum {total:,.1f} "
              f"vs national {national:,.1f} MW", flush=True)
        if national is not None and abs(total - national) > 1:
            raise SystemExit(f"{m:%Y-%m}: technology rows do not add up to the national total - layout changed?")
        rows[m] = {f"{f}_MW": v for f, v in fuels.items()}
        techs[m] = dict(tech, source_url=url)

    # Before the BNEE workbooks on arconel.gob.ec start: December values from ARCONEL's annual statistics PDFs.
    bnee_months = [m for m in list(rows) + list(have) if m >= pd.Timestamp("2024-01-01")]
    first_bnee = min(bnee_months) if bnee_months else min(months)
    annual_years = [y for y in range(pcc.START.year, first_bnee.year)
                    if pd.Timestamp(year=y, month=12, day=1) not in have]
    if annual_years:
        pdfs = annual_pdf_urls()
        print(f"Annual statistics PDFs: {sorted(pdfs)}; fetching {annual_years}", flush=True)
        for y in annual_years:
            if y not in pdfs:
                print(f"  {y}: no annual statistics PDF listed", flush=True)
                continue
            try:
                r = S.get(pdfs[y], timeout=(15, 600))
                r.raise_for_status()
                fuels, tech = parse_annual_pdf(r.content)
            except Exception as e:  # noqa: BLE001
                print(f"  {y}: annual PDF failed: {type(e).__name__}: {str(e)[:150]}", flush=True)
                failed.append(pd.Timestamp(year=y, month=12, day=1))
                continue
            m = pd.Timestamp(year=y, month=12, day=1)
            print(f"  {y}-12 (annual PDF {r.url}): {', '.join(f'{k} {v:,.1f}' for k, v in tech.items())}; "
                  f"sum {sum(tech.values()):,.1f} MW", flush=True)
            rows[m] = {f"{f}_MW": v for f, v in fuels.items()}
            techs[m] = dict(tech, source_url=r.url)

    new = pd.DataFrame.from_dict(rows, orient="index")
    monthly = pcc.standardise(new) if not new.empty else pd.DataFrame()
    if not old.empty:
        old = pcc.standardise(old)
        monthly = pd.concat([old[~old.index.isin(monthly.index)], monthly]).sort_index() if not monthly.empty else old
    if monthly.empty:
        sys.exit("No data fetched and no archive - nothing to write")
    tech_df = pd.DataFrame.from_dict(techs, orient="index")
    if not old_src.empty:
        old_src.index = pd.to_datetime(old_src.index)
        tech_df = pd.concat([old_src[~old_src.index.isin(tech_df.index)], tech_df]).sort_index()
    tech_df.index.name = "date"
    print(monthly.to_string(), flush=True)

    val = pcc.validation_lines(monthly, "Ecuador")
    print("\n".join(val), flush=True)
    monthly_part = [m for m in monthly.index if m >= pd.Timestamp("2024-01-01")]
    gaps = ([m for m in pd.date_range(min(monthly_part), monthly.index.max(), freq="MS") if m not in monthly.index]
            if monthly_part else [])
    notes = pcc.unit_notes("month") + [
        "Nominal (nameplate) capacity, as ARCONEL's BNEE reports it ('Potencia Nominal en Generacion'), national "
        "total (S.N.I. plus isolated systems and self-generators).",
        "",
        "COVERAGE",
        f"{monthly.index.min():%Y-%m} to {monthly.index.max():%Y-%m} ({len(monthly)} rows). Monthly BNEE workbooks "
        "are on arconel.gob.ec from April 2024; earlier months were published on the retired ARCERNNR site "
        "(www.controlrecursosyenergia.gob.ec, which no longer serves them). For 2021-2023 the series has one row "
        "per year, dated 1 December: the December 'Balance nacional de energia electrica' table reprinted in "
        "ARCONEL's 'Estadistica Anual y Multianual del Sector Electrico Ecuatoriano' PDF for that year "
        f"({ANNUAL_PAGE}). ARCONEL publishes each month about two to three months later."
        + (f" Months missing inside the range: {', '.join(f'{m:%Y-%m}' for m in gaps)}." if gaps else ""),
        "",
        "SOURCE",
        f"ARCONEL - Balance Nacional de Energia Electrica (BNEE), monthly workbook: {PAGE} (latest month linked; "
        f"older months found through the media listing {MEDIA}?search=BNEE and the same file names). "
        "Per-month workbook URLs are in the 'BNEE by technology' tab.",
        "Script: south_america/ECUADOR_ARCONEL_CAPACITY.py (scheduled by .github/workflows/ecuador_power_capacity.yml).",
        "",
        "MAPPING",
        "Hydro_MW = Hidraulica; Wind_MW = Eolica; Solar_MW = Fotovoltaica; Bioenergy_MW = Biomasa + Biogas.",
        "Gas_MW = Turbogas (gas turbines). BNEE gives technology, not fuel: Ecuador's gas turbines include the "
        "natural-gas Termogas Machala plant and several diesel-fired units, so this is a technology proxy.",
        "Oil_MW = MCI (internal-combustion engines, fuel oil / diesel, some on oil-field associated gas) + Turbovapor "
        "(fuel-oil steam plants).",
        "Other_MW = any other BNEE technology row (none now). Importacion (Colombia and Peru interconnector "
        "capacity) is not generation and is left out.",
        "",
        "VALIDATION",
    ] + val
    pcc.write(args.out, monthly, notes, {"BNEE by technology": tech_df})
    print(f"Saved {args.out}; failed months: {[f'{m:%Y-%m}' for m in failed]}", flush=True)
    if failed and len(failed) > max(3, len(todo) // 3):
        sys.exit(f"{len(failed)} of {len(todo)} months failed")


if __name__ == "__main__":
    main()
