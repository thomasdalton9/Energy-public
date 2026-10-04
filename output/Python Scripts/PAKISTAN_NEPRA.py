"""
Pakistan electricity generation by fuel, monthly, from the Fuel Charges Adjustment (FCA) filings of CPPA-G (Central
Power Purchasing Agency Guarantee Ltd) and NEPRA (National Electric Power Regulatory Authority). Replaces Ember for
Pakistan. Found via discovery_archive/asia/SCADA_DISCOVERY2-4.py and PAKISTAN_DISCOVERY1-5.py.

Sources, best first (each month comes from the best one that has it):
  1. CPPA-G "XWDISCOs Energy Purchase Data" workbooks (Excel; one per fiscal year or quarter, a sheet per month, from
     July 2024): https://cppa.gov.pk/downloads/xwdiscos-energy-purchase-data -> one page per fiscal year
     (/downloads/xwdiscos-energy-purchase-data/<id>) linking /storage/uploads/downloads/<key>.xlsx / .pdf.
     Each month's "Summary" block gives energy purchased (kWh) by fuel: Hydel, Coal-Local, Coal-Imported, HSD,
     RFO (F.O.), Gas, RLNG, Nuclear, Import (from Iran), Wind, Solar, Bagasse, Mixed, and the month's total.
  2. The same data as monthly PDFs on that page (the months not yet in a workbook; text layer from 2026).
  3. NEPRA's monthly FCA decision for the Ex-WAPDA DISCOs (listed on https://nepra.org.pk/tariff/Distribution%20FESCO.php,
     files under https://nepra.org.pk/tariff/Tariff/Ex-WAPDA%20DISCOS/<year>/TRF-100 ...pdf). Its "Annex-II Source
     Wise Generation" table gives reference and actual GWh by source (CPPA-G's figures). Used for 2021 - June 2024
     (the scanned CPPA-G filings of those years have no usable text). Coal is split local/imported from mid-2023 only.

  4. K-Electric (Karachi), real figures only: NEPRA's State of Industry Reports
     (https://nepra.org.pk/publications/State%20of%20Industry%20Reports.php; statistical tables 'K-Electric (Fuel-wise
     Own Generation and Fuel Costs)' and '(Fuel-wise Power Purchase(s) and Fuel Costs)', monthly, one report per fiscal
     year, July 2020 - June 2025 so far); for months no report covers yet, NEPRA's monthly KE FCA decisions
     (https://nepra.org.pk/tariff/Distribution%20K-Electric.php, TRF-362: own sent-out and external purchases) with the
     fuel split / IPP purchases from KE's own filings (NEPRA Admission Notices) where they add up. KE's consumers moved
     to the uniform national FCA in 2025, so KE's monthly decisions stop in spring 2025. Found via
     PAKISTAN_DISCOVERY9-14.py.

What it covers: Pakistan including K-Electric. The national grid = energy bought by CPPA-G (the Ex-WAPDA DISCOs' pool,
which also supplies part of KE's load), ~120-135 TWh a year; plus KE's own plants and KE's purchases from plants
outside the CPPA-G pool (KE's purchases from CPPA-G are NOT added again). Months without a real KE figure have KE_included = False and
national-grid-only totals (see KE_basis). NOT included: captive / off-grid generation and rooftop solar. Imports from Iran (Tavanir) are kept
apart as Imports_MWh and are not in Total_MWh.

Writes output/Data and Chart Outputs/pakistan_power_generation_daily.xlsx:
  Daily   one row per month, dated the 1st (MWh in the month), PAKISTAN INCL. K-ELECTRIC: Hydro_MWh, Coal_MWh,
          Gas_MWh, Oil_MWh, Nuclear_MWh, Wind_MWh, Solar_MWh, Bioenergy_MWh (bagasse), Other_MWh, Total_MWh (= their sum)
          = national grid + KE own + KE non-CPPA purchases; Imports_MWh (Iran); Total_grid_MWh and <fuel>_grid_MWh
          (national grid alone); grid detail Coal_local / Coal_imported / Gas_local / RLNG / RFO / HSD _MWh; KE detail
          KE_own_<Gas|RLNG|Oil|unsplit>_MWh, KE_own_MWh, KE_IPP_<Oil|Gas|Coal|Solar|Other>_MWh,
          KE_purchases_nonCPPA_MWh, KE_from_CPPA_MWh (check only, not added), KE_basis (source), KE_included
  Grid    the national-grid history store (as filed);  KE  the K-Electric history store (GWh, as filed)
  Months  per month: the source used, its file, and the check of the fuel rows against the filing's own total
  Files   every file read (so a file is not downloaded again), with the months it gave

And, from NEPRA's one-off hourly plant-wise file (CPPA-G's FPA application, NEPRA Admission Notices Sep 2026:
'01- Plant wise Units.xlsx', 129 plants, Dec 2025 - May 2026) plus '03- Hourly Marginal Price.xlsx', when such a
file is linked from NEPRA's home page / news page and not read yet:
output/Data and Chart Outputs/pakistan_power_hourly.xlsx
  Hourly  MW (= MWh in the hour) by fuel, CPPA-G system;  Daily  standard layout, MWh per day;  Prices  Rs/kWh

Incremental: the workbook is the history store. Each run reads the two listing pages (CPPA-G, NEPRA) and downloads
only files that can give a month not saved yet (or saved from a lower-ranked source); every file read is recorded on
the Files sheet. Runs on the 1st and 15th.

    python3 asia/PAKISTAN_NEPRA.py [--out PATH] [--hourly-out PATH]
"""
import argparse
import html
import io
import os
import re
import sys
import tempfile
import time
from concurrent.futures import ThreadPoolExecutor
from urllib.parse import unquote, urljoin

import pandas as pd
import pdfplumber
import requests
import urllib3

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)
import xlsx_notes  # noqa: E402

urllib3.disable_warnings()
H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 "
                   "Safari/537.36"}
T = (20, 180)
DATA_START = pd.Timestamp("2021-01-01")
CPPA_LIST = "https://cppa.gov.pk/downloads/xwdiscos-energy-purchase-data"
NEPRA_LIST = "https://nepra.org.pk/tariff/Distribution%20FESCO.php"
NEPRA_KE_LIST = "https://nepra.org.pk/tariff/Distribution%20K-Electric.php"
SOIR_LIST = "https://nepra.org.pk/publications/State%20of%20Industry%20Reports.php"   # every XWDISCO page lists the same FCA decisions
NEPRA_HOURLY_PAGES = ["https://nepra.org.pk/", "https://nepra.org.pk/news.php"]
KNOWN_HOURLY = ["https://nepra.org.pk/Admission%20Notices/2026/09%20Sep/01-%20Plant%20wise%20Units.xlsx"]
OUT_DIR = os.path.join(ROOT, "output", "Data and Chart Outputs")
OUT = os.path.join(OUT_DIR, "pakistan_power_generation_daily.xlsx")
OUT_HOURLY = os.path.join(OUT_DIR, "pakistan_power_hourly.xlsx")
RANK = {"CPPA-G workbook": 3, "CPPA-G PDF": 2, "NEPRA FCA decision": 1}
MON = {m: i for i, m in enumerate(["jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov",
                                   "dec"], start=1)}
FULL_MONTHS = ["january", "february", "march", "april", "may", "june", "july", "august", "september", "october",
               "november", "december"]
MONTH_YEAR = re.compile(r"\b(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)[a-z]*\.?[\s\-_',]*(20\d\d|\d\d)(?!\d)",
                        re.I)
DETAIL = ["Hydro", "Coal", "Coal_local", "Coal_imported", "Gas_local", "RLNG", "RFO", "HSD", "Nuclear", "Wind",
          "Solar", "Bagasse", "Mixed", "Imports"]
COLS = ["Hydro_MWh", "Coal_MWh", "Gas_MWh", "Oil_MWh", "Nuclear_MWh", "Wind_MWh", "Solar_MWh", "Bioenergy_MWh",
        "Other_MWh", "Total_MWh", "Imports_MWh", "Coal_local_MWh", "Coal_imported_MWh", "Gas_local_MWh", "RLNG_MWh",
        "RFO_MWh", "HSD_MWh"]


def out(*a):
    print(*a, flush=True)


SESSION = requests.Session()
SESSION.headers.update(H)


def get(url, tries=4):
    for i in range(tries):
        try:
            r = SESSION.get(url, timeout=T, verify=False)
            r.raise_for_status()
            return r
        except requests.RequestException as e:
            if i == tries - 1:
                raise
            out(f"  retry {unquote(url)[-90:]}: {e}")
            time.sleep(5 * (i + 1))


# ---------------------------------------------------------------------------------------------------- helpers
def month_tokens(text):
    """'1st QTR (JUL 25 to SEP 25)' -> [2025-07-01, 2025-09-01]."""
    res = []
    for m in MONTH_YEAR.finditer(text or ""):
        word = re.match(r"[A-Za-z]+", m.group(0)).group(0).lower()
        if not FULL_MONTHS[MON[word[:3]] - 1].startswith(word) and word not in ("sept", "sepetember"):
            continue   # 'separately 32', 'decision 20' ...
        y = int(m.group(2))
        y = y + 2000 if y < 100 else y
        if 2000 <= y <= 2100:
            res.append(pd.Timestamp(y, MON[m.group(1).lower()[:3]], 1))
    return res


def month_range(text):
    t = month_tokens(text)
    if not t:
        return []
    if len(t) >= 2 and re.search(r"\bto\b|-\s*\w+\s*\d", text, re.I) and t[1] > t[0]:
        return list(pd.date_range(t[0], t[1], freq="MS"))
    return [t[0]]


def fuel_key(label):
    """A filing's row label -> detail fuel key (or 'TOTAL' / None). NEPRA's decision PDFs carry an OCR text layer,
    so the usual misreadings are accepted ('Hyde!', '1-lydel', 'FISD', 'Baggasse', 'RING')."""
    s = re.sub(r"[^a-z ]", " ", str(label).lower())
    words = s.split()
    if not words:
        return None
    w0, s = words[0], "".join(words)
    if w0.startswith("total") or s.startswith("grandtotal") or s.startswith("gtotal"):
        return "TOTAL"
    if s.startswith("hyd") or "ydel" in s[:6]:
        return "Hydro"
    if w0.startswith("coal") or w0.startswith("cool"):
        if "local" in s:
            return "Coal_local"
        if "import" in s:
            return "Coal_imported"
        return "Coal"
    if w0.startswith("gas"):
        return "Gas_local"
    if "rlng" in s or w0 == "ring":
        return "RLNG"
    if re.search(r"bag+as+e", s):
        return "Bagasse"
    if w0 in ("hsd", "fisd", "isd", "lisd", "diesel"):
        return "HSD"
    if w0 in ("rfo", "fo", "furnace", "efo", "rio") or s.startswith("fo"):
        return "RFO"
    if w0.startswith("nuclear"):
        return "Nuclear"
    if "iran" in s or w0.startswith("import"):
        return "Imports"
    if w0.startswith("wind"):
        return "Wind"
    if w0.startswith("solar") or w0 == "soar":
        return "Solar"
    if w0.startswith("mixed") or w0.startswith("misc"):
        return "Mixed"
    return None


def add(d, k, v):
    d[k] = d.get(k, 0.0) + (v or 0.0)


def check(rows, total):
    s = sum(v for k, v in rows.items() if k != "TOTAL")
    return (s / total - 1.0) if total else None


# ---------------------------------------------------------------------------------------------------- CPPA-G
def cppa_files():
    """[(url, link text, kind, months)] from every fiscal-year page."""
    r = get(CPPA_LIST)
    subs = sorted(set(re.findall(r'href="(https://cppa\.gov\.pk/downloads/xwdiscos-energy-purchase-data/\d+)"', r.text)))
    files = []
    for sp in subs:
        t = get(sp).text
        t = t[t.find('<div class="col-md-9">'):]
        for h, txt in re.findall(r'<a[^>]+href="([^"]+\.(?:xlsx?|pdf))"[^>]*>(.*?)</a>', t, re.S | re.I):
            txt = html.unescape(re.sub(r"<[^>]+>|\s+", " ", txt)).strip()
            kind = "CPPA-G workbook" if re.search(r"\.xlsx?$", h, re.I) else "CPPA-G PDF"
            files.append((urljoin(sp, h), txt, kind, month_range(re.sub(r"\.\w+$", "", txt))))
    out(f"CPPA-G: {len(subs)} fiscal-year pages, {len(files)} files "
        f"({sum(f[2] == 'CPPA-G workbook' for f in files)} workbooks)")
    return files


def parse_cppa_sheet(df):
    """One month's sheet -> ({fuel: kWh}, total kWh) from its Summary block."""
    st = df.astype(str)
    hit = st.apply(lambda c: c.str.strip().str.fullmatch(r"summary", case=False)).any(axis=1)
    if not hit.any():
        return None, None
    i0 = hit[hit].index[-1]
    ecol = None
    for i in range(i0, -1, -1):   # nearest header above the Summary with an 'Energy ... kWh' cell
        for j, v in enumerate(df.iloc[i].tolist()):
            if isinstance(v, str) and re.search(r"energy", v, re.I) and re.search(r"kwh|units", v, re.I):
                ecol = j
                break
        if ecol is not None:
            break
    if ecol is None:
        return None, None
    rows, total = {}, None
    for i in range(i0 + 1, min(i0 + 40, len(df))):
        row = df.iloc[i].tolist()
        key = None
        for j, v in enumerate(row[:ecol]):
            if isinstance(v, str) and re.search(r"[A-Za-z]", v):
                key = fuel_key(v)
                if key:
                    break
        if not key:
            continue
        v = pd.to_numeric(row[ecol], errors="coerce")
        v = 0.0 if pd.isna(v) else float(v)
        if key == "TOTAL":
            total = v
            break
        add(rows, key, v)
    return rows, total


def sheet_month(name, df):
    t = month_tokens(name)
    if t:
        return t[0]
    head = " ".join(str(v) for v in df.head(8).values.ravel() if isinstance(v, str))
    m = re.search(r"month of\s+([A-Za-z]+\s*,?\s*\d{4})", head, re.I)
    t = month_tokens(m.group(1)) if m else []
    return t[0] if t else None


def read_cppa_workbook(url):
    x = pd.ExcelFile(io.BytesIO(get(url).content))
    res = {}
    for sh in x.sheet_names:
        df = x.parse(sh, header=None)
        mo = sheet_month(sh, df)
        rows, total = parse_cppa_sheet(df)
        if mo is None or not rows:
            out(f"  {sh}: no month / Summary block")
            continue
        res[mo] = (rows, total, 1e-3)   # kWh -> MWh
    return res


NUM = re.compile(r"^\s*(-|\(?\d(?:\s?\d){0,2}(?:\s?,\s?\d{3})*)")


def read_cppa_pdf(url, months=None):
    """Monthly PDF with a text layer: the 'Summary' block after the plant tables."""
    with pdfplumber.open(io.BytesIO(get(url).content)) as p:
        text = "\n".join(pg.extract_text() or "" for pg in p.pages)
    k = text.rfind("\nSummary")
    if k < 0:
        return {}
    mo = None
    m = re.search(r"FOR THE MONTH OF\s+([A-Za-z]+\s*,?\s*\d{4})", text, re.I)
    if m:
        t = month_tokens(m.group(1))
        mo = t[0] if t else None
    mo = mo or (months[0] if months and len(months) == 1 else None)
    if mo is None:
        return {}
    rows, total = {}, None
    for line in text[k:].splitlines()[1:40]:
        lm = re.match(r"\s*([A-Za-z][A-Za-z .\-/]*?)\s+(?=[\d(\-])", line)
        if not lm:
            continue
        key = fuel_key(lm.group(1))
        if not key:
            continue
        nm = NUM.match(line[lm.end():])
        if not nm:
            continue
        v = 0.0 if nm.group(1) == "-" else float(re.sub(r"[^\d]", "", nm.group(1)))
        if key == "TOTAL":
            total = v
            break
        add(rows, key, v)
    return {mo: (rows, total, 1e-3)} if rows else {}


# ---------------------------------------------------------------------------------------------------- NEPRA
def nepra_decisions():
    t = get(NEPRA_LIST).text
    res = []
    for h in sorted(set(re.findall(r'href\s*=\s*["\']([^"\']+\.pdf)["\']', t, re.I))):
        name = unquote(h)
        if "ex-wapda discos" not in name.lower():
            continue
        fn = name.rsplit("/", 1)[-1]
        if not re.search(r"FCA|FPA|MFPA|Fuel", fn, re.I) or \
                re.search(r"QUARTER|QTR|QTA|HIGH COURT|REQUEST|REBASING|ANNUAL|INDEXATION|CORRIGENDUM", fn, re.I):
            continue
        fy = re.search(r"/(20\d\d)/", name)
        guess = month_tokens(re.sub(r"\d{1,2}-\d{1,2}-\d{4}", " ", fn))
        guess = [g for g in guess if not fy or abs(g.year - int(fy.group(1))) <= 1]
        if fy and int(fy.group(1)) < DATA_START.year:
            continue
        if not fy and (not guess or guess[0] < DATA_START):
            continue
        res.append((urljoin(NEPRA_LIST, h), fn, "NEPRA FCA decision", guess[:1]))
    out(f"NEPRA: {len(res)} FCA decisions from {DATA_START.year}")
    return res


PAIR = re.compile(r"(-|\d[\d,.]*\d|\d)\s+(\d{1,3}[.,]\d{1,2})\s*%")


def num(s):
    if s == "-":
        return 0.0
    s = s.replace(",", "") if s.count(".") <= 1 else s.replace(",", "").replace(".", "", s.count(".") - 1)
    return float(s)


EXPECTED = ["Hydro", "Coal", "HSD", "RFO", "Gas_local", "RLNG", "Nuclear", "Imports", "Mixed", "Wind", "Bagasse",
            "Solar"]


def decision_month(text):
    """Most frequent 'for the month of <Month> <Year>' in a decision (its pages repeat it as a header)."""
    found = []
    for m in re.finditer(r"month\s*of\s*([A-Za-z]{3,9})\s*,?\s*(20\d\d)", text, re.I):
        found += month_tokens(f"{m.group(1)} {m.group(2)}")
    return max(set(found), key=found.count) if found else None


def parse_annex(block):
    """Lines of one 'Source Wise Generation' table -> ({fuel: actual GWh}, total GWh, {fuel: actual %})."""
    rows, pct, total, unlabelled, unparsed = {}, {}, None, [], []
    for line in block[1:32]:
        if re.search(r"sale\s*to|transmission|net\s*deliver|fuel\s*cost", line, re.I):
            break
        pairs = PAIR.findall(line)
        pm = PAIR.search(line)
        label = line[:pm.start()] if pm else re.match(r"[^\d]*", line).group(0)
        key = fuel_key(label)
        if key is None:
            if len(pairs) >= 2 and not re.search(r"[A-Za-z]{2}", label):
                unlabelled.append(pairs[-1])
            continue
        val = None
        if len(pairs) >= 2:
            try:
                val, pc = num(pairs[-1][0]), float(pairs[-1][1].replace(",", "."))
            except ValueError:
                val = None
        if key == "TOTAL":
            total = val
            break
        if val is None:
            unparsed.append(key)
            continue
        add(rows, key, val)
        pct[key] = pct.get(key, 0.0) + pc
    have = {("Coal" if k.startswith("Coal") else k) for k in rows}
    missing = [k for k in EXPECTED if k not in have and k not in unparsed]
    if len(unlabelled) == 1 and len(missing) == 1:   # a row whose label the OCR lost
        rows[missing[0]], pct[missing[0]] = num(unlabelled[0][0]), float(unlabelled[0][1].replace(",", "."))
    if total is None and not unparsed and abs(sum(pct.values()) - 100) < 0.3:
        total = sum(rows.values())
    if total and len(unparsed) == 1:   # one row's digits unreadable: the remainder of the stated total
        rows[unparsed[0]] = max(total - sum(rows.values()), 0.0)
        pct[unparsed[0]] = 100.0 * rows[unparsed[0]] / total
    return rows, total, pct


def read_decision(url, months=None):
    """'Annex-II Source Wise Generation <Month Year> / Sources Reference Actual / GWh % GWh %' -> actual GWh."""
    with pdfplumber.open(io.BytesIO(get(url).content)) as p:
        text = "\n".join(pg.extract_text() or "" for pg in p.pages)
    hits = list(re.finditer(r"source\s*-?\s*wise\s*generation", text, re.I))
    res = {}
    for m in hits:
        block = text[m.end():m.end() + 3000].splitlines()
        mo = None
        for line in block[:4]:
            t = month_tokens(line)
            if t:
                mo = t[0]
                break
        if mo is None and len(hits) == 1:   # month line unreadable ('TUNE 2021'): the decision's own month
            mo = decision_month(text) or (months[0] if months else None)
        if mo is None or mo in res:
            continue
        rows, total, pct = parse_annex(block)
        if not rows or not total:
            continue
        c = check(rows, total)
        if c is not None and abs(c) > 0.01:   # OCR'd digits: fall back to the share column where a row disagrees
            for k in rows:
                est = pct[k] / 100.0 * total
                if abs(rows[k] - est) > max(0.004 * total, 2.0):
                    out(f"  {mo:%Y-%m} {k}: {rows[k]:,.1f} GWh disagrees with its share {pct[k]}% -> {est:,.1f}")
                    rows[k] = est
        res[mo] = (rows, total, 1e3)   # GWh -> MWh
    return res


# ---------------------------------------------------------------------------------------------------- K-Electric
KE_KEY = re.compile(r"K-?Electric|\bKE\b|\bKEL\b", re.I)
KE_OWN = ["KE_own_Oil", "KE_own_Gas", "KE_own_RLNG", "KE_own_Coal", "KE_own_unsplit"]
KE_IPP = ["KE_IPP_Oil", "KE_IPP_Gas", "KE_IPP_Coal", "KE_IPP_Nuclear", "KE_IPP_Solar", "KE_IPP_Other"]
NUMS = re.compile(r"\d[\d,]*(?:\.\d+)?")


def last_num(line):
    v = NUMS.findall(line)
    return float(v[-1].replace(",", "")) if v else 0.0


def ke_files():
    """K-Electric's monthly fuel cost filings (NEPRA Admission Notices) and NEPRA's KE FCA decisions."""
    res = []
    r = get("https://nepra.org.pk/news.php")
    for h in sorted(set(re.findall(r'href\s*=\s*["\']([^"\']+\.pdf)["\']', r.text, re.I))):
        n = unquote(h).rsplit("/", 1)[-1]
        if "admission notices" not in unquote(h).lower() or not KE_KEY.search(n) or \
                re.search(r"XWDISCO|WAPDA|petit|tariff|quarter|QTR|licen|write|interven|issues", n, re.I):
            continue
        if not re.search(r"fuel cost variation|FCA|FPA data|fuel charges", n, re.I):
            continue
        y = re.search(r"/(20\d\d)/", unquote(h))
        if y and int(y.group(1)) < DATA_START.year:
            continue
        res.append((urljoin("https://nepra.org.pk/news.php", h), n, "KE filing", month_range(re.sub(r"\.\w+$", "", n))))
    r = get(NEPRA_KE_LIST)
    for h in sorted(set(re.findall(r'href\s*=\s*["\']([^"\']+\.pdf)["\']', r.text, re.I))):
        name = unquote(h)
        n = name.rsplit("/", 1)[-1]
        y = re.search(r"/(?:KESC|K-Electric)/(20\d\d)/", name)
        if not y or int(y.group(1)) < DATA_START.year:
            continue
        if not re.search(r"FCA|MFPA|fuel", n, re.I) or not re.search(r"TRF-362|K-?Electric|\bKE\b", n, re.I) or \
                re.search(r"PAR-14|TRF-14[678]|FPCL|FFBL|SNPC|QTR|quarter|insurance|corrigendum|JUL-JUN|WAPDA", n, re.I):
            continue
        res.append((urljoin(NEPRA_KE_LIST, h), n, "NEPRA KE decision", month_tokens(re.sub(r"\d{1,2}-\d{1,2}-\d{4}", " ", n))[:1]))
    r = get(SOIR_LIST)
    for h in sorted(set(re.findall(r'href\s*=\s*["\']([^"\']+\.pdf)["\']', r.text, re.I))):
        y = re.search(r"State of Industry Report\s*(20\d\d)", unquote(h), re.I)
        if y and int(y.group(1)) > DATA_START.year:   # report Y covers fiscal year Jul Y-1 .. Jun Y
            fy = int(y.group(1))
            res.append((urljoin(SOIR_LIST, h).replace(" ", "%20"), f"State of Industry Report {fy}", "NEPRA SOIR",
                        list(pd.date_range(f"{fy - 1}-07-01", f"{fy}-06-01", freq="MS"))))
    out(f"K-Electric: {sum(f[2] == 'NEPRA SOIR' for f in res)} State of Industry Reports, "
        f"{sum(f[2] == 'KE filing' for f in res)} KE filings, {sum(f[2] == 'NEPRA KE decision' for f in res)} "
        f"NEPRA KE decisions from {DATA_START.year}")
    return res


def pdf_text(url, pages=None):
    with pdfplumber.open(io.BytesIO(get(url).content)) as p:
        return "\n".join(pg.extract_text() or "" for pg in (p.pages[:pages] if pages else p.pages))


def parse_ke_filing(text, months=None):
    """KE's 'Provisional request for monthly fuel cost variation': Annexure B (KE's own sent-out by fuel) and the
    'Power Purchase Details' (CPPA-G, and IPPs on KE's network by fuel), GWh."""
    lines = text.splitlines()
    mo = None
    m = re.search(r"(?:variation|month)\s*(?:for|of)\s*(?:the\s*month\s*of\s*)?([A-Za-z]{3,9})\s*,?\s*(20\d\d)", text, re.I)
    if m:
        t = month_tokens(f"{m.group(1)} {m.group(2)}")
        mo = t[0] if t else None
    mo = mo or (months[0] if months else None)
    if months and len(months) > 1:   # a multi-month filing (one column per month): not parsed
        return {}
    mo = months[0] if months else mo
    if mo is None:
        return {}
    pick = {}
    for ln in lines:
        l = ln.lower()
        if not re.search(r"total\s*units\s*sent\s*out", l):
            continue
        if re.search(r"sent\s*out\s*on\s*furnace", l):
            pick.setdefault("fo", last_num(ln))
        elif re.search(r"sent\s*out\s*on\s*indigenous", l):
            pick["gas"] = last_num(ln)
        elif re.search(r"sent\s*out\s*on\s*(r?lng|ing|inc|l\.?n\.?g)\b", l):
            pick["lng"] = last_num(ln)
        elif re.search(r"sent\s*out\s*on\s*hsd", l):
            pick["hsd"] = last_num(ln)
        elif re.search(r"sent\s*out\s*\W*[xk]e\b", l):
            pick.setdefault("total", last_num(ln))
    row = {}
    if pick:
        parts = sum(pick.get(k, 0.0) for k in ("fo", "gas", "lng", "hsd"))
        total = pick.get("total") or parts
        if total and abs(parts / total - 1) <= 0.03:
            row.update({"KE_own_Oil": pick.get("fo", 0.0) + pick.get("hsd", 0.0), "KE_own_Gas": pick.get("gas", 0.0),
                        "KE_own_RLNG": pick.get("lng", 0.0), "KE_own_filed": total})
    # purchases
    k = next((i for i, ln in enumerate(lines) if re.search(r"power\s*purchase\s*details", ln, re.I)), None)
    if k is not None:
        cat, subs, cppa, total = None, {}, None, None
        for ln in lines[k:k + 60]:
            l = ln.lower()
            if re.search(r"total\s*units\s*purchased", l):
                total = last_num(ln)
                break
            if re.search(r"sub.{0,4}tot|[s5]ub.{0,3}ota", l):
                if cat:
                    subs[cat] = subs.get(cat, 0.0) + last_num(ln)
                continue
            if re.match(r"\s*cppa\b", l) and re.search(r"g\s*w", l):
                cppa = last_num(ln)
            if re.search(r"\bgw|\bgvv|\bcwh|\bgwb", l):   # an item line, not a category header
                continue
            if re.search(r"\bmix\b", l):
                cat = "CPPA"
            elif re.search(r"furnace|\bfo\b|hsd|diesel", l):
                cat = "KE_IPP_Oil"
            elif re.search(r"natural\s*gas|\bng\b|rlng", l):
                cat = "KE_IPP_Gas"
            elif "coal" in l:
                cat = "KE_IPP_Coal"
            elif re.search(r"re[nr]e[wv]|solar|wind", l):
                cat = "KE_IPP_Solar"
            elif re.match(r"\s*[a-h]\s+[a-z]", l):
                cat = "KE_IPP_Other"
        cppa = cppa if cppa is not None else subs.pop("CPPA", None)
        subs.pop("CPPA", None)
        if cppa is not None and total:
            other = sum(subs.values())
            if abs((cppa + other) / total - 1) <= 0.03:
                row.update(subs)
            row["KE_from_CPPA_filed"] = cppa
            row["KE_purchases_filed"] = total
    return {mo: row} if row else {}


def parse_ke_decision(text, months=None):
    """NEPRA's KE FCA decision: the closing table 'Own Generation / Own sent outs GWh' and 'External Purchases GWh'
    (external = CPPA-G + IPPs on KE's network, not split), one column per month after the reference column."""
    lines = text.splitlines()
    dm = decision_month(text) or (months[0] if months else None)
    res = {}
    for i, ln in enumerate(lines):
        l = ln.lower()
        if "rs" in l.split() or "rs." in l or "mln" in l or "min rs" in l or "cost" in l:
            continue
        if re.search(r"own\s*(generation|se\w*\s*outs?)", l):
            key = "own"
        elif re.search(r"ext\w*\W*purchases", l):
            key = "ext"
        else:
            continue
        tail = re.split(r"\bGW\S*", ln, flags=re.I)[-1] if re.search(r"\bgw", l) else \
            re.split(r"outs?|purchases", ln, flags=re.I)[-1]
        toks = NUMS.findall(tail)
        if not re.search(r"\bgw", l) and (len(toks) < 3 or any("." in t for t in toks)):
            continue   # without a GWh unit only an integer row of several months is taken (the JUL-MAR table)
        # OCR'd thousands separator read as a decimal point ('1.131' = 1,131 GWh)
        nums = [float(t.replace(",", "")) * (1000 if re.fullmatch(r"\d\.\d{3}", t) else 1) for t in toks]
        if not nums:
            continue
        head = " ".join(lines[max(0, i - 12):i])
        hm = [t for t in month_tokens(re.sub(r"\b(20\d\d)\b", " ", head))]
        ref = re.search(r"\bref", head, re.I) is not None
        if ref and len(nums) >= 2:
            nums, hm = nums[1:], hm[1:] if len(hm) == len(nums) else hm
        if len(hm) == len(nums) and len(nums) > 2:
            pairs = list(zip(hm, nums))
        elif dm is not None and dm in hm and len(hm) == len(nums):
            pairs = [(dm, nums[hm.index(dm)])]
        elif dm is not None:
            pairs = [(dm, nums[-1])]
        else:
            continue
        for mo, v in pairs:
            if 50 <= v <= 3000:
                res.setdefault(mo, {})[key] = v
    return {mo: {"KE_own_dec": d["own"], "KE_external_dec": d["ext"]} for mo, d in res.items() if "own" in d and "ext" in d}


SOIR_OWN = {"gas": "KE_soir_own_Gas", "rfo": "KE_soir_own_Oil", "hsd": "KE_soir_own_Oil", "rlng": "KE_soir_own_RLNG",
            "coal": "KE_soir_own_Coal", "total": "KE_soir_own"}


def soir_key(label, own):
    l = re.sub(r"[^a-z/ ]", " ", label.lower()).strip()
    if l.startswith("total"):
        return "KE_soir_own" if own else "KE_soir_purch"
    if own:
        for k, v in SOIR_OWN.items():
            if l.startswith(k):
                return v
        return "KE_soir_own_Other"
    if l.startswith("cppa"):
        return "KE_soir_CPPA"
    if "rlng" in l or l.startswith("gas"):
        return "KE_soir_IPP_Gas"
    if l.startswith(("rfo", "hsd", "furnace", "oil")):
        return "KE_soir_IPP_Oil"
    if l.startswith("coal"):
        return "KE_soir_IPP_Coal"
    if l.startswith("nuclear"):
        return "KE_soir_IPP_Nuclear"
    if l.startswith(("solar", "wind", "net", "renew")):
        return "KE_soir_IPP_Solar"
    return "KE_soir_IPP_Other"


def soir_values(line):
    """'Generation GWh 761.717 743.386 - ... % ...' -> the 12 monthly values ('-' = 0)."""
    rest = re.split(r"GWh", line, maxsplit=1, flags=re.I)[-1]
    rest = re.split(r"\s%\s|\s%$", " " + rest + " ")[0]
    toks = re.findall(r"(?<![\d.])-(?![\d])|\d[\d,]*\.?\d*%?", rest.replace(" ,", ",").replace(", ", ","))
    vals = []
    for t in toks:
        if t.endswith("%"):
            break
        vals.append(0.0 if t == "-" else float(t.replace(",", "")))
    return vals[:12] if len(vals) >= 12 else None


def parse_soir_table(text, own):
    """One 'K-Electric (Fuel-wise Own Generation / Power Purchase ...) (YYYY-YY)' page -> {month: {field: GWh}}."""
    m = re.search(r"[({](20\d\d)\s*-\s*(\d\d)[)}]", text)
    if not m:
        return {}
    months = list(pd.date_range(f"{m.group(1)}-07-01", periods=12, freq="MS"))
    lines = [l.strip() for l in text.splitlines() if l.strip()]
    rows, label = {}, None
    for l in lines:
        if re.match(r"generation\s*gwh", l, re.I):
            v = soir_values(l)
            if v and label:
                k = soir_key(label, own)
                rows[k] = [a + b for a, b in zip(rows[k], v)] if k in rows and k not in ("KE_soir_own", "KE_soir_purch") else v
            label = None
        elif not re.match(r"(fuel\s*cost|cost\s*mil|rs\.?/kwh|%|source|company|july|table|k-?electric|state of)", l, re.I) \
                and re.search(r"[A-Za-z]{2}", l) and len(re.findall(r"\d[\d,.]*", l)) <= 2:
            label = l
    tot = rows.get("KE_soir_own" if own else "KE_soir_purch")
    if not tot:
        return {}
    parts = [v for k, v in rows.items() if k not in ("KE_soir_own", "KE_soir_purch")]
    out_ = {}
    for i, mo in enumerate(months):
        s_ = sum(p[i] for p in parts)
        if tot[i] and abs(s_ / tot[i] - 1) > 0.02:
            continue   # a month whose fuel rows do not add up (bad text layer): left out
        out_[mo] = {k: v[i] for k, v in rows.items()}
    return out_


SOIR_LOCK = __import__("threading").Lock()   # pdfium is not thread-safe


def read_soir(url):
    with SOIR_LOCK:
        return _read_soir(url)


def _read_soir(url):
    """NEPRA State of Industry Report: KE's monthly fuel-wise own generation and fuel-wise power purchases tables."""
    import pypdfium2 as pdfium
    res = {}
    with tempfile.NamedTemporaryFile(suffix=".pdf") as f:
        with SESSION.get(url, timeout=(20, 900), verify=False, stream=True) as r:
            r.raise_for_status()
            for ch in r.iter_content(1 << 20):
                f.write(ch)
        f.flush()
        doc = pdfium.PdfDocument(f.name)
        hits = []
        for i in range(len(doc)):
            t = doc[i].get_textpage().get_text_range()[:700]
            m = re.search(r"Fuel-?\s*wise\s*(Own\s*Generation|Power\s*Purchase)", t, re.I)
            if m and re.search(r"K-?\s*Electric", t, re.I):
                hits.append((i, "own" in m.group(1).lower()))
        doc.close()
        with pdfplumber.open(f.name) as p:
            for i, own in hits:
                for mo, row in parse_soir_table(p.pages[i].extract_text() or "", own).items():
                    res.setdefault(mo, {}).update(row)
    return res


def _safe(fn, f):
    try:
        return fn(f), None
    except Exception as e:  # noqa: BLE001
        return {}, f"{type(e).__name__}: {e}"


def read_ke(f):
    url, name, kind, mos = f
    if kind == "NEPRA SOIR":
        return read_soir(url)
    if kind == "KE filing":
        return parse_ke_filing(pdf_text(url, pages=14), mos)
    return parse_ke_decision(pdf_text(url), mos)


def ke_monthly(grid_index, ke):
    """KE history (GWh, as filed) -> MWh per grid month; real figures only, nothing copied or estimated.
    own sent-out: NEPRA's KE decision (closing table or 'Company Wide Mix'), else KE's own filing summary;
    its fuel split: the decision's mix table, else KE's filing when its plant blocks add up to that total within 5%,
    else KE_own_unsplit (counted as gas); purchases from IPPs on KE's network: the mix table or KE's filing (by fuel)
    when consistent, else external purchases minus the CPPA-G drawl where both are filed (KE_purchases_other,
    no fuel), else blank. KE_included = KE's own sent-out is known for the month."""
    g = lambda r, c: (None if r is None or pd.isna(r.get(c + "_GWh", float("nan"))) else float(r[c + "_GWh"]))  # noqa
    rec = {mo: ke.loc[mo].to_dict() for mo in ke.index} if not ke.empty else {}
    rows = {}
    for mo in grid_index:
        r = rec.get(mo)
        d, notes = {}, []
        tot = None
        if r and g(r, "KE_soir_own") is not None and g(r, "KE_soir_purch") is not None:
            # NEPRA State of Industry Report, KE's own fuel-wise monthly tables: complete and exact
            for c, k in (("KE_own_Gas", "KE_soir_own_Gas"), ("KE_own_RLNG", "KE_soir_own_RLNG"),
                         ("KE_own_Oil", "KE_soir_own_Oil"), ("KE_own_unsplit", "KE_soir_own_Other"),
                         ("KE_own_Coal", "KE_soir_own_Coal")):
                if g(r, k) is not None:
                    d[c] = g(r, k)
            for c in KE_IPP:
                v = g(r, c.replace("KE_IPP_", "KE_soir_IPP_"))
                if v is not None:
                    d[c] = v
            d["KE_from_CPPA"] = g(r, "KE_soir_CPPA") or 0.0
            d["KE_external"] = g(r, "KE_soir_purch")
            row = {f"{c}_MWh": round(v * 1000, 1) for c, v in d.items()}
            row["KE_basis"] = "NEPRA State of Industry Report (KE fuel-wise monthly tables)"
            row["KE_included"] = True
            rows[mo] = row
            continue
        if r:
            for c, lab in (("KE_own_mix", "NEPRA KE decision (mix table)"), ("KE_own_dec", "NEPRA KE decision"),
                           ("KE_own_sum", "KE filing summary")):
                if g(r, c) is not None:
                    tot, src = g(r, c), lab
                    break
        if tot is None:
            rows[mo] = {"KE_basis": "no KE figure", "KE_included": False}
            continue
        mix = [g(r, "KE_mix_" + c) for c in ("Oil", "Gas", "RLNG")]
        parts = [g(r, c) for c in ("KE_own_Oil", "KE_own_Gas", "KE_own_RLNG")]
        if None not in mix and sum(mix) > 0 and abs(sum(mix) / tot - 1) <= 0.05:
            d.update(dict(zip(("KE_own_Oil", "KE_own_Gas", "KE_own_RLNG"), mix)))
            notes.append(f"own: {src}, by fuel from the decision's mix table")
        elif None not in parts and sum(parts) > 0 and abs(sum(parts) / tot - 1) <= 0.05:
            f = tot / sum(parts)
            d.update({c: v * f for c, v in zip(("KE_own_Oil", "KE_own_Gas", "KE_own_RLNG"), parts)})
            notes.append(f"own: {src}, by fuel from KE's filing")
        else:
            d["KE_own_unsplit"] = tot
            notes.append(f"own: {src}, no fuel split (counted as gas)")
        ext = g(r, "KE_external_dec")
        if ext is None:
            ext = g(r, "KE_purchases_sum")
        cppa = g(r, "KE_mix_CPPA")
        if cppa is None:
            cppa = g(r, "KE_from_CPPA_filed")
        subs = {c: g(r, c.replace("KE_IPP_", "KE_mix_IPP_")) for c in KE_IPP}
        subs = {c: v for c, v in subs.items() if v is not None}
        src_ipp = "decision's mix table"
        if not subs:
            subs = {c: g(r, c) for c in KE_IPP if g(r, c) is not None}
            src_ipp = "KE's filing"
        ok = subs and cppa is not None and (ext is None or abs((cppa + sum(subs.values())) / ext - 1) <= 0.05)
        if ok:
            d.update(subs)
            notes.append(f"IPP purchases by fuel: {src_ipp}")
        elif ext is not None and cppa is not None and 0 <= ext - cppa <= 0.5 * ext:
            d["KE_purchases_other"] = ext - cppa
            notes.append("IPP purchases: external minus CPPA-G (no fuel split)")
        else:
            notes.append("IPP purchases: no figure")
        if cppa is not None:
            d["KE_from_CPPA"] = cppa
        if ext is not None:
            d["KE_external"] = ext
        row = {f"{c}_MWh": round(v * 1000, 1) for c, v in d.items()}
        row["KE_basis"] = "; ".join(notes)
        row["KE_included"] = True
        rows[mo] = row
    k = pd.DataFrame.from_dict(rows, orient="index")
    if not k.empty:
        k["KE_own_MWh"] = k[[c + "_MWh" for c in KE_OWN if c + "_MWh" in k]].sum(axis=1, min_count=1)
        k["KE_purchases_nonCPPA_MWh"] = k[[c + "_MWh" for c in KE_IPP + ["KE_purchases_other"] if c + "_MWh" in k]]\
            .sum(axis=1, min_count=1)
    return k


# ---------------------------------------------------------------------------------------------------- frames
def to_row(rows, scale):
    g = {k: rows.get(k) for k in DETAIL}
    f = lambda *ks: sum(g[k] for k in ks if g[k] is not None) if any(g[k] is not None for k in ks) else None  # noqa
    coal = f("Coal", "Coal_local", "Coal_imported")
    r = {"Hydro_MWh": f("Hydro"), "Coal_MWh": coal, "Gas_MWh": f("Gas_local", "RLNG"), "Oil_MWh": f("RFO", "HSD"),
         "Nuclear_MWh": f("Nuclear"), "Wind_MWh": f("Wind"), "Solar_MWh": f("Solar"), "Bioenergy_MWh": f("Bagasse"),
         "Other_MWh": f("Mixed"), "Imports_MWh": f("Imports"), "Coal_local_MWh": f("Coal_local"),
         "Coal_imported_MWh": f("Coal_imported"), "Gas_local_MWh": f("Gas_local"), "RLNG_MWh": f("RLNG"),
         "RFO_MWh": f("RFO"), "HSD_MWh": f("HSD")}
    r = {k: (None if v is None else round(v * scale, 1)) for k, v in r.items()}
    r["Total_MWh"] = round(sum(r[c] or 0.0 for c in COLS[:9]), 1)
    return r


def read_sheet(path, sheet, index_col=0):
    try:
        d = pd.read_excel(path, sheet_name=sheet, index_col=index_col)
    except (FileNotFoundError, ValueError, KeyError):
        return pd.DataFrame()
    if index_col is not None:
        d.index = pd.to_datetime(d.index)
        d = d.sort_index()
    return d


STD = ["Hydro", "Coal", "Gas", "Oil", "Nuclear", "Wind", "Solar", "Bioenergy", "Other"]


def combine(grid, kem):
    """Daily = national grid (CPPA-G) + K-Electric's own plants + KE's purchases from plants outside the CPPA-G pool."""
    d = pd.DataFrame(index=grid.index)
    k = kem.reindex(grid.index) if not kem.empty else pd.DataFrame(index=grid.index)
    kg = lambda *cs: sum(k[c + "_MWh"].fillna(0.0) for c in cs if c + "_MWh" in k) if any(  # noqa: E731
        c + "_MWh" in k for c in cs) else 0.0
    add_ke = {"Gas": kg("KE_own_Gas", "KE_own_RLNG", "KE_own_unsplit", "KE_IPP_Gas"),
              "Oil": kg("KE_own_Oil", "KE_IPP_Oil"), "Coal": kg("KE_own_Coal", "KE_IPP_Coal"),
              "Nuclear": kg("KE_IPP_Nuclear"), "Solar": kg("KE_IPP_Solar"),
              "Other": kg("KE_IPP_Other", "KE_purchases_other")}
    for f in STD:
        g = grid.get(f"{f}_MWh")
        d[f"{f}_MWh"] = (g.fillna(0.0) if g is not None else 0.0) + add_ke.get(f, 0.0)
    d["Total_MWh"] = d[[f"{f}_MWh" for f in STD]].sum(axis=1)
    d["Imports_MWh"] = grid.get("Imports_MWh")
    d["Total_grid_MWh"] = grid["Total_MWh"]
    for f in STD:
        d[f"{f}_grid_MWh"] = grid.get(f"{f}_MWh")
    for c in COLS[11:]:
        d[c] = grid.get(c)
    for c in k.columns:
        d[c] = k[c]
    return d.round(1)


def save(path, grid, months, files, ke):
    grid = grid[[c for c in COLS if c in grid]].sort_index()
    kem = ke_monthly(grid.index, ke)
    daily = combine(grid, kem)
    for x in (daily, grid, months, ke):
        x.index.name = "date"
    first, last = daily.index.min(), daily.index.max()
    yr = daily.groupby(daily.index.year).agg({"Total_MWh": ["sum", "count"], "Total_grid_MWh": "sum"})
    full = {y: (round(r[("Total_MWh", "sum")] / 1e6, 1), round(r[("Total_grid_MWh", "sum")] / 1e6, 1))
            for y, r in yr.iterrows() if r[("Total_MWh", "count")] == 12}
    src = months["Source"].value_counts().to_dict() if "Source" in months else {}
    kb = kem["KE_basis"].str.split(" ").str[0].str.strip(":(").value_counts().to_dict() if "KE_basis" in kem else {}
    notes = [
        "UNITS",
        "Daily: one row per MONTH dated the 1st; MWh in the month. PAKISTAN INCLUDING K-ELECTRIC: each standard fuel "
        "column (Hydro, Coal, Gas, Oil, Nuclear, Wind, Solar, Bioenergy, Other) = the national grid (CPPA-G pool, "
        "<fuel>_grid_MWh) + K-Electric's own plants + KE's purchases from plants outside the CPPA-G pool. Total_MWh = "
        "their sum; Total_grid_MWh = the national grid alone (the energy CPPA-G buys, which already includes what the "
        "grid supplies to KE). KE's purchases from CPPA-G (KE_from_CPPA_MWh) are a detail column only and are NOT "
        "added (that would double count). Imports_MWh (Tavanir, Iran, CPPA-G) is kept apart.",
        "National grid fuels: Hydro (WAPDA and IPP hydel), Coal (Coal_local_MWh, Thar lignite etc., + "
        "Coal_imported_MWh where the filing splits them; NEPRA's decisions before mid-2023 give coal unsplit), Gas = "
        "Gas_local_MWh (pipeline gas) + RLNG_MWh, Oil = RFO_MWh + HSD_MWh, Nuclear (Chashma, Karachi K-2/K-3), Wind, "
        "Solar, Bioenergy (bagasse), Other (CPPA-G's 'Mixed'). Coal_local/Coal_imported/Gas_local/RLNG/RFO/HSD columns "
        "are national-grid only.",
        "K-Electric: KE_own_Gas_MWh (indigenous gas), KE_own_RLNG_MWh, KE_own_Oil_MWh (furnace oil at BQPS-I + HSD at "
        "KCCPP), KE_own_unsplit_MWh (KE's own sent-out where the source gives no fuel split - counted under Gas, as "
        "KE's plants run mainly on gas/RLNG; this overstates gas and understates oil in those months), KE_own_MWh = "
        "their sum (sent-out, net of auxiliaries); KE_IPP_<fuel>_MWh = purchases from plants on KE's network outside "
        "the CPPA-G pool (Oil: Gul Ahmed / Tapal; Gas: SNPC I/II, Lucky, ISL, Lotte; Coal: FFBL/FPCL; Nuclear: "
        "KANUPP until Aug 2021; Solar: Oursun, Gharo and net metering), KE_purchases_other_MWh = such purchases "
        "without a fuel split (counted under Other), KE_purchases_nonCPPA_MWh = their sum; KE_from_CPPA_MWh = drawn "
        "from the national grid (detail / check only, never added); KE_external_MWh = all of KE's purchases.",
        "KE_basis: where each month's KE figures come from - REAL FIGURES ONLY, nothing copied or estimated. First "
        "NEPRA's State of Industry Report (KE's fuel-wise own-generation and power-purchase tables, by month, one "
        "report per fiscal year); else NEPRA's monthly KE FCA decision (own sent-out; its 'Company Wide Mix' table "
        "where given) with the fuel split from KE's own filing when its plant blocks add up within 5%; IPP purchases "
        "only where filed (else blank). KE_included = True where KE's own sent-out is known. Months with no KE figure "
        "(KE_basis 'no KE figure', KE_included False): KE columns blank and Total_MWh / the fuel columns are the "
        "national grid only - the master fills such months for the whole country from Ember.",
        "Months: the national-grid source used for each month and Check_pct = (sum of the fuel rows / the filing's "
        "own total - 1) x 100. Grid / KE: the history stores (as filed). Files: every file read.",
        "",
        "COVERAGE",
        f"Monthly from {first:%b %Y} to {last:%b %Y} ({len(daily)} months; national-grid sources: {src}; KE basis: {kb}). "
        f"Calendar-year totals, TWh (Pakistan incl. KE, national grid alone): {full}.",
        "Scope: all of Pakistan's utility generation - the national grid (NTDC/NGC, ISMO) and K-Electric (Karachi). "
        "NOT included: captive / off-grid generation and rooftop solar other than net-metered exports bought by KE.",
        "",
        "SOURCE",
        f"CPPA-G (Central Power Purchasing Agency Guarantee Ltd), XWDISCOs Energy Purchase Data: {CPPA_LIST} "
        "(Excel workbooks and monthly PDFs, July 2024 on).",
        f"NEPRA (National Electric Power Regulatory Authority), monthly Fuel Charges Adjustment decisions for the "
        f"Ex-WAPDA DISCOs, Annex-II 'Source Wise Generation' (actual GWh): {NEPRA_LIST}",
        f"NEPRA State of Industry Reports (statistical tables: K-Electric fuel-wise own generation and fuel-wise power "
        f"purchases, monthly): {SOIR_LIST}",
        "K-Electric's monthly 'Provisional request for monthly fuel cost variation' filings (Annexure B sent-out by "
        "fuel; power purchase details), NEPRA Admission Notices: https://nepra.org.pk/news.php ; NEPRA's monthly FCA "
        f"decisions for K-Electric (own sent-out and external purchases): {NEPRA_KE_LIST}",
    ]
    files = files.copy()
    files.index = range(1, len(files) + 1)
    files.index.name = "n"
    xlsx_notes.write_workbook(path, {"Daily": daily, "Grid": grid, "KE": ke.sort_index(), "Months": months.sort_index(),
                                     "Files": files}, notes, {"UNITS", "COVERAGE", "SOURCE"})
    out(f"Saved {path}: {len(daily)} months {first:%Y-%m}..{last:%Y-%m}; TWh (incl. KE, grid only): {full}")
    return daily


# ---------------------------------------------------------------------------------------------------- hourly
HOURLY_FUEL = {"hydel": "Hydro", "hydro": "Hydro", "coal": "Coal", "rlng": "RLNG", "gas": "Gas_local", "rfo": "RFO",
               "hsd": "HSD", "nuclear": "Nuclear", "wind": "Wind", "solar": "Solar", "bagasse": "Bagasse"}


def hourly_files():
    urls = set(KNOWN_HOURLY)
    for page in NEPRA_HOURLY_PAGES:
        try:
            t = get(page, tries=2).text
        except requests.RequestException as e:
            out(f"  {page}: {e}")
            continue
        for h in re.findall(r'href\s*=\s*["\']([^"\']+\.xlsx?)["\']', t, re.I):
            if re.search(r"plant\s*wise", unquote(h), re.I):
                urls.add(urljoin(page, h).replace(" ", "%20"))
    return sorted(urls)


def read_hourly(url):
    d = pd.read_excel(io.BytesIO(get(url).content))
    c = {k.lower().replace(" ", "").replace("_", ""): k for k in d.columns}
    date_c = next(v for k, v in c.items() if "date" in k)
    hour_c = next(v for k, v in c.items() if "hour" in k)
    mw_c = next(v for k, v in c.items() if k in ("mw", "mwh", "energy", "units") or k.startswith("mw"))
    fuel_c = next(v for k, v in c.items() if "fuel" in k)
    hr = pd.to_numeric(d[hour_c], errors="coerce")
    off = 1 if hr.max() >= 24 else 0   # Hour_No 1..24 = the hour ending
    ts = pd.to_datetime(d[date_c]).dt.normalize() + pd.to_timedelta(hr - off, unit="h")
    fuel = d[fuel_c].astype(str).str.strip().str.lower().map(HOURLY_FUEL).fillna("Other")
    x = pd.DataFrame({"ts": ts, "fuel": fuel, "MW": pd.to_numeric(d[mw_c], errors="coerce")})
    w = x.pivot_table(index="ts", columns="fuel", values="MW", aggfunc="sum").sort_index()
    w.columns = [f"{c}_MW" for c in w.columns]
    out(f"  hourly {unquote(url)[-40:]}: {len(d)} rows, {d[fuel_c].nunique()} fuels, {w.index.min()}..{w.index.max()}")
    return w


def read_prices(url):
    try:
        d = pd.read_excel(io.BytesIO(get(url, tries=2).content))
    except Exception as e:  # noqa: BLE001
        out(f"  prices {url}: {e}")
        return pd.DataFrame()
    c = {k.lower().replace(" ", "").replace("_", ""): k for k in d.columns}
    try:
        date_c = next(v for k, v in c.items() if "date" in k)
        hour_c = next(v for k, v in c.items() if "hour" in k)
        p_c = next(v for k, v in c.items() if "price" in k)
    except StopIteration:
        return pd.DataFrame()
    hr = pd.to_numeric(d[hour_c], errors="coerce")
    off = 1 if hr.max() >= 24 else 0
    ts = pd.to_datetime(d[date_c]).dt.normalize() + pd.to_timedelta(hr - off, unit="h")
    return pd.DataFrame({"Marginal_price_Rs_per_kWh": pd.to_numeric(d[p_c], errors="coerce").values},
                        index=pd.DatetimeIndex(ts, name="datetime")).sort_index()


def update_hourly(path):
    done = read_sheet(path, "Files", index_col=None)
    seen = set(done["url"]) if "url" in done else set()
    todo = [u for u in hourly_files() if u not in seen]
    if not todo:
        out("Hourly plant-wise file: nothing new")
        return
    old = read_sheet(path, "Hourly")
    oldp = read_sheet(path, "Prices")
    frames, prices, files = [old] if not old.empty else [], [oldp] if not oldp.empty else [], [done]
    for u in todo:
        try:
            w = read_hourly(u)
        except Exception as e:  # noqa: BLE001
            out(f"  hourly {u}: {type(e).__name__}: {e}")
            continue
        frames.append(w)
        pu = re.sub(r"01-%20Plant%20[Ww]ise%20Units", "03-%20Hourly%20Marginal%20Price", u)
        if pu != u:
            p = read_prices(pu)
            if not p.empty:
                prices.append(p)
        files.append(pd.DataFrame({"url": [u], "from": [f"{w.index.min():%Y-%m-%d}"], "to": [f"{w.index.max():%Y-%m-%d}"]}))
    if not frames:
        return
    h = pd.concat(frames)
    h = h[~h.index.duplicated(keep="last")].sort_index().fillna(0.0)
    h.index.name = "datetime"
    day = h.resample("D").sum()
    g = lambda *ks: sum(day[f"{k}_MW"] for k in ks if f"{k}_MW" in day)   # noqa: E731
    daily = pd.DataFrame({"Hydro_MWh": g("Hydro"), "Coal_MWh": g("Coal"), "Gas_MWh": g("Gas_local", "RLNG"),
                          "Oil_MWh": g("RFO", "HSD"), "Nuclear_MWh": g("Nuclear"), "Wind_MWh": g("Wind"),
                          "Solar_MWh": g("Solar"), "Bioenergy_MWh": g("Bagasse"), "Other_MWh": g("Other")},
                         index=day.index)
    daily = daily.loc[:, (daily != 0).any()] if not daily.empty else daily
    daily["Total_MWh"] = daily.sum(axis=1)
    daily["Gas_local_MWh"], daily["RLNG_MWh"] = g("Gas_local"), g("RLNG")
    daily.index.name = "date"
    h["Total_MW"] = h.sum(axis=1)
    sheets = {"Hourly": h.round(2), "Daily": daily.round(1)}
    if prices:
        pr = pd.concat(prices)
        sheets["Prices"] = pr[~pr.index.duplicated(keep="last")].sort_index()
    f = pd.concat([x for x in files if not x.empty], ignore_index=True)
    sheets["Files"] = f.set_index("url")
    notes = [
        "UNITS",
        "Hourly: MW by fuel for each hour (hourly energy, = MWh in the hour), summed over CPPA-G's dispatched plants "
        "(129 plants in the first file). Fuels as CPPA-G labels them: Hydro (Hydel), Coal, RLNG, Gas_local (GAS/Gas), "
        "RFO, HSD, Nuclear, Wind, Solar, Bagasse. Timestamps are the hour's start (Hour_No 1 = 00:00-01:00), PKT.",
        "Daily: standard layout, MWh per day (the hourly values summed); Gas_MWh = Gas_local + RLNG, Oil_MWh = RFO + "
        "HSD, Bioenergy_MWh = bagasse.",
        "Prices: CPPA-G's hourly marginal price, Rs/kWh (filed with the same application).",
        "",
        "COVERAGE",
        f"{h.index.min():%Y-%m-%d} to {h.index.max():%Y-%m-%d}: a one-off snapshot filed with CPPA-G's fuel price "
        "adjustment application (NEPRA Admission Notices, Sep 2026). New files of the same name linked from NEPRA's "
        "home or news page are added when they appear. CPPA-G system only (no K-Electric own generation).",
        "",
        "SOURCE",
        "NEPRA Admission Notices: '01- Plant wise Units.xlsx' and '03- Hourly Marginal Price.xlsx', "
        f"{', '.join(todo)}",
    ]
    xlsx_notes.write_workbook(path, sheets, notes, {"UNITS", "COVERAGE", "SOURCE"})
    m = (daily.resample("MS").sum() / 1000).round(0)
    out(f"Saved {path}: {len(h)} hours; monthly GWh:\n{m.to_string()}")


# ---------------------------------------------------------------------------------------------------- main
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", default=OUT)
    ap.add_argument("--hourly-out", default=OUT_HOURLY)
    ap.add_argument("--no-hourly", action="store_true")
    args = ap.parse_args()

    daily, months = read_sheet(args.out, "Grid"), read_sheet(args.out, "Months")
    if daily.empty:   # first layout (grid only) kept the national grid on the Daily sheet
        d0 = read_sheet(args.out, "Daily")
        daily = d0 if "Total_grid_MWh" not in d0 else pd.DataFrame()
    ke = read_sheet(args.out, "KE")
    try:
        files = pd.read_excel(args.out, sheet_name="Files", index_col=0).reset_index(drop=True)
    except (FileNotFoundError, ValueError, KeyError):
        files = pd.DataFrame()
    if files.empty:
        files = pd.DataFrame(columns=["url", "kind", "name", "months", "status"])
    seen = set(files["url"])
    rank = {d: RANK.get(s, 0) for d, s in months["Source"].items()} if "Source" in months else {}
    out(f"{len(daily)} months saved ({daily.index.min() if len(daily) else '-'}..{daily.index.max() if len(daily) else '-'}); "
        f"{len(seen)} files read before")

    listing = []
    for fn in (cppa_files, nepra_decisions):
        try:
            listing += fn()
        except requests.RequestException as e:
            out(f"  listing {fn.__name__} failed: {e}")
    # best source first, so a month got from a workbook does not also pull a PDF or a decision
    listing.sort(key=lambda f: (-RANK[f[2]], f[3][0] if f[3] else pd.Timestamp.max))
    new_rows, new_meta, new_files = {}, {}, []

    def needed(f):
        url, _, kind, mos = f
        if url in seen:
            return False
        if not mos:
            return True
        cur = {**rank, **{m: RANK[new_meta[m]["Source"]] for m in new_meta}}
        return any(m >= DATA_START and cur.get(m, 0) < RANK[kind] for m in mos)

    def fetch(f):
        url, name, kind, mos = f
        try:
            if kind == "CPPA-G workbook":
                return f, read_cppa_workbook(url), None
            if kind == "CPPA-G PDF":
                return f, read_cppa_pdf(url, mos), None
            return f, read_decision(url, mos), None
        except Exception as e:  # noqa: BLE001
            return f, {}, f"{type(e).__name__}: {e}"

    for kind in RANK:   # one source rank at a time (each later rank sees what the better ones gave)
        todo = [f for f in listing if f[2] == kind and needed(f)]
        out(f"{kind}: {len(todo)} files to read")
        with ThreadPoolExecutor(4) as ex:
            for f, res, err in ex.map(fetch, todo):
                url, name, k, mos = f
                got = []
                for mo, (rows, total, scale) in sorted(res.items()):
                    c = check(rows, total)
                    if total and c is not None and abs(c) > 0.02:
                        out(f"  {mo:%Y-%m} from {name[:60]}: fuel rows off the total by {c:+.1%} - skipped")
                        continue
                    if mo < DATA_START:
                        continue
                    have = RANK.get(new_meta[mo]["Source"], 0) if mo in new_meta else rank.get(mo, 0)
                    if have > RANK[k] or (have == RANK[k] and mo in new_meta):
                        continue
                    new_rows[mo] = to_row(rows, scale)
                    new_meta[mo] = {"Source": k, "File": name[:120], "URL": url,
                                    "Check_pct": None if c is None else round(100 * c, 2),
                                    "Filing_total_MWh": None if not total else round(total * scale, 1)}
                    got.append(f"{mo:%Y-%m}")
                status = err or ("ok" if got else "no month parsed")
                if err or not got:
                    out(f"  {name[:80]}: {status}")
                new_files.append({"url": url, "kind": k, "name": name[:150], "months": " ".join(got), "status": status})
    if new_rows:
        nd = pd.DataFrame.from_dict(new_rows, orient="index")
        nm = pd.DataFrame.from_dict(new_meta, orient="index")
        daily = pd.concat([daily[~daily.index.isin(nd.index)], nd]) if not daily.empty else nd
        months = pd.concat([months[~months.index.isin(nm.index)], nm]) if not months.empty else nm
    # K-Electric: KE's own filings first, then NEPRA's KE decisions, for months not saved from a better KE source
    try:
        kl = ke_files()
    except requests.RequestException as e:
        out(f"  KE listing failed: {e}")
        kl = []
    new_ke = {}
    todo = [f for f in kl if f[0] not in seen]
    out(f"K-Electric: {len(todo)} files to read")
    with ThreadPoolExecutor(4) as ex:
        for f, res, err in ex.map(lambda f: (f, *_safe(read_ke, f)), todo):
            url, name, k, mos = f
            got = []
            for mo, row in sorted(res.items()):
                if mo < DATA_START:
                    continue
                cur = new_ke.setdefault(mo, {})
                cur.update({c + "_GWh": v for c, v in row.items() if v is not None})
                cur["KE_files"] = (cur.get("KE_files", "") + " | " + name[:80]).strip(" |")
                got.append(f"{mo:%Y-%m}")
            status = err or ("ok" if got else "no month parsed")
            if err or not got:
                out(f"  {name[:80]}: {status}")
            if k == "NEPRA SOIR" and not got:
                continue   # not recorded: tried again next run (the table reader still needs work)
            new_files.append({"url": url, "kind": k, "name": name[:150], "months": " ".join(got), "status": status})
    if new_ke:
        old_ke = {pd.Timestamp(m): {c: v for c, v in r.items() if not pd.isna(v)} for m, r in ke.to_dict("index").items()} \
            if not ke.empty else {}
        for mo, row in new_ke.items():
            old_ke.setdefault(mo, {}).update(row)
        ke = pd.DataFrame.from_dict(old_ke, orient="index").sort_index()
    if not ke.empty:
        ke.index = pd.to_datetime(ke.index)
    files = pd.concat([files, pd.DataFrame(new_files)], ignore_index=True) if new_files else files
    if daily.empty:
        raise SystemExit("No Pakistan monthly generation parsed")
    daily.index = pd.to_datetime(daily.index)
    months.index = pd.to_datetime(months.index)
    combined = save(args.out, daily.sort_index(), months, files, ke)
    gaps = pd.date_range(DATA_START, daily.index.max(), freq="MS").difference(daily.index)
    if len(gaps):
        out(f"Months missing: {[f'{g:%Y-%m}' for g in gaps]}")
    show = combined[["Hydro_MWh", "Coal_MWh", "Gas_MWh", "Oil_MWh", "Nuclear_MWh", "Wind_MWh", "Solar_MWh",
                     "Bioenergy_MWh", "Total_MWh", "Total_grid_MWh", "Imports_MWh"]] / 1000
    out("GWh, last 6 months:\n" + show.tail(6).round(0).to_string())
    if not args.no_hourly:
        try:
            update_hourly(args.hourly_out)
        except Exception as e:  # noqa: BLE001
            out(f"Hourly file failed: {type(e).__name__}: {e}")


if __name__ == "__main__":
    main()
