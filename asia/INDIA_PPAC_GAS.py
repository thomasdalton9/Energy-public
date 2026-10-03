"""
India natural gas from PPAC (Petroleum Planning & Analysis Cell, Ministry of Petroleum and Natural Gas),
https://ppac.gov.in/natural-gas/. Found via discovery_archive/subcontinent/SEA_DISCOVERY_INDIA.py and
SEA_DISCOVERY_SOUTH_ASIA2/3.py.

  Sectoral consumption   monthly, MMSCM, by sector (power, CGD, fertiliser, refinery, petrochemical, LPG
                         shrinkage, sponge iron, industrial, manufacturing, pipeline internal use, ...), each
                         split into domestic gas and RLNG. History workbook NG-H_Sectoral_Consumption.xlsx
                         (one sheet per financial year, from FY 2015-16) + the current-FY file
                         NG-C-Sectoral-Consumption.xlsx (sheet 'NG-H-SC').
  LNG imports            monthly, MMSCM (and MMT): NG-H_LNG_Import.xlsx ('Month wise <FY>' sheets) + the
                         current-FY NG-C-LNG-Import.xls.
  Production/consumption monthly, MMSCM: net production, LNG imports and total consumption, from the page's
                         own data call (AjaxController/getGasConsumption, per financial year).

File names carry an upload stamp, so links are read from the pages each run. The workbook keeps every month
already saved (PPAC overwrites the current-FY files in place); each run re-reads the current and previous
financial year (revisions) and fills any month not yet saved from the history files.

Writes output/Data and Chart Outputs/india_gas.xlsx: 'Sectoral' (month x sector, total MMSCM), 'Sectoral RLNG'
(RLNG part), 'Balance' (net production, LNG imports, total consumption), 'LNG imports'.

    python3 asia/INDIA_PPAC_GAS.py
"""
import argparse
import io
import os
import re
import sys
import time
from datetime import date

import pandas as pd
import requests

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)
import xlsx_notes  # noqa: E402

BASE = "https://ppac.gov.in"
H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0 Safari/537.36"}
T = (20, 120)
FIRST_FY = 2018
OUT = os.path.join(ROOT, "output", "Data and Chart Outputs", "india_gas.xlsx")
MONTHS = {m: i for i, m in enumerate(["jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov",
                                      "dec"], start=1)}
SKIP_ROWS = ("energy sector", "non energy", "sectoral consumption", "source", "note", "mmt", "mmscm", "*", "1 mmt")


def out(*a):
    print(*a, flush=True)


def fy_now(d=None):
    d = d or date.today()
    return d.year if d.month >= 4 else d.year - 1


def fy_month(fy_start, month_name):
    m = MONTHS[month_name.strip().lower()[:3]]
    return pd.Timestamp(fy_start + (0 if m >= 4 else 1), m, 1)


def page_links(session, page):
    r = session.get(f"{BASE}/natural-gas/{page}", timeout=T)
    r.raise_for_status()
    links = []
    for u in re.findall(r'["\']([^"\']+\.xlsx?)["\']', r.text):
        if not u.startswith("http"):
            u = f"{BASE}/uploads/page-images/{u.lstrip('/')}"
        links.append(u)
    # oldest upload first, so the newest file's months win when files overlap (names start with an upload stamp)
    stamp = lambda u: int((re.match(r"(\d{9,})", u.rsplit("/", 1)[-1]) or [None, 0])[1])  # noqa: E731
    return sorted(dict.fromkeys(links), key=stamp)


def get_xl(session, url):
    r = session.get(url, timeout=T)
    r.raise_for_status()
    if r.content[:2] not in (b"PK", b"\xd0\xcf"):
        raise ValueError(f"not a workbook: {url}")
    return pd.ExcelFile(io.BytesIO(r.content))


def sectoral_sheet(df):
    """FY sheet: a row of month dates over (Domestic, RLNG, Total) triples, sector rows below ->
    (total, rlng) frames indexed by month, one column per sector."""
    sub_i = next(i for i in range(len(df)) if (df.iloc[i].astype(str).str.strip().str.lower() == "domestic").sum() >= 2)
    months = df.iloc[sub_i - 1]
    month_at, cur = {}, None
    for j in range(df.shape[1]):
        v = months.iloc[j]
        t = pd.to_datetime(v, errors="coerce") if not isinstance(v, str) else pd.to_datetime(v, format="mixed",
                                                                                              errors="coerce")
        if pd.notna(t) and 2000 < t.year < 2100:
            cur = t.to_period("M").to_timestamp()
        elif isinstance(v, str) and re.search(r"fy|total", v, re.I):
            cur = None   # FY total block
        month_at[j] = cur
    sub = df.iloc[sub_i].astype(str).str.strip().str.lower()
    total, rlng = {}, {}
    for i in range(sub_i + 1, len(df)):
        name = str(df.iat[i, 0]).strip()
        if name == "nan" or name.lower().startswith(SKIP_ROWS):
            continue
        name = re.sub(r"\s+", " ", name.replace("?", "/"))
        for j in range(1, df.shape[1]):
            m = month_at.get(j)
            v = pd.to_numeric(df.iat[i, j], errors="coerce")
            if m is None or pd.isna(v):
                continue
            if sub.iloc[j] == "total":
                total.setdefault(m, {})[name] = v
            elif sub.iloc[j] == "rlng":
                rlng.setdefault(m, {})[name] = v
    return (pd.DataFrame.from_dict(total, orient="index").sort_index(),
            pd.DataFrame.from_dict(rlng, orient="index").sort_index())


def lng_sheet(df, fy_start):
    """Month-wise LNG import sheet (months across, rows MMT / MMSCM / value) -> MMSCM and MMT by month."""
    hdr_i = next(i for i in range(len(df)) if (df.iloc[i].astype(str).str.strip().str.lower().str[:3]
                                               .isin(["apr", "may", "jun"])).sum() >= 3)
    rows = {}
    for i in range(hdr_i + 1, len(df)):
        label = str(df.iat[i, 0]).lower()
        unit = "MMSCM" if "mmscm" in label else "MMT" if "mmt" in label else None
        if unit is None or unit in rows:
            continue
        rows[unit] = {fy_month(fy_start, str(df.iat[hdr_i, j])): pd.to_numeric(df.iat[i, j], errors="coerce")
                      for j in range(1, df.shape[1])
                      if str(df.iat[hdr_i, j]).strip().lower()[:3] in MONTHS}
    return pd.DataFrame(rows).dropna(how="all").sort_index()


def balance(session):
    """getGasConsumption per FY -> Net production, LNG import, Total consumption (MMSCM) by month."""
    rows = {}
    for fy in range(FIRST_FY, fy_now() + 1):
        r = session.post(f"{BASE}/AjaxController/getGasConsumption",
                         data={"financialYear": f"{fy}-{fy + 1}", "reportBy": 4, "pageId": 138},
                         headers={"X-Requested-With": "XMLHttpRequest", "Referer": f"{BASE}/natural-gas/consumption"},
                         timeout=T)
        try:
            res = r.json().get("result") or {}
        except ValueError:
            continue
        for v in (res.values() if isinstance(res, dict) else res):
            title = re.sub(r"<[^>]+>", "", str(v.get("title", ""))).strip()
            key = ("Net_production" if title.lower().startswith("net production") else
                   "LNG_imports" if title.lower().startswith("lng import") else
                   "Total_consumption" if title.lower().startswith("total consumption") else None)
            if not key:
                continue
            for mon in ("april", "may", "june", "july", "august", "september", "october", "november", "december",
                        "january", "february", "march"):
                val = pd.to_numeric(v.get(mon), errors="coerce")
                if pd.notna(val):
                    rows.setdefault(fy_month(fy, mon), {})[key] = val
        time.sleep(0.5)
    return pd.DataFrame.from_dict(rows, orient="index").sort_index()


def read_sheet(path, sheet):
    try:
        df = pd.read_excel(path, sheet_name=sheet, index_col=0)
    except (FileNotFoundError, ValueError):
        return pd.DataFrame()
    df.index = pd.to_datetime(df.index)
    return df.sort_index()


def merge(old, new):
    """New values win; months only in the saved workbook are kept."""
    if old.empty or new.empty:
        return (new if old.empty else old).sort_index()
    both = new.combine_first(old)
    return both[sorted(both.columns, key=lambda c: (c not in new.columns, list(both.columns).index(c)))].sort_index()


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", default=OUT)
    args = ap.parse_args()
    s = requests.Session()
    s.headers.update(H)
    stamps = []

    tot, rl = [], []
    for u in page_links(s, "sectoral-consumption"):
        if "sectoral" not in u.lower():
            continue
        try:
            xl = get_xl(s, u)
        except Exception as e:  # noqa: BLE001
            out(f"  {u}: {e}")
            continue
        stamps.append(u.rsplit("/", 1)[-1])
        for sh in xl.sheet_names:
            if "statewise" in sh.lower() or sh.strip() == "NG-H Sectoral":
                continue
            try:
                a, b = sectoral_sheet(pd.read_excel(xl, sh, header=None))
                out(f"  sectoral {sh!r}: {len(a)} months")
                tot.append(a)
                rl.append(b)
            except Exception as e:  # noqa: BLE001
                out(f"  sectoral {sh!r}: {type(e).__name__}: {e}")
    sect = pd.concat(tot) if tot else pd.DataFrame()
    sect = sect[~sect.index.duplicated(keep="last")].sort_index() if not sect.empty else sect
    rlng = pd.concat(rl) if rl else pd.DataFrame()
    rlng = rlng[~rlng.index.duplicated(keep="last")].sort_index() if not rlng.empty else rlng

    lng = []
    for u in page_links(s, "import"):
        try:
            xl = get_xl(s, u)
        except Exception as e:  # noqa: BLE001
            out(f"  {u}: {e}")
            continue
        stamps.append(u.rsplit("/", 1)[-1])
        for sh in xl.sheet_names:
            m = re.search(r"(20\d\d)-\d\d", sh)
            raw = pd.read_excel(xl, sh, header=None)
            if not m:   # current-FY file: 'Financial Year 2026-27' in the sheet
                txt = " ".join(raw.head(10).fillna("").astype(str).values.ravel())
                m = re.search(r"Financial Year\s*(20\d\d)", txt)
                if not m or "month" not in txt.lower():
                    continue
            try:
                f = lng_sheet(raw, int(m.group(1)))
                out(f"  LNG {sh!r}: {len(f)} months")
                lng.append(f)
            except StopIteration:
                pass
    lng = pd.concat(lng) if lng else pd.DataFrame()
    lng = lng[~lng.index.duplicated(keep="last")].sort_index() if not lng.empty else lng

    bal = balance(s)
    out(f"balance: {len(bal)} months")

    sheets = {}
    for name, new in (("Sectoral", sect), ("Sectoral RLNG", rlng), ("Balance", bal), ("LNG imports", lng)):
        d = merge(read_sheet(args.out, name), new)
        d = d[d.index >= f"{FIRST_FY - 3}-04-01"] if not d.empty else d
        d.index.name = "month"
        sheets[name] = d
        out(f"{name}: {len(d)} months" + (f" {d.index.min():%Y-%m}..{d.index.max():%Y-%m}" if len(d) else ""))
    if all(v.empty for v in sheets.values()):
        raise SystemExit("No PPAC data")
    notes = [
        "UNITS",
        "MMSCM (million standard cubic metres) per month, as published by PPAC (1 MMT LNG = 1,325 MMSCM). Sectoral: "
        "consumption by sector (domestic gas + RLNG); Sectoral RLNG: the RLNG (regasified LNG) part. Balance: "
        "Net_production (gross production less flaring and losses), LNG_imports, Total_consumption (= the two). "
        "LNG imports: MMSCM and MMT by month (DGCIS trade data).",
        "",
        "COVERAGE",
        "India, financial years April-March; recent months are provisional and revised. Files read this run: "
        + "; ".join(stamps),
        "",
        "SOURCE",
        "PPAC (Petroleum Planning & Analysis Cell, Ministry of Petroleum and Natural Gas): "
        "https://ppac.gov.in/natural-gas/sectoral-consumption, /natural-gas/import, /natural-gas/consumption.",
    ]
    xlsx_notes.write_workbook(args.out, sheets, notes, {"UNITS", "COVERAGE", "SOURCE"})
    out(f"Saved {args.out}")
    for k, v in sheets.items():
        if not v.empty:
            out(f"--- {k}\n{v.tail(3).T.to_string()}")


if __name__ == "__main__":
    main()
