"""
Taiwan Energy Administration (MOEA) Monthly Energy Statistics via the E-STAT open API (ESIST, ea01.moeaea.gov.tw/a0303/02,
swagger at /api/pages/database/api, keyless, reachable from GitHub Actions) -> output/Data and Chart Outputs/
taiwan_esist_monthly.xlsx.

One pair of sheets per table: '<tag> M' (monthly, index = first day of the month) and '<tag> A' (annual, index = year).
The API serves the last ~24 months and annual rows back to 2007 in every call; the committed workbook is the history store:
rows of the current response replace stored rows of the same period, older stored rows are kept, so the monthly history
grows by a month at every release (the Energy Administration publishes about the 2nd of each month, two months in arrears).
'Series' lists every column with its unit; 'Releases' the report date seen at each run.

Tables (API path /v1/zone/monthly/<n>/<m>, section):
  GEN   3-02 electricity generation, nationwide, by fuel (GWh)         CAP   3-03 installed capacity, nationwide (MW)
  CONS  3-04 electricity consumption by sector (GWh)                   REN   4-01 renewable generation (MWh)
  GAS   6-01 natural gas supply and consumption (thousand m3)          LNG   6-02 LNG imports by origin (thousand tonnes)
  SUP   2-02 primary energy supply by fuel (thousand tonnes oil eq.)   CRUDE 5-01 crude oil supply and refining (kTOE)
  CRUDESRC 5-02 crude imports by origin (thousand barrels)             COAL  7-01 coal supply and consumption (kt)
  COALSRC 7-02 coal imports by origin (kt)                             IMPPRICE 9-02 import prices (crude, LNG, coal; US$)
  OILPRICE 9-03 WTI, Brent, Dubai monthly (US$/barrel)

    python3 taiwan/TAIWAN_ESIST_MONTHLY.py --out "output/Data and Chart Outputs/taiwan_esist_monthly.xlsx"
"""
import argparse
import os
import re
import sys

import pandas as pd
import requests

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import xlsx_notes  # noqa: E402

API = "https://ea01.moeaea.gov.tw/a0303/02/api"
UA = {"User-Agent": "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 Chrome/124 Safari/537.36"}
DEFAULT_OUT = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "output",
                           "Data and Chart Outputs", "taiwan_esist_monthly.xlsx")
# tag, path, section, unit, description
TABLES = [
    ("GEN", "3/2", "全國", "GWh", "3-02 Electricity generation, nationwide (Taipower + IPPs + self-generation)"),
    ("CAP", "3/3", "全國", "MW", "3-03 Installed generating capacity, nationwide (end of period)"),
    ("CONS", "3/4", "電力消費", "GWh", "3-04 Electricity consumption by sector"),
    ("REN", "4/1", "再生能源", "MWh", "4-01 Renewable electricity generation"),
    ("GAS", "6/1", "天然氣", "thousand m3", "6-01 Natural gas supply and consumption"),
    ("LNG", "6/2", "LNG進口來源", "thousand tonnes", "6-02 LNG imports by origin"),
    ("SUP", "2/2", "總供給", "thousand tonnes oil equivalent", "2-02 Energy supply by fuel"),
    ("CRUDE", "5/1", "原油供給與煉製", "thousand tonnes oil equivalent", "5-01 Crude oil supply and refining"),
    ("CRUDESRC", "5/2", "原油進口來源", "thousand barrels", "5-02 Crude oil imports by origin"),
    ("COAL", "7/1", "煤炭", "thousand tonnes", "7-01 Coal supply and consumption"),
    ("COALSRC", "7/2", "煤炭", "thousand tonnes", "7-02 Coal imports by origin"),
    ("IMPPRICE", "9/2", "平均價格", "US$ (crude per barrel, LNG and coal per tonne)", "9-02 Imported energy prices"),
    ("OILPRICE", "9/3", "油價", "US$ per barrel", "9-03 International crude oil prices"),
]
CJK = re.compile(r"[　-鿿＀-￯]+")
PERIOD_M = re.compile(r"^(\d{4})/(\d{2})$")
PERIOD_A = re.compile(r"^(\d{4})$")


def get_json(url):
    last = None
    for i in range(3):
        try:
            r = requests.get(url, headers=UA, timeout=(15, 90))
            r.raise_for_status()
            return r.json()
        except (requests.RequestException, ValueError) as e:
            last = e
            print(f"  attempt {i + 1}/3 {url}: {type(e).__name__}", file=sys.stderr)
    raise last


def english(text):
    t = CJK.sub(" ", str(text if text is not None else "")).replace("\n", " ")
    t = re.sub(r"\s+", " ", t).strip(" /:：")
    return t if re.search(r"[A-Za-z]", t) else ""


def num(v):
    try:
        x = float(str(v).replace(",", ""))
        return x
    except ValueError:
        return float("nan")


def parse_section(rows):
    """Rows of one section -> (monthly frame, annual frame). Column names from the header rows (English text)."""
    first = list(rows[0].keys())[0]
    numeric = sorted({k for r in rows for k in r if k != first}, key=lambda s: int(re.sub(r"\D", "", s) or 0))
    last = numeric[-1]
    data_i = [i for i, r in enumerate(rows) if PERIOD_M.match(str(r.get(last, "")).strip()) or PERIOD_A.match(str(r.get(last, "")).strip())]
    if not data_i:
        raise ValueError("no period rows")
    header = [r for r in rows[:data_i[0]] if sum(1 for k in r if k != first) > 0]
    share = {k for r in header for k, v in r.items() if k != first and str(v).strip() in ("占比", "(% )")}
    allk = sorted({k for r in rows for k in r if k != first}, key=lambda s: int(re.sub(r"\D", "", s) or 0))
    cols = [k for k in allk if k != last and k not in share]
    # row 0 of the header is the group row (forward filled); later rows are sub-labels (not filled)
    unit_rows = [r for r in header if any(str(v).strip() in ("占比", "(% )") or re.match(r"^\(.*\)$", str(v).strip()) for v in r.values())]
    label_rows = [r for r in header if r not in unit_rows]
    group, g = {}, ""
    for k in allk:
        v = english(label_rows[0].get(k, "")) if label_rows else ""
        if v:
            g = v
        group[k] = g
    names = {}
    for k in cols:
        subs = [english(r.get(k, "")) for r in label_rows[1:]]
        subs = [s for s in subs if s]
        sub = subs[-1] if subs else ""
        grp = group[k]
        if grp and sub and sub.lower() != grp.lower():
            nm = f"{grp} - {sub}"
        else:
            nm = sub or grp
        names[k] = nm or k
    # make names unique
    seen = {}
    for k in cols:
        n = names[k]
        seen[n] = seen.get(n, 0) + 1
        if seen[n] > 1:
            names[k] = f"{n} ({seen[n]})"
    monthly, annual = {}, {}
    for i in data_i:
        r = rows[i]
        p = str(r.get(last, "")).strip()
        vals = {names[k]: num(r.get(k)) for k in cols}
        m = PERIOD_M.match(p)
        if m:
            monthly[pd.Timestamp(int(m.group(1)), int(m.group(2)), 1)] = vals
        else:
            annual[int(p)] = vals
    mf = pd.DataFrame.from_dict(monthly, orient="index").sort_index().dropna(how="all")
    af = pd.DataFrame.from_dict(annual, orient="index").sort_index().dropna(how="all")
    mf.index.name, af.index.name = "month", "year"
    return mf, af


def read_old(path, sheet, dates):
    if not os.path.exists(path):
        return None
    try:
        d = pd.read_excel(path, sheet_name=sheet, index_col=0)
        d.index = pd.to_datetime(d.index) if dates else d.index.astype(int)
        return d
    except Exception as e:  # noqa: BLE001
        print(f"  stored sheet {sheet} unreadable ({type(e).__name__})", file=sys.stderr)
        return None


def merge(old, new):
    if old is None or old.empty:
        return new
    keep = old[~old.index.isin(new.index)]
    out = pd.concat([keep, new]).sort_index()
    out.index.name = new.index.name
    return out


def roc_date(s):
    m = re.match(r"^(\d+)\.(\d+)\.(\d+)$", str(s))
    return f"{int(m.group(1)) + 1911}-{m.group(2)}-{m.group(3)}" if m else str(s)


def main():
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--out", default=DEFAULT_OUT)
    args = ap.parse_args()
    sheets, series, problems = {}, [], []
    try:
        rel = get_json(API + "/pages/newest/monthly")
        report = roc_date(rel.get("date")), roc_date(rel.get("nextDate"))
    except Exception as e:  # noqa: BLE001
        report = ("unknown", "unknown")
        print(f"  release date unavailable ({type(e).__name__})", file=sys.stderr)
    for tag, path, section, unit, desc in TABLES:
        om, oa = read_old(args.out, f"{tag} M", True), read_old(args.out, f"{tag} A", False)
        mf, af = om, oa
        try:
            j = get_json(f"{API}/v1/zone/monthly/{path}")
            new_m, new_a = parse_section(j[section])
            mf, af = merge(om, new_m), merge(oa, new_a)
            print(f"  {tag}: {len(new_m)} monthly / {len(new_a)} annual rows; {len(new_m.columns)} columns; latest "
                  f"{new_m.index.max():%Y-%m}")
        except Exception as e:  # noqa: BLE001
            problems.append(f"{tag} ({path}): {type(e).__name__}: {e}")
            print(f"  {tag} FAILED ({type(e).__name__}: {e}); keeping stored rows", file=sys.stderr)
        for nm, fr in ((f"{tag} M", mf), (f"{tag} A", af)):
            if fr is not None and not fr.empty:
                sheets[nm] = fr
        if mf is not None:
            for c in mf.columns:
                series.append({"table": tag, "column": c, "unit": unit, "description": desc, "api_path": f"/v1/zone/monthly/{path}",
                               "section": section})
    if not sheets:
        raise SystemExit("nothing pulled and nothing stored")
    old_rel = pd.read_excel(args.out, sheet_name="Releases") if os.path.exists(args.out) and "Releases" in pd.ExcelFile(args.out).sheet_names else pd.DataFrame()
    rel_df = pd.concat([old_rel, pd.DataFrame([{"report_date": report[0], "next_report": report[1],
                                                "problems": "; ".join(problems)}])]).drop_duplicates("report_date", keep="last")
    sheets["Series"] = pd.DataFrame(series).drop_duplicates(["table", "column"]).set_index("table")
    sheets["Releases"] = rel_df.set_index("report_date")
    notes = list(NOTES) + [""] + [f"Latest report seen: {report[0]} (next {report[1]})."] + ([f"Problems in this run: {'; '.join(problems)}"] if problems else [])
    os.makedirs(os.path.dirname(os.path.abspath(args.out)), exist_ok=True)
    xlsx_notes.write_workbook(args.out, sheets, notes, {"UNITS", "SOURCE", "UPDATES"})
    print(f"Saved {len(sheets)} sheets -> {args.out}")


NOTES = [
    "UNITS",
    "Per table, see the Series sheet: generation GWh (source 'million kWh'), capacity MW (source 'thousand kW'), consumption GWh, "
    "renewable generation MWh (source 'thousand kWh'), natural gas thousand m3, LNG thousand tonnes, crude thousand barrels, coal "
    "thousand tonnes, energy supply thousand tonnes oil equivalent, prices US$. Share (%) columns of the source are dropped. "
    "'M' sheets are monthly (index = first day of the month), 'A' sheets annual (index = year).",
    "",
    "SOURCE",
    "Energy Administration, Ministry of Economic Affairs (Taiwan): Monthly Energy Statistics (E-STAT / ESIST open API, "
    "https://ea01.moeaea.gov.tw/a0303/02/ , swagger /a0303/02/api/pages/database/api). Published about the 2nd of each month "
    "with a two-month lag; the API returns about 24 months and the annual rows from 2007.",
    "",
    "UPDATES",
    "Incremental: stored rows are kept, rows of the current response replace stored rows of the same period (revisions). The "
    "monthly history therefore starts about 24 months before the first run and grows by one month per release. Gaps are not filled.",
]

if __name__ == "__main__":
    main()
