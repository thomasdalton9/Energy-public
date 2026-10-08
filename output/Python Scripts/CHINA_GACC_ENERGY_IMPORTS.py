"""
Pull China's monthly imports of coal, crude oil, refined products and natural gas (total and liquefied) from the General
Administration of Customs (GACC) English-language monthly bulletin, table "(14) Major Import Commodities in Quantity and
Value" (http://english.customs.gov.cn/Statistics/Statistics?ColumnId=2). The English site answers from GitHub Actions
(the Chinese site customs.gov.cn returns a JavaScript challenge / TLS error and stats.customs.gov.cn the same; found by
discovery_archive/china/CHINA_SOURCES_PROBE*.py).

Each month's table has the month's quantity (10,000 tonnes) and value (US$ 1,000) per commodity plus the year-to-date
figures. Quantities are weights: natural gas is in tonnes, not cubic metres, and no conversion to bcm is made here.
January and February are published separately from 2021 on (the table for Jan-Feb 2020 was the last combined one).
Pipeline and other non-liquefied gas is shown as "total less LNG": both rows are published, the difference is not.

History: the list holds the table from Sep 2019; this pull keeps Jan 2021 on (the NBS workbooks start in 2021).
Incremental: months already in the workbook are not re-fetched except the latest two; the list is read from the newest
page until two consecutive known months are met.

    python3 asia/CHINA_GACC_ENERGY_IMPORTS.py --out "output/Data and Chart Outputs/china_gacc_energy_imports_monthly.xlsx"
"""
import argparse
import html as htmllib
import os
import re
import sys
import time

import pandas as pd
import requests

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import xlsx_notes  # noqa: E402

REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DEFAULT_OUT = os.path.join(REPO_ROOT, "output", "Data and Chart Outputs", "china_gacc_energy_imports_monthly.xlsx")
SITE = "http://english.customs.gov.cn"
LIST = SITE + "/Statistics/Statistics?ColumnId=2&page={}"
HISTORY_START = pd.Timestamp("2021-01-01")
MAX_PAGES = 220
HEADERS = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) "
                         "Chrome/124.0 Safari/537.36", "Accept-Language": "en-US,en;q=0.9"}
SESSION = requests.Session()
SESSION.headers.update(HEADERS)
TITLE_RE = re.compile(r"Major\s+Import\s+Commodities", re.I)
PERIOD_RE = re.compile(r"Value\s*,?\s*(\d{1,2})(?:\s*-\s*(\d{1,2}))?\s*\.\s*(\d{4})")
# (row name in the table, column stem, label)
ROWS = [("Coal and lignite", "Coal", "Coal and lignite"),
        ("Crude petroleum oils", "Crude_Oil", "Crude oil"),
        ("Refined petroleum products", "Refined_Products", "Refined petroleum products"),
        ("Natural gases", "Natural_Gas", "Natural gas (total)"),
        ("Natural gases in liquefied state", "LNG", "LNG")]
last_request = [0.0]


def log(msg):
    print(msg, file=sys.stderr, flush=True)


def fetch(url, attempts=4):
    err = None
    for attempt in range(attempts):
        time.sleep(max(0.0, 1.0 - (time.time() - last_request[0])))
        try:
            r = SESSION.get(url, timeout=(10, 40))
            last_request[0] = time.time()
        except requests.RequestException as e:
            err = e
            time.sleep(8 * (attempt + 1))
            continue
        if r.status_code == 404:
            return None
        if r.status_code != 200:
            err = requests.HTTPError(f"{r.status_code} {url}")
            time.sleep(8 * (attempt + 1))
            continue
        r.encoding = r.encoding if r.encoding and r.encoding.lower() != "iso-8859-1" else "utf-8"
        return r.text
    raise err


def clean(cell):
    return re.sub(r"\s+", " ", htmllib.unescape(re.sub(r"<[^>]+>", "", cell)).replace("\xa0", " ")).strip()


def num(s):
    s = s.replace(",", "").strip()
    try:
        return float(s)
    except ValueError:
        return None


def period_of(title):
    m = PERIOD_RE.search(title)
    if not m or m.group(2):
        return None   # combined periods are not used (none from 2021)
    return pd.Timestamp(int(m.group(3)), int(m.group(1)), 1)


def listing(held, redo):
    """(period, url) of table (14) for months not yet held (newest first); stops after two known months in a row."""
    out, known_run = [], 0
    for page in range(1, MAX_PAGES + 1):
        html = fetch(LIST.format(page))
        items = re.findall(r'<a[^>]+href="([^"]*Statics[^"]*)"[^>]*>([^<]+)</a>', html or "")
        if not items:
            break
        done = False
        for href, title in items:
            title = htmllib.unescape(title)
            if not TITLE_RE.search(title):
                continue
            p = period_of(title)
            if p is None:
                continue
            if p < HISTORY_START:
                done = True
                break
            if p in held and p not in redo:
                known_run += 1
                if known_run >= 2:
                    done = True
                    break
                continue
            known_run = 0
            out.append((p, requests.compat.urljoin(SITE + "/", href)))
        if done:
            break
    return out


def parse_table(html):
    """{column: value} from table (14): month quantity (10,000 t) -> Mt, month value (US$ 1,000) -> US$ million, and the
    year-to-date quantity (kept only for the check)."""
    rows = {}
    for tr in re.findall(r"(?s)<tr[^>]*>(.*?)</tr>", html):
        cells = [clean(c) for c in re.findall(r"(?s)<t[dh][^>]*>(.*?)</t[dh]>", tr)]
        if len(cells) >= 6 and re.search(r"10000\s*T", cells[1], re.I):
            rows.setdefault(cells[0].lower(), cells)
    out, ytd = {}, {}
    for name, stem, _ in ROWS:
        c = rows.get(name.lower())
        if c is None:
            continue
        q, v, yq = num(c[2]), num(c[3]), num(c[4])
        if q is not None:
            out[f"{stem}_Mt"] = round(q / 100.0, 2)
        if v is not None:
            out[f"{stem}_USD_m"] = round(v / 1000.0, 1)
        ytd[stem] = yq
    if "Natural_Gas_Mt" in out and "LNG_Mt" in out:
        out["Natural_Gas_Other_Than_LNG_Mt"] = round(out["Natural_Gas_Mt"] - out["LNG_Mt"], 2)
    return out, ytd


def columns():
    cols = []
    for _, stem, _ in ROWS:
        cols += [f"{stem}_Mt", f"{stem}_USD_m"]
    return cols + ["Natural_Gas_Other_Than_LNG_Mt"]


def main():
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--out", default=DEFAULT_OUT)
    args = ap.parse_args()
    cols = columns()
    try:
        data = pd.read_excel(args.out, sheet_name="Data", index_col=0)
        data.index = pd.to_datetime(data.index, errors="coerce")
        data = data[data.index.notna()].reindex(columns=cols)
    except (FileNotFoundError, ValueError, KeyError):
        data = pd.DataFrame(columns=cols, dtype=float)
    held = set(data.index)
    redo = set(sorted(held)[-2:])
    todo = listing(held, redo)
    log(f"  held {len(held)} months; {len(todo)} table(s) to read")
    new, prev_ytd = {}, {}
    for p, url in sorted(todo):
        try:
            html = fetch(url)
        except Exception as e:  # noqa: BLE001 - next run retries
            log(f"  [{p:%Y-%m}] fetch failed ({type(e).__name__}) - skipped")
            continue
        row, ytd = parse_table(html or "")
        if "Natural_Gas_Mt" not in row or "Crude_Oil_Mt" not in row:
            log(f"  [{p:%Y-%m}] energy rows not found - skipped ({url})")
            continue
        new[p] = row
        prev_ytd[p] = ytd
        log(f"  [{p:%Y-%m}] gas {row.get('Natural_Gas_Mt')} Mt (LNG {row.get('LNG_Mt')}), crude {row.get('Crude_Oil_Mt')} Mt, "
            f"coal {row.get('Coal_Mt')} Mt")
    if new:
        add = pd.DataFrame.from_dict(new, orient="index").reindex(columns=cols)
        data = add.combine_first(data) if len(data) else add
    data = data.reindex(columns=cols).sort_index()
    data.index.name = "month"
    if data.dropna(how="all").empty:
        raise SystemExit("No data at all - nothing to save.")
    # check: each month's quantity vs the change in the bulletin's own year-to-date quantity
    for p in sorted(prev_ytd):
        q = p - pd.DateOffset(months=1)
        if p.month > 1 and q in prev_ytd:
            for stem, label in (("Natural_Gas", "gas"), ("Crude_Oil", "crude"), ("Coal", "coal")):
                a, b = prev_ytd[p].get(stem), prev_ytd[q].get(stem)
                m = new[p].get(f"{stem}_Mt")
                if None not in (a, b, m) and abs((a - b) / 100.0 - m) > 0.03:
                    log(f"  CHECK [{p:%Y-%m}] {label}: month {m} Mt vs YTD change {(a - b) / 100.0:.2f} Mt")
    series = pd.DataFrame([
        ("LNG_Mt", "LNG", "Mt per month", "natural gas imports", "stacked_bar"),
        ("Natural_Gas_Other_Than_LNG_Mt", "Pipeline and other non-liquefied gas (total less LNG)", "Mt per month",
         "natural gas imports", "stacked_bar"),
        ("Crude_Oil_Mt", "Crude oil", "Mt per month", "crude oil and product imports", "line"),
        ("Refined_Products_Mt", "Refined petroleum products", "Mt per month", "crude oil and product imports",
         "line"),
        ("Coal_Mt", "Coal and lignite", "Mt per month", "coal imports", "line")],
        columns=["column", "label", "unit", "chart", "kind"]).set_index("column")
    os.makedirs(os.path.dirname(os.path.abspath(args.out)), exist_ok=True)
    xlsx_notes.write_workbook(args.out, {"Data": data, "Series": series}, NOTES_LINES, NOTES_SECTION_TITLES)
    print(f"Saved {len(data)} month(s) ({data.index.min():%Y-%m}..{data.index.max():%Y-%m}), {len(new)} read, to {args.out}")


NOTES_LINES = [
    "UNITS",
    "*_Mt: the calendar month's import quantity, million tonnes (source unit 10,000 tonnes, / 100; GACC rounds to "
    "10,000 t). Natural gas is a weight, not a volume: no conversion to cubic metres is applied. *_USD_m: import value, "
    "US$ million (source US$ 1,000, / 1000). Natural_Gas_Other_Than_LNG_Mt = 'Natural gases' less 'Natural gases in "
    "liquefied state' (both published rows; the remainder is pipeline gas).",
    "",
    "SOURCE",
    "General Administration of Customs of China (GACC), monthly bulletin, table (14) Major Import Commodities in "
    "Quantity and Value, English site http://english.customs.gov.cn/Statistics/Statistics?ColumnId=2.",
    "",
    "UPDATES",
    "Incremental: months already held are not re-fetched except the latest two. History kept from January 2021 "
    "(the site lists the table from September 2019).",
]
NOTES_SECTION_TITLES = {"UNITS", "SOURCE", "UPDATES"}

if __name__ == "__main__":
    main()
