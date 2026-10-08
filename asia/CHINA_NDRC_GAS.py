"""
Pull China's monthly natural gas APPARENT CONSUMPTION from the National Development and Reform Commission (NDRC)
"national natural gas operation bulletin" (全国天然气运行快报, issued by NDRC's Operation Bureau, list page
https://www.ndrc.gov.cn/fggz/jjyxtj/). Found reachable from GitHub Actions by discovery_archive/china/CHINA_SOURCES_PROBE*.py.

Each bulletin's body is one or two sentences: the month's apparent consumption (production + imports, 亿立方米) with
its y/y change, and the year-to-date (or full-year, in the December bulletin) figure with its y/y change. January and
February are published only combined ("1—2月份"), so they go on the 'Jan-Feb' sheet, as in the NBS workbooks. The
bulletins also carry a ranking attachment (排行榜) that is not read. The list reaches back to April 2023 (the pages
behind that no longer hold gas bulletins).

NDRC publishes consumption only: there is no production, import or sector split in the bulletins, so none is invented.

Incremental: months already in the workbook are not re-fetched except the last two (a bulletin can be revised).

    python3 asia/CHINA_NDRC_GAS.py --out "output/Data and Chart Outputs/china_ndrc_gas_monthly.xlsx"
"""
import argparse
import os
import re
import sys
import time

import pandas as pd
import requests

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import xlsx_notes  # noqa: E402

REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DEFAULT_OUT = os.path.join(REPO_ROOT, "output", "Data and Chart Outputs", "china_ndrc_gas_monthly.xlsx")
BASE = "https://www.ndrc.gov.cn/fggz/jjyxtj/"
PAGES = ["", *[f"index_{i}.html" for i in range(1, 12)]]
HEADERS = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) "
                         "Chrome/124.0 Safari/537.36", "Accept-Language": "zh-CN,zh;q=0.9,en;q=0.8"}
SESSION = requests.Session()
SESSION.headers.update(HEADERS)
TITLE_RE = re.compile(r"(\d{4})年(\d{1,2})(?:[—\-～~](\d{1,2}))?月份全国天然气运行快报")
MONTH_RE = re.compile(r"(\d{4})年(\d{1,2})(?:[—\-～~](\d{1,2}))?月[，,]\s*全国天然气表观消费量\s*([\d.]+)\s*亿立方米[，,]\s*"
                      r"同比(增长|下降)\s*([\d.]+)\s*%")
YTD_RE = re.compile(r"(?:(\d{4})年|1[—\-～~](\d{1,2})月)[，,]\s*全国天然气表观消费量\s*([\d.]+)\s*亿立方米[，,]\s*"
                    r"同比(增长|下降)\s*([\d.]+)\s*%")
COLS = ["Apparent_Consumption_Bcm", "Consumption_YoY_pct", "YTD_Consumption_Bcm", "YTD_YoY_pct"]
last_request = [0.0]


def log(msg):
    print(msg, file=sys.stderr, flush=True)


def fetch(url, attempts=4):
    err = None
    for attempt in range(attempts):
        time.sleep(max(0.0, 1.0 - (time.time() - last_request[0])))
        try:
            r = SESSION.get(url, timeout=(10, 30))
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
        r.encoding = "utf-8"
        return r.text
    raise err


def bulletins():
    """[(title, absolute url)] from the list pages (newest first), de-duplicated."""
    out, seen = [], set()
    for page in PAGES:
        html = fetch(BASE + page)
        if html is None:
            break
        found = 0
        for href, text in re.findall(r'<a[^>]+href="([^"]+)"[^>]*>([^<]{3,90})</a>', html):
            title = text.strip()
            if TITLE_RE.search(title):
                found += 1
                url = requests.compat.urljoin(BASE + page, href)
                if url not in seen:
                    seen.add(url)
                    out.append((title, url))
        if not found and page:
            break
    return out


def body_text(html):
    """Page text with every tag and all whitespace removed (some bulletins split a sentence over several tags)."""
    html = re.sub(r"(?s)<(script|style).*?</\1>", "", html)
    t = re.sub(r"<[^>]+>", "", html).replace("&nbsp;", "").replace("&emsp;", "")
    return re.sub(r"[\s\u3000\xa0]+", "", t)


def parse(text):
    """-> (period start Timestamp, combined JanFeb flag, row dict) from a bulletin's text, or None."""
    m = MONTH_RE.search(text)
    if not m:
        return None
    year, m1, m2 = int(m.group(1)), int(m.group(2)), m.group(3)
    sign = 1 if m.group(5) == "增长" else -1
    row = {"Apparent_Consumption_Bcm": round(float(m.group(4)) / 10.0, 2),
           "Consumption_YoY_pct": sign * float(m.group(6))}
    rest = text[m.end():]
    y = YTD_RE.search(rest)
    if y:
        row["YTD_Consumption_Bcm"] = round(float(y.group(3)) / 10.0, 2)
        row["YTD_YoY_pct"] = (1 if y.group(4) == "增长" else -1) * float(y.group(5))
    jan_feb = m2 is not None
    if jan_feb and (m1, int(m2)) != (1, 2):
        return None
    return pd.Timestamp(year, 2 if jan_feb else m1, 1), jan_feb, row


def main():
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--out", default=DEFAULT_OUT)
    args = ap.parse_args()
    held = {}
    for sheet in ("Data", "Jan-Feb"):
        try:
            d = pd.read_excel(args.out, sheet_name=sheet, index_col=0)
            d.index = pd.to_datetime(d.index, errors="coerce")
            held[sheet] = d[d.index.notna()].reindex(columns=COLS)
        except (FileNotFoundError, ValueError, KeyError):
            held[sheet] = pd.DataFrame(columns=COLS, dtype=float)
    have = set(held["Data"].index) | set(held["Jan-Feb"].index)
    redo = set(sorted(have)[-2:])   # the two latest periods are re-read in case of a revision
    new = {"Data": {}, "Jan-Feb": {}}
    for title, url in bulletins():
        t = TITLE_RE.search(title)
        period = pd.Timestamp(int(t.group(1)), 2 if t.group(3) else int(t.group(2)), 1)
        if period in have and period not in redo:
            continue
        html = fetch(url)
        text = body_text(html or "")
        parsed = parse(text)
        if parsed is None or parsed[0] != period:
            i = text.find("全国天然气表观消费量")
            log(f"  [{period:%Y-%m}] not parsed - skipped ({url}): ...{text[max(0, i - 40):i + 80]}")
            continue
        _, jan_feb, row = parsed
        new["Jan-Feb" if jan_feb else "Data"][period] = row
        log(f"  [{period:%Y-%m}{' Jan-Feb' if jan_feb else ''}] {row}")
    out = {}
    for sheet in ("Data", "Jan-Feb"):
        add = pd.DataFrame.from_dict(new[sheet], orient="index", columns=COLS) if new[sheet] else None
        d = held[sheet]
        if add is not None:
            d = add.combine_first(d) if len(d) else add     # new values win
        d = d.reindex(columns=COLS).sort_index()
        d.index.name = "month"
        out[sheet] = d
    if out["Data"].empty and out["Jan-Feb"].empty:
        raise SystemExit("No data at all - nothing to save.")
    series = pd.DataFrame([
        ("Apparent_Consumption_Bcm", "Natural gas apparent consumption", "bcm per month",
         "natural gas apparent consumption", "line"),
        ("Consumption_YoY_pct", "Natural gas apparent consumption, y/y", "% y/y",
         "natural gas apparent consumption y/y", "line")],
        columns=["column", "label", "unit", "chart", "kind"]).set_index("column")
    out["Series"] = series
    os.makedirs(os.path.dirname(os.path.abspath(args.out)), exist_ok=True)
    xlsx_notes.write_workbook(args.out, out, NOTES_LINES, NOTES_SECTION_TITLES)
    d = out["Data"]
    print(f"Saved {len(d)} month(s) ({d.index.min():%Y-%m}..{d.index.max():%Y-%m}) + {len(out['Jan-Feb'])} Jan-Feb "
          f"to {args.out}")


NOTES_LINES = [
    "UNITS",
    "Apparent_Consumption_Bcm: the calendar month's apparent natural gas consumption (production + imports), bcm "
    "(source unit 亿立方米, 100 million m3, / 10). Consumption_YoY_pct: the bulletin's % change on the same month a year "
    "earlier. YTD_Consumption_Bcm / YTD_YoY_pct: the year-to-date figure and its y/y change (the December bulletin "
    "gives the full year). January and February are published only combined: the 'Jan-Feb' sheet holds the two-month "
    "total, dated 1 February; the Data sheet has no January or February row.",
    "",
    "SOURCE",
    "National Development and Reform Commission, Operation Bureau, 全国天然气运行快报 (national natural gas operation "
    "bulletin), https://www.ndrc.gov.cn/fggz/jjyxtj/. Consumption only: the bulletins give no production, import "
    "or sector split.",
    "",
    "UPDATES",
    "Incremental: months already held are not re-fetched, except the latest two. The list holds bulletins from April 2023.",
]
NOTES_SECTION_TITLES = {"UNITS", "SOURCE", "UPDATES"}

if __name__ == "__main__":
    main()
