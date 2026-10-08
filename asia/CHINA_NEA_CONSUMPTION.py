"""
Pull China's monthly electricity CONSUMPTION by sector from the National Energy Administration (NEA, nea.gov.cn) releases
"N月份全社会用电量同比增长x%" (全社会用电量, total electricity consumption of the whole society).

Each release gives, in plain text, the month and the year to date: total, primary industry (第一产业), secondary industry
(第二产业), tertiary industry (第三产业) and urban and rural residential use (城乡居民生活用电量), each with its y/y change.
Source unit 亿千瓦时 (100 million kWh) / 10 = TWh. Release finding (the list page is script-rendered; its ds_*.json holds
the whole list) is in asia/nea_common.py.

  Data      calendar month (as published; January and February are given only as a two-month total from 2024 on, and in
            some other years the January-February release has no separate February), TWh and y/y by sector
  Jan-Feb   the January-February total (dated 1 February), same columns
  YTD       the year-to-date figures as published, dated by the last month of the period (December row = full year)
  Releases  one row per release: date, title, url, status (months with no release - May 2023, May 2026 at the time of
            writing - are simply absent: nothing is derived or interpolated)

Checks per release block: the four sectors must add up to the total (+/- 4 TWh*10 = 4 亿千瓦时 of rounding).
Incremental: releases already read with status 'ok' are not fetched again, except the two latest.

    python3 asia/CHINA_NEA_CONSUMPTION.py --out "output/Data and Chart Outputs/china_nea_consumption_monthly.xlsx"
"""
import argparse
import os
import re
import sys

import pandas as pd

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import nea_common as nc  # noqa: E402
import xlsx_notes  # noqa: E402

REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DEFAULT_OUT = os.path.join(REPO_ROOT, "output", "Data and Chart Outputs", "china_nea_consumption_monthly.xlsx")
DASH = r"[-–—~～－﹣至]"
HEAD_RE = re.compile(
    rf"(?:(?P<y1>\d{{4}})年)?(?P<m>\d{{1,2}})月份?[，,][^。；]{{0,30}}?全社会用电量"   # N月份，全社会用电量 / 2026年8月，...
    rf"|1月?\s*{DASH}+\s*(?P<n>\d{{1,2}})月份?[，,][^。；]{{0,30}}?全社会用电量"   # 1-N月，全社会用电量累计 / 1月至8月，...
    rf"|(?P<y2>\d{{4}})年[，,][^。；]{{0,30}}?全社会用电量")                       # 2024年，全社会用电量 (full year)
VAL = r"([\d.]+)\s*(万亿|亿)千瓦时"
CHG = r"同比\s*(增长|下降)\s*([\d.]+)\s*%"
TOTAL_RE = re.compile(rf"全社会用电量[^。；]{{0,30}}?{VAL}\s*[，,]\s*{CHG}")
SECTORS = [("Primary", "第一产业用电量"), ("Secondary", "第二产业用电量"), ("Tertiary", "第三产业用电量"),
           ("Residential", "城乡居民生活用电量")]
SECTOR_RES = {k: re.compile(rf"{lab}\s*{VAL}(?:\s*[，,]\s*{CHG})?") for k, lab in SECTORS}
SECTOR_NOVAL = {k: re.compile(rf"{lab}\s*{CHG}") for k, lab in SECTORS}   # y/y only (value not stated)
COLS = ["Total_TWh", "Primary_TWh", "Secondary_TWh", "Tertiary_TWh", "Residential_TWh",
        "Total_YoY_pct", "Primary_YoY_pct", "Secondary_YoY_pct", "Tertiary_YoY_pct", "Residential_YoY_pct",
        "Release_Date", "Release_URL"]


def twh(num, unit):
    v = float(num) * (10000.0 if unit == "万亿" else 1.0)
    return round(v / 10.0, 1)     # 亿千瓦时 -> TWh


def signed(sign, num):
    return (1 if sign == "增长" else -1) * float(num)


def block_values(seg):
    """Total and the four sectors (TWh and y/y %) from one text segment (month or year to date)."""
    row = {}
    m = TOTAL_RE.search(seg)
    if not m:
        return None
    row["Total_TWh"] = twh(m.group(1), m.group(2))
    row["Total_YoY_pct"] = signed(m.group(3), m.group(4))
    for k, _ in SECTORS:
        s = SECTOR_RES[k].search(seg)
        if s:
            row[f"{k}_TWh"] = twh(s.group(1), s.group(2))
            if s.group(3):
                row[f"{k}_YoY_pct"] = signed(s.group(3), s.group(4))
        else:
            n = SECTOR_NOVAL[k].search(seg)
            if n:
                row[f"{k}_YoY_pct"] = signed(n.group(1), n.group(2))
    return row


def check(row):
    """None if the sectors add up to the total (rounding 0.4 TWh), else a reason."""
    parts = [row.get(f"{k}_TWh") for k, _ in SECTORS]
    if any(p is None for p in parts):
        return None if all(p is None for p in parts) else "sector value missing"
    d = sum(parts) - row["Total_TWh"]
    return None if abs(d) <= 0.45 else f"sectors sum to {sum(parts):.1f} vs total {row['Total_TWh']:.1f} TWh"


def parse(text, release_date):
    """-> list of (kind, period Timestamp, row) with kind 'month' | 'ytd'."""
    body = nc.body_of(text)
    body = re.sub(r"\s+", "", body)
    pub = pd.Timestamp(release_date)
    marks = list(HEAD_RE.finditer(body))
    out = []
    for i, m in enumerate(marks):
        seg = body[m.start(): marks[i + 1].start() if i + 1 < len(marks) else len(body)]
        row = block_values(seg)
        if row is None:
            continue
        if m.group("y2"):                                   # full year
            out.append(("ytd", pd.Timestamp(int(m.group("y2")), 12, 1), row))
        elif m.group("n"):                                  # 1-N月
            n = int(m.group("n"))
            y = pub.year if n < pub.month or (n == pub.month and pub.day > 28) else pub.year - 1
            out.append(("ytd", pd.Timestamp(y, n, 1), row))
        else:
            mo = int(m.group("m"))
            y = int(m.group("y1")) if m.group("y1") else (pub.year if mo < pub.month else pub.year - 1)
            out.append(("month", pd.Timestamp(y, mo, 1), row))
    return out


def read_sheet(path, sheet, cols):
    try:
        d = pd.read_excel(path, sheet_name=sheet, index_col=0)
        d.index = pd.to_datetime(d.index, errors="coerce")
        return d[d.index.notna()].reindex(columns=cols)
    except (FileNotFoundError, ValueError, KeyError):
        return pd.DataFrame(columns=cols)


def main():
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--out", default=DEFAULT_OUT)
    args = ap.parse_args()
    held = {s: read_sheet(args.out, s, COLS) for s in ("Data", "Jan-Feb", "YTD")}
    try:
        rel_old = pd.read_excel(args.out, sheet_name="Releases")
    except (FileNotFoundError, ValueError):
        rel_old = pd.DataFrame(columns=["Release_Date", "Title", "URL", "Status"])
    status = dict(zip(rel_old["URL"], rel_old["Status"]))
    rels = [r for r in nc.releases() if r["kind"] == "cons" and r["date"] >= "2021-01-01"]
    if not rels:
        raise SystemExit("NEA release list is empty - nothing to do.")
    latest = {r["url"] for r in rels[-2:]}
    new = {"Data": {}, "Jan-Feb": {}, "YTD": {}}
    log_rows = []
    for r in rels:
        if status.get(r["url"]) == "ok" and r["url"] not in latest:
            log_rows.append((r["date"], r["title"], r["url"], "ok"))
            continue
        html = nc.fetch(r["url"])
        res, st = [], "page not found"
        if html:
            res = parse(nc.page_text(html), r["date"])
            st = "ok" if res else "no consumption figures found"
            for kind, period, row in res:
                bad = check(row)
                if bad:
                    st = f"check failed ({kind} {period:%Y-%m}): {bad}"
                    nc.log(f"  {r['date']} {st}")
                    break
        if st == "ok":
            for kind, period, row in res:
                row = {**row, "Release_Date": r["date"], "Release_URL": r["url"]}
                if kind == "month":
                    new["Data"].setdefault(period, row)      # earliest release of a period wins
                else:
                    new["YTD"].setdefault(period, row)
                    if period.month == 2:
                        new["Jan-Feb"].setdefault(period, row)
            nc.log(f"  {r['date']} {r['title'][:30]}: " + ", ".join(f"{k} {p:%Y-%m}" for k, p, _ in res))
        else:
            nc.log(f"  {r['date']} {r['title'][:30]} -> {st}")
        log_rows.append((r["date"], r["title"], r["url"], st))
    out = {}
    for sheet in ("Data", "Jan-Feb", "YTD"):
        add = pd.DataFrame.from_dict(new[sheet], orient="index", columns=COLS) if new[sheet] else None
        d = held[sheet]
        if add is not None:
            d = add.combine_first(d) if len(d) else add
        d = d.reindex(columns=COLS).sort_index()
        d.index.name = "month"
        out[sheet] = d
    if out["Data"].empty:
        raise SystemExit("No monthly data at all - nothing to save.")
    out["Releases"] = pd.DataFrame(log_rows, columns=["Release_Date", "Title", "URL", "Status"]).set_index("Release_Date")
    labels = {"Primary": "Primary industry", "Secondary": "Secondary industry", "Tertiary": "Tertiary industry",
              "Residential": "Urban and rural residential"}
    ser = [(f"{k}_TWh", labels[k], "TWh per month", "electricity consumption by sector", "stacked_bar") for k, _ in SECTORS]
    ser += [("Total_YoY_pct", "Total", "% y/y", "electricity consumption growth y/y", "line")]
    ser += [(f"{k}_YoY_pct", labels[k], "% y/y", "electricity consumption growth y/y", "line") for k, _ in SECTORS]
    out["Series"] = pd.DataFrame(ser, columns=["column", "label", "unit", "chart", "kind"]).set_index("column")
    os.makedirs(os.path.dirname(os.path.abspath(args.out)), exist_ok=True)
    xlsx_notes.write_workbook(args.out, out, NOTES_LINES, NOTES_SECTION_TITLES)
    d = out["Data"]
    miss = [f"{p:%Y-%m}" for p in pd.date_range(d.index.min(), d.index.max(), freq="MS") if p not in d.index]
    print(f"Saved {len(d)} month(s) ({d.index.min():%Y-%m}..{d.index.max():%Y-%m}); months without a monthly figure: {miss}; "
          f"{len(out['Jan-Feb'])} Jan-Feb, {len(out['YTD'])} YTD rows to {args.out}")


NOTES_LINES = [
    "UNITS",
    "*_TWh: the calendar month's electricity consumption of the whole society (全社会用电量) by sector, TWh (source unit "
    "亿千瓦时, 100 million kWh, / 10). Sectors: primary (第一产业), secondary (第二产业, incl. industry), tertiary (第三产业), "
    "urban and rural residential (城乡居民生活用电量); they add up to the total within rounding (checked on every release). "
    "*_YoY_pct: the release's own y/y % change. Data = months NEA publishes a monthly figure for: January and February are "
    "given only as a two-month total from 2024 on (and in other years only February is published separately), so those "
    "months are gaps - the 'Jan-Feb' sheet holds the two-month total, dated 1 February. The YTD sheet holds the year-to-date "
    "figures as published (December row = full year). NOTHING is derived from year-to-date differences or interpolated.",
    "",
    "SOURCE",
    "National Energy Administration (国家能源局), monthly releases 'N月份全社会用电量同比增长x%' on https://www.nea.gov.cn/ "
    "(press list https://www.nea.gov.cn/xwfb/ and the 综合司 column https://www.nea.gov.cn/sjzz/ghs/). Whole-society "
    "consumption (全口径), the series NBS does not publish by sector monthly. Missing months in the release list (e.g. "
    "May 2023, May 2026 when last checked) are shown on the 'Releases' sheet by their absence.",
    "",
    "UPDATES",
    "Incremental: releases already read are not fetched again, except the latest two. The list page is script-rendered; "
    "the pull reads the ds_*.json files behind it (asia/nea_common.py).",
]
NOTES_SECTION_TITLES = {"UNITS", "SOURCE", "UPDATES"}

if __name__ == "__main__":
    main()
