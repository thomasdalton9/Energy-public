"""
Vietnam electricity generation by source from EVN's daily national power-system operation summaries
("Thong tin chung ve van hanh he thong dien Quoc gia ngay DD/MM/YYYY"), published by EVN from NSMO (the national
power system and market operator) data: https://www.evn.com.vn/vi-VN/news-l/Thong-tin-tom-tat-van-hanh-HTD-Quoc-gia-60-2015
Found via discovery_archive/asia/SCADA_DISCOVERY1-3.py (nsmo.vn itself does not answer from GitHub's runners).

Each day's post gives, in million kWh (= GWh): hydro, coal, gas turbines (gas + DO oil), oil-fired thermal, wind,
farm solar, rooftop solar (two estimates: delivered / at the generator terminals), imports and other (biomass,
southern diesel); the day's total production + imports and peak demand (MW); and the dispatch at the midday low and
the evening peak by source (MW). The list is paged with ?page=N, newest first.

Writes output/Data and Chart Outputs/vietnam_power_generation_daily.xlsx:
  Daily    standard layout, MWh per day: Hydro, Coal, Gas, Oil, Wind, Solar (farm + rooftop at the terminals), Other;
           Total_MWh = their sum (domestic generation); Imports_MWh; Solar_farm_MWh, Solar_rooftop_MWh,
           Solar_rooftop_delivered_MWh; Total_incl_imports_MWh as EVN states it
  Demand   Demand_peak_MW (day's maximum, rooftop at the terminals), Midday_MW / Evening_MW (dispatch at the midday
           low and the evening peak)

Incremental: the Daily sheet is the history store; only posts for days not saved yet (plus REVISION_DAYS) are
downloaded. The list (~120 pages, from May 2023) is read back to the earliest unsaved day, so gaps are refilled. Runs on the 1st and 15th.

    python3 asia/VIETNAM_EVN.py
"""
import argparse
import html
import os
import re
import sys
import time
from concurrent.futures import ThreadPoolExecutor
from datetime import date, timedelta

import pandas as pd
import requests

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)
import xlsx_notes  # noqa: E402

BASE = "https://www.evn.com.vn"
LIST = BASE + "/vi-VN/news-l/Thong-tin-tom-tat-van-hanh-HTD-Quoc-gia-60-2015"
H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 "
                   "Safari/537.36", "Accept-Language": "vi-VN,vi;q=0.9,en;q=0.8"}
T = (20, 90)
DATA_START = date(2021, 1, 1)
REVISION_DAYS = 18   # runs are 14-17 days apart: re-read everything since the last run, plus spare
MAX_PAGES = 600
OUT = os.path.join(ROOT, "output", "Data and Chart Outputs", "vietnam_power_generation_daily.xlsx")
# label in the post (lower case, accents kept) -> column; first match wins, so the rooftop lines come first
SOURCES = [("đmt mái nhà (ước tính đầu cực)", "Solar_rooftop"), ("đmt mái nhà (ước tính thương phẩm)", "Solar_rooftop_delivered"),
           ("đmt mái nhà", "Solar_rooftop"), ("đmt trang trại", "Solar_farm"), ("thủy điện", "Hydro"),
           ("nhiệt điện than", "Coal"), ("tuabin khí", "Gas"), ("nhiệt điện dầu", "Oil"), ("điện gió", "Wind"),
           ("nhập khẩu", "Imports"), ("khác", "Other")]
FUELS = ["Hydro", "Coal", "Gas", "Oil", "Wind", "Solar", "Other"]


def out(*a):
    print(*a, flush=True)


def get(url):
    for i in range(4):
        try:
            r = requests.get(url, headers=H, timeout=T)
            r.raise_for_status()
            return r.text
        except requests.RequestException as e:
            if i == 3:
                out(f"  {url[-70:]}: {e}")
                return None
            time.sleep(5 * (i + 1))


def num(s):
    """'1114,4' / '53478,5' / '1.114,4' / '2417' -> float (Vietnamese decimal comma)."""
    s = s.strip()
    if "," in s:
        s = s.replace(".", "").replace(",", ".")
    elif re.fullmatch(r"\d{1,3}(\.\d{3})+", s):
        s = s.replace(".", "")
    return float(s)


def plain(page):
    t = re.sub(r"<script.*?</script>|<style.*?</style>", " ", page, flags=re.S | re.I)
    t = re.sub(r"<br\s*/?>|</p>|</li>|</tr>", "\n", t, flags=re.I)
    t = re.sub(r"<[^>]+>", " ", t)
    return re.sub(r"[ \t ]+", " ", html.unescape(t))


def listing(page):
    """{date: post url} on one list page."""
    t = get(f"{LIST}?page={page}")
    if t is None:
        return None
    posts = {}
    # the date in the link text ('... ngày 11/1/2026'); the url slug drops leading zeros ('ngay-1112026') so is
    # ambiguous, and is used only when it has all eight digits (the sidebar's links carry no text)
    for u, title in re.findall(r'<a[^>]+href="(/d/vi-VN/news/[^"]*?van-hanh-he-thong-dien-Quoc-gia[^"]*)"[^>]*>(.*?)</a>',
                               t, re.I | re.S):
        m = re.search(r"ngày\s*(\d{1,2})/(\d{1,2})/(\d{4})", re.sub(r"<[^>]+>", " ", title))
        s = re.search(r"ngay-(\d{8})-", u)
        try:
            if m:
                day = date(int(m.group(3)), int(m.group(2)), int(m.group(1)))
            elif s:
                d = s.group(1)
                day = date(int(d[4:]), int(d[2:4]), int(d[:2]))
            else:
                continue
        except ValueError:
            continue
        posts.setdefault(day, BASE + u)
    return posts


def parse(page):
    """One post -> {column: value}: GWh by source, totals, MW."""
    t = plain(page)
    i = t.find("Thông tin chung về vận hành")
    t = t[i:] if i >= 0 else t
    row = {}
    tot = re.findall(r"Sản lượng điện sản xuất và nhập khẩu\s*:\s*([\d.,]+)", t)
    peak = re.findall(r"Công suất lớn nhất trong ngày\s*:\s*([\d.,]+)", t)
    if tot:
        row["Total_incl_imports_GWh"] = num(tot[-1])   # the last one is on the at-terminals rooftop basis
    if peak:
        row["Demand_peak_MW"] = max(num(p) for p in peak)
    mix = t[t.find("Cơ cấu sản lượng"):]
    j = mix.upper().find("CÔNG SUẤT HUY ĐỘNG")
    table, mix = (mix[j:], mix[:j]) if j > 0 else ("", mix)
    for line in re.findall(r"-\s*([^\n\-]+?)\s+([\d.,]+)\s*triệu kWh", mix):
        label, v = line[0].strip().lower(), num(line[1])
        for key, col in SOURCES:
            if key in label:
                row.setdefault(col + "_GWh", v)
                break
    # dispatch table: 'Quốc gia + ĐMT mái nhà (ước tính đầu cực) 50416,8 49855,6'
    m = re.search(r"Quốc gia \+ ĐMT mái nhà \(ước tính đầu cực\)\s+([\d.,]+)\s+([\d.,]+)", table)
    if m:
        row["Midday_MW"], row["Evening_MW"] = num(m.group(1)), num(m.group(2))
    return row


def fetch(item):
    day, url = item
    page = get(url)
    if page is None:
        return day, None
    try:
        return day, parse(page)
    except Exception as e:  # noqa: BLE001
        out(f"  {day}: {type(e).__name__}: {e}")
        return day, None


def read_sheet(path, sheet):
    try:
        d = pd.read_excel(path, sheet_name=sheet, index_col=0)
    except (FileNotFoundError, ValueError):
        return pd.DataFrame()
    d.index = pd.to_datetime(d.index)
    return d.sort_index()


def to_frames(rows):
    """{date: parsed row} -> (Daily MWh, Demand MW)."""
    r = pd.DataFrame.from_dict(rows, orient="index")
    if r.empty:
        return pd.DataFrame(), pd.DataFrame()
    r.index = pd.to_datetime(r.index)
    g = lambda c: r.get(c + "_GWh")   # noqa: E731
    solar = pd.concat([g("Solar_farm"), g("Solar_rooftop")], axis=1).sum(axis=1, min_count=1) \
        if g("Solar_farm") is not None else None
    daily = pd.DataFrame({"Hydro_MWh": g("Hydro"), "Coal_MWh": g("Coal"), "Gas_MWh": g("Gas"), "Oil_MWh": g("Oil"),
                          "Wind_MWh": g("Wind"), "Solar_MWh": solar, "Other_MWh": g("Other")}, index=r.index) * 1000
    daily["Total_MWh"] = daily.sum(axis=1, min_count=3)
    for c in ("Imports", "Solar_farm", "Solar_rooftop", "Solar_rooftop_delivered", "Total_incl_imports"):
        if g(c) is not None:
            daily[c + "_MWh"] = g(c) * 1000
    demand = r[[c for c in ("Demand_peak_MW", "Midday_MW", "Evening_MW") if c in r]]
    return daily.round(0), demand.round(1)


def merge(old, new):
    if old.empty or new.empty:
        return (new if old.empty else old).sort_index()
    return pd.concat([old[~old.index.isin(new.index)], new]).sort_index()


def save(path, daily, demand, archive_start=None):
    daily.index.name = demand.index.name = "date"
    check = (daily["Total_MWh"] + daily.get("Imports_MWh", 0) - daily.get("Total_incl_imports_MWh")).abs()
    off = int((check > 0.01 * daily["Total_MWh"]).sum()) if "Total_incl_imports_MWh" in daily else 0
    notes = [
        "UNITS",
        "Daily: MWh per day (EVN reports million kWh = GWh; x 1000). Hydro, Coal (coal-fired thermal), Gas (gas "
        "turbines, burning gas or DO oil), Oil (oil-fired thermal), Wind, Solar = farm solar + rooftop solar "
        "ESTIMATED at the generator terminals (EVN's 'uoc tinh dau cuc'), Other (biomass, southern diesel, ...). "
        "Total_MWh = their sum (domestic generation); Imports_MWh (from Laos and China) is kept apart. "
        "Solar_rooftop_delivered_MWh = EVN's other rooftop estimate (energy delivered, 'thuong pham'). "
        "Total_incl_imports_MWh = production + imports as EVN states it (rooftop at the terminals).",
        "Demand: Demand_peak_MW = the day's maximum (rooftop at the terminals); Midday_MW / Evening_MW = national "
        "dispatch at the midday low and at the evening peak.",
        f"Check: Total_MWh + Imports_MWh matches EVN's stated total within 1% on all but {off} days.",
        "",
        "COVERAGE",
        f"Daily from {daily.index.min():%Y-%m-%d} to {daily.index.max():%Y-%m-%d} ({len(daily)} days; EVN posts the "
        "summary the next morning and skips the odd day). National power system (EVN / NSMO dispatch); rooftop solar "
        "is NSMO's estimate, not metered.",
        "",
        "SOURCE",
        f"EVN (Vietnam Electricity), 'Thong tin tom tat van hanh HTD Quoc gia' daily posts, data from NSMO (National "
        f"Power System and Market Operator, https://www.nsmo.vn/HeThongDien): {LIST}",
    ]
    sheets = {"Daily": daily, "Demand": demand}
    if archive_start:
        sheets["Archive"] = pd.DataFrame({"note": ["earliest post in EVN's list when last read to its end"]},
                                         index=pd.DatetimeIndex([pd.Timestamp(archive_start)], name="date"))
    xlsx_notes.write_workbook(path, sheets, notes, {"UNITS", "COVERAGE", "SOURCE"})
    out(f"Saved {path}: {len(daily)} days {daily.index.min():%Y-%m-%d}..{daily.index.max():%Y-%m-%d}")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", default=OUT)
    args = ap.parse_args()
    old_d, old_m = read_sheet(args.out, "Daily"), read_sheet(args.out, "Demand")
    have = set(old_d.dropna(subset=["Total_MWh"]).index.date) if "Total_MWh" in old_d else set()
    today = date.today()
    revise = {today - timedelta(days=k) for k in range(REVISION_DAYS + 1)}
    # every day from EVN's archive start (recorded once the list has been read to its end) that is not saved yet
    # counts as missing, interior gaps included; the whole list (~120 pages) is then read in about 3 minutes
    arch = read_sheet(args.out, "Archive")
    archive_start = arch.index.min().date() if not arch.empty else None   # None: not read to the end yet
    first = max(DATA_START, archive_start or DATA_START)
    missing = {first + timedelta(days=k) for k in range((today - first).days + 1)} - (have - revise)
    earliest = min(missing) if missing else today
    out(f"{len(have)} days saved; {len(missing)} candidate days back to {earliest}")
    posts, empty = {}, 0
    for page in range(1, MAX_PAGES + 1):
        p = listing(page)
        if p is None:
            break
        new = {d: u for d, u in p.items() if d not in posts}
        posts.update(new)
        if not new:   # past the last page only the sidebar's links remain
            empty += 1
            if empty >= 3:
                break
            continue
        empty = 0
        if min(new) < earliest:
            break
        if page % 25 == 0:
            out(f"  page {page}: back to {min(posts)}")
    if empty >= 3 and posts:   # read to the end of the list: record where EVN's archive starts
        archive_start = min(posts)
    todo = sorted((d, u) for d, u in posts.items() if d in missing and d >= DATA_START)
    out(f"{len(posts)} posts listed ({min(posts) if posts else '-'}..{max(posts) if posts else '-'}, list read to page "
        f"{page}); fetching {len(todo)}")
    rows = {}
    for b in range(0, len(todo), 120):
        with ThreadPoolExecutor(6) as ex:
            for d, row in ex.map(fetch, todo[b:b + 120]):
                if row and len([k for k in row if k.endswith("_GWh")]) >= 5:
                    rows[d] = row
                elif row is not None:
                    out(f"  {d}: too few sources parsed ({sorted(row)})")
        daily, demand = to_frames(rows)
        d_all, m_all = merge(old_d, daily), merge(old_m, demand)
        if not d_all.empty:
            save(args.out, d_all, m_all, archive_start)
    if not todo and not old_d.empty:
        save(args.out, old_d, old_m, archive_start)
    d = read_sheet(args.out, "Daily")
    if d.empty:
        raise SystemExit("No EVN daily posts parsed")
    out(d.tail(3).T.to_string())


if __name__ == "__main__":
    main()
