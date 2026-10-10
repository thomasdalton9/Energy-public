"""
Generic incremental pull for NBS releases that are one wide table per period (CPI, industrial profits, fixed-asset
investment, retail sales): see CHINA_NBS_CPI.py, CHINA_NBS_PROFITS.py, CHINA_NBS_FAI.py, CHINA_NBS_RETAIL.py.
Not a pull script itself. Source and anti-bot handling: china_nbs_common.py.

pull(spec) reads the held workbook, crawls the release list only as deep as needed, parses new releases and writes
Data (+ 'Jan-Feb' where January and February come only combined) and Series. The committed workbook is the history store.
"""

import os
import re

import pandas as pd

import china_nbs_common as nbs
import xlsx_notes

_OLD = "https://www.stats.gov.cn/sj/zxfb/202302/t20230203_{}.html"


def old_releases(pairs):
    """[(title, id)] -> [(title, url)] for releases of the Feb-2023 site migration that are no longer on the list."""
    return [(t, _OLD.format(i)) for t, i in pairs]


def ytd_period(title):
    """Period of a year-to-date style release -> (Timestamp of the last month covered, is_jan_feb_combined).

    '2026年1—8月份...' -> 2026-08; '2026年1—2月份...' -> (2026-02, True); '2025年上半年' -> 06; '2025年前三季度' -> 09;
    '2022年一季度' -> 03; '2022年12月...' -> 12; '2025年全国...' (full year) -> 12."""
    m = re.search(r"(\d{4})年1[—\-－~～至](\d{1,2})月", title)
    if m:
        n = int(m.group(2))
        return pd.Timestamp(int(m.group(1)), n, 1), n == 2
    for word, month in (("上半年", 6), ("前三季度", 9), ("一季度", 3)):
        m = re.search(r"(\d{4})年" + word, title)
        if m:
            return pd.Timestamp(int(m.group(1)), month, 1), False
    m = re.search(r"(\d{4})年(\d{1,2})月", title)
    if m:
        return pd.Timestamp(int(m.group(1)), int(m.group(2)), 1), False
    m = re.search(r"(\d{4})年(?:全国|全年|社会|规模)", title)
    if m:
        return pd.Timestamp(int(m.group(1)), 12, 1), False
    return None, None


def clean_name(s):
    """Row label as a plain key: spaces, list numbering and the 其中： prefix removed."""
    s = re.sub(r"\s+", "", str(s)).replace("（", "(").replace("）", ")").replace("：", ":")
    s = re.sub(r"^[一二三四五六七八九十]+、", "", s)
    return re.sub(r"^其中:", "", s)


def pull(out_path, *, title_re, near_re, period_fn, extra, parse_fn, columns, series_rows, notes_lines, notes_titles,
         history_start, probe_col, jan_feb_sheet=False, skip_gap_months=(), index_name="month"):
    """parse_fn(html, is_jan_feb) -> {column: value}. A release counts as held once probe_col is filled for its period."""
    data = janfeb = None
    if nbs.has_sheet(out_path, "Series"):
        data = nbs.read_sheet(out_path, "Data")
        janfeb = nbs.read_sheet(out_path, "Jan-Feb") if jan_feb_sheet else None
    data = pd.DataFrame(columns=columns, dtype=float) if data is None else data.reindex(columns=columns)
    janfeb = pd.DataFrame(columns=columns, dtype=float) if janfeb is None else janfeb.reindex(columns=columns)
    held = {(p.year, p.month) for p in data.index[data[probe_col].notna()]} if len(data) else set()
    held_jf = {p.year for p in janfeb.index[janfeb[probe_col].notna()]} if len(janfeb) else set()
    span = data.index[data.notna().any(axis=1)]
    gaps = [g for g in nbs.recent_gaps(span, "MS") if g.month not in skip_gap_months] if len(span) else []
    deep = (not held and not held_jf) or bool(gaps)
    nbs.log(f"  held {len(held)} periods (+{len(held_jf)} Jan-Feb); gaps {len(gaps)} -> {'full' if deep else 'shallow'} crawl")

    def wanted(title):
        period, is_jf = period_fn(title)
        if period is None or period < pd.Timestamp(history_start):
            return False
        if is_jf and jan_feb_sheet:
            return period.year not in held_jf
        return (period.year, period.month) not in held

    releases = nbs.crawl_index(title_re, wanted, deep, extra, near_re=near_re)
    nbs.log(f"  {len(releases)} release(s) to fetch")
    new, new_jf = {}, {}
    for title, url in releases:
        period, is_jf = period_fn(title)
        try:
            html = nbs.fetch(url)
        except Exception as e:  # noqa: BLE001 - next run retries
            nbs.log(f"  [{period:%Y-%m}] fetch failed ({type(e).__name__}) - skipped")
            continue
        row = parse_fn(html or "", bool(is_jf and jan_feb_sheet))
        row = {k: v for k, v in row.items() if k in columns and v is not None}
        if not row:
            nbs.log(f"  [{period:%Y-%m}] nothing parsed - skipped ({url})")
            continue
        nbs.log(f"  [{period:%Y-%m}{' Jan-Feb' if is_jf and jan_feb_sheet else ''}] {len(row)} values")
        (new_jf if is_jf and jan_feb_sheet else new)[period] = row

    if new:
        add = pd.DataFrame.from_dict(new, orient="index")
        data = data.combine_first(add) if len(data) else add
    if new_jf:
        add = pd.DataFrame.from_dict(new_jf, orient="index")
        janfeb = janfeb.combine_first(add) if len(janfeb) else add
    if data.dropna(how="all").empty and janfeb.dropna(how="all").empty:
        raise SystemExit("No data at all - nothing to save.")
    data = data.reindex(columns=columns).sort_index()
    data = data[data.index >= pd.Timestamp(history_start)]
    janfeb = janfeb.reindex(columns=columns).sort_index()
    data.index.name = janfeb.index.name = index_name
    sheets = {"Data": data, "Series": nbs.series_sheet(series_rows)}
    if jan_feb_sheet:
        sheets["Jan-Feb"] = janfeb
    os.makedirs(os.path.dirname(os.path.abspath(out_path)), exist_ok=True)
    xlsx_notes.write_workbook(out_path, sheets, notes_lines, notes_titles)
    print(f"Saved {len(data)} period(s) ({data.index.min():%Y-%m}..{data.index.max():%Y-%m}), {len(janfeb)} Jan-Feb, "
          f"{len(new) + len(new_jf)} new, to {out_path}")
