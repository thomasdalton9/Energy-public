"""
Shared helpers for the China National Bureau of Statistics (NBS) pulls
(asia/CHINA_NBS_*.py). Not a pull script itself.

SOURCE (found via discovery_archive/asia/CHINA_NBS_DISCOVERY*.py):
  - data.stats.gov.cn (NBS's query database, incl. the new /dg/ portal)
    answers every automated request from GitHub Actions with 403
    "reason:UrlACL" - unusable.
  - The Chinese "latest releases" list, www.stats.gov.cn/sj/zxfb/ (page 1
    = base URL, page N = index_{N-1}.html), is reachable and holds ~1000
    releases (67 pages), reaching back to Oct 2021. Releases from before
    the Feb-2023 site migration keep URLs of the form
    /sj/zxfb/202302/t20230203_<id>.html (IDs roughly in date order).
  - The English press-release list only reaches back to Apr 2024, so the
    Chinese releases are the primary source; their tables carry the same
    numbers.
  - stats.gov.cn serves an anti-bot JavaScript challenge page ("Please
    enable JavaScript", HTTP 200) when hit too fast - fetch() requests
    sequentially with a pause, detects the challenge and backs off.

Release types used (each an HTML table in the release body):
  - "YYYY年M月份规模以上工业增加值增长X%" (industrial production): table
    "规模以上工业主要产品产量" - output of ~37 industrial products incl. raw
    coal, crude oil, crude runs, natural gas, coke, electricity by source,
    steel, cement, aluminium, chemicals, vehicles, solar cells. Columns:
    [product (unit), month value, month y/y %, YTD value, YTD y/y %]. The
    January and February figures are only published combined
    ("1—2月份" release: [product, Jan-Feb value, y/y %]).
  - "YYYY年M月X旬流通领域重要生产资料市场价格变动情况" (10-day market prices
    of ~50 producer goods incl. coal grades, coke, LNG, LPG, gasoline,
    diesel, steel, metals, chemicals).
  - "YYYY年X季度全国(规模以上)工业产能利用率为X%" (quarterly capacity
    utilisation by industry).
  - "YYYY年M月份工业生产者出厂价格..." (PPI: m/m and y/y % by sector).
"""

import os
import re
import sys
import time

import pandas as pd
import requests

CN_INDEX = "https://www.stats.gov.cn/sj/zxfb/"
HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                  "(KHTML, like Gecko) Chrome/124.0 Safari/537.36",
    "Accept-Language": "zh-CN,zh;q=0.9,en;q=0.8",
}
TIMEOUT = (10, 30)
PAUSE_SECONDS = 0.8           # between requests - the site challenges fast crawlers
MAX_INDEX_PAGES = 80          # list is ~67 pages (capped at ~1000 items)
CHALLENGE_MARK = "Please enable JavaScript"

_session = requests.Session()
_session.headers.update(HEADERS)
_last = [0.0]


class Blocked(Exception):
    """stats.gov.cn kept serving its anti-bot challenge."""


def log(msg):
    print(msg, file=sys.stderr, flush=True)


def fetch(url, attempts=4):
    """Text of url, or None on 404. Raises Blocked / RequestException on persistent failure."""
    last_err = None
    for attempt in range(attempts):
        wait = PAUSE_SECONDS - (time.time() - _last[0])
        if wait > 0:
            time.sleep(wait)
        try:
            r = _session.get(url, timeout=TIMEOUT)
            _last[0] = time.time()
        except requests.RequestException as e:
            _last[0] = time.time()
            last_err = e
            log(f"    attempt {attempt + 1}/{attempts} {url}: {type(e).__name__}")
            time.sleep(10 * (attempt + 1))
            continue
        if r.status_code == 404:
            return None
        r.encoding = "utf-8"
        if r.status_code == 200 and CHALLENGE_MARK in r.text[:3000]:
            last_err = Blocked(url)
            log(f"    anti-bot challenge on {url} - backing off")
            time.sleep(30 * (attempt + 1))
            continue
        if r.status_code != 200:
            last_err = requests.HTTPError(f"{r.status_code} {url}")
            time.sleep(10 * (attempt + 1))
            continue
        return r.text
    raise last_err


def _links(html, base):
    out = []
    for m in re.finditer(r"<a\b([^>]*)>(.*?)</a>", html, re.S):
        attrs, inner = m.group(1), m.group(2)
        h = re.search(r'href="([^"]*)"', attrs)
        t = re.search(r'title="([^"]*)"', attrs)
        title = (t.group(1) if t else re.sub(r"<[^>]+>", "", inner)).strip()
        if not h or not title or not re.search(r"/t\d{8}_\d+\.html", h.group(1)):
            continue
        href = h.group(1)
        if href.startswith("./"):
            href = base + href[2:]
        elif href.startswith("/"):
            href = "https://www.stats.gov.cn" + href
        out.append((re.sub(r"\s+", "", title), href))
    return out


def crawl_index(title_re, wanted, deep, extra=(), stop_after_known=3, near_re=None):
    """Releases on the /sj/zxfb/ list whose title matches title_re and wanted(title), newest first.

    deep=False (routine run, nothing missing): stop once `stop_after_known` matching releases in a
    row are already held. deep=True (backfill): read the whole list.
    extra: (title, url) pairs for releases known to exist but no longer listed.
    near_re: titles matching this but not title_re are logged (spots renamed releases).
    """
    rx = re.compile(title_re)
    found, seen, near_logged = [], set(), set()
    known_streak = 0
    for i in range(MAX_INDEX_PAGES):
        url = CN_INDEX if i == 0 else f"{CN_INDEX}index_{i}.html"
        try:
            html = fetch(url)
        except (requests.RequestException, Blocked) as e:
            log(f"  index page {i}: {type(e).__name__} - stopping crawl")
            break
        if html is None:
            log(f"  index page {i}: 404 - end of list")
            break
        for title, href in _links(html, url.rsplit("/", 1)[0] + "/"):
            if href in seen:
                continue
            if not rx.search(title):
                if near_re and re.search(near_re, title) and title not in near_logged:
                    near_logged.add(title)
                    log(f"  (similar title not matched: {title})")
                continue
            seen.add(href)
            if wanted(title):
                found.append((title, href))
                known_streak = 0
            else:
                known_streak += 1
        if not deep and known_streak >= stop_after_known:
            log(f"  index page {i}: reached releases already held - stopping crawl")
            break
    for title, href in extra:
        if href not in seen and rx.search(title) and wanted(title):
            found.append((title, href))
    return found


def table_rows(html):
    """All <tr> rows of the page as lists of non-empty cell texts."""
    rows = []
    for tr in re.findall(r"<tr[^>]*>(.*?)</tr>", html, re.S | re.I):
        cells = []
        for c in re.findall(r"<t[dh][^>]*>(.*?)</t[dh]>", tr, re.S | re.I):
            c = re.sub(r"<[^>]+>", "", c).replace("&nbsp;", " ")
            c = re.sub(r"\s+", "", c)
            if c:
                cells.append(c)
        if cells:
            rows.append(cells)
    return rows


def num(s):
    s = str(s).replace(",", "").replace("，", "").strip()
    try:
        return float(s)
    except ValueError:
        return None


def norm_label(label):
    """'其中：原铝（电解铝）（万吨）' -> ('原铝(电解铝)', '万吨')."""
    s = label.replace("（", "(").replace("）", ")").replace("：", ":").replace(" ", "")
    s = re.sub(r"^其中:", "", s)
    m = re.match(r"^(.*?)\(([^()]*)\)$", s)
    if m and re.search(r"吨|米|箱|台|套|辆|千瓦|块|立方米|件|%|元", m.group(2)):
        return m.group(1), m.group(2)
    return s, ""


def month_period(title):
    """'2026年8月份...' -> (Timestamp 2026-08-01, False); '2026年1—2月份...' -> (2026-02-01, True)."""
    m = re.search(r"(\d{4})年1[—\-－~～至]2月", title)
    if m:
        return pd.Timestamp(int(m.group(1)), 2, 1), True
    m = re.search(r"(\d{4})年(\d{1,2})月", title)
    if m:
        return pd.Timestamp(int(m.group(1)), int(m.group(2)), 1), False
    return None, None


def read_sheet(path, sheet):
    """Existing data sheet (index = first column as dates) or None."""
    if not os.path.exists(path):
        return None
    try:
        df = pd.read_excel(path, sheet_name=sheet, index_col=0)
    except (ValueError, KeyError):
        return None
    df.index = pd.to_datetime(df.index, errors="coerce")
    return df[df.index.notna()]


def has_sheet(path, sheet):
    if not os.path.exists(path):
        return False
    try:
        return sheet in pd.ExcelFile(path).sheet_names
    except Exception:  # noqa: BLE001 - unreadable file -> rebuild
        return False


# Gaps older than this are not chased with a full crawl every run: the release list only
# holds ~1000 items, so an old gap that is not on it now never will be.
GAP_WINDOW = pd.Timedelta(days=400)


def recent_gaps(held, freq):
    """Periods (Timestamps) missing between the first and last held, within GAP_WINDOW of the last."""
    held = set(pd.DatetimeIndex(list(held)))
    if not held:
        return []
    if freq == "10D":
        span = [pd.Timestamp(d.year, d.month, x) for d in pd.date_range(min(held).replace(day=1), max(held), freq="MS")
                for x in (1, 11, 21)]
    else:
        span = list(pd.date_range(min(held), max(held), freq=freq))
    return [p for p in span if min(held) <= p <= max(held) and p not in held and p >= max(held) - GAP_WINDOW]


def series_sheet(rows):
    """'Series' sheet read by add_charts.china_nbs_series(): column, label, unit, chart group, chart kind."""
    return pd.DataFrame(rows, columns=["column", "label", "unit", "chart", "kind"]).set_index("column")
