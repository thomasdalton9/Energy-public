"""
Jordan's daily peak load from NEPCO (National Electric Power Company) 'Max Load Archive'
(https://www.nepco.com.jo/en/MaxLoadArchive.aspx): one maximum MW figure per day, no hour, no minimum, from
2 Oct 2019. Days NEPCO did not post are simply absent (often 14-27 of 31 days a month posted in 2021-24); they stay
gaps, nothing is filled.

HOW THE PAGE WORKS (probe discovery_archive/middle_east/JORDAN_NEPCO_PROBE.py, rounds 1-8, Oct 2026):
  * ASP.NET WebForms. A plain GET returns the latest 10 days, newest first.
  * A date search is a POST replaying __VIEWSTATE / __EVENTVALIDATION with txtDayFrom / fromMonth / txtYearFrom /
    txtDayTo / toMonth / txtYearTo + btnSearch. The results page holds 10 rows; the DataPager's page-2 postback
    DROPS the filter, so every search must fit on one page: calendar windows of at most 9 days
    (1-9, 10-18, 19-27, 28-end of month), 4 searches a month.
  * After about 32 searches (about 65 requests) in one run the server cuts the runner's IP off for the rest of the
    run: first pages without the table, then connection resets. It is a COUNT limit, not a rate limit; a new run
    (new runner) starts clean. So a run stops at the first such page / reset, SAVES what it has and does not count
    the failed window as searched. Each search reuses the hidden fields of the previous answer when it holds the form
    (one request per search instead of a GET + POST).

HISTORY STORE: the committed workbook. 'Windows' lists every searched window (rows found, status): 'ok' (rows),
'empty' (NEPCO posted nothing in it) - neither is searched again unless it was searched less than 3 days after
the window ended (NEPCO may still post). 'Daily' keeps every day found.
  normal run      one GET (latest 10 days) + the windows from the last stored day minus 10 days to today
  --backfill      one GET + windows not yet searched, newest first, back to 2019-10-01, at most --max-searches

VALUES: Raw_MW is the figure as published; Peak_Load_MW is the same figure, blanked (and `flag` set) when it lies
outside 1,000-6,000 MW (Jordan's peak is about 2,500-4,800 MW; e.g. 24600 MW on 02/01/2020 is a data-entry slip).
Nothing is corrected by guessing.
'Monthly': maximum of the daily peaks per month with the number of days posted; `Charted` = month ended and at
least 20 days posted (a month with fewer posted days may miss its true peak).

Workflow: .github/workflows/jordan_peak_load.yml (1st and 15th; manual input `backfill`).
"""

import argparse
import calendar
import os
import re
import sys
import time
from datetime import date, datetime, timedelta, timezone

import pandas as pd
import requests

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))  # repo root, for xlsx_notes

import xlsx_notes

URL = "https://www.nepco.com.jo/en/MaxLoadArchive.aspx"
P = "ctl00$ContentPlaceHolder1$"
H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) "
                   "Chrome/124.0 Safari/537.36", "Accept-Language": "en-US,en;q=0.9"}
START = date(2019, 10, 1)
PLAUSIBLE_MW = (1000, 6000)
MIN_DAYS_MONTH = 20
REVISION_DAYS = 10
PAUSE_S = 3
POST_LAG_DAYS = 2      # a window searched within this many days of its end is searched again later

REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DEFAULT_OUT = os.path.join(REPO_ROOT, "output", "Data and Chart Outputs", "jordan_peak_load_daily.xlsx")
DAILY_COLS = ["Peak_Load_MW", "Raw_MW", "flag", "Fetched_UTC"]
WIN_COLS = ["Window_end", "Searched_UTC", "Rows", "Status", "Note"]


class CutOff(Exception):
    """The server stopped answering properly (page without the table, or a connection reset)."""


# ---------------------------------------------------------------------------------------------------- source
def hidden(html):
    return {k: v for k, v in re.findall(r'<input[^>]+type="hidden"[^>]+name="([^"]+)"[^>]*value="([^"]*)"', html)}


def parse_rows(html):
    """[(date, raw text)] from the results table (dd/mm/yyyy then the MW figure in the next cell)."""
    out = []
    for d, v in re.findall(r'(\d{2}/\d{2}/\d{4})\s*</[^>]+>\s*<[^>]+>\s*([\d,]+(?:\.\d+)?)', html):
        try:
            out.append((datetime.strptime(d, "%d/%m/%Y").date(), v.replace(",", "")))
        except ValueError:
            pass
    return out


def page_ok(html):
    """A real answer carries the archive page and its search form; the cut-off pages do not."""
    return "Max Load" in html and (P + "txtDayFrom") in html and "__VIEWSTATE" in html


class Nepco:
    def __init__(self):
        self.s = requests.Session()
        self.s.headers.update(H)
        self.verify = True
        self.state = None      # hidden fields of the last good page
        self.requests = 0

    def _req(self, method, data=None):
        self.requests += 1
        try:
            r = self.s.request(method, URL, data=data, timeout=(15, 60), verify=self.verify)
        except requests.exceptions.SSLError as e:
            if "CERTIFICATE_VERIFY_FAILED" in str(e) and self.verify:
                print("  certificate not verifiable from the runner - retrying without verification", flush=True)
                requests.packages.urllib3.disable_warnings()
                self.verify = False
                return self._req(method, data)
            raise CutOff(f"SSL error: {str(e)[:160]}")
        except (requests.exceptions.ConnectionError, requests.exceptions.Timeout,
                requests.exceptions.ChunkedEncodingError) as e:
            raise CutOff(f"{type(e).__name__}: {str(e)[:160]}")
        if r.status_code >= 400:
            raise CutOff(f"HTTP {r.status_code}")
        if not page_ok(r.text):
            raise CutOff(f"page without the table/form (HTTP {r.status_code}, {len(r.content)} B)")
        self.state = hidden(r.text)
        return r.text

    def latest(self):
        return parse_rows(self._req("GET"))

    def search(self, d1, d2):
        if not self.state:
            self._req("GET")
        f = dict(self.state)
        f.pop("__EVENTTARGET", None)
        f.pop("__EVENTARGUMENT", None)
        f.update({P + "txtDayFrom": str(d1.day), P + "fromMonth": str(d1.month), P + "txtYearFrom": str(d1.year),
                  P + "txtDayTo": str(d2.day), P + "toMonth": str(d2.month), P + "txtYearTo": str(d2.year),
                  P + "btnSearch": "Search"})
        return parse_rows(self._req("POST", f))


# ---------------------------------------------------------------------------------------------------- windows
def windows(first, last):
    """Calendar windows 1-9, 10-18, 19-27, 28-end overlapping [first, last], oldest first, as (start, end)."""
    out = []
    y, m = first.year, first.month
    while (y, m) <= (last.year, last.month):
        nd = calendar.monthrange(y, m)[1]
        for a, b in ((1, 9), (10, 18), (19, 27), (28, nd)):
            ws, we = date(y, m, a), date(y, m, b)
            if we >= first and ws <= last:
                out.append((ws, we))
        y, m = (y + 1, 1) if m == 12 else (y, m + 1)
    return out


# ---------------------------------------------------------------------------------------------------- store
def load(path):
    try:
        d = pd.read_excel(path, sheet_name="Daily", index_col=0)
        d.index = pd.to_datetime(d.index).date
        d = d.reindex(columns=DAILY_COLS).astype(object)
    except (FileNotFoundError, ValueError):
        d = pd.DataFrame(columns=DAILY_COLS, dtype=object)
    try:
        w = pd.read_excel(path, sheet_name="Windows", index_col=0)
        w.index = pd.to_datetime(w.index).date
        w = w.reindex(columns=WIN_COLS).astype(object)
        w["Window_end"] = pd.to_datetime(w["Window_end"]).dt.date
        w["Searched_UTC"] = pd.to_datetime(w["Searched_UTC"]).dt.date
    except (FileNotFoundError, ValueError):
        w = pd.DataFrame(columns=WIN_COLS, dtype=object)
    return d, w


def check(raw):
    """(Peak_Load_MW, flag) for a published figure."""
    try:
        v = float(raw)
    except (TypeError, ValueError):
        return None, f"not a number as published: {raw!r}"
    if not PLAUSIBLE_MW[0] <= v <= PLAUSIBLE_MW[1]:
        return None, f"outside {PLAUSIBLE_MW[0]:,}-{PLAUSIBLE_MW[1]:,} MW as published; blanked, not corrected"
    return v, ""


def store_rows(daily, rows, today):
    new = changed = 0
    for d, raw in rows:
        rawv = float(raw) if re.fullmatch(r"\d+(\.\d+)?", raw) else raw
        v, flag = check(raw)
        if d in daily.index:
            if daily.at[d, "Raw_MW"] == rawv or str(daily.at[d, "Raw_MW"]) == str(rawv):
                continue
            changed += 1
            print(f"  revised {d}: {daily.at[d, 'Raw_MW']} -> {rawv}")
        else:
            new += 1
        daily.loc[d] = [v, rawv, flag, today.isoformat()]
    return new, changed


def monthly(daily, today):
    d = daily.copy()
    d.index = pd.to_datetime(d.index)
    v = pd.to_numeric(d["Peak_Load_MW"], errors="coerce")
    g = v.resample("MS")
    m = pd.DataFrame({"Max_Peak_MW": g.max(), "Days_posted": g.count()})
    m["Days_in_month"] = m.index.days_in_month
    month_end = (m.index + pd.offsets.MonthEnd(0)).date
    m["Charted"] = ["yes" if (n >= MIN_DAYS_MONTH and e < today) else "no"
                    for n, e in zip(m["Days_posted"], month_end)]
    m.index = pd.Index(m.index.date, name="month")
    return m


NOTES = [
    "UNITS",
    "Peak_Load_MW: Jordan's daily maximum load, MW, as NEPCO publishes it in its 'Max Load Archive' (one figure per "
    "day; NEPCO does not give the hour or a minimum). Raw_MW: the figure exactly as published. A figure outside "
    f"{PLAUSIBLE_MW[0]:,}-{PLAUSIBLE_MW[1]:,} MW is blanked in Peak_Load_MW and explained in `flag` (e.g. 24600 on "
    "2 Jan 2020, evidently 2460 mistyped) - it is NOT corrected.",
    "Monthly: Max_Peak_MW = highest daily peak posted in the month; Days_posted = days with a usable figure; "
    f"Charted = 'yes' when the month has ended and at least {MIN_DAYS_MONTH} days are posted (fewer posted days may "
    "miss the month's true peak), and only those months are charted.",
    "",
    "COVERAGE",
    "From 2 Oct 2019 (nothing earlier in the archive). Days NEPCO did not post are missing (often 14-27 of 31 days a "
    "month posted in 2021-24, mostly Fridays/holidays and odd weekdays); they are gaps, nothing is filled.",
    "Windows: every searched date window (searches must fit one 10-row results page, so <= 9-day windows: 1-9, 10-18, "
    "19-27, 28-end), with the date searched, rows found and status (ok / empty = NEPCO posted nothing in it).",
    "",
    "SOURCE",
    f"National Electric Power Company (NEPCO), Jordan: Max Load Archive, {URL}",
]
NOTES_TITLES = {"UNITS", "COVERAGE", "SOURCE"}


def save(path, daily, wins, today):
    daily = daily.sort_index()
    daily.index = pd.Index(daily.index, name="date")
    wins = wins.sort_index(ascending=False)
    wins.index = pd.Index(wins.index, name="Window_start")
    os.makedirs(os.path.dirname(os.path.abspath(path)), exist_ok=True)
    xlsx_notes.write_workbook(path, {"Daily": daily[DAILY_COLS], "Monthly": monthly(daily, today),
                                     "Windows": wins[WIN_COLS]}, NOTES, NOTES_TITLES)


# ---------------------------------------------------------------------------------------------------- main
def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--out", default=DEFAULT_OUT)
    ap.add_argument("--backfill", action="store_true", help="search unsearched windows back to 2019-10-01")
    ap.add_argument("--max-searches", type=int, default=26)
    args = ap.parse_args()

    today = datetime.now(timezone.utc).date()
    daily, wins = load(args.out)
    print(f"Stored: {len(daily)} days, {len(wins)} windows searched")

    def done(ws):
        if ws not in wins.index:
            return False
        r = wins.loc[ws]
        return (r["Status"] in ("ok", "empty") and pd.notna(r["Searched_UTC"])
                and r["Searched_UTC"] > r["Window_end"] + timedelta(days=POST_LAG_DAYS))

    if args.backfill:
        todo = [w for w in reversed(windows(START, today)) if not done(w[0])]
    else:
        first = (max(daily.index) - timedelta(days=REVISION_DAYS)) if len(daily) else START
        todo = list(reversed(windows(max(first, START), today)))
    print(f"{len(todo)} window(s) to search, at most {args.max_searches} this run")

    src = Nepco()
    n_search, cut, total_new = 0, None, 0
    try:
        rows = src.latest()
        new, ch = store_rows(daily, rows, today)
        total_new += new
        print(f"Latest page: {len(rows)} rows ({rows[0][0] if rows else '-'} newest), {new} new, {ch} revised")
        for ws, we in todo:
            if n_search >= args.max_searches:
                break
            time.sleep(PAUSE_S)
            end = min(we, today)
            n_search += 1
            rows = src.search(ws, end)
            inside = [(d, v) for d, v in rows if ws <= d <= end]
            note = ""
            if len(inside) != len(rows):
                note = f"{len(rows) - len(inside)} row(s) outside the window ignored"
            if len(rows) >= 10 and len(inside) < len(rows):
                # the filter was not applied (unfiltered list): not a valid answer for this window
                src.state = None
                print(f"  {ws}..{end}: filter not applied ({note}); not recorded", flush=True)
                continue
            new, ch = store_rows(daily, inside, today)
            total_new += new
            wins.loc[ws] = [we, today, len(inside), "ok" if inside else "empty", note]
            print(f"  search {n_search:2d} {ws}..{end}: {len(inside)} rows, {new} new, {ch} revised "
                  f"(requests {src.requests}){' ' + note if note else ''}", flush=True)
    except CutOff as e:
        cut = str(e)
        print(f"CUT-OFF after {n_search} search(es), {src.requests} request(s): {cut} - saving and stopping", flush=True)
    except KeyboardInterrupt:
        cut = "interrupted"

    save(args.out, daily, wins, today)
    left = [w for w in windows(START, today) if not done(w[0])]
    print(f"Saved {len(daily)} days ({min(daily.index) if len(daily) else '-'} to "
          f"{max(daily.index) if len(daily) else '-'}), {total_new} new; {len(wins)} windows searched, "
          f"{len(left)} not yet searched/complete -> {args.out}")
    if cut and n_search == 0 and not total_new:
        sys.exit(f"Blocked before any search: {cut}")


if __name__ == "__main__":
    main()
