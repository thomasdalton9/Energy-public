"""
Nepal electricity, daily, from the Nepal Electricity Authority (NEA), System Operation Department, Load Dispatch
Centre (LDC): the "Daily Operational Report of Nepal" (NDOR, one PDF per day) plus, for days after the published
archive ends, the "Energy Details" panel on NEA's home page. Found via discovery_archive/asia/NEPAL_DISCOVERY*.py.

Where the reports are
  NEA's old site (www.nea.org.np/admin/assets/uploads/ldc/NDOR <BS date>.pdf, also mirrored as cms.nea.org.np) is
  gone: those paths now 404. The new site lists the reports under the Transmission Directorate:
    https://transd.nea.org.np/en/category/daily-operational-reports-1    (64 listing pages)
  each detail page linking  https://transd.nea.org.np/uploads/shares/Daily_op_reports/NDOR%20YYYY_MM_DD.pdf
  where YYYY_MM_DD is the Bikram Sambat (BS) date. *.nea.org.np sits behind an F5 bot challenge that plain requests
  cannot pass (they get a 5.6 KB JavaScript page); a headless browser can. The site's developer hosts an identical
  copy at td.neasite.dryicesolutions.net (same files, byte for byte, no challenge), which is tried first; the
  browser route to transd.nea.org.np is the fallback.
  The published archive runs 2080-01-01 .. 2081-09-27 BS = 14 Apr 2023 .. 11 Jan 2025 (a few days missing); NEA has
  posted none since. No NDOR file exists for earlier dates on the new site.

  BS dates: the `nepali-datetime` package (nepali_datetime.date.from_datetime_date) builds each day's file name;
  each report also prints its own date ("For Date: 2081/09/27 ( 2025/01/11 )"), which is checked against it (the
  BS date rules: the printed AD date is hand-typed and wrong in 34 reports - see read_pdf).

Each report gives
  Daily energy (MWh): NEA, NEA Subsidiary, IPP generation, Import, Total Energy Available, Energy Export,
  Net Energy Met within the country (INPS demand), Energy Interruption, Energy not served / Generation Deficit,
  Energy Requirement, net exchange with India.
  Peak (MW): peak time, generation, import, recorded peak availability, export, demand met at peak, interruption,
  deficit, peak demand (requirement), net exchange with India.

Monthly reports (NMOR): https://transd.nea.org.np/en/category/monthly-operational-reports, 20 files, Shrawan 2079 ..
  Falgun 2080 BS (Jul 2022 .. Mar 2024), PDFs uploads/shares/Monthly_op_Report/NMOR%20YYYY_MM.pdf ('-rev1' files
  replace originals). Each lists every day of the month with the same energy fields and a peak table. Their daily
  rows fill days with no daily report (Source 'NMOR'): this extends the series back to 17 Jul 2022. On days both
  exist the daily report is kept and differences are printed. Nothing after Jan 2025 exists in any category.

After the archive: NEA's home page (https://nea.org.np/en, behind the same challenge) shows "Energy Details" for the
latest day - NEA / NEA Subsidiary Companies / IPP / Import / Export / Interruption (MWh), Total Energy Demand and
National Energy Demand (MWh), Total Peak Demand and National Peak Demand (MW) - with no history and no date. Each
run saves it as the previous Nepal day (Source column says so); a panel identical to the last one saved is not
saved again. That part needs playwright + chromium and runs daily.

Writes output/Data and Chart Outputs/nepal_power_generation_daily.xlsx:
  Daily   date (Gregorian), MWh: Hydro_MWh (= NEA + NEA_subsidiary + IPP: Nepal's generation is hydro bar a few
          tens of MW of solar, which the reports do not split out), Total_MWh (the same), NEA_MWh,
          NEA_subsidiary_MWh, IPP_MWh, Imports_MWh, Exports_MWh, Net_imports_MWh, Energy_available_MWh,
          Energy_met_MWh, Interruption_MWh, Deficit_MWh, Energy_requirement_MWh, Source
  Demand  date, MW: Demand_peak_MW (peak demand requirement, national), Demand_met_peak_MW, Demand_avg_MW
          (Energy_met / 24), Peak_generation_MW, Peak_import_MW, Peak_export_MW, Peak_interruption_MW,
          Peak_deficit_MW, Peak_time

Incremental: the workbook is the history store. The first run (or --rebuild) tries every day from 14 Apr 2023;
after that only the last GAP_DAYS days without a report are tried. Checkpoint every BATCH reports.

    python3 asia/NEPAL_NEA_LDC.py                     # reports for missing days + today's home-page panel
    python3 asia/NEPAL_NEA_LDC.py --no-panel          # reports only (no browser needed unless the mirror fails)
    python3 asia/NEPAL_NEA_LDC.py --no-pdfs           # home-page panel only
    python3 asia/NEPAL_NEA_LDC.py --start 2024-01-01 --end 2024-01-20 --out /tmp/x.xlsx   # test window
"""
import argparse
import io
import os
import re
import sys
import time
from concurrent.futures import ProcessPoolExecutor, ThreadPoolExecutor
from datetime import date, datetime, timedelta, timezone

import nepali_datetime
import pandas as pd
import pdfplumber
import requests
import urllib3

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)
import xlsx_notes  # noqa: E402

urllib3.disable_warnings()
UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36"
H = {"User-Agent": UA}
T = (20, 120)
MIRROR = "https://td.neasite.dryicesolutions.net"
TRANSD = "https://transd.nea.org.np"
PDF_PATH = "/uploads/shares/Daily_op_reports/NDOR%20{}.pdf"
LISTING = TRANSD + "/en/category/daily-operational-reports-1"
HOME = "https://nea.org.np/en"
DATA_START = date(2023, 4, 14)   # 2080-01-01 BS, the first report on the new site
GAP_DAYS = 45
BATCH = 150
NPT = timezone(timedelta(hours=5, minutes=45))
OUT = os.path.join(ROOT, "output", "Data and Chart Outputs", "nepal_power_generation_daily.xlsx")
SRC_PDF = "NDOR"
SRC_PANEL = "NEA home page (latest day)"
SRC_NMOR = "NMOR"
NMOR_LISTING = "/en/category/monthly-operational-reports"
NMOR_ENERGY = ["NEA_MWh", "NEA_subsidiary_MWh", "IPP_MWh", "Imports_MWh", "Energy_available_MWh", "Exports_MWh",
               "Energy_met_MWh", "Interruption_MWh", "Deficit_MWh", "Energy_requirement_MWh"]
# peak rows: time, generation, import, recorded peak availability, export, demand met at peak, interruption,
# system peak demand, national peak demand, net exchange with India
NMOR_PEAK = ["Peak_generation_MW", "Peak_import_MW", "Peak_availability_MW", "Peak_export_MW", "Demand_met_peak_MW",
             "Peak_interruption_MW", "System_peak_MW", "Demand_peak_MW", "Peak_net_exchange_MW"]

DAILY = ["Hydro_MWh", "Total_MWh", "NEA_MWh", "NEA_subsidiary_MWh", "IPP_MWh", "Imports_MWh", "Exports_MWh",
         "Net_imports_MWh", "Energy_available_MWh", "Energy_met_MWh", "Interruption_MWh", "Deficit_MWh",
         "Energy_requirement_MWh", "Source"]
DEMAND = ["Demand_peak_MW", "System_peak_MW", "Demand_met_peak_MW", "Demand_avg_MW", "Peak_generation_MW", "Peak_import_MW",
          "Peak_export_MW", "Peak_interruption_MW", "Peak_deficit_MW", "Peak_time"]
# header text -> field, first match wins (more specific phrases first)
ENERGY_KEYS = [("net value", "Net_exchange_MWh"), ("subsidiary", "NEA_subsidiary_MWh"), ("ipp", "IPP_MWh"),
               ("total energy", "Energy_available_MWh"), ("net energy met", "Energy_met_MWh"),
               ("interruption", "Interruption_MWh"), ("not served", "Deficit_MWh"), ("deficit", "Deficit_MWh"),
               ("requirement", "Energy_requirement_MWh"), ("export", "Exports_MWh"), ("import", "Imports_MWh"),
               ("nea", "NEA_MWh")]
PEAK_KEYS = [("net value", "Peak_net_exchange_MW"), ("demand met", "Demand_met_peak_MW"), ("peak time", "Peak_time"),
             ("recorded", "Peak_availability_MW"), ("availability", "Peak_availability_MW"),
             ("requirement", "Demand_peak_MW"), ("peak demand", "Demand_peak_MW"), ("generation", "Peak_generation_MW"),
             ("interruption", "Peak_interruption_MW"), ("deficit", "Peak_deficit_MW"), ("export", "Peak_export_MW"),
             ("import", "Peak_import_MW")]
# the order the 2080-2081 reports print their two value rows in (fallback when the tables do not extract)
ENERGY_ORDER = ["NEA_MWh", "NEA_subsidiary_MWh", "IPP_MWh", "Imports_MWh", "Energy_available_MWh", "Exports_MWh",
                "Energy_met_MWh", "Interruption_MWh", "Deficit_MWh", "Energy_requirement_MWh", "Net_exchange_MWh"]
PEAK_ORDER = ["Peak_generation_MW", "Peak_import_MW", "Peak_availability_MW", "Peak_export_MW",
              "Demand_met_peak_MW", "Peak_interruption_MW", "Peak_deficit_MW", "Demand_peak_MW",
              "Peak_net_exchange_MW"]
NUM = r"-?\d[\d,]*(?:\.\d+)?"


def out(*a):
    print(*a, flush=True)


def bs(g):
    d = nepali_datetime.date.from_datetime_date(g)
    return d.year, d.month, d.day


def bs_name(g):
    y, m, d = bs(g)
    return f"{y}_{m:02d}_{d:02d}"


def num(x):
    try:
        return float(str(x).replace(",", "").strip())
    except (TypeError, ValueError):
        return None


# ---------------------------------------------------------------- fetching
class Browser:
    """Headless Chromium for the F5-protected *.nea.org.np pages (started only when needed)."""

    def __init__(self):
        self.pw = self.b = self.ctx = None

    def start(self):
        if self.ctx is None:
            from playwright.sync_api import sync_playwright
            self.pw = sync_playwright().start()
            self.b = self.pw.chromium.launch(args=["--disable-blink-features=AutomationControlled"])
            self.ctx = self.b.new_context(user_agent=UA, ignore_https_errors=True)
        return self.ctx

    def page(self, url, marker, tries=4):
        """HTML of url once `marker` shows (the challenge page reloads itself into the real one)."""
        ctx = self.start()
        for k in range(tries):
            pg = ctx.new_page()
            try:
                pg.goto(url, wait_until="domcontentloaded", timeout=60000)
                for _ in range(25):
                    pg.wait_for_timeout(1500)
                    try:
                        html = pg.content()
                    except Exception:  # noqa: BLE001 - the challenge page is still navigating
                        continue
                    if marker in html:
                        return html
                out(f"  {url}: '{marker}' not shown after 37 s (try {k + 1})")
            except Exception as e:  # noqa: BLE001
                out(f"  {url}: {type(e).__name__}: {str(e)[:120]}")
            finally:
                pg.close()
        return None

    def get(self, url):
        ctx = self.start()
        for i in range(3):
            try:
                r = ctx.request.get(url, timeout=90000)
                return r.status, r.body()
            except Exception as e:  # noqa: BLE001
                out(f"  retry {url[-40:]}: {type(e).__name__}")
                time.sleep(5 * (i + 1))
        return None, b""

    def close(self):
        if self.b:
            self.b.close()
        if self.pw:
            self.pw.stop()


def mirror_get(url):
    """(status, body) from the plain-HTTP mirror; status None on a network failure."""
    for i in range(3):
        try:
            r = requests.get(url, headers=H, timeout=T, verify=False)
            return r.status_code, r.content
        except requests.RequestException as e:
            if i == 2:
                out(f"  {url[-40:]}: {type(e).__name__}")
                return None, b""
            time.sleep(5 * (i + 1))


def mirror_ok():
    status, body = mirror_get(MIRROR + PDF_PATH.format("2080_01_01"))
    return status == 200 and body[:4] == b"%PDF"


# ---------------------------------------------------------------- parsing
def pdf_date(text):
    """'For Date: 2081/09/27 ( 2025/01/11 )' -> ((2081, 9, 27), date(2025, 1, 11))."""
    m = re.search(r"For\s*Date\s*:?\s*(\d{4})\s*[/\-.]\s*(\d{1,2})\s*[/\-.]\s*(\d{1,2})\s*\(\s*(\d{4})\s*[/\-.]\s*"
                  r"(\d{1,2})\s*[/\-.]\s*(\d{1,2})", text)
    if not m:
        return None, None
    v = [int(x) for x in m.groups()]
    try:
        return tuple(v[:3]), date(v[3], v[4], v[5])
    except ValueError:
        return tuple(v[:3]), None


def map_table(table, keys):
    """pdfplumber table (header rows, then one value row) -> {field: value}. Each column is labelled by its
    lowest non-empty header cell (the spanning group titles sit above)."""
    rows = [r for r in table if r and any(c not in (None, "") for c in r)]
    if len(rows) < 2:
        return {}
    vals = rows[-1]
    rec = {}
    for j, v in enumerate(vals):
        label = ""
        for r in reversed(rows[:-1]):
            if j < len(r) and r[j] not in (None, ""):
                label = re.sub(r"\s+", " ", str(r[j])).lower()
                break
        if not label or v in (None, ""):
            continue
        for k, field in keys:
            if k in label:
                if field not in rec:
                    rec[field] = str(v).strip() if field == "Peak_time" else num(v)
                break
    return rec


def parse_text(text):
    """The two value rows straight from the text (fast): the energy row has 11 numbers, the peak row a time and 9."""
    rec = {}
    for line in text.split("\n"):
        v = re.findall(NUM, line)
        if len(v) == len(ENERGY_ORDER) and not re.search(r"\d:\d", line):
            rec.update({k: num(x) for k, x in zip(ENERGY_ORDER, v)})
            break
    m = re.search(r"^\s*(\d{1,2}:\d{2})\s+((?:" + NUM + r"\s+){8}" + NUM + r")\s*$", text, re.M)
    if m:
        rec["Peak_time"] = m.group(1)
        rec.update({k: num(x) for k, x in zip(PEAK_ORDER, m.group(2).split())})
    return rec


def parse_tables(tables):
    """Columns matched by their header text (slow: only when the text rows do not add up)."""
    rec = {}
    for t in tables:
        flat = " ".join(str(c) for r in t for c in r if c).lower()
        if "ipp" in flat and "subsidiary" in flat:
            rec.update(map_table(t, ENERGY_KEYS))
        elif "peak time" in flat:
            rec.update(map_table(t, PEAK_KEYS))
    return rec


def consistent(rec):
    """The report's own identities: NEA + subsidiary + IPP + import = available; available - export = met;
    met + interruption + deficit = requirement (each within 1% or 5 MWh)."""
    def close(a, b):
        return abs(a - b) <= max(5, 0.01 * abs(b))
    try:
        return (close(rec["NEA_MWh"] + rec["NEA_subsidiary_MWh"] + rec["IPP_MWh"] + rec["Imports_MWh"],
                      rec["Energy_available_MWh"])
                and close(rec["Energy_available_MWh"] - rec["Exports_MWh"], rec["Energy_met_MWh"])
                and close(rec["Energy_met_MWh"] + rec["Interruption_MWh"] + rec["Deficit_MWh"],
                          rec["Energy_requirement_MWh"]))
    except (KeyError, TypeError):
        return False


def peak_consistent(rec):
    """Peak row identities: generation + import = availability; availability - export = demand met;
    met + interruption + deficit = peak demand (within 2% or 5 MW)."""
    def close(a, b):
        return abs(a - b) <= max(5, 0.02 * abs(b))
    try:
        return (close(rec["Peak_generation_MW"] + rec["Peak_import_MW"], rec["Peak_availability_MW"])
                and close(rec["Peak_availability_MW"] - rec["Peak_export_MW"], rec["Demand_met_peak_MW"])
                and close(rec["Demand_met_peak_MW"] + rec["Peak_interruption_MW"] + rec["Peak_deficit_MW"],
                          rec["Demand_peak_MW"]))
    except (KeyError, TypeError):
        return False


def check(rec, label):
    """Report (do not drop) rows whose parts do not add up."""
    g = [rec.get(k) for k in ("NEA_MWh", "NEA_subsidiary_MWh", "IPP_MWh", "Imports_MWh")]
    if all(x is not None for x in g) and rec.get("Energy_available_MWh"):
        if abs(sum(g) - rec["Energy_available_MWh"]) > max(5, 0.01 * rec["Energy_available_MWh"]):
            out(f"  {label}: NEA+subsidiary+IPP+import {sum(g):.0f} vs total available "
                f"{rec['Energy_available_MWh']:.0f}")
    if None not in (rec.get("Energy_available_MWh"), rec.get("Exports_MWh"), rec.get("Energy_met_MWh")):
        if abs(rec["Energy_available_MWh"] - rec["Exports_MWh"] - rec["Energy_met_MWh"]) > \
                max(5, 0.01 * rec["Energy_met_MWh"]):
            out(f"  {label}: available - export {rec['Energy_available_MWh'] - rec['Exports_MWh']:.0f} vs energy met "
                f"{rec['Energy_met_MWh']:.0f}")


def read_pdf(body, g):
    """(date, record) from one report; the PDF's own AD date wins (mismatches are printed). Values come from the
    text rows when they satisfy the report's identities, else from the tables matched by header."""
    with pdfplumber.open(io.BytesIO(body)) as pdf:
        p = pdf.pages[0]
        text = p.extract_text() or ""
        bsd, ad = pdf_date(text)
        rec = parse_text(text)
        if not consistent(rec) or not peak_consistent(rec):
            tab = parse_tables(p.extract_tables())
            if consistent(tab) or not consistent(rec):
                rec.update({k: v for k, v in tab.items() if k in ENERGY_ORDER})
            if peak_consistent(tab) or not peak_consistent(rec):
                rec.update({k: v for k, v in tab.items() if k not in ENERGY_ORDER})
            if not (consistent(rec) and peak_consistent(rec)):
                out(f"  {bs_name(g)}: report identities do not hold after the table read - kept, check")
    # Which date: the report prints 'For Date: <BS> ( <AD> )'. Its AD date is typed by hand and is sometimes a
    # month or a few days off (e.g. 2080/05/15 printed with 2023/08/01 instead of 2023/09/01), so the BS date rules:
    #  - report BS = file BS            -> that day (an AD typo is printed and ignored)
    #  - report BS != file BS, but the report's BS and AD agree -> the report's day (a file posted under the
    #    wrong name, usually a copy of the previous day; duplicates are resolved by the caller)
    #  - otherwise                      -> the file's day, printed
    when = g
    if bsd is None:
        out(f"  file {bs_name(g)}: no 'For Date' line, date from the file name")
    elif bsd == bs(g):
        if ad and ad != g:
            out(f"  file {bs_name(g)} = {g}: report prints AD {ad} (typo), {g} used")
    else:
        try:
            bsd_ad = nepali_datetime.date(*bsd).to_datetime_date()
        except Exception:  # noqa: BLE001
            bsd_ad = None
        if bsd_ad and bsd_ad == ad:
            out(f"  file {bs_name(g)} = {g} holds the report for {bsd} = {ad}: report date used")
            when = ad
        else:
            out(f"  file {bs_name(g)} = {g}: report prints {bsd} / {ad}, inconsistent - file date used")
    if rec.get("NEA_MWh") is None:
        out(f"  {bs_name(g)}: values not found; first lines: {text[:300]!r}")
        return when, None
    if not rec.get("NEA_MWh") and not rec.get("IPP_MWh"):
        out(f"  {when} ({bs_name(g)}): report shows zero generation - not saved")
        return when, None
    rec["Source"] = SRC_PDF
    check(rec, f"{when} ({bs_name(g)})")
    return when, rec


def fetch_mirror(g):
    status, body = mirror_get(MIRROR + PDF_PATH.format(bs_name(g)))
    if status is None:
        return g, None, "error"
    if body[:4] != b"%PDF":
        return g, None, "missing"
    try:
        when, rec = read_pdf(body, g)
        return when, rec, "ok" if rec else "unparsed"
    except Exception as e:  # noqa: BLE001
        out(f"  {bs_name(g)}: {type(e).__name__}: {e}")
        return g, None, "unparsed"


def fetch_browser(br, g):
    status, body = br.get(TRANSD + PDF_PATH.format(bs_name(g)))
    if status is None:
        return g, None, "error"
    if body[:4] != b"%PDF":
        return g, None, "missing"
    try:
        when, rec = read_pdf(body, g)
        return when, rec, "ok" if rec else "unparsed"
    except Exception as e:  # noqa: BLE001
        out(f"  {bs_name(g)}: {type(e).__name__}: {e}")
        return g, None, "unparsed"


# ---------------------------------------------------------------- monthly reports (NMOR)
def nmor_files(br=None):
    """{(BS year, month): pdf url} from the monthly-report listing (a '-revN' file replaces the original)."""
    files, rank = {}, {}
    for page in range(1, 10):
        url = f"{MIRROR}{NMOR_LISTING}?page={page}"
        if br is None:
            status, body = mirror_get(url)
            html = body.decode("utf-8", "replace") if status == 200 else ""
        else:
            html = br.page(url.replace(MIRROR, TRANSD), "/detail/") or ""
        slugs = [x for x in dict.fromkeys(re.findall(r'href="[^"]+/detail/(nepal-monthly-operational-report-[^"]+)"',
                                                    html))]
        new = [x for x in slugs if all(x not in v for v in rank.values())]
        if not new:
            break
        for slug in new:
            m = re.search(r"(\d{4})-(\d{2})(?:-rev(\d+))?$", slug)
            if not m:
                out(f"  NMOR listing: unrecognised item {slug}")
                continue
            key, rev = (int(m.group(1)), int(m.group(2))), int(m.group(3) or 0)
            if key in rank and rank[key][0] >= rev:
                continue
            durl = f"{MIRROR}/en/detail/{slug}"
            if br is None:
                status, body = mirror_get(durl)
                dhtml = body.decode("utf-8", "replace") if status == 200 else ""
            else:
                dhtml = br.page(durl.replace(MIRROR, TRANSD), ".pdf") or ""
            pdfs = re.findall(r'href="([^"]+\.pdf)"', dhtml)
            if pdfs:
                files[key] = pdfs[0]
                rank[key] = (rev, slug)
    return files


def parse_nmor(body, key):
    """Daily rows of one monthly report: {date: record}. Rows are 'dd/mm/yyyy(BS) dd/mm/yyyy(AD) <10 energy values>'
    and, in the peak table, 'BS AD hh:mm <9 values>'. The BS date rules (the AD column is typed by hand)."""
    rows, peaks, total = {}, {}, None
    with pdfplumber.open(io.BytesIO(body)) as pdf:
        text = "\n".join((p.extract_text() or "") for p in pdf.pages)
    dpat = r"(\d{1,2})/(\d{1,2})/(\d{4})\s+(\d{1,2})/(\d{1,2})/(\d{4})\s+"
    for line in text.split("\n"):
        m = re.match(r"\s*" + dpat + r"(.*)$", line)
        if not m:
            t = re.match(r"\s*Total\s+(.*)$", line)
            if t and total is None:
                total = [num(x) for x in re.findall(NUM, t.group(1))]
            continue
        bd, bm, by, ad_d, ad_m, ad_y = (int(x) for x in m.groups()[:6])
        rest = m.group(7)
        try:
            ad = date(ad_y, ad_m, ad_d)
        except ValueError:
            ad = None
        # the BS date rules, but it must belong to the report's month: a row typed '29/11/2080' in the Falgun 2079
        # report is 29/11/2079 (its AD column confirms it)
        when = None
        for y, mo in ([(by, bm)] if (by, bm) == key else [key, (by, bm)]):
            try:
                cand = nepali_datetime.date(y, mo, bd).to_datetime_date()
            except Exception:  # noqa: BLE001
                continue
            if (y, mo) == key or cand == ad:
                when = cand
                break
        if when is None:
            out(f"  NMOR {key}: row {line[:40]!r} is not in the report's month - skipped")
            continue
        if (by, bm) != key:
            out(f"  NMOR {key}: row typed {by}/{bm:02d}/{bd:02d} read as {key[0]}/{key[1]:02d}/{bd:02d} = {when}")
        elif ad and ad != when:
            out(f"  NMOR {key}: row {by}/{bm:02d}/{bd:02d} prints AD {ad}, {when} used")
        pm = re.match(r"(\d{1,2}:\d{2})\s+(.*)$", rest)
        if pm:
            v = [num(x) for x in re.findall(NUM, pm.group(2))]
            if len(v) == len(NMOR_PEAK):
                peaks[when] = dict(zip(NMOR_PEAK, v), Peak_time=pm.group(1))
            else:
                out(f"  NMOR {key}: peak row with {len(v)} values: {line[:90]!r}")
            continue
        v = [num(x) for x in re.findall(NUM, rest)]
        if len(v) in (10, 11):
            rows[when] = dict(zip(NMOR_ENERGY, v[:10]))
        elif len(v) == 3:
            pass   # the 2079 reports add an import / export / net exchange table: already in the energy rows
        else:
            out(f"  NMOR {key}: energy row with {len(v)} values: {line[:90]!r}")
    if total and len(total) >= 10 and rows:
        s = sum(r["Energy_met_MWh"] or 0 for r in rows.values())
        if abs(s - total[6]) > max(5, 0.005 * total[6]):
            out(f"  NMOR {key}: daily energy met sums to {s:.0f}, Total row says {total[6]:.0f}")
    outp = {}
    for when, rec in rows.items():
        rec.update(peaks.get(when, {}))
        if not consistent(rec):
            out(f"  NMOR {key} {when}: energy identities do not hold - row skipped")
            continue
        if not rec.get("NEA_MWh") and not rec.get("IPP_MWh"):
            continue
        if not rec.get("System_peak_MW"):   # the 2079 reports leave the system peak column at 0
            rec.pop("System_peak_MW", None)
        rec["Source"] = SRC_NMOR
        outp[when] = rec
    return outp


def fetch_nmor(item):
    key, url = item
    status, body = mirror_get(url)
    if body[:4] != b"%PDF":   # some detail pages link a wrong folder ('Annual Report'): try the monthly folder
        alt = f"{MIRROR}/uploads/shares/Monthly_op_Report/{url.rsplit('/', 1)[-1]}"
        status, body = mirror_get(alt)
    if body[:4] != b"%PDF":
        out(f"  NMOR {key}: no PDF at {url}")
        return key, {}
    try:
        return key, parse_nmor(body, key)
    except Exception as e:  # noqa: BLE001
        out(f"  NMOR {key}: {type(e).__name__}: {e}")
        return key, {}


# ---------------------------------------------------------------- home-page panel
PANEL = {"nea": "NEA_MWh", "nea subsidiary companies": "NEA_subsidiary_MWh", "ipp": "IPP_MWh",
         "import": "Imports_MWh", "export": "Exports_MWh", "interruption": "Interruption_MWh",
         "total energy demand": "Panel_total_demand_MWh", "national energy demand": "Energy_requirement_MWh",
         "total peak demand": "Panel_total_peak_MW", "national peak demand": "Demand_peak_MW"}


def panel(br):
    html = br.page(HOME, "Energy Details")
    if not html:
        out("Home page 'Energy Details' panel not reached")
        return None
    i = html.find("Energy Details")
    seg = html[i:i + 12000]
    rec = {}
    for v, unit, label in re.findall(r"<span[^>]*>\s*(" + NUM + r")\s*(MWh|MW)\s*</span>\s*<span[^>]*>([^<]+)</span>",
                                     seg):
        k = PANEL.get(re.sub(r"\s+", " ", label).strip().lower())
        if k:
            rec[k] = num(v)
    if not all(k in rec for k in ("NEA_MWh", "IPP_MWh", "Energy_requirement_MWh")):
        flat = re.sub(r"<[^>]+>|\s+", " ", seg)
        out(f"Panel not parsed: {flat[:400]}")
        return None
    rec["Energy_available_MWh"] = sum(rec.get(k) or 0 for k in ("NEA_MWh", "NEA_subsidiary_MWh", "IPP_MWh",
                                                                 "Imports_MWh"))
    if rec.get("Interruption_MWh") is not None:
        rec["Energy_met_MWh"] = rec["Energy_requirement_MWh"] - rec["Interruption_MWh"]
    tot = rec.get("Panel_total_demand_MWh")
    if tot and abs(rec["Energy_available_MWh"] + (rec.get("Interruption_MWh") or 0) - tot) > 0.01 * tot:
        out(f"  panel: generation + import + interruption {rec['Energy_available_MWh']:.0f} vs Total Energy Demand {tot}")
    rec["Source"] = SRC_PANEL
    return rec


# ---------------------------------------------------------------- workbook
def read_saved(path):
    try:
        d = pd.read_excel(path, sheet_name="Daily", index_col=0)
        try:
            dm = pd.read_excel(path, sheet_name="Demand", index_col=0)
            d = d.join(dm[[c for c in dm.columns if c not in d.columns]], how="outer")
        except ValueError:
            pass
    except (FileNotFoundError, ValueError):
        return pd.DataFrame()
    d.index = pd.to_datetime(d.index)
    return d.sort_index()


def save(path, new_rows, rebuild=False):
    old = pd.DataFrame() if rebuild else read_saved(path)
    new = pd.DataFrame.from_dict(new_rows, orient="index")
    if not new.empty:
        new.index = pd.to_datetime(new.index)
    d = new if old.empty else (old if new.empty else pd.concat([old[~old.index.isin(new.index)], new]))
    if d.empty:
        out("Nothing saved yet and nothing parsed this run")
        return d
    d = d.sort_index()
    for c in ("NEA_MWh", "NEA_subsidiary_MWh", "IPP_MWh"):
        if c not in d:
            d[c] = None
    d["Hydro_MWh"] = d[["NEA_MWh", "NEA_subsidiary_MWh", "IPP_MWh"]].sum(axis=1, min_count=1)
    d["Total_MWh"] = d["Hydro_MWh"]
    if "Imports_MWh" in d and "Exports_MWh" in d:
        d["Net_imports_MWh"] = d["Imports_MWh"] - d["Exports_MWh"]
    if "Energy_met_MWh" in d:
        d["Demand_avg_MW"] = d["Energy_met_MWh"] / 24
    daily = d[[c for c in DAILY if c in d]].copy()
    demand = d[[c for c in DEMAND if c in d]].copy()
    num_cols = [c for c in daily if c != "Source"]
    daily[num_cols] = daily[num_cols].apply(pd.to_numeric, errors="coerce").round(1)
    nd = [c for c in demand if c != "Peak_time"]
    demand[nd] = demand[nd].apply(pd.to_numeric, errors="coerce").round(1)
    demand = demand.dropna(how="all", subset=nd) if nd else demand
    daily.index.name = demand.index.name = "date"
    pdf_days = daily.index[daily["Source"] == SRC_PDF]
    nmor_days = daily.index[daily["Source"] == SRC_NMOR]
    panel_days = daily.index[daily["Source"] == SRC_PANEL]
    notes = [
        "UNITS",
        "Daily: MWh per day (Nepal days, 00:00-24:00 NPT). Hydro_MWh = NEA_MWh + NEA_subsidiary_MWh + IPP_MWh = all "
        "domestic generation on the integrated grid (Nepal's generation is hydro apart from a few tens of MW of solar, "
        "which the reports do not show separately); Total_MWh the same. NEA_MWh = NEA's own plants; "
        "NEA_subsidiary_MWh = NEA subsidiary companies (e.g. Chilime, Upper Tamakoshi); IPP_MWh = independent power "
        "producers. Imports_MWh / Exports_MWh = cross-border energy with India (Net_imports_MWh = import - export). "
        "Energy_available_MWh = generation + import; Energy_met_MWh = energy met within Nepal (INPS demand = available "
        "- export); Interruption_MWh = load shed / interrupted; Deficit_MWh = energy not served (generation deficit); "
        "Energy_requirement_MWh = met + interruption + deficit.",
        "Demand: MW. Demand_peak_MW = peak demand (requirement) of the day; Demand_met_peak_MW = demand met at the peak "
        "time; Demand_avg_MW = Energy_met_MWh / 24; Peak_generation / import / export / interruption / deficit at the "
        "peak time (Peak_time, NPT).",
        "Source column: 'NDOR' = the LDC daily operational report (PDF); 'NMOR' = a daily row of the LDC monthly "
        "operational report (used only for days with no daily report; same fields, no deficit at peak; its "
        "System_peak_MW = system peak demand, Demand_peak_MW = national peak demand); '" + SRC_PANEL + "' = the 'Energy Details' panel "
        "on NEA's home page, which shows one undated day and keeps no history: saved each run as the previous Nepal "
        "day (an assumption - NEA publishes the previous day's figures). From the panel: Energy_requirement_MWh = its "
        "'National Energy Demand', Demand_peak_MW = 'National Peak Demand', Energy_met_MWh = requirement - "
        "interruption (deficit not shown); its 'Total' demand figures include exports and are not kept.",
        "Dates: each report's file name carries the Bikram Sambat date (NDOR 2081_09_27 = 11 Jan 2025); the report "
        "prints both dates. The BS date rules: the printed Gregorian date is hand-typed and was a month or two days off "
        "in Sep 2023 and Sep 2024 (ignored); a file holding another day's report (2080_08_19 = copy of 2080_08_18) "
        "counts for the day it states. Reports showing zero generation (9-12 Aug 2024) are left out. BS->AD "
        "conversion: nepali-datetime package.",
        "",
        "COVERAGE",
        f"Daily from {daily.index.min():%Y-%m-%d} to {daily.index.max():%Y-%m-%d}: {len(pdf_days)} days from reports"
        + (f" ({pdf_days.min():%Y-%m-%d}..{pdf_days.max():%Y-%m-%d})" if len(pdf_days) else "")
        + f", {len(nmor_days)} from monthly reports"
        + (f" ({nmor_days.min():%Y-%m-%d}..{nmor_days.max():%Y-%m-%d})" if len(nmor_days) else "")
        + f", {len(panel_days)} from the home-page panel"
        + (f" ({panel_days.min():%Y-%m-%d}..{panel_days.max():%Y-%m-%d})" if len(panel_days) else "") + ". "
        "NEA's published report archive covers 2080-01-01..2081-09-27 BS (14 Apr 2023 - 11 Jan 2025) with a few days "
        "missing; the monthly reports (NMOR, daily rows) cover Shrawan 2079 - Falgun 2080 BS (Jul 2022 - Mar 2024). "
        "Nothing earlier is online since NEA's site moved, and nothing after Jan 2025 has been posted (checked: daily, "
        "monthly and yearly categories, Oct 2026). Days between the end of the archive and the first panel snapshot "
        "are empty.",
        "",
        "SOURCE",
        "Nepal Electricity Authority, System Operation Department, Load Dispatch Centre: Daily Operational Report of "
        f"Nepal, listed at {LISTING} (PDFs {TRANSD}/uploads/shares/Daily_op_reports/NDOR%20<BS date>.pdf; fetched from "
        f"the site developer's identical copy {MIRROR} where reachable, because *.nea.org.np blocks non-browser "
        f"clients). Latest day: 'Energy Details' on {HOME}.",
    ]
    xlsx_notes.write_workbook(path, {"Daily": daily, "Demand": demand}, notes, {"UNITS", "COVERAGE", "SOURCE"})
    out(f"Saved {path}: {len(daily)} days {daily.index.min():%Y-%m-%d}..{daily.index.max():%Y-%m-%d} "
        f"({len(pdf_days)} reports, {len(panel_days)} panel)")
    return d


# ---------------------------------------------------------------- main
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", default=OUT)
    ap.add_argument("--rebuild", action="store_true", help="ignore the saved workbook, re-read every report")
    ap.add_argument("--start", help="test window start (YYYY-MM-DD): try only these days, ignore what is saved")
    ap.add_argument("--end", help="test window end (YYYY-MM-DD)")
    ap.add_argument("--no-pdfs", action="store_true", help="skip the report PDFs (home-page panel only)")
    ap.add_argument("--no-panel", action="store_true", help="skip the home-page panel (no browser)")
    args = ap.parse_args()
    saved = pd.DataFrame() if args.rebuild else read_saved(args.out)
    have = set(saved.index[saved["Source"] == SRC_PDF].date) if "Source" in saved else set()
    yesterday = datetime.now(NPT).date() - timedelta(days=1)
    br = Browser()
    done = {}
    if args.rebuild:   # panel days cannot be fetched again: carry them over
        old = read_saved(args.out)
        if "Source" in old:
            for ts, row in old[old["Source"] == SRC_PANEL].iterrows():
                done[ts.date()] = row.dropna().to_dict()
    target = args.out.replace(".xlsx", "_rebuild.xlsx") if args.rebuild else args.out
    try:
        if not args.no_pdfs:
            if args.start:
                first = date.fromisoformat(args.start)
                last = date.fromisoformat(args.end) if args.end else yesterday
                todo = [first + timedelta(days=k) for k in range((last - first).days + 1)]
            else:
                last = yesterday
                # once reports are saved only the last GAP_DAYS are retried: days NEA never posted (and the long gap
                # after the archive ends) must not be re-tried on every run
                first = max(DATA_START, yesterday - timedelta(days=GAP_DAYS)) if have else DATA_START
                todo = [first + timedelta(days=k) for k in range((yesterday - first).days + 1)]
                todo = [g for g in todo if g not in have]
            out(f"{len(have)} report days saved; trying {len(todo)} days {todo[0] if todo else ''}.."
                f"{todo[-1] if todo else ''}")
            use_mirror = mirror_ok()
            out(f"Mirror {MIRROR}: {'reachable' if use_mirror else 'not reachable - using a browser on ' + TRANSD}")
            tally = {}
            for b in range(0, len(todo), BATCH):
                chunk = todo[b:b + BATCH]
                if use_mirror:
                    # one process per core: parsing the PDFs is CPU-bound
                    with ProcessPoolExecutor(max_workers=max(2, os.cpu_count() or 2)) as ex:
                        res = list(ex.map(fetch_mirror, chunk))
                else:
                    if b == 0:
                        br.page(TRANSD + "/en", "Transmission")   # pass the challenge once, cookies are reused
                    res = [fetch_browser(br, g) for g in chunk]
                for when, rec, status in res:
                    tally[status] = tally.get(status, 0) + 1
                    if rec:
                        if when in done:
                            out(f"  two reports for {when}: keeping the first (its file name matches)")
                            continue
                        done[when] = rec
                save(target, done, args.rebuild)
            out(f"Reports: {tally}")
            # monthly reports: their daily rows fill days with no daily report (mostly Jul 2022 - Apr 2023)
            if use_mirror:
                files = nmor_files()
                out(f"Monthly reports listed: {len(files)} ({min(files) if files else ''}..{max(files) if files else ''})")
                have_ndor = set(have) | set(done)
                added, diffs = 0, []
                with ThreadPoolExecutor(max_workers=6) as ex:
                    for key, recs in ex.map(fetch_nmor, sorted(files.items())):
                        for when, rec in recs.items():
                            if when in have_ndor:
                                ref = done.get(when)
                                if ref is None and when in saved.index:
                                    ref = saved.loc[pd.Timestamp(when)].to_dict()
                                if ref and ref.get("IPP_MWh") is not None and rec.get("IPP_MWh") is not None and \
                                        abs(ref["IPP_MWh"] - rec["IPP_MWh"]) > 1:
                                    diffs.append((when, ref["IPP_MWh"], rec["IPP_MWh"]))
                                continue
                            if args.start and not (first <= when <= last):
                                continue
                            done[when] = rec
                            added += 1
                out(f"Monthly reports: {added} days added; {len(diffs)} overlap days where the monthly and daily "
                    f"reports differ (IPP MWh): {diffs[:10]}")
                save(target, done, args.rebuild)
            else:
                out("Monthly reports skipped (mirror not reachable)")
        if not args.no_panel:
            rec = panel(br)
            if rec:
                prev = saved[saved["Source"] == SRC_PANEL].iloc[-1] if "Source" in saved and \
                    (saved["Source"] == SRC_PANEL).any() else None
                same = prev is not None and all(prev.get(k) == rec.get(k) for k in ("NEA_MWh", "IPP_MWh",
                                                                                    "Exports_MWh", "Imports_MWh"))
                out(f"Panel (saved as {yesterday}): " + ", ".join(f"{k}={v}" for k, v in rec.items()))
                if same:
                    out("  identical to the last panel saved: not updated yet, nothing saved")
                elif yesterday in have or yesterday in done:
                    out(f"  a report already covers {yesterday}: panel not saved")
                else:
                    done[yesterday] = rec
                save(target, done, args.rebuild)
    finally:
        br.close()
    if args.rebuild:
        before = len(read_saved(args.out))
        if len(done) < 0.95 * before:
            raise SystemExit(f"Rebuild parsed {len(done)} days vs {before} saved: workbook left as it was ({target})")
        os.replace(target, args.out)
    d = read_saved(args.out)
    if d.empty:
        raise SystemExit("No NEA data parsed")
    out(d[[c for c in ("Hydro_MWh", "Imports_MWh", "Exports_MWh", "Energy_met_MWh", "Demand_peak_MW", "Source")
           if c in d]].tail(4).to_string())


if __name__ == "__main__":
    main()
