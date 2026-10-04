"""
Pakistan's main reservoirs - Tarbela (Indus), Mangla (Jhelum) and Chashma (Indus) - water level, mean inflow and
mean outflow, from the Indus River System Authority's (IRSA) Daily Water Situation reports, with WAPDA's daily
river-flow workbook as the history before them. Found via discovery_archive/asia/HYDRO_SSEA_DISCOVERY1-4.py.

  IRSA      http://pakirsa.gov.pk/DailyData.aspx lists the last ~8 daily PDFs (Doc/DataDD-MM-YYYY.pdf, one page):
            per reservoir LEVEL (ft), DEAD LEVEL (ft), MEAN INFLOW / MEAN OUTFLOW (cusecs, 24-hour means to 06:00),
            plus total rim-station inflows / outflows. Older PDFs are deleted from the server, so this runs DAILY and
            the workbook is the history store (every listed PDF not saved yet is read, so a few missed runs leave no gap).
  WAPDA     https://wapda.gov.pk/river-flow/ embeds GRAPH-DG-16-for-MAIL-2.xls ("River flows and levels", sheet
            GHRAPH): daily Tarbela and Mangla level (ft), inflow and outflow (1000 cusecs) from 2016-03-01 to Dec 2024.
            WAPDA stopped updating it in Dec 2024; it is read once, to seed the history (and again only if WAPDA
            replaces the file - its Last-Modified is recorded on the Units sheet).
  (wapda.gov.pk/river-flow-data/ is only a link page; PMD's Flood Forecasting Division, ffd.pmd.gov.pk, sits behind a
  Cloudflare challenge that GitHub's runners do not pass.)

Writes output/Data and Chart Outputs/pakistan_hydro_reservoirs.xlsx:
  Daily    date x: Tarbela_level_ft, Tarbela_inflow_cusecs, Tarbela_outflow_cusecs, the same for Mangla and Chashma
           (Chashma from IRSA only), Rim_inflow_cusecs / Rim_outflow_cusecs (Indus at Tarbela + Jhelum at Mangla + Kabul
           at Nowshera + Chenab at Marala), Source (WAPDA / IRSA)
  Limits   per reservoir: dead level and maximum conservation level (ft)
  Water year charts (Oct-Sep) of the Tarbela and Mangla levels (add_charts.py)

    python3 asia/PAKISTAN_IRSA_RESERVOIRS.py
"""
import argparse
import io
import os
import re
import sys
import time
from datetime import date, datetime

import pandas as pd
import pdfplumber
import requests
import urllib3

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)
import xlsx_notes  # noqa: E402

urllib3.disable_warnings()
IRSA = "http://pakirsa.gov.pk/"            # https does not answer
CDX = "http://web.archive.org/cdx/search/cdx"
WAYBACK_BUDGET = int(os.environ.get("WAYBACK_BUDGET", 900))   # seconds per run for the one-off Internet Archive backfill
WAPDA_XLS = "https://wapda.gov.pk/wp-content/uploads/2024/12/GRAPH-DG-16-for-MAIL-2.xls"
H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0 Safari/537.36"}
T = (20, 90)
OUT = os.path.join(ROOT, "output", "Data and Chart Outputs", "pakistan_hydro_reservoirs.xlsx")
DAMS = ["Tarbela", "Mangla", "Chashma"]
COLS = [f"{d}_{k}" for d in DAMS for k in ("level_ft", "inflow_cusecs", "outflow_cusecs")] + \
       ["Rim_inflow_cusecs", "Rim_outflow_cusecs", "Source"]
# plausible water levels (ft above mean sea level): reject a misread number
RANGE = {"Tarbela": (1370, 1560), "Mangla": (1030, 1250), "Chashma": (630, 652)}
# maximum conservation levels per WAPDA (raised Mangla since 2012); dead levels as IRSA reports them
MAX_LEVEL = {"Tarbela": 1550.0, "Mangla": 1242.0, "Chashma": 649.0}
SEG = {"Tarbela": (r"INDUS\s*@\s*TARBELA", r"CHASHMA|KALABAGH"), "Chashma": (r"CHASHMA", r"TAUNSA"),
       "Mangla": (r"JHELUM\s*@\s*MANGLA", r"PANJNAD|IRSA RELEASES|$")}


def out(*a):
    print(*a, flush=True)


def get(u, **k):
    for i in range(4):
        try:
            r = requests.get(u, headers=H, timeout=T, verify=False, **k)
            if r.status_code == 404:
                return r
            r.raise_for_status()
            return r
        except requests.RequestException as e:
            if i == 3:
                out(f"  {u} failed: {e}")
                return None
            time.sleep(5 * (i + 1))


def num(pat, txt):
    m = re.search(pat, txt)
    return float(m.group(1).replace(",", "")) if m else None


def parse_pdf(content):
    """-> ({column: value}, {dam: dead level})"""
    with pdfplumber.open(io.BytesIO(content)) as pdf:
        txt = "\n".join((p.extract_text() or "") for p in pdf.pages)
    txt = re.sub(r"[ \t]+", " ", txt)
    row, dead = {}, {}
    for dam, (a, b) in SEG.items():
        m = re.search(a, txt)
        if not m:
            continue
        seg = txt[m.end():]
        e = re.search(b, seg)
        seg = seg[:e.start()] if e and e.start() > 0 else seg
        lv = num(r"(?<!DEAD )LEVEL\s*=\s*([\d,.]+)", seg)
        lo, hi = RANGE[dam]
        row[f"{dam}_level_ft"] = lv if lv is not None and lo <= lv <= hi else None
        row[f"{dam}_inflow_cusecs"] = num(r"MEAN INFLOW\s*=\s*([\d,.]+)", seg)
        row[f"{dam}_outflow_cusecs"] = num(r"MEAN OUTFLOW\s*=\s*([\d,.]+)", seg)
        dl = num(r"DEAD LEVEL\s*=\s*([\d,.]+)", seg)
        if dl is not None and lo <= dl <= hi:
            dead[dam] = dl
    row["Rim_inflow_cusecs"] = num(r"rim station inflows are\s*([\d,]+)", txt)
    row["Rim_outflow_cusecs"] = num(r"rim station outflows are\s*([\d,]+)", txt)
    return row, dead


def irsa(have):
    r = get(IRSA + "DailyData.aspx")
    if r is None or not r.ok:
        out("IRSA list page unavailable")
        return {}, {}
    files = sorted(set(re.findall(r'href=["\']?(Doc/Data(\d\d)-(\d\d)-(\d{4})\.pdf)', r.text, re.I)))
    out(f"IRSA lists {len(files)} PDFs: {[f[0][4:] for f in files]}")
    rows, dead = {}, {}
    for f, dd, mm, yy in files:
        d = pd.Timestamp(int(yy), int(mm), int(dd))
        if d in have:
            continue
        p = get(IRSA + f)
        if p is None or not p.ok or p.content[:4] != b"%PDF":
            out(f"  {f}: not a PDF")
            continue
        try:
            row, dl = parse_pdf(p.content)
        except Exception as e:  # noqa: BLE001  (one corrupt PDF must not stop the run)
            out(f"  {f}: unreadable ({e})")
            continue
        if row.get("Tarbela_level_ft") is None and row.get("Mangla_level_ft") is None:
            out(f"  {f}: no levels parsed")
            continue
        row["Source"] = "IRSA"
        rows[d] = row
        dead.update(dl)
        out(f"  {d:%Y-%m-%d}: " + ", ".join(f"{k}={v:g}" for k, v in row.items() if isinstance(v, float)))
    return rows, dead


def wapda_seed():
    """WAPDA's 2016-2024 daily river flows and levels workbook -> (DataFrame, Last-Modified)."""
    r = get(WAPDA_XLS)
    if r is None or not r.ok:
        out("WAPDA workbook unavailable")
        return pd.DataFrame(), None
    raw = pd.read_excel(io.BytesIO(r.content), sheet_name="GHRAPH", header=None)
    raw = raw[raw[0].map(lambda x: isinstance(x, (pd.Timestamp, datetime)))]
    df = pd.DataFrame({
        "Tarbela_level_ft": pd.to_numeric(raw[1], errors="coerce"),
        "Tarbela_inflow_cusecs": pd.to_numeric(raw[2], errors="coerce") * 1000,
        "Tarbela_outflow_cusecs": pd.to_numeric(raw[3], errors="coerce") * 1000,
        "Mangla_level_ft": pd.to_numeric(raw[5], errors="coerce"),
        "Mangla_inflow_cusecs": pd.to_numeric(raw[6], errors="coerce") * 1000,
        "Mangla_outflow_cusecs": pd.to_numeric(raw[7], errors="coerce") * 1000,
    })
    df.index = pd.to_datetime(raw[0]).dt.normalize()
    df = df[~df.index.duplicated(keep="last")].sort_index()
    for dam in ("Tarbela", "Mangla"):
        lo, hi = RANGE[dam]
        bad = ~df[f"{dam}_level_ft"].between(lo, hi)
        df.loc[bad, f"{dam}_level_ft"] = None
    df = df.dropna(how="all")
    df["Source"] = "WAPDA"
    out(f"WAPDA seed: {len(df)} days {df.index.min():%Y-%m-%d}..{df.index.max():%Y-%m-%d}")
    return df, r.headers.get("Last-Modified")


def saved_note(path, key):
    """Value after `key` on the Units sheet (one note per row), or None."""
    try:
        u = pd.read_excel(path, sheet_name="Units", header=None)
    except (FileNotFoundError, ValueError):
        return None
    for v in u.stack().map(str):
        if v.startswith(key):
            return v[len(key):].strip()
    return None


def wayback(have, after):
    """IRSA reports the Internet Archive happened to save, for dates after `after` (the end of WAPDA's workbook) not
    saved yet: a one-off backfill of the gap between WAPDA's workbook and IRSA's rolling list. The archive is slow, so
    each run spends at most WAYBACK_BUDGET seconds; -> (rows, finished)."""
    try:
        r = requests.get(CDX, params={"url": "pakirsa.gov.pk/Doc/Data*", "output": "json", "collapse": "original",
                                      "fl": "timestamp,original,statuscode", "limit": 20000}, timeout=(20, 180))
        caps = r.json()[1:]
    except (requests.RequestException, ValueError) as e:
        out(f"Wayback index unavailable: {e}")
        return {}, False
    rows, t0 = {}, time.time()
    for ts, orig, st in caps:
        if time.time() - t0 > WAYBACK_BUDGET:
            out(f"Wayback: time budget used, {len(rows)} new days; the rest next run")
            return rows, False
        m = re.search(r"Data(\d\d)-(\d\d)-(\d{4})\.pdf$", orig, re.I)
        if st != "200" or not m:
            continue
        try:
            d = pd.Timestamp(int(m.group(3)), int(m.group(2)), int(m.group(1)))
        except ValueError:
            continue
        if d <= after or d in have or d in rows:
            continue
        try:
            p = requests.get(f"http://web.archive.org/web/{ts}id_/{orig}", headers=H, timeout=(20, 60))
        except requests.RequestException:
            continue
        if not p.ok or p.content[:4] != b"%PDF":
            continue
        try:
            row, _ = parse_pdf(p.content)
        except Exception:  # noqa: BLE001
            continue
        if row.get("Tarbela_level_ft") is not None or row.get("Mangla_level_ft") is not None:
            row["Source"] = "IRSA (Wayback)"
            rows[d] = row
    out(f"Wayback: {len(caps)} captures, {len(rows)} new days")
    return rows, True


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", default=OUT)
    args = ap.parse_args()
    try:
        old = pd.read_excel(args.out, sheet_name="Daily", index_col=0)
        old.index = pd.to_datetime(old.index)
    except (FileNotFoundError, ValueError):
        old = pd.DataFrame()
    seed_mod = saved_note(args.out, "WAPDA workbook Last-Modified:")
    have_seed = not old.empty and "Source" in old and (old["Source"] == "WAPDA").any()
    head = None
    if have_seed:
        try:
            h = requests.head(WAPDA_XLS, headers=H, timeout=T, verify=False)
            head = h.headers.get("Last-Modified") if h.ok else None
        except requests.RequestException:
            head = None
    seed = pd.DataFrame()
    if not have_seed or (head and seed_mod and head != seed_mod):
        seed, seed_mod = wapda_seed()
    rows, dead = irsa(set(old.index) if not old.empty else set())
    new = pd.DataFrame.from_dict(rows, orient="index")
    wb_done = saved_note(args.out, "Wayback backfill:") == "done"
    wb = pd.DataFrame()
    if not wb_done:
        have = set(old.index) | set(seed.index) | set(new.index)
        wapda_days = old.index[old["Source"].eq("WAPDA")] if "Source" in old else []
        after = max([*seed.index, *wapda_days], default=pd.Timestamp("2016-01-01"))
        wb_rows, wb_done = wayback(have, after)
        wb = pd.DataFrame.from_dict(wb_rows, orient="index")
    daily = old
    for part in (wb, seed, new):   # later parts win on the same date (IRSA over WAPDA)
        if not part.empty:
            daily = part if daily.empty else pd.concat([daily[~daily.index.isin(part.index)], part])
    if daily.empty:
        raise SystemExit("No Pakistan reservoir data")
    daily = daily.sort_index().reindex(columns=COLS)
    daily.index.name = "date"
    try:
        lim = pd.read_excel(args.out, sheet_name="Limits", index_col=0)
    except (FileNotFoundError, ValueError):
        lim = pd.DataFrame({"Dead_level_ft": {"Tarbela": 1402.0, "Mangla": 1050.0, "Chashma": 638.15}})
    for dam, v in dead.items():
        lim.loc[dam, "Dead_level_ft"] = v
    lim["Max_conservation_level_ft"] = pd.Series(MAX_LEVEL)
    lim = lim.reindex(DAMS)
    lim.index.name = "reservoir"
    notes = [
        "UNITS",
        "Daily: reservoir water level at 06:00, feet above mean sea level; mean inflow / outflow over the 24 hours to "
        "06:00, cusecs (cubic feet per second; 1 cusec for a day = 1.98 acre-feet). Rim_inflow / Rim_outflow: total "
        "of the four rim stations (Indus at Tarbela, Jhelum at Mangla, Kabul at Nowshera, Chenab at Marala). "
        "Source: which publisher the day's row comes from. Limits: dead level (lowest drawdown level) and maximum "
        "conservation level (full), ft.",
        "",
        "COVERAGE",
        f"Daily from {daily.index.min():%Y-%m-%d} to {daily.index.max():%Y-%m-%d}. 2016-03-01 to Dec 2024 from WAPDA's "
        "river-flow workbook (Tarbela and Mangla only); from late Sep 2026 IRSA's daily reports (adds Chashma and the "
        "rim-station totals). Between them only the days whose IRSA report the Internet Archive saved (scattered, "
        "a few a month); IRSA deletes its reports after about a week and WAPDA stopped its workbook. IRSA's reports give "
        "Tarbela's dead level as 1,402 ft (older WAPDA figures use 1,380 ft).",
        "",
        "SOURCE",
        "Indus River System Authority (IRSA), Daily Water Situation: http://pakirsa.gov.pk/DailyData.aspx",
        "WAPDA, River Flow (river flows and levels workbook): https://wapda.gov.pk/river-flow/",
        "Internet Archive copies of IRSA reports (Source 'IRSA (Wayback)'): https://web.archive.org/",
        f"WAPDA workbook Last-Modified: {seed_mod or ''}",
        f"Wayback backfill: {'done' if wb_done else ''}",
    ]
    xlsx_notes.write_workbook(args.out, {"Daily": daily, "Limits": lim}, notes, {"UNITS", "COVERAGE", "SOURCE"})
    out(f"Saved {args.out}: {len(daily)} days {daily.index.min():%Y-%m-%d}..{daily.index.max():%Y-%m-%d}")
    out(daily.tail(5).to_string())
    out(lim.to_string())


if __name__ == "__main__":
    main()
