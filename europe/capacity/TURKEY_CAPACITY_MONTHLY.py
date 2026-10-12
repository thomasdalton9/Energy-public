"""
Monthly installed generation capacity for Turkey from EPDK's monthly Electricity Market Sector Reports (nothing interpolated).

  output/Data and Chart Outputs/turkey_power_capacity_monthly.xlsx

Source: EPDK (Enerji Piyasasi Duzenleme Kurumu, the energy regulator), 'Elektrik Piyasasi Aylik Sektor Raporu', one PDF per
month from https://www.epdk.gov.tr/Detay/Icerik/3-0-23-3/elektrikaylik-sektor-raporlar (reports to Dec 2025 are Word .docx files, the
2026 ones PDF with an Excel annex; both are read). Each report carries the table '<Y-1> <Month>-<Y> <Month> Elektrik Kurulu Gucu ve Uretim
Miktari': TOTAL installed capacity (MW) by source, LICENSED AND UNLICENSED plants together, at the end of the report month
and of the same month a year earlier (the report text says licensed capacity and generation come from TEIAS's YTBS and are
provisional until TEIAS publishes). RAW official source, MONTHLY, all fuels. About 6-8 weeks lag.

Date convention (same as the ENTSO-E annual files and the other monthly capacity files): the row dated 1 May holds the
capacity at the end of April.

Two table layouts exist. 2022 and later: one table of TOTAL capacity (licensed + unlicensed) at the report month and the same month a
year earlier (two points per report). Up to 2021: the LICENSED capacity table (previous and current year) plus a separate table of
UNLICENSED capacity (current month only); the file then sums the two for the current month only ('basis' column on Reports).
Each report therefore gives up to two month-end points. A period takes the value from the LATEST
report that contains it (the next year's report revises the previous year's figure); 'Revisions' lists the differences.
Hence the history reaches back one year before the oldest report read (2016 reports -> from Jan 2015 if the table exists in that
layout; reports whose table cannot be parsed or whose parts do not add up to the reported total within 0.3% are left out and
listed on 'Releases').

Fuel mapping (EPDK source -> column): HIDROLIK -> Hydro_MW (dam + river), RUZGAR -> Wind_MW, GUNES -> Solar_MW, BIYOKUTLE ->
Bioenergy_MW, JEOTERMAL -> Geothermal_MW, DOGAL GAZ + LNG -> Gas_MW, ITHAL KOMUR + TAS KOMURU -> Coal_MW (imported and hard coal),
LINYIT + ASFALTIT -> Lignite_MW, FUEL OIL + NAFTA + MOTORIN -> Oil_MW, nuclear -> Nuclear_MW (none before Akkuyu), anything
else unmapped -> Other_MW. Total_MW is EPDK's own TOPLAM. Pumped storage and batteries are not listed separately by EPDK.

Incremental: the committed workbook is the history store ('Reports' sheet, one row per report month, period, source). Report
months stored with status ok are not downloaded again except the last 3 (revisions); parse failures are retried only with --retry (download failures every run).

Usage: python3 TURKEY_CAPACITY_MONTHLY.py [--out-dir DIR] [--retry] [--max-reports N]
"""
import argparse
import hashlib
import html
import os
import re
import sys
import time
from datetime import datetime, timezone

import numpy as np
import pandas as pd
import requests

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(os.path.dirname(HERE))
sys.path.insert(0, ROOT)
import xlsx_notes  # noqa: E402

OUT_DEFAULT = os.path.join(ROOT, "output", "Data and Chart Outputs")
FILE = "turkey_power_capacity_monthly.xlsx"
LIST_URL = "https://www.epdk.gov.tr/Detay/Icerik/3-0-23-3/elektrikaylik-sektor-raporlar"
BASE = "https://www.epdk.gov.tr/"
UA = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36",
      "Accept-Language": "tr,en;q=0.8"}
MONTHS = {"Ocak": 1, "Şubat": 2, "Mart": 3, "Nisan": 4, "Mayıs": 5, "Haziran": 6, "Temmuz": 7, "Ağustos": 8, "Eylül": 9,
          "Ekim": 10, "Kasım": 11, "Aralık": 12}
# label as printed -> output column ('' = subtotal row, not used)
LABELS = {"HİDROLİK": "Hydro_MW", "HIDROLIK": "Hydro_MW", "RÜZGÂR": "Wind_MW", "RÜZGAR": "Wind_MW", "RÜZGÃ¢R": "Wind_MW",
          "GÜNEŞ": "Solar_MW", "GUNES": "Solar_MW", "BİYOKÜTLE": "Bioenergy_MW", "BIYOKUTLE": "Bioenergy_MW",
          "JEOTERMAL": "Geothermal_MW", "YENİLENEBİLİR": "", "YENILENEBILIR": "", "DOĞAL GAZ": "Gas_MW", "DOGAL GAZ": "Gas_MW",
          "LNG": "Gas_MW", "İTHAL KÖMÜR": "Coal_MW", "ITHAL KOMUR": "Coal_MW", "TAŞ KÖMÜRÜ": "Coal_MW", "TAS KOMURU": "Coal_MW",
          "LİNYİT": "Lignite_MW", "LINYIT": "Lignite_MW", "ASFALTİT": "Lignite_MW", "ASFALTIT": "Lignite_MW",
          "FUEL OİL": "Oil_MW", "FUEL OIL": "Oil_MW", "NAFTA": "Oil_MW", "MOTORİN": "Oil_MW", "MOTORIN": "Oil_MW",
          "NÜKLEER": "Nuclear_MW", "NUKLEER": "Nuclear_MW", "URANYUM": "Nuclear_MW", "ATIK ISI": "Other_MW", "ATIK ISI*": "Other_MW",
          "TERMİK": "", "TERMIK": "", "TOPLAM": "TOTAL"}
FUEL_COLS = ["Hydro_MW", "Gas_MW", "Wind_MW", "Solar_MW", "Coal_MW", "Lignite_MW", "Nuclear_MW", "Oil_MW", "Bioenergy_MW",
             "Geothermal_MW", "Other_MW"]
COLS = FUEL_COLS + ["Total_MW"]
NUM = re.compile(r"\d{1,3}(?:\.\d{3})*,\d{2}")


def num(s):
    return float(s.replace(".", "").replace(",", "."))


def get(url, tries=3, **kw):
    for i in range(tries):
        try:
            r = requests.get(url, headers=UA, timeout=(15, 120), **kw)
            if r.status_code == 200:
                return r
            print(f"  http {r.status_code} {url[:90]}")
        except Exception as e:  # noqa: BLE001
            print(f"  {type(e).__name__} {url[:90]} (try {i + 1})")
        time.sleep(5 * (i + 1))
    return None


def is_docx(c):
    import io
    import zipfile
    try:
        return "word/document.xml" in zipfile.ZipFile(io.BytesIO(c)).namelist()
    except Exception:  # noqa: BLE001
        return False


def list_reports():
    """{report month: [download urls]} from the link titles (title wording varies by year: '2025 Yili Elektrik Piyasasi Ocak Ayi
    Sektor Raporu', '2017 Yili Ocak Ayi Elektrik Piyasasi Sektor Raporu', 'Elektrik Piyasasi Sektor Raporu Mayis 2022', ...)."""
    r = get(LIST_URL)
    if r is None:
        raise RuntimeError("EPDK list page not reachable")
    t = r.text
    out = {}
    for m in re.finditer(r'<a[^>]*href="([^"]*DownloadDocument\?id=[^"]+)"[^>]*?title="([^"]*)"', t):
        title = html.unescape(m.group(2))
        yr = re.findall(r"(?<!\d)(20\d\d)(?!\d)", title)
        mo = [k for k in MONTHS if k in title]
        if len(yr) != 1 or len(mo) != 1 or not re.search(r"Elektrik|Ocak|Şubat|Mart|Nisan|Mayıs|Haziran|Temmuz|Ağustos|Eylül|Ekim|Kasım|Aralık", title):
            continue
        key = pd.Timestamp(int(yr[0]), MONTHS[mo[0]], 1)
        u = BASE + html.unescape(m.group(1)).lstrip("/")
        out.setdefault(key, [])
        if u not in out[key]:
            out[key].append(u)
    return out


def parse_pdf(content):
    """-> (rows [(label, col, cap_prev, share_prev, cap_cur, share_cur)], note). Needs pymupdf."""
    import fitz
    d = fitz.open(stream=content, filetype="pdf")
    for i in range(len(d)):
        tx = d[i].get_text()
        if not re.search(r"TOPLAM\s+KURULU\s+G[ÜU]Ç", tx):
            continue
        s = tx[re.search(r"TOPLAM\s+KURULU\s+G[ÜU]Ç", tx).start():]
        end = re.search(r"\*\s*Lisanslı ve lisanssız", s)
        if end:
            s = s[:end.start()]
        labs = sorted(LABELS, key=len, reverse=True)
        pat = re.compile(r"(?<![A-ZÇĞİÖŞÜÂa-zçğıöşü])(" + "|".join(re.escape(x) for x in labs) + r")(?![A-ZÇĞİÖŞÜÂa-zçğıöşü])")
        ms = list(pat.finditer(s))
        rows = []
        for j, m in enumerate(ms):
            seg = s[m.end(): ms[j + 1].start() if j + 1 < len(ms) else len(s)]
            ns = NUM.findall(seg)
            if len(ns) >= 4:
                rows.append((m.group(1), LABELS[m.group(1)], num(ns[0]), num(ns[1]), num(ns[2]), num(ns[3])))
        if len(rows) >= 8:
            return rows, f"page {i + 1}"
    return None, "table not found"


UNLIC = {"GÜNEŞ": "Solar_MW", "GÜNEŞ (FOTOVOLTAİK)": "Solar_MW", "GÜNEŞ (YOĞUNLAŞTIRILMIŞ)": "Solar_MW", "DOĞAL GAZ": "Gas_MW",
         "BİYOKÜTLE": "Bioenergy_MW", "RÜZGÂR": "Wind_MW", "RÜZGAR": "Wind_MW", "HİDROLİK": "Hydro_MW", "JEOTERMAL": "Geothermal_MW"}
LIC_ALIASES = {"BARAJLI HİDROLİK": "Hydro_MW", "AKARSU": "Hydro_MW"}


def _upper(x):
    return re.sub(r"\s+", " ", x.replace("i", "İ").replace("ı", "I").upper()).strip()


def _tonum(x):
    x = x.strip()
    if re.fullmatch(r"-?\d{1,3}(?:\.\d{3})*,\d+|-?\d+,\d+|-?\d+", x):
        return float(x.replace(".", "").replace(",", "."))
    return None


def docx_tables(content):
    import io
    import docx
    d = docx.Document(io.BytesIO(content))
    out = []
    for tb in d.tables:
        rows = []
        for row in tb.rows:
            cells, prev = [], None
            for c in row.cells:
                if prev is not None and c._tc is prev:
                    continue
                prev = c._tc
                cells.append(re.sub(r"\s+", " ", c.text).strip())
            rows.append(cells)
        out.append(rows)
    return out


def parse_docx(content):
    """-> (rows [(label, col, cap_prev|None, share_prev|None, cap_cur, share_cur)], basis, note) or (None, '', why)."""
    tabs = docx_tables(content)
    # A: combined table 'TOPLAM KURULU GUC' (licensed + unlicensed), previous and current month
    for ti, rows in enumerate(tabs):
        flat = _upper(" ".join(" ".join(r) for r in rows))
        if "TOPLAM KURULU" in flat and ("HİDROLİK" in flat or "HIDROLIK" in flat) and "TOPLAM" in flat:
            out = []
            for r in rows:
                if not r:
                    continue
                lab = _upper(r[0])
                key = next((k for k in LABELS if _upper(k) == lab), None)
                nums = [_tonum(x) for x in r[1:]]
                nums = [x for x in nums if x is not None]
                if key and len(nums) >= 4:
                    out.append((key, LABELS[key], nums[0], nums[1], nums[2], nums[3]))
            if len(out) >= 8:
                return out, "combined table (licensed + unlicensed)", f"table {ti}"
    # B: licensed table (KAYNAK TURU, prev, cur, change) + unlicensed table (Kaynak Turu | Kurulu Guc | Oran)
    lic = unl = None
    for ti, rows in enumerate(tabs):
        flat = _upper(" ".join(" ".join(r) for r in rows))
        if lic is None and "KAYNAK TÜRÜ" in flat and "KURULU GÜÇ (MW)" in flat and "BARAJLI" in flat and "TOPLAM" in flat:
            out = []
            for r in rows:
                if not r:
                    continue
                lab = _upper(r[0])
                nums = [_tonum(x) for x in r[1:]]
                nums = [x for x in nums if x is not None]
                if len(nums) < 4:
                    continue
                col = LIC_ALIASES.get(lab)
                if col is None:
                    key = next((k for k in LABELS if _upper(k) == lab), None)
                    col = LABELS.get(key) if key else None
                if col is not None:
                    out.append((lab, col, None, None, nums[2], nums[3]))
            lic = (out, ti)
        if unl is None and len(rows) >= 4 and all(len(r) == 3 for r in rows if r) and "GENEL TOPLAM" in flat and "GÜNEŞ" in flat \
                and "KURULU GÜÇ" in flat and "KAYNAK TÜRÜ" in flat:
            out = []
            for r in rows:
                lab = _upper(r[0])
                v = _tonum(r[1])
                if lab == "GENEL TOPLAM" and v is not None:
                    out.append(("GENEL TOPLAM", "UNL_TOTAL", None, None, v, 100.0))
                elif lab in UNLIC and v is not None:
                    out.append(("LİSANSSIZ " + lab, UNLIC[lab], None, None, v, None))
            unl = (out, ti)
    if lic and unl:
        lrows, urows = lic[0], unl[0]
        ltot = [r for r in lrows if r[1] == "TOTAL"]
        utot = [r for r in urows if r[1] == "UNL_TOTAL"]
        if not ltot or not utot:
            return None, "", "licensed/unlicensed totals missing"
        out = [r for r in lrows if r[1] not in ("TOTAL", "")]
        out += [r for r in urows if r[1] != "UNL_TOTAL"]
        out.append(("TOPLAM (lisanslı + lisanssız)", "TOTAL", None, None, ltot[0][4] + utot[0][4], 100.0))
        chk = [("licensed", sum(r[4] for r in lrows if r[1] not in ("TOTAL", "")), ltot[0][4]),
               ("unlicensed", sum(r[4] for r in urows if r[1] != "UNL_TOTAL"), utot[0][4])]
        bad = [f"{n}: parts {a:.1f} vs total {b:.1f}" for n, a, b in chk if b <= 0 or abs(a - b) / b > 0.003]
        if bad:
            return None, "", "; ".join(bad)
        return out, "licensed table + unlicensed table (summed)", f"tables {lic[1]}+{unl[1]}"
    return None, "", "no capacity table found"


def validate(rows):
    tot = [r for r in rows if r[1] == "TOTAL"]
    if not tot:
        return "no TOPLAM row"
    parts = [r for r in rows if r[1] not in ("", "TOTAL")]
    for k, nm in ((2, "prev"), (4, "cur")):
        if tot[0][k] is None:
            continue
        s, t = sum(r[k] for r in parts), tot[0][k]
        if t <= 0 or abs(s - t) / t > 0.003:
            return f"{nm}: parts {s:.1f} vs total {t:.1f}"
    return ""


def read_sheet(path, sheet):
    if not os.path.exists(path):
        return None
    try:
        return pd.read_excel(path, sheet_name=sheet)
    except Exception as e:  # noqa: BLE001
        print(f"  could not read {sheet}: {type(e).__name__}")
        return None


def build_monthly(rep):
    """rep: long frame report_month, period, col, MW -> monthly (row date = 1st of the month after the period)."""
    rep = rep.copy()
    rep["report_month"] = pd.to_datetime(rep["report_month"])
    rep["period"] = pd.to_datetime(rep["period"])
    rep = rep.groupby(["report_month", "period", "col"], as_index=False)["MW"].sum()
    latest = rep.sort_values("report_month").groupby(["period", "col"], as_index=False).last()
    wide = latest.pivot_table(index="period", columns="col", values="MW", aggfunc="sum")
    for c in COLS:
        if c not in wide:
            wide[c] = 0.0 if c != "Total_MW" else np.nan
    tot = latest[latest["col"] == "TOTAL"].set_index("period")["MW"]
    wide["Total_MW"] = tot
    wide = wide[COLS].round(2)
    src = latest.groupby("period")["report_month"].max()
    wide["source_report"] = src.reindex(wide.index).dt.strftime("%Y-%m")
    wide["source_lag_months"] = ((src.reindex(wide.index).dt.year - wide.index.year) * 12 + (src.reindex(wide.index).dt.month - wide.index.month))
    wide.index = wide.index + pd.DateOffset(months=1)
    wide.index.name = "date"
    return wide.sort_index()


def revisions(rep):
    rep = rep.copy()
    rep["report_month"] = pd.to_datetime(rep["report_month"])
    rep["period"] = pd.to_datetime(rep["period"])
    rep = rep.groupby(["report_month", "period", "col"], as_index=False)["MW"].sum()
    rows = []
    for (per, col), g in rep.groupby(["period", "col"]):
        g = g.sort_values("report_month")
        if len(g) > 1 and col != "":
            a, b = g.iloc[0], g.iloc[-1]
            rows.append({"period_end_of": per.strftime("%Y-%m"), "series": col, "first_report": a["report_month"].strftime("%Y-%m"),
                         "first_MW": a["MW"], "latest_report": b["report_month"].strftime("%Y-%m"), "latest_MW": b["MW"],
                         "revision_MW": round(b["MW"] - a["MW"], 2)})
    d = pd.DataFrame(rows)
    return d.set_index("period_end_of") if len(d) else d


def entsoe_check(out_dir, monthly):
    p = os.path.join(out_dir, "turkey_power_capacity.xlsx")
    e = read_sheet(p, "Monthly")
    if e is None:
        return pd.DataFrame({"note": ["no turkey_power_capacity.xlsx (ENTSO-E A68) in the repo: Turkey is not in the ENTSO-E annual pull"]})
    e["date"] = pd.to_datetime(e["date"])
    e = e.set_index("date")
    rows = []
    for d in e.index:
        if d in monthly.index:
            for c in ("Hydro_MW", "Gas_MW", "Wind_MW", "Solar_MW"):
                if c in e and pd.notna(e.at[d, c]):
                    a, b = monthly.at[d, c], e.at[d, c]
                    rows.append({"date": d, "series": c, "this_file_MW": a, "ENTSOE_A68_MW": b, "difference_MW": round(a - b, 1),
                                 "difference_pct": round((a - b) / b * 100, 1) if b else np.nan})
    return pd.DataFrame(rows).set_index("date") if rows else pd.DataFrame()


NOTES = ["Turkey - monthly installed generation capacity (MW) from EPDK monthly Electricity Market Sector Reports", "",
         "Source",
         "EPDK (Energy Market Regulatory Authority) 'Elektrik Piyasasi Aylik Sektor Raporu', one PDF a month: https://www.epdk.gov.tr/Detay/Icerik/3-0-23-3/elektrikaylik-sektor-raporlar",
         "Reports to Dec 2025 are Word files, 2026 PDF. Table 'Elektrik Kurulu Gucu ve Uretim Miktari' (2022 on): TOTAL installed capacity by source, licensed AND unlicensed plants, end of the report month and of the same month a year earlier. Licensed capacity comes from TEIAS (YTBS) and is provisional until TEIAS publishes (EPDK's own note).",
         "Raw official source, monthly, all fuels. The row dated 1 May is the capacity at the end of April. About 6-8 weeks lag.",
         "", "Method",
         "Up to 2021 only a LICENSED table and a separate UNLICENSED table exist: the current month is their sum (basis column on Reports). A period takes the value of the latest report that contains it (next year's report revises it). Reports whose parts do not add to EPDK's total within 0.3%, or whose table cannot be read, are left out and listed on Releases. Nothing is interpolated or filled.",
         "Incremental: the Reports sheet is the history store; stored report months are not downloaded again except the last three; parse failures are retried only with --retry (download failures every run).",
         "", "Columns",
         "Hydro_MW = dam + river. Coal_MW = imported + hard coal; Lignite_MW = lignite + asphaltite; Gas_MW = natural gas + LNG; Oil_MW = fuel oil + naphtha + diesel; Bioenergy_MW = biomass (incl. biogas); Geothermal_MW; Nuclear_MW (none before Akkuyu). Other_MW = any source EPDK lists that is not mapped. Total_MW = EPDK's TOPLAM.",
         "Pumped storage and batteries are not listed by EPDK.",
         "", "Sheets", "Monthly, Revisions (same period in two reports), Check_ENTSOE, Reports (long, history store), Releases"]
TITLES = ["Source", "Method", "Columns", "Sheets"]


def save(path, rep, rel, out_dir):
    if rep.empty:
        raise RuntimeError("no EPDK report could be read")
    rep = rep.copy()
    rep["report_month"] = pd.to_datetime(rep["report_month"])
    rep["period"] = pd.to_datetime(rep["period"])
    rep = rep.sort_values(["report_month", "period", "col"]).reset_index(drop=True)
    monthly = build_monthly(rep)
    resid = (monthly["Total_MW"] - monthly[FUEL_COLS].sum(axis=1)).abs()
    print("max |total - sum of fuels| MW:", round(resid.max(), 2))
    rel = rel.copy()
    rel["report_month"] = pd.to_datetime(rel["report_month"])
    rel = rel.sort_values("report_month").set_index("report_month")
    sheets = {"Monthly": monthly, "Revisions": revisions(rep), "Check_ENTSOE": entsoe_check(out_dir, monthly),
              "Reports": rep.set_index("report_month"), "Releases": rel}
    xlsx_notes.write_workbook(path, sheets, NOTES, TITLES)
    print(f"Turkey: {len(monthly)} months {monthly.index.min().date()} - {monthly.index.max().date()}")
    print(monthly.tail(3).T)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out-dir", default=OUT_DEFAULT)
    ap.add_argument("--retry", action="store_true")
    ap.add_argument("--max-reports", type=int, default=400)
    ap.add_argument("--budget-min", type=float, default=55.0)
    a = ap.parse_args()
    os.makedirs(a.out_dir, exist_ok=True)
    path = os.path.join(a.out_dir, FILE)
    rep = read_sheet(path, "Reports")
    rel = read_sheet(path, "Releases")
    rep = rep if rep is not None else pd.DataFrame(columns=["report_month", "period", "label", "col", "MW", "share_pct", "basis"])
    rel = rel if rel is not None else pd.DataFrame(columns=["report_month", "status", "detail", "bytes", "sha256", "checked_utc"])
    rel["report_month"] = pd.to_datetime(rel["report_month"])
    listing = list_reports()
    print(f"EPDK list: {len(listing)} report months {min(listing).date()} - {max(listing).date()}")
    months = sorted(listing)
    last3 = set(months[-3:])
    done = set(rel.loc[rel["status"] == "ok", "report_month"])
    failed = set(rel.loc[rel["status"] == "parse-failed", "report_month"])
    todo = [m for m in months if (m not in done and (m not in failed or a.retry)) or m in last3]
    print(f"to fetch: {len(todo)}")
    t0 = time.time()
    stop = {"flag": False}

    def fetch(m):
        if stop["flag"]:
            return m, None, None
        for u in listing[m]:
            r = get(u)
            if r is None:
                continue
            c = r.content
            if c[:4] == b"%PDF":
                return m, "pdf", c
            if c[:2] == b"PK" and is_docx(c):
                return m, "docx", c
            time.sleep(0.5)
        return m, "none", None

    from concurrent.futures import ThreadPoolExecutor, as_completed
    ex = ThreadPoolExecutor(max_workers=3)
    futs = [ex.submit(fetch, m) for m in sorted(todo, reverse=True)[:a.max_reports]]
    n = 0
    for fu in as_completed(futs):
        m, kind, content = fu.result()
        if kind is None:
            continue
        if (time.time() - t0) / 60 > a.budget_min and not stop["flag"]:
            stop["flag"] = True
            print("budget reached, stopping; the next run continues")
        n += 1
        if n % 10 == 0 and not rep.empty:
            save(path, rep, rel, a.out_dir)
        now = datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M")
        rel = rel[rel["report_month"] != m]
        if content is None:
            rel.loc[len(rel)] = [m, "no-file", f"{len(listing[m])} links, none a PDF or Word file", 0, "", now]
            print(f"{m:%Y-%m}: no usable file")
            continue
        basis = "combined table (licensed + unlicensed)"
        if kind == "pdf":
            rows, note = parse_pdf(content)
            bad = "table not found" if rows is None else validate(rows)
        else:
            rows, basis, note = parse_docx(content)
            bad = note if rows is None else validate(rows)
        sha = hashlib.sha256(content).hexdigest()[:16]
        if rows is None or bad:
            rel.loc[len(rel)] = [m, "parse-failed", f"{kind}: {bad or note}", len(content), sha, now]
            print(f"{m:%Y-%m}: parse failed ({kind}: {bad or note})")
            continue
        per_cur = m
        per_prev = m - pd.DateOffset(years=1)
        rep = rep[pd.to_datetime(rep["report_month"]) != m]
        new = []
        for lab, col, cp, sp, cc, sc in rows:
            if col == "":
                continue
            if cp is not None:
                new.append({"report_month": m, "period": per_prev, "label": lab, "col": col, "MW": cp, "share_pct": sp, "basis": basis})
            new.append({"report_month": m, "period": per_cur, "label": lab, "col": col, "MW": cc, "share_pct": sc, "basis": basis})
        rep = pd.concat([rep, pd.DataFrame(new)], ignore_index=True)
        rel.loc[len(rel)] = [m, "ok", f"{kind}; {basis}; {note}", len(content), sha, now]
        tot = [r for r in rows if r[1] == "TOTAL"][0]
        print(f"{m:%Y-%m}: [{time.time() - t0:.0f}s] ok {kind} total {tot[4]:.0f} MW (prev-year {tot[2]}) {basis}")
    ex.shutdown(wait=False, cancel_futures=True)
    save(path, rep, rel, a.out_dir)


if __name__ == "__main__":
    main()
