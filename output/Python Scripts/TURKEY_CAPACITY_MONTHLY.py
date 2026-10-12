"""
Monthly installed generation capacity for Turkey from EPDK's monthly Electricity Market Sector Reports (nothing interpolated).

  output/Data and Chart Outputs/turkey_power_capacity_monthly.xlsx

Source: EPDK (Enerji Piyasasi Duzenleme Kurumu, the energy regulator), 'Elektrik Piyasasi Aylik Sektor Raporu', one PDF per
month from https://www.epdk.gov.tr/Detay/Icerik/3-0-23-3/elektrikaylik-sektor-raporlar (the list page links a PDF and an Excel
annex per month; the PDF is read). Each report carries the table '<Y-1> <Month>-<Y> <Month> Elektrik Kurulu Gucu ve Uretim
Miktari': TOTAL installed capacity (MW) by source, LICENSED AND UNLICENSED plants together, at the end of the report month
and of the same month a year earlier (the report text says licensed capacity and generation come from TEIAS's YTBS and are
provisional until TEIAS publishes). RAW official source, MONTHLY, all fuels. About 6-8 weeks lag.

Date convention (same as the ENTSO-E annual files and the other monthly capacity files): the row dated 1 May holds the
capacity at the end of April.

Each report gives two month-end points (this year and the same month last year). A period takes the value from the LATEST
report that contains it (the next year's report revises the previous year's figure); 'Revisions' lists the differences.
Hence the history reaches back one year before the oldest report read (2016 reports -> from Jan 2015 if the table exists in that
layout; reports whose table cannot be parsed or whose parts do not add up to the reported total within 0.3% are left out and
listed on 'Releases').

Fuel mapping (EPDK source -> column): HIDROLIK -> Hydro_MW (dam + river), RUZGAR -> Wind_MW, GUNES -> Solar_MW, BIYOKUTLE ->
Bioenergy_MW, JEOTERMAL -> Geothermal_MW, DOGAL GAZ + LNG -> Gas_MW, ITHAL KOMUR + TAS KOMURU -> Coal_MW (imported and hard coal),
LINYIT + ASFALTIT -> Lignite_MW, FUEL OIL + NAFTA + MOTORIN -> Oil_MW, nuclear -> Nuclear_MW (none before Akkuyu), anything
else unmapped -> Other_MW. Total_MW is EPDK's own TOPLAM. Pumped storage and batteries are not listed separately by EPDK.

Incremental: the committed workbook is the history store ('Reports' sheet, one row per report month, period, source). Report
months stored with status ok are not downloaded again except the last 3 (revisions); failed ones are retried only with --retry.

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


def get(url, tries=4, **kw):
    for i in range(tries):
        try:
            r = requests.get(url, headers=UA, timeout=(20, 180), **kw)
            if r.status_code == 200:
                return r
            print(f"  http {r.status_code} {url[:90]}")
        except Exception as e:  # noqa: BLE001
            print(f"  {type(e).__name__} {url[:90]} (try {i + 1})")
        time.sleep(5 * (i + 1))
    return None


def list_reports():
    r = get(LIST_URL)
    if r is None:
        raise RuntimeError("EPDK list page not reachable")
    t = r.text
    out = {}
    for m in re.finditer(r'href="([^"]*DownloadDocument\?id=[^"]+)"', t):
        pre = re.sub(r"\s+", " ", html.unescape(re.sub(r"<[^>]+>", " ", t[max(0, m.start() - 700):m.start()])))
        mm = re.findall(r"(\d{4}) Yılı Elektrik Piyasası (\w+) Ayı Sektör Raporu", pre)
        if not mm or mm[-1][1] not in MONTHS:
            continue
        key = pd.Timestamp(int(mm[-1][0]), MONTHS[mm[-1][1]], 1)
        out.setdefault(key, [])
        u = BASE + html.unescape(m.group(1)).lstrip("/")
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


def validate(rows):
    tot = [r for r in rows if r[1] == "TOTAL"]
    if not tot:
        return "no TOPLAM row"
    parts = [r for r in rows if r[1] not in ("", "TOTAL")]
    for k, nm in ((2, "prev"), (4, "cur")):
        s, t = sum(r[k] for r in parts), tot[0][k]
        if t <= 0 or abs(s - t) / t > 0.003:
            return f"{nm}: parts {s:.1f} vs total {t:.1f}"
        if abs(tot[0][k] - tot[0][k]) > 0:
            pass
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
         "Table 'Elektrik Kurulu Gucu ve Uretim Miktari': TOTAL installed capacity by source, licensed AND unlicensed plants, end of the report month and of the same month a year earlier. Licensed capacity comes from TEIAS (YTBS) and is provisional until TEIAS publishes (EPDK's own note).",
         "Raw official source, monthly, all fuels. The row dated 1 May is the capacity at the end of April. About 6-8 weeks lag.",
         "", "Method",
         "A period takes the value of the latest report that contains it (next year's report revises it). Reports whose parts do not add to EPDK's total within 0.3%, or whose table cannot be read, are left out and listed on Releases. Nothing is interpolated or filled.",
         "Incremental: the Reports sheet is the history store; stored report months are not downloaded again except the last three; failed ones are retried only with --retry.",
         "", "Columns",
         "Hydro_MW = dam + river. Coal_MW = imported + hard coal; Lignite_MW = lignite + asphaltite; Gas_MW = natural gas + LNG; Oil_MW = fuel oil + naphtha + diesel; Bioenergy_MW = biomass (incl. biogas); Geothermal_MW; Nuclear_MW (none before Akkuyu). Other_MW = any source EPDK lists that is not mapped. Total_MW = EPDK's TOPLAM.",
         "Pumped storage and batteries are not listed by EPDK.",
         "", "Sheets", "Monthly, Revisions (same period in two reports), Check_ENTSOE, Reports (long, history store), Releases"]
TITLES = ["Source", "Method", "Columns", "Sheets"]


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out-dir", default=OUT_DEFAULT)
    ap.add_argument("--retry", action="store_true")
    ap.add_argument("--max-reports", type=int, default=400)
    a = ap.parse_args()
    os.makedirs(a.out_dir, exist_ok=True)
    path = os.path.join(a.out_dir, FILE)
    rep = read_sheet(path, "Reports")
    rel = read_sheet(path, "Releases")
    rep = rep if rep is not None else pd.DataFrame(columns=["report_month", "period", "label", "col", "MW", "share_pct"])
    rel = rel if rel is not None else pd.DataFrame(columns=["report_month", "status", "detail", "bytes", "sha256", "checked_utc"])
    rel["report_month"] = pd.to_datetime(rel["report_month"])
    listing = list_reports()
    print(f"EPDK list: {len(listing)} report months {min(listing).date()} - {max(listing).date()}")
    months = sorted(listing)
    last3 = set(months[-3:])
    done = set(rel.loc[rel["status"] == "ok", "report_month"])
    failed = set(rel.loc[rel["status"] != "ok", "report_month"])
    todo = [m for m in months if (m not in done and (m not in failed or a.retry)) or m in last3]
    print(f"to fetch: {len(todo)}")
    n = 0
    for m in sorted(todo, reverse=True):
        if n >= a.max_reports:
            break
        n += 1
        content, used = None, None
        for u in listing[m]:
            r = get(u)
            if r is not None and r.content[:4] == b"%PDF":
                content, used = r.content, u
                break
            time.sleep(1)
        now = datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M")
        rel = rel[rel["report_month"] != m]
        if content is None:
            rel.loc[len(rel)] = [m, "no-pdf", f"{len(listing[m])} links, none a PDF", 0, "", now]
            print(f"{m:%Y-%m}: no PDF")
            continue
        rows, note = parse_pdf(content)
        sha = hashlib.sha256(content).hexdigest()[:16]
        bad = "table not found" if rows is None else validate(rows)
        if rows is None or bad:
            rel.loc[len(rel)] = [m, "parse-failed", bad or note, len(content), sha, now]
            print(f"{m:%Y-%m}: parse failed ({bad or note})")
            continue
        per_cur = m
        per_prev = m - pd.DateOffset(years=1)
        rep = rep[rep["report_month"].pipe(pd.to_datetime) != m]
        new = []
        for lab, col, cp, sp, cc, sc in rows:
            if col == "":
                continue
            new.append({"report_month": m, "period": per_prev, "label": lab, "col": col, "MW": cp, "share_pct": sp})
            new.append({"report_month": m, "period": per_cur, "label": lab, "col": col, "MW": cc, "share_pct": sc})
        rep = pd.concat([rep, pd.DataFrame(new)], ignore_index=True)
        rel.loc[len(rel)] = [m, "ok", note, len(content), sha, now]
        tot = [r for r in rows if r[1] == "TOTAL"][0]
        print(f"{m:%Y-%m}: ok total {tot[4]:.0f} MW (prev-year {tot[2]:.0f}) {note}")
        time.sleep(1)
    if rep.empty:
        raise RuntimeError("no EPDK report could be read")
    rep["report_month"] = pd.to_datetime(rep["report_month"])
    rep["period"] = pd.to_datetime(rep["period"])
    rep = rep.sort_values(["report_month", "period", "col"]).reset_index(drop=True)
    monthly = build_monthly(rep)
    # fuel columns: sources absent in a period are 0; unmapped total check
    resid = (monthly["Total_MW"] - monthly[FUEL_COLS].sum(axis=1)).abs()
    print("max |total - sum of fuels| MW:", round(resid.max(), 2))
    rel = rel.sort_values("report_month").set_index("report_month")
    sheets = {"Monthly": monthly, "Revisions": revisions(rep), "Check_ENTSOE": entsoe_check(a.out_dir, monthly),
              "Reports": rep.set_index("report_month"), "Releases": rel}
    xlsx_notes.write_workbook(path, sheets, NOTES, TITLES)
    print(f"Turkey: {len(monthly)} months {monthly.index.min().date()} - {monthly.index.max().date()}")
    print(monthly.tail(3).T)


if __name__ == "__main__":
    main()
