"""
Nigeria gas production and sales from the NNPC Ltd Monthly Report Summary (one-page PDF).

Each monthly summary (nnpcgroup.com/insights/nnpc-limited-monthly-report-summary-<month>-<year>, PDF on
cms1977.nnpcgroup.com or the Azure front-door CMS host) carries trailing-12-month charts of
  - Natural gas production (mmscf/d)
  - Gas sales (mmscf/d)
  - Crude oil + condensate production (mmbopd), with the crude and condensate split.
The numbers are chart data labels, read from the PDF text layer. A series is accepted only if exactly 12 values
are found for 12 month labels (and, for oil, total = crude + condensate within 0.03). Otherwise the release is
listed as failed on 'Releases'; nothing is guessed.

Incremental: the committed workbook is the history store. Releases already read are not downloaded again except the
newest two (revision window). Where two releases cover a month, the most recent release wins ('release' column).
No gas-to-power / domestic / export split is in the summary (checked; see 'Releases' notes).

Usage: python3 NIGERIA_NNPC_GAS.py [--out nigeria_nnpc_gas_monthly.xlsx] [--full]
"""
import argparse
import io
import os
import re
import sys
from datetime import date

import pandas as pd
import requests

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import xlsx_notes  # noqa: E402

BASE = "https://www.nnpcgroup.com/insights/"
HEADERS = {"User-Agent": "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 Chrome/124 Safari/537.36"}
TIMEOUT = (15, 90)
MONTHS = ["january", "february", "march", "april", "may", "june", "july", "august", "september", "october",
          "november", "december"]
MON3 = {m[:3].title(): i + 1 for i, m in enumerate(MONTHS)}
SLUGS = ["nnpc-limited-monthly-report-summary-{m}-{y}", "nnpc-ltd-monthly-report-summary-{m}-{y}",
         "nnpc-limited-monthly-report-{m}-{y}"]
START = (2024, 6)  # earliest report month tried; months without a page are listed as missing
OUT_DEFAULT = "nigeria_nnpc_gas_monthly.xlsx"
COLS = ["gas_production_mmscfd", "gas_sales_mmscfd", "crude_condensate_mmbopd", "crude_mmbopd", "condensate_mmbopd"]


def get(url):
    return requests.get(url, headers=HEADERS, timeout=TIMEOUT)


def find_release(y, m):
    """(slug, pdf_url) for the report of month m (1-12) of year y, or None."""
    for pat in SLUGS:
        slug = pat.format(m=MONTHS[m - 1], y=y)
        try:
            r = get(BASE + slug)
        except requests.RequestException:
            continue
        if r.status_code != 200:
            continue
        pdfs = [u for u in re.findall(r'https://[^"\' <>]+\.pdf', r.text) if "onthly" in u]
        if pdfs:
            return slug, pdfs[0]
    return None


def pdf_text(url):
    from pypdf import PdfReader
    r = get(url)
    r.raise_for_status()
    return "\n".join((p.extract_text() or "") for p in PdfReader(io.BytesIO(r.content)).pages)


def block(text, start_pat, end_pat):
    s = re.search(start_pat, text)
    if not s:
        return None
    rest = text[s.end():]
    e = re.search(end_pat, rest)
    return rest[:e.start()] if e else rest


def month_labels(seg):
    labs = re.findall(r"([A-Z][a-z]{2})-(\d\d)", seg)
    out = [pd.Timestamp(2000 + int(yy), MON3[mm], 1) for mm, yy in labs if mm in MON3]
    return out[:12]


def parse(text):
    """-> DataFrame indexed by month with COLS, or raises ValueError."""
    t = text.replace(" ", " ")
    data = {}
    gp = block(t, r"Gas Production \(mmscf/d\)", r"Crude Oil & Condensate Production")
    gs = block(t, r"Gas Sales \(mmscf/d\)", r"Crude Oil & Condensate Sales")
    oil = block(t, r"Crude Oil & Condensate Production \(mmbopd\)", r"Upstream Pipeline|Gas Sales")
    for key, seg, pat in (("gas_production_mmscfd", gp, r"\b\d{4,5}\b"), ("gas_sales_mmscfd", gs, r"\b\d{4,5}\b")):
        if seg is None:
            raise ValueError(f"{key}: block not found")
        labs = month_labels(seg)
        nums = [float(x) for x in re.findall(pat, re.sub(r"[A-Z][a-z]{2}-\d\d", " ", seg))]
        if len(labs) != 12 or len(nums) != 12:
            raise ValueError(f"{key}: {len(labs)} labels, {len(nums)} values")
        data[key] = pd.Series(nums, index=labs)
    if oil is not None:
        labs = month_labels(oil)
        nums = [float(x) for x in re.findall(r"\b\d\.\d\d\b", re.sub(r"[A-Z][a-z]{2}-\d\d", " ", oil))]
        if len(labs) == 12 and len(nums) >= 36:
            tot, cr, co = (pd.Series(nums[i * 12:(i + 1) * 12], index=labs) for i in range(3))
            if ((tot - cr - co).abs() <= 0.03).all():
                data.update(crude_condensate_mmbopd=tot, crude_mmbopd=cr, condensate_mmbopd=co)
    df = pd.DataFrame(data)
    for c in COLS:
        if c not in df:
            df[c] = float("nan")
    df = df[COLS]
    if not df["gas_production_mmscfd"].between(1000, 15000).all() or not df["gas_sales_mmscfd"].between(500, 15000).all():
        raise ValueError("gas values outside plausible range")
    return df


def main():
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--out", default=OUT_DEFAULT)
    ap.add_argument("--full", action="store_true")
    args = ap.parse_args()

    monthly = pd.DataFrame()
    rel = pd.DataFrame({"report_month": pd.Series(dtype="datetime64[ns]"), "slug": pd.Series(dtype=object),
                        "pdf_url": pd.Series(dtype=object), "status": pd.Series(dtype=object),
                        "months_read": pd.Series(dtype="int64")})
    if os.path.exists(args.out) and not args.full:
        monthly = pd.read_excel(args.out, sheet_name="Monthly", index_col=0, parse_dates=True)
        rel = pd.read_excel(args.out, sheet_name="Releases")
        rel["report_month"] = pd.to_datetime(rel["report_month"])
        print(f"archive: {len(monthly)} months, {len(rel)} releases", flush=True)

    today = date.today()
    months = pd.period_range(f"{START[0]}-{START[1]:02d}", f"{today.year}-{today.month:02d}", freq="M")
    done = set(rel.loc[rel["status"] == "ok", "report_month"].dt.to_period("M"))
    newest_two = set(sorted(done)[-2:])
    todo = [p for p in months if p not in done or p in newest_two]
    print(f"{len(todo)} report months to try", flush=True)

    new_rows = {}
    for p in todo:
        hit = find_release(p.year, p.month)
        key = p.to_timestamp()
        if hit is None:
            if p not in done:
                rel = rel[rel["report_month"] != key]
                rel.loc[len(rel)] = [key, "", "", "no page", 0]
            continue
        slug, url = hit
        try:
            df = parse(pdf_text(url))
            status = "ok"
        except Exception as e:  # noqa: BLE001
            df, status = None, f"failed: {type(e).__name__}: {str(e)[:100]}"
        print(f"  {p}: {slug} -> {status}", flush=True)
        rel = rel[rel["report_month"] != key]
        rel.loc[len(rel)] = [key, slug, url, status, 0 if df is None else len(df)]
        if df is not None:
            new_rows[key] = df

    # newest report wins for each month it covers
    if "release" not in monthly.columns:
        monthly["release"] = pd.NaT
    for key in sorted(new_rows):
        df = new_rows[key].copy()
        df["release"] = key
        for idx in df.index:
            if idx in monthly.index and pd.notna(monthly.at[idx, "release"]) and monthly.at[idx, "release"] > key:
                df = df.drop(idx)
        monthly = pd.concat([monthly.drop(df.index, errors="ignore"), df])
    monthly = monthly.sort_index()
    monthly.index.name = "date"
    rel = rel.sort_values("report_month").reset_index(drop=True)
    if monthly.empty:
        raise SystemExit("no NNPC data read")
    print(monthly.tail(14).to_string(), flush=True)

    notes = [
        "UNITS",
        "gas_production_mmscfd and gas_sales_mmscfd: million standard cubic feet per day (monthly average rate, as "
        "published in the chart labels). crude_condensate_mmbopd, crude_mmbopd, condensate_mmbopd: million barrels "
        "of oil per day. 'release' = the report month whose summary supplied the value (the latest wins).",
        "",
        "SOURCE",
        "NNPC Ltd 'Monthly Report Summary' (one-page PDF, nnpcgroup.com/insights). Values are the data labels of "
        "the trailing-12-month charts, read from the PDF text layer; a series is used only when 12 labels and 12 "
        "values are found (oil: total = crude + condensate within 0.03), else the release is listed as failed on "
        "'Releases'.",
        "",
        "COVERAGE",
        "Reports found from June 2024 on (a month with no page is listed 'no page'); each report holds 12 months, so "
        "history starts about 12 months before the first report read. Gas sales is total sales (domestic + export "
        "together): the summary gives NO gas-to-power, domestic or export split. NNPC production is NNPC-reported "
        "(operator and JV gas); it is not NUPRC's national total.",
        "Checked and not usable (Oct 2026): NUPRC (nuprc.gov.ng, Next.js client-rendered, no data files in the "
        "sitemap), NMDPRA (nmdpra.gov.ng Blazor app, no files), NLNG (nlng.com Strapi CMS, no production tables), "
        "NBS elibrary (no gas production table; petroleum price watches and product distribution only), JODI-Gas "
        "(CSV/API need a login).",
    ]
    xlsx_notes.write_workbook(args.out, {"Monthly": monthly, "Releases": rel}, notes, {"UNITS", "SOURCE", "COVERAGE"})
    print(f"Saved {args.out}", flush=True)


if __name__ == "__main__":
    main()
