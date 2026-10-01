"""
Jamaica power generation by source and natural gas use - ANNUAL, from the
Ministry's "Jamaica Energy Statistics" release (the finest granularity any
Jamaican official source publishes: JPS, the OUR and STATIN publish no
monthly or quarterly generation-by-fuel or gas series).

Source: Ministry of Energy, Transport and Telecommunications (formerly MSET),
Energy Division, Energy Economics and Planning Unit - "Jamaica Energy
Statistics <year>" (PDF, published each July, five years per edition), listed
on https://www.mset.gov.jm/document-category/statistics-data/
  Table 8  JPS Electricity Statistics - electricity generation (MWh) by
           Petroleum, Natural Gas, Hydro, Wind, Solar (JPS + IPPs, grid)
  Table 3  Petroleum consumption by products (bbl) - the Natural Gas row
           (barrels of oil equivalent, all uses: power and alumina)

  --what power -> jamaica_power_generation_daily.xlsx, sheet "Daily": one row per
                  YEAR dated 1 January, annual MWh, standard columns
  --what gas   -> jamaica_gas.xlsx, sheet "Gas use": one row per year

Incremental: only the newest edition is downloaded each run; its five years
replace the saved ones (later editions revise), older saved years are kept.

Usage: python3 JAMAICA_MSET_ENERGY_STATS.py --what power|gas [--out PATH] [--pdf URL]
"""

print("STARTING", flush=True)

import argparse
import os
import re
import sys

import pandas as pd
import requests

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import central_america_power_common as C  # noqa: E402

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))  # repo root, for xlsx_notes
import xlsx_notes  # noqa: E402

LIST_PAGE = "https://www.mset.gov.jm/document-category/statistics-data/"
OUTS = {"power": "output/Data and Chart Outputs/jamaica_power_generation_daily.xlsx",
        "gas": "output/Data and Chart Outputs/jamaica_gas.xlsx"}
START_YEAR = 2021
POWER_ROWS = {"Petroleum": "Oil", "Natural Gas": "Gas", "Hydro": "Hydro", "Wind": "Wind", "Solar": "Solar"}
MMBTU_PER_BOE = 5.8          # EIA's standard barrel-of-oil-equivalent heat content
BTU_PER_CF = 1037            # EIA average heat content of natural gas consumed
M3_PER_CF = 0.0283168466
M3_PER_BOE = MMBTU_PER_BOE * 1e6 / BTU_PER_CF * M3_PER_CF   # ~158 m3 per boe


def latest_edition():
    r = requests.get(LIST_PAGE, headers={"User-Agent": C.USER_AGENT}, timeout=60)
    r.raise_for_status()
    found = {}
    for href in re.findall(r"href=[\"']([^\"']+\.pdf)[\"']", r.text, re.I):
        m = re.search(r"ENERGY-STATISTICS-(\d{4})", href, re.I)
        if m:
            found.setdefault(int(m.group(1)), href)
    if not found:
        raise RuntimeError("no 'Jamaica Energy Statistics' PDF listed on " + LIST_PAGE)
    year = max(found)
    return year, found[year]


def tokens_of(pdf_bytes):
    import pymupdf
    doc = pymupdf.open(stream=pdf_bytes, filetype="pdf")
    return [[t.strip() for t in page.get_text().splitlines() if t.strip()] for page in doc]


def num(t):
    t = t.replace(",", "").strip()
    return float(t) if re.fullmatch(r"-?\d+(\.\d+)?", t) else None


def table(pages, title_pat, header, rows):
    """{row label: {year label: value}} from the page whose text matches title_pat. Each cell is its own line."""
    for toks in pages:
        if not any(re.search(title_pat, t, re.I) for t in toks):
            continue
        i = toks.index(header)
        years = []
        for t in toks[i + 1:]:
            if re.fullmatch(r"\d{4}[rp]?", t):
                years.append(t)
            else:
                break
        out = {}
        for label in rows:
            j = next(k for k, t in enumerate(toks) if t.strip() == label)
            vals = [num(t) for t in toks[j + 1:j + 1 + 3 * len(years)] if num(t) is not None][:len(years)]
            out[label] = dict(zip(years, vals))
        return out
    raise RuntimeError(f"table {title_pat!r} not found")


def to_frame(tbl):
    rows = []
    for label, by_year in tbl.items():
        for y, v in by_year.items():
            rows.append({"year": int(y[:4]), "flag": {"p": "preliminary", "r": "revised"}.get(y[4:], ""),
                         "label": label, "value": v})
    d = pd.DataFrame(rows)
    w = d.pivot_table(index="year", columns="label", values="value", aggfunc="last")
    flags = d.groupby("year")["flag"].agg(lambda s: next((f for f in s if f), ""))
    w.index = pd.to_datetime(w.index.astype(str) + "-01-01")
    flags.index = w.index
    return w, flags


def build_power(pages, edition):
    tbl = table(pages, r"TABLE 8\. JPS ELECTRICITY", "Category", [*POWER_ROWS, "Total Net Generation"])
    w, flags = to_frame(tbl)
    fuels = w[list(POWER_ROWS)].rename(columns=POWER_ROWS)
    std = C.standardise(fuels)
    check = (std["Total_MWh"] - w["Total Net Generation"]).abs().max()
    print(f"Table 8 sum vs published total: max difference {check:,.0f} MWh", flush=True)
    detail = w.copy()
    detail.columns = [f"{c.replace(' ', '_')}_MWh" for c in detail.columns]
    detail["status"] = flags
    detail["edition"] = edition
    return std, detail


def build_gas(pages, edition):
    tbl = table(pages, r"TABLE 3\. PETROLEUM CONSUMPTION BY PRODUCTS", "Product", ["Natural Gas"])
    w, flags = to_frame(tbl)
    days = pd.Series(w.index.is_leap_year, index=w.index).map({True: 366, False: 365})
    boe = w["Natural Gas"]
    g = pd.DataFrame(index=w.index)
    g["Total_mcm_per_day"] = (boe * M3_PER_BOE / 1e6 / days).round(4)
    g["Total_mcm_per_year"] = (boe * M3_PER_BOE / 1e6).round(2)
    g["Natural_gas_boe_per_year"] = boe
    g["status"] = flags
    g["edition"] = edition
    g.index.name = "Month"
    return g


def load(path, sheet):
    try:
        d = pd.read_excel(path, sheet_name=sheet, index_col=0)
        d.index = pd.to_datetime(d.index)
        return d
    except (FileNotFoundError, ValueError, KeyError, OSError):
        return pd.DataFrame()


def merge(new, old):
    """New edition's years win over saved ones; nothing before START_YEAR."""
    out = new if old.empty else pd.concat([old[~old.index.isin(new.index)], new])
    return out[out.index >= pd.Timestamp(f"{START_YEAR}-01-01")].sort_index()


def power_notes(daily, url):
    return [
        "UNITS",
        "MWh of net generation per YEAR. Jamaica publishes no monthly or quarterly generation by fuel, so each row is "
        "one calendar year, dated 1 January, holding the whole year's MWh (sheet name 'Daily' kept for the standard "
        "layout the master reads). Detail: the table as published, with its preliminary / revised flags.",
        "",
        "COVERAGE",
        f"{daily.index.min():%Y} to {daily.index.max():%Y} ({len(daily)} years), JPS grid: JPS plants and the "
        "independent power producers selling to JPS. Off-grid self-generation (alumina, cement) is not included.",
        "The latest year is preliminary until the next edition (published each July).",
        "",
        "SOURCE",
        "Ministry of Energy, Transport and Telecommunications (Jamaica), Energy Division - 'Jamaica Energy "
        "Statistics', Table 8 JPS Electricity Statistics (source: Jamaica Public Service and IPPs).",
        f"Edition used: {url}",
        f"List of editions: {LIST_PAGE}",
        "",
        "MAPPING",
        "Petroleum -> Oil; Natural Gas -> Gas; Hydro -> Hydro; Wind -> Wind; Solar -> Solar.",
    ]


def gas_notes(g, url):
    return [
        "UNITS",
        "ANNUAL data (Jamaica publishes no monthly or quarterly gas statistics). Month = 1 January of the year.",
        "Natural_gas_boe_per_year: the ministry's figure, barrels of oil equivalent (the table is in 'bbl').",
        f"Total_mcm_per_year / Total_mcm_per_day: converted at {MMBTU_PER_BOE} MMBtu per boe and {BTU_PER_CF} Btu per "
        f"cubic foot (EIA standard factors) = {M3_PER_BOE:.1f} m3 per boe; per day = year / days in year.",
        "",
        "COVERAGE",
        f"{g.index.min():%Y} to {g.index.max():%Y}. All Jamaican natural gas use - imported LNG (New Fortress "
        "terminals at Old Harbour and Montego Bay) burned for power (JPS Old Harbour / Bogue, JPS-JEP 190 MW, "
        "Jamalco CHP) and at the Jamalco alumina refinery. The ministry does not publish the split by use.",
        "The latest year is preliminary until the next edition.",
        "",
        "SOURCE",
        "Ministry of Energy, Transport and Telecommunications (Jamaica), Energy Division - 'Jamaica Energy "
        "Statistics', Table 3 Petroleum Consumption by Products, row Natural Gas.",
        f"Edition used: {url}",
        f"List of editions: {LIST_PAGE}",
        "Cross-check: EIA international data gives Jamaica natural gas consumption of 26 Bcf (0.7 bcm) in 2024.",
    ]


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--what", choices=["power", "gas"], required=True)
    ap.add_argument("--out", default=None)
    ap.add_argument("--pdf", default=None, help="edition URL (default: newest listed)")
    args = ap.parse_args()
    out = args.out or OUTS[args.what]

    if args.pdf:
        edition, url = int(re.search(r"(\d{4})", args.pdf.split("/")[-1]).group(1)), args.pdf
    else:
        edition, url = latest_edition()
    print(f"Edition {edition}: {url}", flush=True)
    r = requests.get(url, headers={"User-Agent": C.USER_AGENT}, timeout=120)
    r.raise_for_status()
    pages = tokens_of(r.content)

    if args.what == "power":
        std, detail = build_power(pages, edition)
        daily = merge(std, C.load_sheet(out, "Daily"))
        detail = merge(detail, load(out, "Detail"))
        C.write(out, daily, power_notes(daily, url), detail)
        print((daily / 1000).round(1).to_string(), flush=True)
    else:
        g = merge(build_gas(pages, edition), load(out, "Gas use"))
        os.makedirs(os.path.dirname(os.path.abspath(out)), exist_ok=True)
        xlsx_notes.write_workbook(out, {"Gas use": g}, gas_notes(g, url), {"UNITS", "COVERAGE", "SOURCE"})
        print(f"Saved {out}", flush=True)
        print(g.to_string(), flush=True)


if __name__ == "__main__":
    main()
