"""
Philippines installed and dependable generating capacity by plant type, annual (year end), from the Department of
Energy's Power Statistics: 'Installed and Dependable Capacity per Grid and per technology' PDF, linked from
https://doe.gov.ph/data-and-prices/energy-statistics/electric-power-industry/<YEAR>-power-statistics (files on
DOE's CloudFront media store). Found via discovery_archive/asia/CAPPRICE_DISCOVERY4-5.py.

The PDF has one page per measure and scope - Installed / Dependable capacity x On Grid / Off Grid / On Grid + Off
Grid - each with blocks for Luzon, Visayas, Mindanao and the Philippines: MW by plant type (Coal, Oil Based,
Natural Gas, Renewable Energy = Geothermal + Hydro + Biomass + Solar + Wind, Total, BESS) for the last five years.
Based on DOE's List of Existing Power Plants as of December of each year.

Writes output/Data and Chart Outputs/philippines_power_capacity.xlsx (standard capacity layout,
south_america/power_capacity_std.py; annual rows dated 1 January of the year, holding the year-end value):
  Monthly     installed capacity, Philippines, on grid + off grid (2021 on; the 'Grid' table, which then included
              off-grid plant, for 2003-2020): Coal, Oil (oil based), Gas, Hydro, Solar, Wind,
              Bioenergy (biomass), Other (geothermal), Total_MW; Battery_MW (BESS) outside the total;
              Geothermal_MW detail column
  Dependable  the same for dependable capacity
  By grid     every figure in the PDF (measure, scope, grid, year, plant type, MW), long format
  Release     the PDF read (URL, Last-Modified)

Whole-file source: the PDF is downloaded only when its URL or Last-Modified changes (recorded on the Release sheet);
years already saved are kept, so the history grows past the PDF's five-year window. Runs on the 1st and 15th.

    python3 asia/PHILIPPINES_DOE_CAPACITY.py [--out PATH]
"""
import argparse
import hashlib
import io
import os
import re
import sys
from datetime import date

import pandas as pd
import requests

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)
sys.path.insert(0, os.path.join(ROOT, "south_america"))
import power_capacity_std as cap_std  # noqa: E402

PAGE = "https://doe.gov.ph/data-and-prices/energy-statistics/electric-power-industry/{year}-power-statistics"
H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0 Safari/537.36"}
T = (20, 120)
OUT = os.path.join(ROOT, "output", "Data and Chart Outputs", "philippines_power_capacity.xlsx")
TYPE_FUEL = {"coal": "Coal", "oil based": "Oil", "natural gas": "Gas", "geothermal": "Other", "hydro": "Hydro",
             "biomass": "Bioenergy", "solar": "Solar", "wind": "Wind"}
GRIDS = ("Luzon", "Visayas", "Mindanao", "Philippines")


def out(*a):
    print(*a, flush=True)


def find_pdf():
    """Latest power-statistics page that links the capacity PDF -> (pdf url, page url)."""
    for y in range(date.today().year, date.today().year - 4, -1):
        page = PAGE.format(year=y)
        try:
            r = requests.get(page, headers=H, timeout=T)
        except requests.RequestException as e:
            out(f"  {page}: {e}")
            continue
        links = re.findall(r'https://[^"\'\s<>\\]+?\.pdf(?:\?[^"\'\s<>\\]*)?', r.text)
        sp = r"(?:%20|\s|\+)+"
        hit = [l for l in links if re.search(rf"Installed{sp}and{sp}Dependable{sp}Capacity{sp}per{sp}Grid", l, re.I)]
        out(f"  {page}: HTTP {r.status_code}, {len(links)} PDFs, capacity PDF {'found' if hit else 'not found'}")
        if hit:
            return hit[0].replace("&amp;", "&"), page
    return None, None


def parse(content):
    import pdfplumber   # only needed when a new release is downloaded
    rows = []
    with pdfplumber.open(io.BytesIO(content)) as pdf:
        for p in pdf.pages:
            text = p.extract_text() or ""
            # the page title names the measure (the notes mention both, so the title line is used)
            mt = re.search(r"(Installed|Dependable) Capacity per Grid", text)
            if not mt:
                continue
            measure = mt.group(1)
            st = re.search(r"In MW\s*\n\s*(On Grid\s*\+\s*Off Grid|Off Grid|On Grid)", text, re.I)
            sc = (st.group(1).lower() if st else "on grid").replace(" ", "")
            scope = "On + off grid" if "+" in sc else "Off grid" if sc.startswith("off") else "On grid"
            grid, years = None, []
            for line in text.splitlines():
                line = line.strip()
                ys = re.findall(r"\b((?:19|20)\d{2})\b", re.sub(r"%\s*Share.*$", "", line))
                head = next((g for g in GRIDS if line.startswith(g)), None)
                if head and len(ys) >= 2:
                    grid, years = head, [int(y) for y in ys]
                    continue
                if not grid:
                    continue
                m = re.match(r"^([A-Za-z][A-Za-z ()]*?)\s+((?:-?[\d,]*\.?\d+\s*)+?)(?:\s+-?[\d.]+%)?$", line)
                if not m:
                    continue
                vals = [float(v.replace(",", "")) for v in m.group(2).split()]
                if len(vals) != len(years):
                    continue
                for y, v in zip(years, vals):
                    rows.append({"measure": measure, "scope": scope, "grid": grid, "year": y,
                                 "plant_type": m.group(1).strip(), "MW": v})
    return pd.DataFrame(rows)


def standard(long, measure):
    """National series: 'On + off grid' where DOE gives it (2021 on); earlier years from the 'Grid' pages, which
    included off-grid plant until DOE separated it in 2021."""
    nat = long[(long["measure"] == measure) & (long["grid"] == "Philippines")]
    both = nat[nat["scope"] == "On + off grid"]
    early = nat[(nat["scope"] == "On grid") & ~nat["year"].isin(both["year"].unique())]
    d = pd.concat([early, both])
    if d.empty:
        return pd.DataFrame()
    w = d.pivot_table(index="year", columns="plant_type", values="MW", aggfunc="last")
    by = pd.DataFrame(index=w.index)
    for col in w.columns:
        f = TYPE_FUEL.get(col.lower())
        if f:
            by[f] = by.get(f, 0.0) + w[col].fillna(0.0)
    by.index = pd.to_datetime([f"{y}-01-01" for y in by.index])
    std = cap_std.standard(by)
    std["Geothermal_MW"] = w.get("Geothermal", pd.Series(0.0, index=w.index)).values
    std["Battery_MW"] = w.get("BESS", pd.Series(0.0, index=w.index)).values
    tot = w.get("Total")
    if tot is not None:
        std["Total_reported_MW"] = tot.values
    return std


def merge(old, new):
    if old.empty:
        return new
    if new.empty:
        return old
    return pd.concat([old[~old.index.isin(new.index)], new]).sort_index()


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", default=OUT)
    a = ap.parse_args()
    pdf_url, page = find_pdf()
    if not pdf_url:
        raise SystemExit("DOE capacity PDF not found")
    head = requests.head(pdf_url, headers=H, timeout=T, allow_redirects=True)
    lm = head.headers.get("last-modified") or head.headers.get("etag") or ""
    rel = cap_std.load_sheet(a.out, "Release")
    rel = rel[[c for c in rel.columns if not str(c).startswith("Unnamed")]]
    last = rel.iloc[-1].fillna("").astype(str) if not rel.empty else None
    if last is not None and lm and last.get("url") == pdf_url and last.get("last_modified") == lm:
        out(f"No new release (Last-Modified {lm}); workbook unchanged")
        return
    r = requests.get(pdf_url, headers=H, timeout=T)
    r.raise_for_status()
    md5 = hashlib.md5(r.content).hexdigest()
    if last is not None and last.get("md5") == md5:
        out(f"Same PDF as the last release read (md5 {md5}); workbook unchanged")
        return
    out(f"  PDF {pdf_url} -> {r.headers.get('content-type')} {len(r.content)} bytes")
    long = parse(r.content)
    if long.empty:   # layout changed: show what the PDF holds
        import pdfplumber
        with pdfplumber.open(io.BytesIO(r.content)) as pdf:
            for i, pg in enumerate(pdf.pages[:6]):
                out(f"=== page {i + 1}\n{(pg.extract_text() or '')[:1500]}")
        raise SystemExit("No capacity tables parsed from the PDF")
    out(long.groupby(["measure", "scope"])["year"].agg(["min", "max", "count"]).to_string())
    inst, dep = standard(long, "Installed"), standard(long, "Dependable")
    if inst.empty:
        raise SystemExit("No national installed capacity block found")
    for name, d in (("installed", inst), ("dependable", dep)):
        if not d.empty and "Total_reported_MW" in d:
            bad = d[(d["Total_MW"] - d["Total_reported_MW"]).abs() > 2]
            if len(bad):
                out(f"  !! {name}: fuel sum differs from DOE total in {list(bad.index.year)}")
    monthly = merge(cap_std.load_monthly(a.out), inst)
    old_dep = cap_std.load_sheet(a.out, "Dependable", index_col=0)
    if not old_dep.empty:
        old_dep.index = pd.to_datetime(old_dep.index)
    dependable = merge(old_dep, dep)
    old_long = cap_std.load_sheet(a.out, "By grid")
    if not old_long.empty:
        old_long = old_long[~old_long["year"].isin(long["year"].unique())]
        long = pd.concat([old_long, long], ignore_index=True)
    long = long.sort_values(["measure", "scope", "grid", "plant_type", "year"]).reset_index(drop=True)
    monthly.index.name = dependable.index.name = "date"
    release = pd.concat([rel, pd.DataFrame([{"read_on": date.today().isoformat(), "url": pdf_url, "page": page,
                                             "last_modified": lm, "md5": md5}])], ignore_index=True) if not rel.empty else \
        pd.DataFrame([{"read_on": date.today().isoformat(), "url": pdf_url, "page": page, "last_modified": lm, "md5": md5}])
    notes = [
        "UNITS",
        "MW. Monthly: installed (nameplate) capacity, Philippines, grid-connected + embedded + off-grid, at the end of "
        "each year; rows dated 1 January of that year (annual series). Coal; Oil = oil based; Gas = natural gas; Hydro; "
        "Solar; Wind; Bioenergy = biomass; Other = geothermal (also in Geothermal_MW). Total_MW = sum of fuels; "
        "Total_reported_MW = DOE's total (a check); Battery_MW = BESS, not in the total.",
        "Dependable: dependable capacity (capacity adjusted for ambient limitations), same layout.",
        "By grid: every figure in the PDF - measure (installed / dependable), scope (on grid / off grid / both), grid "
        "(Luzon, Visayas, Mindanao, Philippines), year, DOE plant type, MW.",
        "Release: the PDF read and its Last-Modified date; the file is downloaded again only when these change.",
        "",
        "COVERAGE",
        f"Annual, {monthly.index.min():%Y} to {monthly.index.max():%Y}. From 2021 the national figure is DOE's 'On Grid + "
        "Off Grid' table; earlier years come from DOE's 'Grid' table (2003 on), which included off-grid plant until DOE "
        "separated it in 2021 (off-grid is about 0.9 GW installed, mostly oil). Years already saved are kept when DOE's "
        "file drops them. Based on DOE's List of Existing Power Plants as of December of each year.",
        "",
        "SOURCE",
        f"Department of Energy (Philippines), Power Statistics - Installed and Dependable Capacity per Grid and per "
        f"technology: {page} (PDF {pdf_url.split('?')[0]}).",
    ]
    cap_std.write(a.out, monthly, {"Dependable": dependable, "By grid": long, "Release": release}, notes,
                  {"UNITS", "COVERAGE", "SOURCE"})


if __name__ == "__main__":
    main()
