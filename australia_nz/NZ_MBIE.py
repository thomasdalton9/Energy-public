"""
New Zealand gas and generating capacity from MBIE's energy statistics
webtables (the official national statistics, xlsx, no key):

  monthly gas    https://www.mbie.govt.nz/assets/Data-Files/Energy/monthly-gas-webtable-<month>-<year>.xlsx
                 sheet Monthly_PJ: gross/net production, reinjection, LPG extracted, flared, stock change
  quarterly gas  .../nz-energy-quarterly-and-energy-in-nz/gas-quarterly-webtable-<month>-<year>.xlsx
                 consumption by sector
  electricity    .../nz-energy-quarterly-and-energy-in-nz/electricity-quarterly-webtable-<month>-<year>.xlsx
                 sheet "7 - Plant type (MW)": installed capacity by plant type

The file names carry the release month, so the current links are read from
MBIE's gas and electricity statistics pages each run.

Writes:
  nz_gas.xlsx              Monthly (PJ per month), Quarterly consumption (PJ per quarter)
  nz_power_capacity.xlsx   standard capacity layout (sheet "Monthly": date, <Fuel>_MW, Total_MW; one row per
                           period MBIE publishes) plus "By plant type" as published

Each release restates the whole history (MBIE revises back years), so each run reads the latest release
whole - one small file per table, nothing re-downloaded beyond that.

Usage: python3 NZ_MBIE.py [--gas-out ...nz_gas.xlsx] [--capacity-out ...nz_power_capacity.xlsx]
"""
import argparse
import io
import os
import re
import sys

import pandas as pd
import requests

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)  # repo root, for xlsx_notes
import xlsx_notes  # noqa: E402

MBIE = "https://www.mbie.govt.nz"
STATS = MBIE + "/building-and-energy/energy-and-natural-resources/energy-statistics-and-modelling/energy-statistics/"
GAS_PAGE = STATS + "gas-statistics/"
ELEC_PAGE = STATS + "electricity-statistics/"
HEADERS = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0"}
OUT_DIR = os.path.join(ROOT, "output", "Data and Chart Outputs")
# MBIE plant type -> standard capacity fuel (matched by regex on the row label, first match wins)
PLANT = [(r"hydro", "Hydro"), (r"geotherm", "Other"), (r"wind", "Wind"), (r"solar", "Solar"),
         (r"coal|gas|thermal|cogen|combined|peaker|diesel|oil", "Gas"), (r"bio|wood|waste", "Bioenergy"),
         (r"batter", "Battery_storage")]


def get(url):
    r = requests.get(url, headers=HEADERS, timeout=(10, 120))
    r.raise_for_status()
    return r


def link(page, pat):
    html = get(page).text
    found = [u for u in re.findall(r'href="([^"]+\.xlsx)"', html, re.I) if re.search(pat, u, re.I)]
    if not found:
        raise SystemExit(f"No link matching {pat} on {page}")
    u = found[0]
    return u if u.startswith("http") else MBIE + u


def period(v):
    """A header cell -> Timestamp: a date, a year (1950-2100, number or text) or a date string; else NaT."""
    if isinstance(v, pd.Timestamp) or hasattr(v, "year") and not isinstance(v, (int, float, str)):
        return pd.Timestamp(v)
    if isinstance(v, (int, float)) and not pd.isna(v) and 1950 <= v <= 2100 and float(v).is_integer():
        return pd.Timestamp(int(v), 1, 1)
    if isinstance(v, str):
        t = v.strip()
        if re.fullmatch(r"(19|20)\d\d", t):
            return pd.Timestamp(int(t), 1, 1)
        if re.search(r"\d{4}", t) and re.search(r"[-/ ]", t):
            return pd.to_datetime(t, errors="coerce", dayfirst=True)
    return pd.NaT


def wide(raw):
    """A webtable sheet with periods across the columns: the header row is the row with the most period cells; row
    labels are the text column left of the first period column. Returns a frame indexed by period."""
    best, best_n, best_p = None, 0, None
    for i in range(min(len(raw), 40)):
        p = raw.iloc[i].map(period)
        if p.notna().sum() > best_n:
            best, best_n, best_p = i, p.notna().sum(), p
    if best is None or best_n < 4:
        return None
    keep = [c for c in raw.columns if pd.notna(best_p[c])]
    left = [c for c in raw.columns if c < keep[0]]
    body = raw.iloc[best + 1:]
    label_col = max(left, key=lambda c: body[c].map(lambda x: isinstance(x, str)).sum()) if left else None
    if label_col is None:
        return None
    out = {}
    for i in body.index:
        label = raw.at[i, label_col]
        if not isinstance(label, str) or label.strip().lower().startswith(("note", "source", "return to")) \
                or re.match(r"^\d+\s", label.strip()):
            continue
        v = pd.to_numeric(raw.loc[i, keep], errors="coerce")
        if v.notna().sum() == 0:
            continue
        name = re.sub(r"(?<=[A-Za-z)])\d+$", "", re.sub(r"\s+", " ", label).strip()).strip()
        while name in out:
            name += " (2)"
        out[name] = v.values
    d = pd.DataFrame(out, index=pd.DatetimeIndex([best_p[c] for c in keep]))
    return d[~d.index.duplicated()].sort_index()


def long(raw):
    """A sheet with periods down the first column: header row = first row whose next rows start with dates."""
    dates = pd.to_datetime(raw.iloc[:, 0], errors="coerce")
    if dates.notna().sum() < 4:
        years = pd.to_numeric(raw.iloc[:, 0], errors="coerce")
        ok = years.between(1950, 2100)
        if ok.sum() < 4:
            return None
        dates = pd.to_datetime(years.where(ok).astype("Int64").astype(str) + "-01-01", errors="coerce")
    first = dates.first_valid_index()
    header = raw.iloc[:first].ffill().iloc[-1] if first else raw.iloc[0]
    d = raw.loc[dates.notna()].iloc[:, 1:].apply(pd.to_numeric, errors="coerce")
    d.columns = [re.sub(r"\s+", " ", str(h)).strip() for h in header.iloc[1:]]
    d.index = dates[dates.notna()]
    return d.dropna(axis=1, how="all").sort_index()


def table(content, sheet):
    raw = pd.read_excel(io.BytesIO(content), sheet_name=sheet, header=None)
    d = wide(raw)
    if d is None or d.empty:
        d = long(raw)
    if d is None or d.empty:
        print(f"  could not parse {sheet}; head:\n{raw.iloc[:15, :8].to_string()[:2500]}", flush=True)
        return pd.DataFrame()
    print(f"  {sheet}: {d.shape[0]} periods {d.index.min():%Y-%m}..{d.index.max():%Y-%m}; rows {list(d.columns)[:30]}",
          flush=True)
    return d


def gas(out):
    murl = link(GAS_PAGE, r"monthly-gas-webtable")
    print("monthly gas:", murl, flush=True)
    monthly = table(get(murl).content, "Monthly_PJ")
    monthly = monthly[monthly.index >= "2010-01-01"]
    monthly.index = monthly.index.to_period("M").to_timestamp()
    sheets = {"Monthly": monthly}
    qurl = None
    try:
        qurl = link(GAS_PAGE, r"gas-quarterly-webtable")
        print("quarterly gas:", qurl, flush=True)
        content = get(qurl).content
        xl = pd.ExcelFile(io.BytesIO(content))
        print("  sheets:", xl.sheet_names, flush=True)
        for sh in xl.sheet_names:
            if re.fullmatch(r"quarterly_?pj", sh, re.I):
                q = table(content, sh)
                if not q.empty:
                    sheets["Quarterly consumption"] = q[q.index >= "2010-01-01"]
                    break
        if "Quarterly consumption" not in sheets:
            for sh in xl.sheet_names:
                print(f"  --- {sh}\n" + pd.read_excel(xl, sheet_name=sh, header=None).iloc[:14, :6].to_string()[:1500])
    except Exception as e:  # noqa: BLE001 - the monthly table still gets written
        print(f"  quarterly gas failed: {type(e).__name__}: {e}", flush=True)
    for v in sheets.values():
        v.index.name = "date"
    notes = [
        "UNITS",
        "Petajoules (PJ) per month (Monthly) or per quarter (Quarterly consumption), gross calorific value, as "
        "MBIE publishes. 1 PJ is about 26 million m3 of New Zealand gas.",
        "Monthly: MBIE's monthly production and stock table - Gross Production, Gas Reinjected, LPG extracted, "
        "Gas Flared, Net Production, Manufactured Production, Stock Change (Ahuroa storage), production losses "
        "and own use, transmission and distribution losses. Latest months are provisional.",
        "",
        "COVERAGE",
        f"New Zealand. Monthly {monthly.index.min():%b %Y} to {monthly.index.max():%b %Y} (MBIE history from 1974, "
        "kept here from 2010).",
        "",
        "SOURCE",
        f"MBIE (Ministry of Business, Innovation and Employment) energy statistics: {murl}" + (f"; {qurl}" if qurl else ""),
        GAS_PAGE,
    ]
    xlsx_notes.write_workbook(out, {k: v.round(4) for k, v in sheets.items()}, notes, {"UNITS", "COVERAGE", "SOURCE"})
    print(f"Saved {out}", flush=True)


def capacity(out):
    url = link(ELEC_PAGE, r"electricity-quarterly-webtable")
    print("electricity:", url, flush=True)
    content = get(url).content
    sheet = next(s for s in pd.ExcelFile(io.BytesIO(content)).sheet_names if re.search(r"plant type.*MW", s, re.I))
    print(pd.read_excel(io.BytesIO(content), sheet_name=sheet, header=None).iloc[:30, :6].to_string()[:3500], flush=True)
    t = table(content, sheet)
    if t.empty:
        raise SystemExit("capacity table not parsed")
    cols = [c for c in t.columns if not re.search(r"total", c, re.I)]
    fuel = {c: next((f for p, f in PLANT if re.search(p, c, re.I)), "Other") for c in cols}
    print("  plant type -> fuel:", fuel, flush=True)
    m = t[cols].T.groupby(fuel).sum(min_count=1).T
    std = m.drop(columns=["Battery_storage"], errors="ignore").add_suffix("_MW")
    std["Total_MW"] = std.sum(axis=1, min_count=1)
    if "Battery_storage" in m:
        std["Battery_storage_MW"] = m["Battery_storage"]
    std = std[std.index >= "2010-01-01"]
    for x in (std, t):
        x.index.name = "date"
    notes = [
        "UNITS",
        "Installed generating capacity, MW, as at each period MBIE publishes (annual, dated 1 January of the "
        "year or at the period end). Monthly: MBIE plant types grouped - Hydro; Other = geothermal; Wind; "
        "Solar; Gas = all thermal plant (gas, coal/gas Huntly units, cogeneration, diesel peakers) since MBIE "
        "does not split thermal capacity by fuel; Bioenergy; Battery_storage_MW kept out of Total_MW.",
        "Mapping used: " + "; ".join(f"{k} -> {v}" for k, v in fuel.items()),
        "By plant type: as published.",
        "",
        "COVERAGE",
        f"New Zealand, {std.index.min():%Y} to {std.index.max():%Y}.",
        "",
        "SOURCE",
        f"MBIE electricity statistics, quarterly webtable, sheet '{sheet}': {url}",
        ELEC_PAGE,
    ]
    xlsx_notes.write_workbook(out, {"Monthly": std.round(1), "By plant type": t.round(1)}, notes,
                              {"UNITS", "COVERAGE", "SOURCE"})
    print(f"Saved {out}", flush=True)
    print(std.tail(3).to_string(), flush=True)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--gas-out", default=os.path.join(OUT_DIR, "nz_gas.xlsx"))
    ap.add_argument("--capacity-out", default=os.path.join(OUT_DIR, "nz_power_capacity.xlsx"))
    args = ap.parse_args()
    errors = []
    for fn, out in ((gas, args.gas_out), (capacity, args.capacity_out)):
        try:
            fn(out)
        except Exception as e:  # noqa: BLE001
            errors.append(f"{fn.__name__}: {type(e).__name__}: {e}")
            print("FAILED", errors[-1], flush=True)
    if len(errors) == 2:
        raise SystemExit("; ".join(errors))


if __name__ == "__main__":
    main()
