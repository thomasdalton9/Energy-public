"""
Canada power generation and natural gas from Statistics Canada (no key):

  Table 25-10-0015-01  Electric power generation, monthly generation by type of electricity
      -> canada_power_generation_daily.xlsx (standard power layout, one row per MONTH, MWh)
         plus a "Provinces" sheet (total generation by province)
  Table 25-10-0055-01  Supply and disposition of natural gas, monthly
      -> canada_gas.xlsx ("Supply and disposition": every item for Canada, million m3/day)

StatCan publishes each table as one CSV zip; there is no date-range
query, so each run first asks the WDS API for the table's latest month
(getCubeMetadata) and downloads only when it is newer than the saved data
(when one happens, every downloaded month replaces the saved one, since
StatCan revises recent months).

StatCan reports generation by TYPE OF PLANT, not by fuel: hydraulic,
tidal, wind, solar, nuclear, conventional steam (coal, gas, biomass and
oil boilers), combustion turbine (gas) and internal combustion (mostly
diesel in remote communities). These map to the standard columns as:
Hydro (hydraulic + tidal), Wind, Solar, Nuclear, Gas (combustion turbine),
Oil (internal combustion), Steam (conventional steam - mixed fuels, kept
separate), Other.

Usage: python3 CANADA_STATCAN.py [--power-out ...] [--gas-out ...]
"""
import argparse
import io
import os
import re
import sys
import time
import zipfile

import pandas as pd
import requests

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)  # repo root, for xlsx_notes
import xlsx_notes  # noqa: E402

DEFAULT_DIR = os.path.join(ROOT, "output", "Data and Chart Outputs")
CSV_ZIP = "https://www150.statcan.gc.ca/n1/tbl/csv/{pid}-eng.zip"
METADATA = "https://www150.statcan.gc.ca/t1/wds/rest/getCubeMetadata"
POWER_TABLE = 25100015
GAS_TABLE = 25100055
HISTORY_START = "2019-01-01"
HEADERS = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                         "(KHTML, like Gecko) Chrome/124.0 Safari/537.36"}

# Type of electricity generation -> standard column (first match wins; "Total ..." rows are skipped)
POWER_MAP = [
    (r"hydraulic|tidal", "Hydro"),
    (r"wind", "Wind"),
    (r"solar", "Solar"),
    (r"nuclear", "Nuclear"),
    (r"internal combustion", "Oil"),
    (r"combustion turbine", "Gas"),
    (r"conventional steam", "Steam"),
    (r".", "Other"),
]
POWER_COLS = ["Hydro", "Gas", "Wind", "Solar", "Nuclear", "Oil", "Steam", "Other"]


def get(url, **kw):
    last = None
    for attempt in range(3):
        try:
            r = requests.get(url, headers=HEADERS, timeout=(10, 300), **kw)
            r.raise_for_status()
            return r
        except requests.RequestException as e:
            last = e
            print(f"    attempt {attempt + 1}/3 failed: {type(e).__name__}: {str(e)[:200]}", flush=True)
            time.sleep(10 * (attempt + 1))
    raise last


def latest_month(pid):
    """The table's last reference month per StatCan's WDS metadata (None if the call fails)."""
    try:
        r = requests.post(METADATA, json=[{"productId": pid}], headers=HEADERS, timeout=(10, 60))
        r.raise_for_status()
        obj = r.json()[0]["object"]
        end = pd.Timestamp(obj["cubeEndDate"]).to_period("M").to_timestamp()
        print(f"  table {pid}: '{obj.get('cubeTitleEn')}', {obj.get('cubeStartDate')} to {obj['cubeEndDate']}, "
              f"released {obj.get('releaseTime')}", flush=True)
        return end
    except Exception as e:  # noqa: BLE001 - just download then
        print(f"  table {pid}: metadata call failed ({type(e).__name__}: {str(e)[:200]})", flush=True)
        return None


def download(pid):
    """Whole table as a DataFrame (the CSV inside StatCan's zip), from HISTORY_START, values in base units."""
    url = CSV_ZIP.format(pid=pid)
    print(f"  downloading {url}", flush=True)
    z = zipfile.ZipFile(io.BytesIO(get(url).content))
    name = next(n for n in z.namelist() if n.endswith(".csv") and "MetaData" not in n)
    d = pd.read_csv(z.open(name), dtype=str, encoding="utf-8-sig")
    d["date"] = pd.to_datetime(d["REF_DATE"], errors="coerce")
    d = d[d["date"] >= HISTORY_START].copy()
    scale = pd.to_numeric(d.get("SCALAR_ID", "0"), errors="coerce").fillna(0)
    d["value"] = pd.to_numeric(d["VALUE"], errors="coerce") * (10.0 ** scale)
    dims = list(d.columns[d.columns.get_loc("DGUID") + 1:d.columns.get_loc("UOM")])
    print(f"  {len(d)} rows from {HISTORY_START}; dimensions {dims}; UOM {sorted(d['UOM'].dropna().unique())}",
          flush=True)
    for dim in dims:
        print(f"    {dim}: {sorted(d[dim].dropna().unique())}", flush=True)
    return d, dims


def load(path, sheet):
    try:
        d = pd.read_excel(path, sheet_name=sheet, index_col=0)
    except (FileNotFoundError, ValueError, KeyError, OSError):
        return pd.DataFrame()
    d.index = pd.to_datetime(pd.Index(d.index).astype(str), errors="coerce")
    return d[d.index.notna()].sort_index()


def merge(new, saved):
    """Downloaded months win (StatCan revises); saved months the download no longer covers are kept."""
    if saved.empty:
        return new
    return new.combine_first(saved)[list(new.columns) + [c for c in saved.columns if c not in new.columns]]


def needs_update(pid, saved, force):
    end = latest_month(pid)
    if force or saved.empty or end is None:
        return True
    if saved.index.max() >= end:
        print(f"  table {pid}: saved data already runs to {saved.index.max():%Y-%m} - no download", flush=True)
        return False
    return True


def power(d, dims):
    cls = next(c for c in dims if "class" in c.lower())
    typ = next(c for c in dims if "type" in c.lower())
    d = d[d[cls].str.contains("total all classes", case=False, na=False)]
    d = d[d["UOM"].str.contains("megawatt", case=False, na=False)]

    def fuel(t):
        if t.lower().startswith("total"):
            return None
        return next(col for pat, col in POWER_MAP if re.search(pat, t, re.I))

    ca = d[d["GEO"] == "Canada"].copy()
    ca["fuel"] = ca[typ].map(fuel)
    for t, f in sorted(set(zip(ca[typ], ca["fuel"]))):
        print(f"    {t!r} -> {f}", flush=True)
    w = ca.dropna(subset=["fuel"]).pivot_table(index="date", columns="fuel", values="value", aggfunc="sum")
    w = w.reindex(columns=[c for c in POWER_COLS if c in w.columns])
    total = ca[ca[typ].str.match(r"total all types", case=False, na=False)].groupby("date")["value"].sum()
    out = w.add_suffix("_MWh")
    out["Total_MWh"] = total.reindex(out.index) if len(total) else w.sum(axis=1, min_count=1)
    out = out.round(0)
    out.index.name = "date"
    prov = d[d[typ].str.match(r"total all types", case=False, na=False) & (d["GEO"] != "Canada")]
    prov = prov.pivot_table(index="date", columns="GEO", values="value", aggfunc="sum").round(0)
    prov = prov[prov.sum().sort_values(ascending=False).index].add_suffix("_MWh")
    prov.index.name = "date"
    return out, prov


def gas(d, dims):
    d = d[(d["GEO"] == "Canada") & d["UOM"].str.contains("metre", case=False, na=False)].copy()
    d["item"] = d[dims].fillna("").agg(" | ".join, axis=1).str.strip(" |")
    w = d.pivot_table(index="date", columns="item", values="value", aggfunc="sum")
    w = w.div(w.index.days_in_month, axis=0) / 1e6          # m3 per month -> million m3 per day
    w = w[[c for c in d["item"].drop_duplicates() if c in w.columns]].round(3)   # StatCan's own item order
    w.columns = [f"{c} (mcm/d)" for c in w.columns]
    w.index.name = "Month"
    return w


def save(path, sheets, notes):
    os.makedirs(os.path.dirname(os.path.abspath(path)), exist_ok=True)
    xlsx_notes.write_workbook(path, sheets, notes, {"UNITS", "COVERAGE", "SOURCE"})
    print(f"Saved {path}: " + ", ".join(f"{k} {len(v)} rows" for k, v in sheets.items()), flush=True)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--power-out", default=os.path.join(DEFAULT_DIR, "canada_power_generation_daily.xlsx"))
    ap.add_argument("--gas-out", default=os.path.join(DEFAULT_DIR, "canada_gas.xlsx"))
    ap.add_argument("--force", action="store_true", help="download even if the saved data look current")
    args = ap.parse_args()
    ok = 0

    print("Power generation (table 25-10-0015-01)", flush=True)
    saved, saved_prov = load(args.power_out, "Daily"), load(args.power_out, "Provinces")
    try:
        if needs_update(POWER_TABLE, saved, args.force):
            out, prov = power(*download(POWER_TABLE))
            out, prov = merge(out, saved), merge(prov, saved_prov)
            notes = [
                "UNITS",
                "MWh per MONTH (this standard 'Daily' layout holds one row per month, dated the 1st). "
                "Provinces: total generation by province, MWh per month.",
                "Columns are StatCan's types of plant: Hydro = hydraulic + tidal turbines; Gas = combustion "
                "turbines; Oil = internal combustion (mostly diesel, remote communities); Steam = conventional "
                "steam turbines, burning coal, gas, biomass or oil (StatCan does not split this monthly by fuel); "
                "Wind, Solar, Nuclear as named; Total_MWh = StatCan's 'Total all types'.",
                "Class of producer: total all classes (utilities and industry).",
                "",
                "COVERAGE",
                f"Canada, monthly, {out.index.min():%b %Y} to {out.index.max():%b %Y}. StatCan publishes about "
                "two to three months after the month ends.",
                "",
                "SOURCE",
                "Statistics Canada, Table 25-10-0015-01 Electric power generation, monthly generation by type of "
                "electricity: https://www150.statcan.gc.ca/t1/tbl1/en/tv.action?pid=2510001501",
            ]
            save(args.power_out, {"Daily": out, "Provinces": prov}, notes)
        ok += 1
    except Exception as e:  # noqa: BLE001 - still try the gas table
        print(f"POWER FAILED: {type(e).__name__}: {e}", flush=True)

    print("Natural gas (table 25-10-0055-01)", flush=True)
    saved = load(args.gas_out, "Supply and disposition")
    try:
        if needs_update(GAS_TABLE, saved, args.force):
            new = gas(*download(GAS_TABLE))
            new.index = pd.to_datetime(new.index)
            out = merge(new, saved)
            out.index = out.index.strftime("%Y-%m")
            out.index.name = "Month"
            notes = [
                "UNITS",
                "Million cubic metres per day (mcm/d), monthly average = StatCan monthly volume / days in month. "
                "One column per StatCan 'Supply and disposition' item, Canada total, in StatCan's order.",
                "",
                "COVERAGE",
                f"Canada, monthly, {out.index.min()} to {out.index.max()}. StatCan publishes about "
                "two to three months after the month ends.",
                "",
                "SOURCE",
                "Statistics Canada, Table 25-10-0055-01 Supply and disposition of natural gas, monthly: "
                "https://www150.statcan.gc.ca/t1/tbl1/en/tv.action?pid=2510005501",
            ]
            save(args.gas_out, {"Supply and disposition": out}, notes)
        ok += 1
    except Exception as e:  # noqa: BLE001
        print(f"GAS FAILED: {type(e).__name__}: {e}", flush=True)
    if not ok:
        raise SystemExit("Neither StatCan table could be updated")


if __name__ == "__main__":
    main()
