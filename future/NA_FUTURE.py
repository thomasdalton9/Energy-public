"""
North America "future" workbook (one-off, run on demand): what is planned to be built or retired, and the official
outlooks.

  US build / retire  EIA-860M monthly generator inventory workbook (latest month):
                     https://www.eia.gov/electricity/data/eia860m/  -> xls/<month>_generator<year>.xlsx
                     sheets Planned (planned operation year, status), Operating (planned retirement year),
                     Retired (actual retirements)
  US outlook         EIA Short-Term Energy Outlook, API v2 /steo (needs EIA_API_KEY): Henry Hub, dry gas production,
                     consumption, LNG exports, working gas, generation by fuel - monthly to the end of next year
  Canada outlook     Canada Energy Regulator, Canada's Energy Future 2026 open data (electricity capacity and
                     generation by fuel to 2050, by scenario)

Writes north_america_future.xlsx.

Usage: python3 future/NA_FUTURE.py [--out "output/Data and Chart Outputs/north_america_future.xlsx"]
"""
import argparse
import io
import os
import re
import sys

import pandas as pd

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)
sys.path.insert(0, os.path.join(ROOT, "americas"))
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import xlsx_notes  # noqa: E402
from future_common import by_year_fuel, col, get, read_table  # noqa: E402
from US_POWER_CAPACITY_EIA import SOURCE  # noqa: E402

EIA860M = "https://www.eia.gov/electricity/data/eia860m/"
STEO = "https://api.eia.gov/v2/steo/"
CKAN = "https://open.canada.ca/data/api/action/package_search"
DEFAULT_OUT = os.path.join(ROOT, "output", "Data and Chart Outputs", "north_america_future.xlsx")
MONTHS = ["january", "february", "march", "april", "may", "june", "july", "august", "september", "october",
          "november", "december"]
STEO_SERIES = {"NGHHUUS": "Henry Hub spot price ($/MMBtu)", "NGPRPUS": "Dry gas production (Bcf/d)",
               "NGTCPUS": "Gas consumption (Bcf/d)", "NGEXPUS_LNG": "LNG gross exports (Bcf/d)",
               "NGWGPUS": "Working gas in storage (Bcf)", "EPEOPUS": "Power sector generation (BkWh/d)"}
FUEL = {**{k: v for k, v in SOURCE.items()}, "MWH": "Battery_storage"}
NAMES = {"Battery_storage": "Battery storage", "Pumped_storage": "Pumped storage"}


def eia_fuel(row, tech_col, src_col):
    tech = row.get(tech_col, "")
    tech = "" if pd.isna(tech) else str(tech)
    if "Pumped Storage" in tech:
        return "Pumped storage"
    if re.search(r"Batter|Flywheel", tech):
        return "Battery storage"
    src = row.get(src_col, "")
    f = FUEL.get("" if pd.isna(src) else str(src).strip(), "Other")
    return NAMES.get(f, f)


def eia860m():
    html = get(EIA860M).text
    found = []
    for m in re.finditer(r'href="([^"]*?xls/(?!archive)[^"]*?([a-z]+)_generator(\d{4})\.xlsx)"', html, re.I):
        if "archive" in m.group(1) or m.group(2).lower() not in MONTHS:
            continue
        found.append((int(m.group(3)), MONTHS.index(m.group(2).lower()), m.group(1)))
    if not found:
        raise RuntimeError("no current EIA-860M workbook link")
    content = None
    for year, mon, href in sorted(set(found), reverse=True)[:6]:   # newest first; a link can precede its file
        url = href if href.startswith("http") else EIA860M + href.lstrip("/").replace("electricity/data/eia860m/", "")
        try:
            r = get(url)
        except Exception as e:  # noqa: BLE001
            print(f"  EIA-860M {MONTHS[mon].title()} {year}: {type(e).__name__}", flush=True)
            continue
        if r.content[:2] == b"PK":
            content = r.content
            break
        print(f"  EIA-860M {MONTHS[mon].title()} {year}: not an xlsx ({r.content[:60]!r})", flush=True)
    if content is None:
        raise RuntimeError("no EIA-860M workbook could be downloaded")
    print(f"EIA-860M: {MONTHS[mon].title()} {year} -> {url}", flush=True)
    must = [r"Entity ID", r"Technology"]
    planned = read_table(content, "Planned", must)
    operating = read_table(content, "Operating", must)
    retired = read_table(content, "Retired", must)
    print(f"  Planned {len(planned)} units, columns {list(planned.columns)[:30]}", flush=True)
    tech, src = col(planned, r"^Technology$"), col(planned, r"Energy Source Code")
    mw = col(planned, r"Net Summer Capacity", r"Nameplate Capacity")
    for d in (planned, operating, retired):
        d["Fuel"] = d.apply(lambda r: eia_fuel(r, tech, src), axis=1)
    status = col(planned, r"^Status$")
    planned["Stage"] = planned[status].fillna("").astype(str).map(
        lambda s: "Under construction or complete" if re.search(r"\((U|V|TS)\)", s) else "Approved or pending" if
        re.search(r"\((T|L)\)", s) else "Planned, not started")
    add = by_year_fuel(planned, col(planned, r"Planned Operation Year"), "Fuel", mw)
    stage = planned.assign(_y=pd.to_numeric(planned[col(planned, r"Planned Operation Year")], errors="coerce"),
                           _mw=pd.to_numeric(planned[mw], errors="coerce")).pivot_table(
        index="_y", columns="Stage", values="_mw", aggfunc="sum").fillna(0).round(1)
    stage.index = pd.to_datetime(stage.index.astype(int).astype(str) + "-01-01")
    stage.index.name = "Year"
    ret_plan = by_year_fuel(operating, col(operating, r"Planned Retirement Year"), "Fuel",
                            col(operating, r"Net Summer Capacity", r"Nameplate Capacity"))
    ret_done = by_year_fuel(retired, col(retired, r"Retirement Year"), "Fuel",
                            col(retired, r"Net Summer Capacity", r"Nameplate Capacity"), start=2021)
    keep = [c for c in (col(planned, r"Entity Name", required=False), col(planned, r"Plant Name", required=False),
                        col(planned, r"Plant State", required=False), tech, src, mw, status,
                        col(planned, r"Planned Operation Year"), col(planned, r"Planned Operation Month",
                                                                     required=False), "Fuel") if c]
    units = planned[keep].sort_values(keep[-3] if len(keep) > 3 else keep[0]).set_index(keep[0])
    label = f"{MONTHS[mon].title()} {year}"
    return {"US additions by fuel": add, "US additions by stage": stage, "US planned retirements": ret_plan,
            "US retirements since 2021": ret_done, "US planned units": units}, label, url


def steo(key):
    out = {}
    for sid, name in STEO_SERIES.items():
        r = get(STEO + "data/", params={"api_key": key, "frequency": "monthly", "data[0]": "value",
                                        "facets[seriesId][]": sid, "start": "2023-01", "length": 5000})
        rows = r.json()["response"]["data"]
        out[name] = pd.Series({pd.Timestamp(x["period"] + "-01"): float(x["value"]) for x in rows if x["value"]})
    # generation by fuel: STEO series named "... generation ..." for the electric power sector / all sectors
    facets = get(STEO + "facet/seriesId/", params={"api_key": key}).json()["response"]["facets"]
    # "Net generation from <fuel>, all sectors, United States" (the *TOGEN_US series), one per fuel
    gen = [f for f in facets if re.match(r"Net generation from .*, all sectors, United States$", f["name"])
           and not re.search(r"renewables excluding|other fossil", f["name"])]
    print(f"STEO generation series: {[(f['id'], f['name']) for f in gen][:30]}", flush=True)
    g = {}
    for f in gen[:14]:
        short = re.sub(r"^Net generation from |, all sectors, United States$", "", f["name"]).capitalize()
        r = get(STEO + "data/", params={"api_key": key, "frequency": "monthly", "data[0]": "value",
                                        "facets[seriesId][]": f["id"], "start": "2023-01", "length": 5000})
        g[short + " (BkWh/d)"] = pd.Series({pd.Timestamp(x["period"] + "-01"): float(x["value"])
                                  for x in r.json()["response"]["data"] if x["value"]})
    m = pd.DataFrame(out).sort_index()
    m.index.name = "Month"
    gg = pd.DataFrame(g).sort_index()
    gg.index.name = "Month"
    return m, gg


def cer():
    r = get(CKAN, params={"q": "Canada's Energy Future 2026", "rows": 5}).json()["result"]["results"]
    pkg = next((p for p in r if "2026" in p["title"]), None)
    if pkg is None:
        raise RuntimeError("Energy Futures 2026 package not found")
    res = [x["url"] for x in pkg["resources"] if "/open/energy/" in x["url"] and x["url"].endswith(".csv")
           and "dictionary" not in x["url"]]
    print(f"CER EF2026: {len(res)} CSVs: {[u.rsplit('/', 1)[-1] for u in res]}", flush=True)
    sheets = {}
    keep = {"electricity-capacity-technology": "capacity MW", "electricity-generation-technology": "generation GWh"}
    for u in res:
        kind = next((v for k, v in keep.items() if k in u), None)
        if kind is None:
            continue
        d = pd.read_csv(io.BytesIO(get(u).content), low_memory=False)
        print(f"  {u.rsplit('/', 1)[-1]}: regions {d['Region'].unique().tolist()}; scenarios "
              f"{d['Scenario'].unique().tolist()}", flush=True)
        t = d[d["Region"].astype(str).str.fullmatch("Canada", case=False)]
        if t.empty:   # provinces only: sum them
            t = d.groupby(["Scenario", "Variable", "Year"], as_index=False)["Value"].sum()
        for scen in t["Scenario"].dropna().unique():
            if not re.search(r"current|net.?zero", str(scen), re.I):   # the two headline scenarios
                continue
            p = t[t["Scenario"] == scen].pivot_table(index="Year", columns="Variable", values="Value", aggfunc="sum")
            p = p.loc[:, p.abs().sum() > 0]
            p.index = pd.to_datetime(p.index.astype(int).astype(str) + "-01-01")
            p.index.name = "Year"
            tag = "Net-zero" if re.search(r"net.?zero", str(scen), re.I) else "Current"
            sheets[f"CA {kind.split()[0]} {tag}"] = p.round(1)
    return sheets


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", default=DEFAULT_OUT)
    args = ap.parse_args()
    sheets, notes_src = {}, []
    try:
        s, label, url = eia860m()
        sheets.update(s)
        notes_src.append(f"EIA-860M {label}: {url}")
    except Exception as e:  # noqa: BLE001
        print(f"EIA-860M failed: {type(e).__name__}: {e}", flush=True)
    key = os.environ.get("EIA_API_KEY")
    if key:
        try:
            m, g = steo(key)
            sheets["US STEO outlook"] = m
            if not g.empty:
                sheets["US STEO generation"] = g
            notes_src.append(f"EIA Short-Term Energy Outlook, API v2: {STEO} (series {', '.join(STEO_SERIES)})")
        except Exception as e:  # noqa: BLE001
            print(f"STEO failed: {type(e).__name__}: {e}", flush=True)
    try:
        c = cer()
        sheets.update(c)
        notes_src.append("Canada Energy Regulator, Canada's Energy Future 2026 open data (open.canada.ca)")
    except Exception as e:  # noqa: BLE001
        print(f"CER failed: {type(e).__name__}: {e}", flush=True)
    if not sheets:
        raise SystemExit("nothing fetched")
    notes = [
        "UNITS",
        "US additions / retirements: MW (net summer capacity where given, else nameplate) by the year the unit is "
        "planned to start or retire, by fuel. Stage: under construction or complete (EIA status U, V, TS), "
        "approved or pending approval (T, L), planned not started (P). Planned retirements: operating units with a "
        "reported planned retirement year. Developers' dates slip - treat later years as indicative.",
        "US STEO: monthly forecast (and recent history) from EIA's Short-Term Energy Outlook; units in the column "
        "names.",
        "Canada: CER Energy Futures 2026 projections, Current Measures and Canada Net-zero scenarios (capacity MW, generation GWh by technology).",
        "",
        "COVERAGE",
        "United States and Canada. Mexico: SENER publishes its plan (PRODESEN) only as PDF - not included.",
        "",
        "SOURCE",
        *notes_src,
    ]
    os.makedirs(os.path.dirname(os.path.abspath(args.out)), exist_ok=True)
    xlsx_notes.write_workbook(args.out, sheets, notes, {"UNITS", "COVERAGE", "SOURCE"})
    print(f"Saved {args.out}: {list(sheets)}", flush=True)


if __name__ == "__main__":
    main()
