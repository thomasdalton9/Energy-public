"""
Europe natural gas consumption by sector from Eurostat (raw statistical-office source: national
statistics reported by member states under Regulation (EU) 2016/1952 and the annual energy
questionnaires), EU-27 + Norway + UK (to 2019/2020, Brexit) + candidate/neighbour countries where
Eurostat publishes them.

STATUS: UNVERIFIED against the live API. The Claude sandbox blocks ec.europa.eu, so this script was
written from the documented Eurostat dissemination API (JSON-stat 2.0) and its JSON-stat parser was
unit-tested offline on a synthetic payload only (discovery_archive/world/TEST_EUROSTAT_JSONSTAT.py).
First run in GitHub Actions: check the run log (it prints the unit chosen, the codes returned and the
expected sector codes that were missing).

Two Eurostat datasets (JSON-stat API, no key):
  nrg_bal_c   Complete energy balances, ANNUAL, siec G3000 (natural gas), unit TJ (GCV where the
              balance is in GCV; the dimension is read from the response). This is the real sector
              split: transformation input to power and heat, industry, transport, households,
              commerce and public services, agriculture, non-energy use, energy sector, losses.
              History from 1990, latest year about t+11 months.
  nrg_cb_gasm Supply, transformation and consumption of gas, MONTHLY (Reg. 2016/1952), siec G3000.
              Every balance item Eurostat returns is kept with its label; which items exist (e.g.
              whether monthly consumption is split by final-use sector or only gross inland
              deliveries / power input) is read from the response, not assumed. History from 2008,
              latest month about t+2 months.

Incremental (CLAUDE.md): the committed workbook is the history store. Each run reads the long raw
tabs back, fetches per country only periods after the last saved one minus a revision window
(monthly: 6 months, annual: 3 years) and upserts. First run (no workbook) pulls full history.

Outputs: output/Data and Chart Outputs/eurostat_gas_by_sector.xlsx with
  Units, Annual raw (long), Monthly raw (long), EU27 annual by sector, EU27 monthly (wide),
  plus chart sheets added by add_charts.py (registry entry "eurostat_gas_by_sector.xlsx").

    python3 europe/EUROSTAT_GAS_BY_SECTOR.py --out "output/Data and Chart Outputs/eurostat_gas_by_sector.xlsx"
"""
import argparse
import os
import sys
import time

import pandas as pd
import requests

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))  # repo root
import xlsx_notes  # noqa: E402

API = "https://ec.europa.eu/eurostat/api/dissemination/statistics/1.0/data"
HEADERS = {"User-Agent": "gas-demand-scripts/1.0 (+github.com/thomasdalton9/Energy)", "Accept": "application/json"}
TIMEOUT = (15, 120)
ATTEMPTS = 3
SLEEP = [5, 20]

ANNUAL_DS, MONTHLY_DS = "nrg_bal_c", "nrg_cb_gasm"
SIEC = "G3000"                       # natural gas
UNIT_PREFERENCE = ["TJ_GCV", "TJ"]   # first one present in the dataset wins
MONTHLY_REVISION_MONTHS = 6
ANNUAL_REVISION_YEARS = 3
ANNUAL_START = "1990"
MONTHLY_START = "2008-01"

# Core: EU-27 aggregate + members + Norway + UK. Others are tried and skipped quietly if Eurostat has no data.
CORE = ["EU27_2020", "AT", "BE", "BG", "HR", "CY", "CZ", "DK", "EE", "FI", "FR", "DE", "EL", "HU", "IE", "IT", "LV",
        "LT", "LU", "MT", "NL", "PL", "PT", "RO", "SK", "SI", "ES", "SE", "NO", "UK"]
EXTRA = ["TR", "UA", "MD", "RS", "BA", "MK", "AL", "ME", "GE", "XK"]
GEOS = CORE + EXTRA

# nrg_bal_c balance codes kept (annual). Sector view used by the charts: SECTOR_GROUPS.
ANNUAL_CODES = ["GIC", "GAE", "FC_E", "FC_IND_E", "FC_TRA_E", "FC_OTH_CP_E", "FC_OTH_HH_E", "FC_OTH_AF_E",
                "FC_OTH_FISH_E", "FC_OTH_NSP_E", "FC_OTH_E", "FC_NE", "TI_EHG_E", "TI_E", "NRG_E", "DL_E"]
SECTOR_GROUPS = {
    "Power and heat generation (transformation input)": ["TI_EHG_E"],
    "Industry (energy use)": ["FC_IND_E"],
    "Households": ["FC_OTH_HH_E"],
    "Commerce and public services": ["FC_OTH_CP_E"],
    "Transport": ["FC_TRA_E"],
    "Agriculture, fishing, other": ["FC_OTH_AF_E", "FC_OTH_FISH_E", "FC_OTH_NSP_E"],
    "Non-energy use": ["FC_NE"],
    "Energy sector own use and losses": ["NRG_E", "DL_E"],
}

RAW_COLS = ["geo", "code", "label", "period", "value"]
UNITS_SECTIONS = {"UNITS", "SOURCE", "WHAT IS ON EACH TAB", "CAVEATS", "RUN STATE"}


def out(*a):
    print(*a, flush=True)


# ------------------------------------------------------------------ Eurostat JSON-stat

def get_json(dataset, params):
    """GET one JSON-stat document; None when Eurostat has no such data (400/404 for unknown codes)."""
    last = None
    for attempt in range(ATTEMPTS):
        try:
            r = requests.get(f"{API}/{dataset}", params=params, headers=HEADERS, timeout=TIMEOUT)
            if r.status_code in (400, 404):
                out(f"  {dataset} {params.get('geo', '')}: HTTP {r.status_code} {r.text[:160]!r}")
                return None
            r.raise_for_status()
            return r.json()
        except (requests.RequestException, ValueError) as exc:
            last = exc
            if attempt < ATTEMPTS - 1:
                time.sleep(SLEEP[attempt])
    raise RuntimeError(f"{dataset} {params}: {last}")


def _positions(category):
    """JSON-stat category index (dict code->pos, or list of codes) -> list of codes in position order."""
    idx = category.get("index")
    if isinstance(idx, dict):
        return [c for c, _ in sorted(idx.items(), key=lambda kv: kv[1])]
    if isinstance(idx, list):
        return list(idx)
    return list(category.get("label", {}))


def parse_jsonstat(doc):
    """JSON-stat 2.0 (Eurostat) -> long DataFrame with one column per dimension (codes), '<dim>_label'
    columns for the dimensions with labels, and 'value'. Missing cells are dropped (value absent)."""
    ids = doc["id"]
    sizes = doc["size"]
    cats = {d: _positions(doc["dimension"][d]["category"]) for d in ids}
    labels = {d: doc["dimension"][d]["category"].get("label", {}) for d in ids}
    values = doc.get("value", {})
    items = values.items() if isinstance(values, dict) else ((i, v) for i, v in enumerate(values))
    strides = []
    acc = 1
    for s in reversed(sizes):
        strides.append(acc)
        acc *= s
    strides = strides[::-1]
    rows = []
    for k, v in items:
        if v is None:
            continue
        k = int(k)
        row = {}
        for d, stride, size in zip(ids, strides, sizes):
            code = cats[d][(k // stride) % size]
            row[d] = code
            row[f"{d}_label"] = labels[d].get(code, code)
        row["value"] = float(v)
        rows.append(row)
    return pd.DataFrame(rows)


def dimension_codes(doc, dim):
    return _positions(doc["dimension"][dim]["category"]) if dim in doc.get("dimension", {}) else []


def choose_unit(dataset):
    """Read the unit dimension from a small request (no unit filter) and pick the preferred TJ unit."""
    doc = get_json(dataset, {"format": "JSON", "lang": "EN", "geo": "DE", "siec": SIEC,
                             "sinceTimePeriod": "2023-01" if dataset == MONTHLY_DS else "2022"})
    if doc is None:
        raise RuntimeError(f"{dataset}: probe request failed - check the dataset id / filters in the run log")
    units = dimension_codes(doc, "unit")
    out(f"{dataset}: units offered {units}; dimensions {doc['id']}")
    for u in UNIT_PREFERENCE:
        if u in units:
            return u
    tj = [u for u in units if u.startswith("TJ")]
    if tj:
        return tj[0]
    raise RuntimeError(f"{dataset}: no TJ unit among {units}")


def pull(dataset, geo, unit, since):
    params = {"format": "JSON", "lang": "EN", "geo": geo, "siec": SIEC, "unit": unit}
    if since:
        params["sinceTimePeriod"] = since
    doc = get_json(dataset, params)
    if doc is None or not doc.get("value"):
        return pd.DataFrame(columns=RAW_COLS)
    df = parse_jsonstat(doc)
    if df.empty:
        return pd.DataFrame(columns=RAW_COLS)
    df = df.rename(columns={"nrg_bal": "code", "nrg_bal_label": "label", "time": "period"})
    return df[RAW_COLS]


# ------------------------------------------------------------------ incremental store

def load_raw(path, sheet):
    if not os.path.exists(path):
        return pd.DataFrame(columns=RAW_COLS)
    try:
        df = pd.read_excel(path, sheet_name=sheet, dtype={"period": str})
    except (ValueError, KeyError):
        return pd.DataFrame(columns=RAW_COLS)
    return df[RAW_COLS] if set(RAW_COLS) <= set(df.columns) else pd.DataFrame(columns=RAW_COLS)


def _shift(period, monthly, n):
    if monthly:
        p = pd.Period(period, freq="M") - n
        return str(p)
    return str(int(period) - n)


def since_for(saved, geo, monthly):
    s = saved[saved["geo"] == geo]
    if s.empty:
        return MONTHLY_START if monthly else ANNUAL_START
    last = s["period"].max()
    return _shift(last, monthly, MONTHLY_REVISION_MONTHS if monthly else ANNUAL_REVISION_YEARS)


def refresh(saved, dataset, unit, monthly):
    parts = [saved]
    for geo in GEOS:
        new = pull(dataset, geo, unit, since_for(saved, geo, monthly))
        if not monthly:
            new = new[new["code"].isin(ANNUAL_CODES)]
        out(f"  {dataset} {geo}: {len(new)} cells")
        if len(new):
            parts.append(new)
        time.sleep(0.3)
    both = pd.concat(parts, ignore_index=True)
    both["period"] = both["period"].astype(str)
    both = both.drop_duplicates(["geo", "code", "period"], keep="last")   # newer pull wins
    return both.sort_values(["geo", "code", "period"]).reset_index(drop=True)


# ------------------------------------------------------------------ wide views

def annual_sectors(raw, geo):
    d = raw[raw["geo"] == geo]
    wide = d.pivot_table(index="period", columns="code", values="value", aggfunc="first")
    out_df = pd.DataFrame(index=wide.index)
    for name, codes in SECTOR_GROUPS.items():
        have = [c for c in codes if c in wide.columns]
        if have:
            out_df[name] = wide[have].sum(axis=1, min_count=1)
    out_df.index = pd.to_datetime(out_df.index.astype(str) + "-01-01")
    out_df.index.name = "Year"
    return out_df


def monthly_wide(raw, geo):
    d = raw[raw["geo"] == geo]
    d = d.assign(name=d["label"].astype(str) + " [" + d["code"].astype(str) + "]")
    wide = d.pivot_table(index="period", columns="name", values="value", aggfunc="first")
    wide.index = pd.to_datetime(wide.index.astype(str) + "-01")
    wide.index.name = "Month"
    return wide


# ------------------------------------------------------------------ main

def build_notes(unit_a, unit_m, annual, monthly):
    last_a = annual["period"].max() if len(annual) else "n/a"
    last_m = monthly["period"].max() if len(monthly) else "n/a"
    codes_m = sorted(monthly["code"].unique()) if len(monthly) else []
    missing = [c for c in ANNUAL_CODES if c not in set(annual["code"])]
    lines = [
        "UNITS",
        f"Annual: {unit_a} (terajoules; Eurostat nrg_bal_c, natural gas SIEC G3000). Monthly: {unit_m} (nrg_cb_gasm).",
        "Values are as published by Eurostat, not converted. 1 TJ is about 0.0278 million m3 at 36 MJ/m3 (GCV); "
        "use the country's own conversion for precision.",
        "",
        "SOURCE",
        "Eurostat dissemination API (JSON-stat), datasets nrg_bal_c (annual complete energy balances) and "
        "nrg_cb_gasm (monthly gas supply, transformation and consumption). National statistics offices report to "
        "Eurostat; this is the raw official EU statistic, not a secondary compilation.",
        "https://ec.europa.eu/eurostat/api/dissemination/statistics/1.0/data/nrg_bal_c",
        "https://ec.europa.eu/eurostat/api/dissemination/statistics/1.0/data/nrg_cb_gasm",
        "Geographies: EU27_2020 aggregate, 27 members (EL = Greece), Norway, UK (Eurostat holds UK only to the "
        "Brexit cut-off, so use DESNZ Energy Trends for later UK data), plus Turkey and Western Balkans/Ukraine/"
        "Moldova/Georgia where Eurostat publishes them.",
        "",
        "WHAT IS ON EACH TAB",
        "Annual raw / Monthly raw: long tables (geo, balance code, label, period, value) = the history store; each "
        "run re-fetches only periods after the last saved one minus a revision window and upserts.",
        "EU27 annual by sector: sector groups summed from annual balance codes (see SECTOR_GROUPS in the script): "
        "power and heat = TI_EHG_E; industry = FC_IND_E; households = FC_OTH_HH_E; commerce and public = FC_OTH_CP_E; "
        "transport = FC_TRA_E; agriculture/fishing/other = FC_OTH_AF_E + FC_OTH_FISH_E + FC_OTH_NSP_E; non-energy = "
        "FC_NE; energy sector and losses = NRG_E + DL_E.",
        "EU27 monthly: every monthly balance item Eurostat returns for the EU-27 aggregate, labelled by Eurostat.",
        "Chart sheets are rebuilt by add_charts.py on every run (native Excel charts, mmm/yy dates, no borders).",
        "",
        "CAVEATS",
        "Power/heat input includes CHP fuel input and autoproducers; it is gas burned for generation, not "
        "electricity output. Annual sector split lags about 11 months; monthly lags about 2 months.",
        "Whether the monthly dataset splits final consumption by sector depends on what member states report; the "
        "monthly tab shows exactly the items returned (see run state below), nothing is inferred.",
        "Gross vs net calorific value differs between datasets (check the unit suffix above).",
        "",
        "RUN STATE",
        f"Last annual period saved: {last_a}. Last monthly period saved: {last_m}.",
        f"Monthly balance codes returned: {', '.join(codes_m) if codes_m else 'none'}.",
        f"Expected annual codes not returned by Eurostat: {', '.join(missing) if missing else 'none'}.",
        f"Updated: {pd.Timestamp.now('UTC'):%Y-%m-%d %H:%M} UTC.",
    ]
    return lines


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", default=os.path.join("output", "Data and Chart Outputs", "eurostat_gas_by_sector.xlsx"))
    args = ap.parse_args()
    os.makedirs(os.path.dirname(args.out) or ".", exist_ok=True)

    unit_a = choose_unit(ANNUAL_DS)
    unit_m = choose_unit(MONTHLY_DS)
    out(f"units: annual {unit_a}, monthly {unit_m}")

    annual = refresh(load_raw(args.out, "Annual raw"), ANNUAL_DS, unit_a, monthly=False)
    monthly = refresh(load_raw(args.out, "Monthly raw"), MONTHLY_DS, unit_m, monthly=True)
    if annual.empty and monthly.empty:
        raise SystemExit("Eurostat returned nothing - not overwriting the workbook")

    sheets = {"Annual raw": annual, "Monthly raw": monthly}
    sheets = {k: v.set_index("geo") for k, v in sheets.items()}
    if len(annual):
        sheets["EU27 annual by sector"] = annual_sectors(annual, "EU27_2020")
    if len(monthly):
        sheets["EU27 monthly"] = monthly_wide(monthly, "EU27_2020")
    xlsx_notes.write_workbook(args.out, sheets, build_notes(unit_a, unit_m, annual, monthly), UNITS_SECTIONS)
    out(f"wrote {args.out}: annual rows {len(annual)}, monthly rows {len(monthly)}")


if __name__ == "__main__":
    main()
