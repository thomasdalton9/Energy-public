"""
Texas natural gas, monthly, in Bcf/d (EIA API v2, EIA_API_KEY needed - EIA is blocked from the Claude sandbox, so this
runs in GitHub Actions: .github/workflows/texas_gas.yml, 1st and 15th).

  Consumption by sector  natural-gas/cons/sum, area STX: electric power VEU, industrial VIN, residential VRS,
                         commercial VCS, vehicle fuel VDV, delivered to consumers VGT. EIA does NOT publish Texas
                         lease and plant fuel (VGL), pipeline and distribution use (VGP) or total consumption (VC0)
                         (withheld) - those columns stay blank and the balance residual carries them.
  Exports                natural-gas/move/poe2: pipeline exports to Mexico at each Texas border crossing (ENP, ports
                         named ', TX' - Eagle Pass, Roma, Rio Bravo, Hidalgo, El Paso, ...) and LNG exports from the Texas
                         terminals (ENG: Corpus Christi, Freeport, Golden Pass). Sabine Pass, Cameron/Calcasieu Pass and
                         Plaquemines are in LOUISIANA and are not Texas exports.
  Production, storage    prod/sum (marketed production VGM, extraction loss VG9, dry production FPD) and stor/sum (net
                         withdrawals SAN, working gas SAO), area STX.
  Interstate movement    EIA publishes no monthly state-to-state flows (move/state is national only), so the balance
                         residual is the implied net out-of-state flow plus the withheld fuel uses - never plugged.

Incremental: the committed workbook (sheet 'Raw MMcf') is the history store. Each run refetches, per source group,
only months after the last saved one minus a 3-month revision window (a new file backfills from 2015-01).

Usage: python3 TEXAS_GAS.py [--out "output/Data and Chart Outputs/texas_gas_monthly.xlsx"] [--full]
"""
import argparse
import os
import re
import sys
import time

import pandas as pd
import requests

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)
import xlsx_notes  # noqa: E402
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import TEXAS_GAS_FORECAST as fcst  # noqa: E402

API = "https://api.eia.gov/v2/natural-gas/"
HISTORY_START = "2015-01"
REVISION_MONTHS = 3
DEFAULT_OUT = os.path.join(ROOT, "output", "Data and Chart Outputs", "texas_gas_monthly.xlsx")
RAW_SHEET = "Raw MMcf"

CONS = {"VEU": "Electric power", "VIN": "Industrial", "VRS": "Residential", "VCS": "Commercial",
        "VDV": "Vehicle fuel", "VGL": "Lease and plant fuel", "VGP": "Pipeline and distribution use",
        "VGT": "Delivered to consumers", "VC0": "Total consumption"}
PROD = {"VGM": "Marketed production", "VG9": "Extraction loss", "FPD": "Dry production"}
STOR = {"SAN": "Net withdrawals", "SAI": "Injections", "SAW": "Withdrawals", "SAO": "Working gas"}
# Texas border crossings / LNG terminals if the series facet cannot be read (the live facet list is used first)
KNOWN_ENP = ["Y44RB", "Y44RM", "YALA", "YBROWN", "YCLI", "YDRT", "YEGP", "YELIZ", "YELP", "YGRAN", "YHDGO", "YLRD",
             "YMFE", "YPENI", "YPRES"]
KNOWN_ENG = ["YCRP", "YFPT", "YGPT"]
TX = re.compile(r"^(.*?),?\s*\b(TX|Texas)\b", re.I)
# sector order/labels of the consumption sheet
SECTORS = ["Electric power", "Industrial", "Residential", "Commercial", "Vehicle fuel", "Lease and plant fuel",
           "Pipeline and distribution use"]


def get(route, params, tries=4):
    key = os.environ.get("EIA_API_KEY")
    if not key:
        raise RuntimeError("EIA_API_KEY not set")
    last = None
    for i in range(tries):
        try:
            r = requests.get(API + route, params={"api_key": key, **params}, timeout=(10, 120))
            r.raise_for_status()
            return r.json()["response"]
        except (requests.RequestException, KeyError, ValueError) as e:
            last = e
            print(f"    {route} attempt {i + 1}/{tries}: {type(e).__name__}: {str(e)[:150]}", flush=True)
            time.sleep(5 * (i + 1))
    raise last


def rows(route, facets, start):
    out, off = [], 0
    while True:
        p = {"frequency": "monthly", "data[0]": "value", "start": start, "length": 5000, "offset": off,
             "sort[0][column]": "period", "sort[0][direction]": "asc"}
        for k, vals in facets.items():
            for i, v in enumerate(vals):
                p[f"facets[{k}][{i}]"] = v
        j = get(route, p)
        out += j["data"]
        off += 5000
        if off >= int(j["total"]):
            return out


def volume(r):
    """MMcf value of a data row, None when withheld/blank or not a volume."""
    if not re.search(r"^MMCF$", str(r.get("units", "MMCF")), re.I):
        return None
    v = pd.to_numeric(r.get("value"), errors="coerce")
    return None if pd.isna(v) else float(v)


def port_series():
    """Texas series ids in move/poe2: pipeline exports to Mexico by crossing (ENP) and LNG exports by terminal (ENG,
    all-countries total), found from the live series list (descriptions say ', TX' / 'Texas')."""
    ids = {"ENP": [], "ENG": []}
    try:
        facets = get("move/poe2/facet/series", {})["facets"]
        for f in facets:
            sid, name = str(f["id"]), str(f.get("name", ""))
            m = re.match(r"^NGM_EPG0_(ENP|ENG)_(Y\w+?)-(NMX|Z00)_MMCF$", sid)
            if not m or not TX.search(name):
                continue
            if (m[1] == "ENP" and m[3] == "NMX") or (m[1] == "ENG" and m[3] == "Z00"):
                ids[m[1]].append(sid)
        if ids["ENP"] and ids["ENG"]:
            return ids
        print("    series facet gave no Texas ports - using the known list", flush=True)
    except Exception as e:  # noqa: BLE001
        print(f"    series facet failed ({e}) - using the known list", flush=True)
    ids["ENP"] = [f"NGM_EPG0_ENP_{a}-NMX_MMCF" for a in KNOWN_ENP]
    ids["ENG"] = [f"NGM_EPG0_ENG_{a}-Z00_MMCF" for a in KNOWN_ENG]
    return ids


def load_store(path):
    try:
        s = pd.read_excel(path, sheet_name=RAW_SHEET, index_col=0)
        s.index = pd.to_datetime(s.index)
        s.index.name = "Month"
        return s.apply(pd.to_numeric, errors="coerce")
    except (FileNotFoundError, ValueError):
        return pd.DataFrame()


def start_for(store, prefix, full):
    cols = [c for c in store.columns if c.startswith(prefix)]
    if full or not cols or store[cols].dropna(how="all").empty:
        return HISTORY_START
    last = store[cols].dropna(how="all").index.max()
    return max((last - pd.DateOffset(months=REVISION_MONTHS)).strftime("%Y-%m"), HISTORY_START)


def fetch(store, full):
    """New/revised monthly MMcf values per column, merged into the saved store."""
    new = {}   # column -> {period: value}

    def put(col, period, v):
        if v is not None:
            new.setdefault(col, {})[pd.Timestamp(period + "-01")] = v

    for prefix, route, procs, names in (("cons|", "cons/sum/data/", CONS, CONS), ("prod|", "prod/sum/data/", PROD, PROD),
                                        ("stor|", "stor/sum/data/", STOR, STOR)):
        st = start_for(store, prefix, full)
        data = rows(route, {"duoarea": ["STX"], "process": list(procs)}, st)
        print(f"  {prefix} from {st}: {len(data)} rows", flush=True)
        for r in data:
            put(prefix + names[r["process"]], r["period"], volume(r))
    ids = port_series()
    for kind, prefix in (("ENP", "ENP|"), ("ENG", "ENG|")):
        st = start_for(store, prefix, full)
        sids = ids[kind]
        for i in range(0, len(sids), 12):
            data = rows("move/poe2/data/", {"series": sids[i:i + 12]}, st)
            print(f"  {prefix} from {st}: {len(data)} rows ({len(sids[i:i + 12])} series)", flush=True)
            for r in data:
                m = TX.match(str(r.get("series-description", "")))
                port = re.sub(r"\s+", " ", m[1]).strip(" ,") if m else r.get("duoarea")
                port = re.sub(r"\s*,?\s*(Liquefied|Natural|Exports|Pipeline).*$", "", port).strip(" ,")
                put(prefix + port, r["period"], volume(r))
    add = pd.DataFrame({c: pd.Series(v) for c, v in new.items()}).sort_index()
    if store.empty:
        merged = add
    else:
        merged = store.reindex(store.index.union(add.index)).reindex(columns=store.columns.union(add.columns))
        merged.update(add)
    merged.index.name = "Month"
    return merged.sort_index().dropna(how="all")


def bcfd(df):
    """MMcf per month -> Bcf/d."""
    return df.div(df.index.days_in_month, axis=0) / 1000.0


def derive(raw):
    d = bcfd(raw)
    c = lambda p: d[[x for x in d.columns if x.startswith(p)]].rename(columns=lambda x: x.split("|", 1)[1])   # noqa: E731
    cons = c("cons|")
    cons = cons.reindex(columns=SECTORS + [x for x in cons.columns if x not in SECTORS])
    cons["Reported total (delivered to consumers)"] = cons.get("Delivered to consumers")
    cons = cons.drop(columns=["Delivered to consumers"], errors="ignore")
    sec_sum = cons[["Electric power", "Industrial", "Residential", "Commercial", "Vehicle fuel"]].sum(axis=1, min_count=5)
    cons["Sum of published sectors"] = sec_sum
    enp, eng = c("ENP|"), c("ENG|")
    exports = pd.DataFrame({"Pipeline exports to Mexico (Texas crossings)": enp.sum(axis=1, min_count=1)})
    for k in eng.columns:
        exports[f"LNG - {k}"] = eng[k]
    exports["LNG exports (Texas terminals)"] = eng.sum(axis=1, min_count=1)
    exports["Total Texas exports"] = exports[["Pipeline exports to Mexico (Texas crossings)",
                                              "LNG exports (Texas terminals)"]].sum(axis=1, min_count=1)
    exports = exports.dropna(how="all")
    # balance
    p, s = c("prod|"), c("stor|")
    ratio = (p["Extraction loss"] / p["Marketed production"]).dropna().tail(12).mean()
    loss_basis = pd.Series("EIA", index=p.index)
    dry = p["Dry production"].copy()
    est = dry.isna() & p["Marketed production"].notna()
    dry[est] = p["Marketed production"][est] * (1 - ratio)
    loss_basis[est] = f"estimated (extraction loss {ratio:.1%} of marketed, mean of last 12 published months)"
    bal = pd.DataFrame({"Marketed production": p["Marketed production"],
                        "Dry production": dry,
                        "Dry production basis": loss_basis,
                        "Net storage withdrawal (+)": s["Net withdrawals"]})
    bal["Consumption (published sectors)"] = cons["Reported total (delivered to consumers)"]
    bal["Pipeline exports to Mexico"] = exports["Pipeline exports to Mexico (Texas crossings)"]
    bal["LNG exports"] = exports["LNG exports (Texas terminals)"]
    need = ["Dry production", "Net storage withdrawal (+)", "Consumption (published sectors)", "Pipeline exports to Mexico",
            "LNG exports"]
    bal = bal.dropna(subset=need)
    bal["Residual: net interstate outflow + withheld fuel uses"] = (bal["Dry production"] + bal["Net storage withdrawal (+)"]
                                                                   - bal["Consumption (published sectors)"]
                                                                   - bal["Pipeline exports to Mexico"] - bal["LNG exports"])
    return cons.dropna(how="all"), exports, enp.dropna(how="all"), bal


def latest_table(cons, exports, enp, bal):
    rows_ = []
    for label, df in (("Consumption", cons), ("Exports", exports), ("Balance", bal)):
        for col in df.columns:
            s = pd.to_numeric(df[col], errors="coerce").dropna()
            if s.empty or df[col].dtype == object:
                continue
            last = s.index.max()
            if (pd.Timestamp.today() - last).days > 400:
                continue
            row = {"Group": label, "Series": col, "Latest month": last.strftime("%b/%y")}
            for k in (2, 1, 0):
                m = last - pd.DateOffset(months=k)
                row[f"{m:%b/%y} Bcf/d"] = s.get(m)
            prev = s.get(last - pd.DateOffset(years=1))
            row["YoY %"] = (s.iloc[-1] / prev - 1) * 100 if prev and prev == prev else None
            rows_.append(row)
    return pd.DataFrame(rows_)


def notes(raw, bal):
    last = {p: raw[[c for c in raw.columns if c.startswith(p)]].dropna(how="all").index.max()
            for p in ("cons|", "prod|", "stor|", "ENP|", "ENG|")}
    fmt = lambda t: "n/a" if pd.isna(t) else f"{t:%b %Y}"   # noqa: E731
    lines = [
        "UNITS",
        "All flow sheets are in Bcf/d (billion cubic feet per day) = EIA monthly MMcf / days in the month / 1000. 'Raw MMcf' holds EIA's own monthly volumes (the history store).",
        "",
        "SOURCE",
        "EIA Natural Gas Monthly via EIA API v2 (https://api.eia.gov/v2/natural-gas/): cons/sum (consumption by sector, area STX), prod/sum, stor/sum, move/poe2 (exports by port of exit).",
        "https://www.eia.gov/dnav/ng/ng_cons_sum_dcu_STX_m.htm  |  https://www.eia.gov/dnav/ng/ng_move_poe2_a_EPG0_ENP_Mmcf_m.htm",
        "",
        "COVERAGE AND LAG",
        f"Latest month by source: consumption {fmt(last['cons|'])}, production {fmt(last['prod|'])}, storage {fmt(last['stor|'])}, "
        f"Mexico pipeline exports {fmt(last['ENP|'])}, LNG exports {fmt(last['ENG|'])}. EIA's state consumption lags about 2-3 months; the latest months are revised.",
        "Each run refetches only months after the last saved one, minus a 3-month revision window.",
        "",
        "WITHHELD SERIES",
        "EIA publishes no Texas lease and plant fuel (VGL), pipeline and distribution use (VGP) or total consumption (VC0): those columns are blank. 'Reported total' is EIA's",
        "Delivered to Consumers (VGT) = electric power + industrial + residential + commercial + vehicle fuel.",
        "EIA's Texas dry production (FPD) and extraction loss (VG9) stop after Dec 2024; marketed production (VGM) continues. Later months estimate dry production as marketed",
        "production less extraction loss at the mean share of the last 12 published months (column 'Dry production basis' says which).",
        "",
        "EXPORTS",
        "Pipeline exports to Mexico = EIA ports of exit named TX (Eagle Pass, Roma, Rio Bravo, Hidalgo, Rio Grande, Brownsville, El Paso, San Elizario, Presidio, Laredo, ...). They are the gas leaving through",
        "Texas crossings, not only Texas-produced gas. LNG exports = Texas terminals only (Corpus Christi, Freeport, Golden Pass from 2026). Sabine Pass, Cameron/Calcasieu Pass and",
        "Plaquemines are in Louisiana and are NOT counted. EIA has no Texas state-to-state movement (move/state is national only).",
        "",
        "BALANCE",
        "Residual = dry production + net storage withdrawal - consumption (published sectors) - Mexico pipeline exports - LNG exports. It is the implied net gas leaving Texas by pipeline to",
        "other states plus the withheld lease/plant/pipeline fuel and any statistical differences. It is shown, never plugged.",
    ]
    return lines, {"UNITS", "SOURCE", "COVERAGE AND LAG", "WITHHELD SERIES", "EXPORTS", "BALANCE"}


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", default=DEFAULT_OUT)
    ap.add_argument("--full", action="store_true", help="ignore the saved history and refetch from 2015-01")
    args = ap.parse_args()
    store = load_store(args.out)
    raw = fetch(store, args.full)
    if raw.empty:
        sys.exit("no Texas data returned")
    cons, exports, enp, bal = derive(raw)
    latest = latest_table(cons, exports, enp, bal)
    lines, titles = notes(raw, bal)
    try:
        fvals, fctx = fcst.build(args.out, cons, exports, bal)
        fl, ft = fcst.notes_lines(fctx)
        lines, titles = lines + fl, titles | ft
    except Exception as e:  # noqa: BLE001
        print(f"LNG forecast failed ({type(e).__name__}: {e})", flush=True)
        fvals = None
    os.makedirs(os.path.dirname(args.out), exist_ok=True)
    for n, df in (("Consumption by sector", cons), ("Exports", exports), ("Mexico by crossing", enp), ("Balance", bal)):
        df.index.name = "Month"
    xlsx_notes.write_workbook(args.out, {"Consumption by sector": cons, "Exports": exports, "Mexico by crossing": enp,
                                         "Balance": bal, "Latest and YoY": latest.set_index("Group"),
                                         **({"Forecast values": fcst.values_sheet(fvals)} if fvals is not None else {}),
                                         RAW_SHEET: raw},
                              lines, titles)
    if fvals is not None:
        fcst.write_sheets(args.out, fctx["trains"], fctx["par"], fctx["prof"], fctx["over"], fvals, fctx["fs"])
        pd.set_option("display.width", 250)
        print(fvals.loc[[m for m in fvals.index if m.month == 12 and m.year >= 2025],
                        ["lng_b", "lng_d", "cons", "dem_b", "dem_d", "mex"]].astype(float).round(2).to_string())
    pd.set_option("display.width", 250)
    pd.set_option("display.max_columns", 30)
    print(f"Saved {args.out}")
    print(latest.to_string())
    print(cons.tail(4).round(3).to_string())
    print(exports.tail(4).round(3).to_string())
    print(bal.tail(4).round(3).T.to_string())


if __name__ == "__main__":
    main()
