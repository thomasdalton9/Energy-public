"""
Philippines wholesale electricity market (WESM) data from IEMOP, the Independent Electricity Market
Operator of the Philippines, https://www.iemop.ph/market-data/. Found via discovery_archive/asia/SEA_DISCOVERY3.py
and SEA_DISCOVERY4.py.

IEMOP's market-data pages list their files through WordPress admin-ajax (action
display_filtered_market_data_files, one post_id per report); each file is served from
https://www.iemop.ph/wp-content/uploads/downloads/data/<REPORT>/<file>. About 90 days are listed:
  RTDREG  'RTD Regional Summaries', one csv per day: per 5-minute interval and region (CLUZ Luzon, CVIS
          Visayas, CMIN Mindanao) and commodity (En = energy): MKT_REQT (market requirement = demand), LOAD_BID,
          LOAD_CURTAILED, LOSSES, GENERATION, MKT_IMPORT / MKT_EXPORT (inter-regional HVDC flows)
  DIPCER  'DIPC Energy Results - Raw', one zip per hour: per 5-minute interval and resource: LMP, SCHED_MW,
          LMP_SMP (system marginal price), LMP_LOSS, LMP_CONGESTION (PHP/MWh)
  DIPCEF  'DIPC Energy Results - Final', same layout, published about 5-6 weeks after the day (after the market
          operator's price review); used in place of DIPCER wherever it is listed

Writes output/Data and Chart Outputs/philippines_power_market.xlsx:
  Daily demand   per day and region: average and peak market requirement (MW), energy (MWh = mean x 24),
                 generation (MWh), losses; Philippines totals
  Daily prices   per day and region: time-weighted average of the 5-minute system marginal price and of the
                 generation-weighted LMP (PHP/MWh); Philippines generation-weighted average
Writes output/Data and Chart Outputs/philippines_power_generation_daily.xlsx (standard layout):
  Daily          per day: scheduled generation by fuel (MWh), from the 5-minute DIPC schedules (SCHED_MW) of every
                 generating resource, positive values only (loads, battery charging and pumping are negative)
  Demand         per day: Philippines market requirement (average and peak MW) from the RTD regional summaries
IEMOP's files carry no fuel: asia/philippines_resource_fuels.csv maps each WESM resource to a fuel (built from DOE's
'List of Existing Power Plants', July 2026, IEMOP's registered-capacity list and name rules; see
discovery_archive/asia/PH_DISCOVERY3-4.py). A resource not in it is placed by its plant code (the name before the
last '_'), then by name rules (_BAT = storage, WIND, SOL), else counted as Unmapped.

Incremental: the workbooks are the history store (IEMOP lists only ~90 days, so the 1st/15th runs keep them
whole); only days not yet saved, plus REVISION_DAYS, plus days saved from DIPCER whose DIPCEF has since appeared,
are downloaded. Each day's 24 hourly DIPC zips give both the prices and the generation mix.

    python3 asia/PHILIPPINES_IEMOP.py [--out market.xlsx] [--mix-out generation.xlsx] [--max-days N]
"""
import argparse
import base64
import io
import os
import sys
import time
import zipfile
from concurrent.futures import ThreadPoolExecutor
from datetime import date, timedelta

import pandas as pd
import requests

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)
import xlsx_notes  # noqa: E402

AJAX = "https://www.iemop.ph/wp-admin/admin-ajax.php"
BASE = "https://www.iemop.ph"
H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0 Safari/537.36"}
T = (20, 120)
POSTS = {"RTDREG": 5760, "DIPCER": 5754, "DIPCEF": 226576}
REGIONS = {"CLUZ": "Luzon", "CVIS": "Visayas", "CMIN": "Mindanao", "LUZON": "Luzon", "VISAYAS": "Visayas",
           "MINDANAO": "Mindanao"}
REVISION_DAYS = 18   # runs are 14-17 days apart: re-read everything since the last run, plus spare (provisional days get final)
OUT = os.path.join(ROOT, "output", "Data and Chart Outputs", "philippines_power_market.xlsx")
MIX_OUT = os.path.join(ROOT, "output", "Data and Chart Outputs", "philippines_power_generation_daily.xlsx")
FUELS_CSV = os.path.join(ROOT, "asia", "philippines_resource_fuels.csv")
FUELS = ["Coal", "Gas", "Oil", "Hydro", "Geothermal", "Solar", "Wind", "Bioenergy", "Storage", "Unmapped"]


def out(*a):
    print(*a, flush=True)


def listing(report):
    """-> {date: [file urls]} for every file IEMOP currently lists for the report."""
    files, page = {}, 1
    while True:
        r = requests.post(AJAX, data={"action": "display_filtered_market_data_files", "sort": "", "datefilter": "",
                                      "page": page, "post_id": POSTS[report]}, headers=H, timeout=T)
        r.raise_for_status()
        j = r.json()
        src = j.get("source") or []
        for s in src:
            path = base64.b64decode(s).decode()
            name = path.rsplit("/", 1)[-1]
            stamp = name.split("_", 1)[1].split(".")[0]
            # DIPCER_202610030000 is the hour ENDING 00:00 on the 3rd -> belongs to the 2nd
            d = pd.Timestamp(stamp[:8]).date()
            if report.startswith("DIPC") and stamp[8:12] == "0000":
                d -= timedelta(days=1)
            files.setdefault(d, []).append(BASE + path.split("/html", 1)[-1])
        total = int(j.get("count") or 0)
        if not src or sum(len(v) for v in files.values()) >= total or page > 200:
            return files
        page += 1


def get(url):
    for i in range(4):
        try:
            r = requests.get(url, headers=H, timeout=T)
            r.raise_for_status()
            return r.content
        except requests.RequestException as e:
            if i == 3:
                out(f"  {url}: {e}")
                return None
            time.sleep(5 * (i + 1))


def demand_day(content):
    df = pd.read_csv(io.BytesIO(content))
    df.columns = [c.strip() for c in df.columns]
    df = df[df["COMMODITY_TYPE"].astype(str).str.strip().eq("En")]
    df["region"] = df["REGION_NAME"].map(REGIONS)
    for c in ("MKT_REQT", "GENERATION", "LOSSES"):
        df[c] = pd.to_numeric(df[c], errors="coerce")
    df["t"] = pd.to_datetime(df["TIME_INTERVAL"], format="%m/%d/%Y %I:%M:%S %p", errors="coerce")
    g = df.groupby("region")
    row = {}
    for reg, x in g:
        row[f"{reg}_demand_avg_MW"] = x["MKT_REQT"].mean()
        row[f"{reg}_demand_peak_MW"] = x["MKT_REQT"].max()
        row[f"{reg}_demand_MWh"] = x["MKT_REQT"].mean() * 24
        row[f"{reg}_generation_MWh"] = x["GENERATION"].mean() * 24
        row[f"{reg}_losses_MWh"] = x["LOSSES"].mean() * 24
    tot = df.groupby("t")["MKT_REQT"].sum()
    row["Philippines_demand_avg_MW"] = tot.mean()
    row["Philippines_demand_peak_MW"] = tot.max()
    row["Philippines_demand_MWh"] = tot.mean() * 24
    row["Intervals"] = df["t"].nunique()
    return row


class FuelMap:
    """RESOURCE_NAME -> fuel from asia/philippines_resource_fuels.csv, with plant-code and name-rule fallbacks."""

    def __init__(self, path=FUELS_CSV):
        m = pd.read_csv(path, dtype=str)
        self.res = dict(zip(m["RESOURCE_NAME"].str.strip(), m["FUEL"].str.strip()))
        self.plant = {}
        for r, f in self.res.items():
            if f != "Storage":   # a plant's battery does not make its new units storage
                self.plant.setdefault(r.rsplit("_", 1)[0], f)
        self.cache = {}

    def __call__(self, name):
        name = str(name).strip()
        if name not in self.cache:
            code, suf = name.rsplit("_", 1) if "_" in name else (name, "")
            if name in self.res and self.res[name] in FUELS:
                f = self.res[name]
            elif suf.upper().startswith("BAT") or "BESS" in name.upper():
                f = "Storage"
            elif code in self.plant:
                f = self.plant[code]
            elif "WIND" in code.upper():
                f = "Wind"
            elif "SOL" in code.upper():
                f = "Solar"
            else:
                f = "Unmapped"
            self.cache[name] = f
        return self.cache[name]


def dipc_hour(content, fuel_of):
    """One hourly DIPC zip -> (per interval and region: SMP, generation-weighted LMP parts;
    per interval: positive scheduled MW by fuel; per resource: positive scheduled MWh)."""
    z = zipfile.ZipFile(io.BytesIO(content))
    df = pd.read_csv(z.open(z.namelist()[0]))
    df.columns = [c.strip() for c in df.columns]
    df["region"] = df["REGION_NAME"].map(REGIONS)
    for c in ("LMP", "SCHED_MW", "LMP_SMP"):
        if c in df:
            df[c] = pd.to_numeric(df[c], errors="coerce")
    df["t"] = pd.to_datetime(df["TIME_INTERVAL"], format="mixed", errors="coerce")
    gen = df[df["SCHED_MW"] > 0].copy()
    gen["fuel"] = gen["RESOURCE_NAME"].map(fuel_of)
    mix = gen.pivot_table(index="t", columns="fuel", values="SCHED_MW", aggfunc="sum")
    res = gen.groupby("RESOURCE_NAME")["SCHED_MW"].sum() * 5 / 60
    out_rows = []
    if "LMP" in df and "LMP_SMP" in df:
        gen["LMP_w"] = gen["LMP"] * gen["SCHED_MW"]
        smp = df.groupby(["t", "region"])["LMP_SMP"].median()
        g = gen.groupby(["t", "region"])[["LMP_w", "SCHED_MW"]].sum()
        x = pd.concat([smp, g], axis=1).reset_index()
        out_rows = [{"t": r.t, "region": r.region, "SMP": r.LMP_SMP, "LMP_w": r.LMP_w if r.LMP_w == r.LMP_w else 0.0,
                     "MW": r.SCHED_MW if r.SCHED_MW == r.SCHED_MW else 0.0} for r in x.itertuples()]
    return pd.DataFrame(out_rows), mix, res


def mix_row(mixes):
    """The day's interval-by-fuel frames -> MWh per fuel (mean MW over the intervals x 24)."""
    m = pd.concat(mixes)
    m = m[~m.index.duplicated(keep="last")].reindex(columns=FUELS).fillna(0.0)
    mwh = m.mean() * 24
    row = {f"{f}_MWh": mwh[f] for f in ("Coal", "Gas", "Oil", "Hydro", "Solar", "Wind", "Bioenergy")}
    row["Other_MWh"] = mwh["Geothermal"] + mwh["Storage"] + mwh["Unmapped"]
    row["Total_MWh"] = mwh.sum()
    row["Geothermal_MWh"] = mwh["Geothermal"]
    row["Storage_MWh"] = mwh["Storage"]
    row["Unmapped_MWh"] = mwh["Unmapped"]
    row["Intervals"] = len(m)
    return row


def read_sheet(path, sheet):
    try:
        df = pd.read_excel(path, sheet_name=sheet, index_col=0)
    except (FileNotFoundError, ValueError):
        return pd.DataFrame()
    df.index = pd.to_datetime(df.index)
    return df.sort_index()


def merge(old, new):
    if old.empty or new.empty:
        return (new if old.empty else old).sort_index()
    return pd.concat([old[~old.index.isin(new.index)], new]).sort_index()


def write(out_path, dem, pri):
    if dem.empty and pri.empty:
        raise SystemExit("No IEMOP data")
    dem.index.name = pri.index.name = "date"
    notes = [
        "UNITS",
        "Daily demand: per region (Luzon, Visayas, Mindanao) from the 5-minute real-time dispatch (RTD) regional "
        "summaries, energy commodity: demand_avg / demand_peak = average and maximum market requirement (MW); "
        "demand_MWh, generation_MWh, losses_MWh = mean MW x 24. Philippines_* = the three regions summed per interval. "
        "Intervals = 5-minute intervals in the day (287 = complete: the files carry 287 per day; under 280 is re-read).",
        "Daily prices: PHP/MWh (Philippine pesos). <Region>_SMP = time-weighted average of the 5-minute system "
        "marginal price; <Region>_LMP_genweighted = locational marginal price weighted by each resource's scheduled "
        "generation; Philippines_LMP_genweighted = the same across all regions.",
        "",
        "COVERAGE",
        (f"Demand from {dem.index.min():%Y-%m-%d} to {dem.index.max():%Y-%m-%d}; " if not dem.empty else "") +
        (f"prices from {pri.index.min():%Y-%m-%d} to {pri.index.max():%Y-%m-%d}. " if not pri.empty else "") +
        "IEMOP lists about 90 days of files; older days are kept from previous runs. Grid-connected WESM regions "
        "(Mindanao joined WESM in 2023); off-grid islands are not included.",
        "",
        "SOURCE",
        "IEMOP (Independent Electricity Market Operator of the Philippines), market data: RTD Regional Summaries and "
        "DIPC Energy Results - Raw, https://www.iemop.ph/market-data/",
    ]
    xlsx_notes.write_workbook(out_path, {"Daily demand": dem, "Daily prices": pri}, notes,
                              {"UNITS", "COVERAGE", "SOURCE"})
    out(f"Saved {out_path}: demand {len(dem)} days, prices {len(pri)} days")
    out(dem.tail(2).T.to_string())
    out(pri.tail(2).T.to_string())


def write_mix(out_path, mix, dem):
    if mix.empty:
        return
    mix.index.name = "date"
    d = pd.DataFrame(index=dem.index)
    if not dem.empty:
        d["Demand_avg_MW"] = dem.get("Philippines_demand_avg_MW")
        d["Demand_peak_MW"] = dem.get("Philippines_demand_peak_MW")
        d["Demand_MWh"] = dem.get("Philippines_demand_MWh")
    d.index.name = "date"
    final = mix.index[mix.get("Source", pd.Series(dtype=str)).eq("DIPCEF")]
    notes = [
        "UNITS",
        "Daily: MWh per day, scheduled generation (SCHED_MW) of every WESM generating resource in the 5-minute "
        "dispatch-interval (DIPC) energy results, positive values only, summed per fuel per interval; MWh = mean MW "
        "over the day's intervals x 24. Total_MWh = all fuels. Other_MWh = Geothermal_MWh + Storage_MWh (battery "
        "discharge) + Unmapped_MWh (resources with no fuel in the map); those three are shown separately as detail. "
        "Hydro includes the Kalayaan pumped-storage plant's generation (its pumping is excluded with all negative "
        "schedules). Intervals = 5-minute intervals read for the day (288 = complete). Source = DIPCEF (final, "
        "after IEMOP's price review, about 5-6 weeks after the day) or DIPCER (raw, replaced by the final file "
        "once IEMOP lists it).",
        "Demand: Philippines market requirement (Luzon + Visayas + Mindanao), average and peak MW, and MWh = mean x "
        "24, from the RTD regional summaries (same data as philippines_power_market.xlsx).",
        "",
        "FUEL MAP",
        "IEMOP's files carry no fuel. Each resource's fuel comes from asia/philippines_resource_fuels.csv, built from "
        "DOE's List of Existing Power Plants (grid-connected, as of 31 July 2026), IEMOP's registered capacity list "
        "(trading participant per resource) and name rules (SOL = solar, WIND = wind, _BAT = battery). Limay CCGT "
        "(PanAsia) runs on diesel and is counted as Oil, as in DOE's list.",
        "",
        "COVERAGE",
        f"From {mix.index.min():%Y-%m-%d} to {mix.index.max():%Y-%m-%d}; final (DIPCEF) data for "
        f"{len(final)} days, raw (DIPCER) for the rest. Grid-connected WESM regions only (Luzon, Visayas, Mindanao); "
        "off-grid islands (Mindoro, Palawan, etc.) and embedded generation that is not scheduled in WESM (rooftop "
        "solar, small plants inside distribution networks) are not included. Scheduled, not metered, generation: "
        "the dispatch target for each 5-minute interval. IEMOP lists about 90 days of files; older days are kept "
        "from previous runs.",
        "",
        "SOURCE",
        "IEMOP (Independent Electricity Market Operator of the Philippines), market data: DIPC Energy Results "
        "(final and raw) and RTD Regional Summaries, https://www.iemop.ph/market-data/; fuel map from DOE List of "
        "Existing Power Plants, https://doe.gov.ph/data-and-prices/energy-statistics/electric-power-industry",
    ]
    xlsx_notes.write_workbook(out_path, {"Daily": mix, "Demand": d.dropna(how="all")}, notes,
                              {"UNITS", "FUEL MAP", "COVERAGE", "SOURCE"})
    out(f"Saved {out_path}: {len(mix)} days")
    out("GWh:\n" + (mix.tail(3)[[c for c in mix.columns if c.endswith("_MWh")]] / 1000).round(1).T.to_string())


def safe_listing(report):
    """listing(), or nothing when IEMOP's file list fails: the other report still runs and the saved days are kept."""
    try:
        return listing(report)
    except Exception as e:  # noqa: BLE001
        out(f"{report} listing failed: {type(e).__name__}: {e}")
        return {}


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", default=OUT)
    ap.add_argument("--mix-out", default=MIX_OUT)
    ap.add_argument("--max-days", type=int, default=0, help="testing: fetch at most the latest N DIPC days")
    args = ap.parse_args()
    old_d, old_p = read_sheet(args.out, "Daily demand"), read_sheet(args.out, "Daily prices")
    old_m = read_sheet(args.mix_out, "Daily")
    yesterday = date.today() - timedelta(days=1)
    revise = {yesterday - timedelta(days=k) for k in range(REVISION_DAYS)}

    rtd = safe_listing("RTDREG")
    # a saved day well short of a full set of five-minute intervals (287 a day in the files) was read part-published:
    # read it again while listed
    short = set(old_d.index[old_d["Intervals"] < 280].date) if "Intervals" in old_d else set()
    todo = sorted(d for d in rtd if d <= yesterday and (d in revise or d in short or old_d.empty
                                                        or pd.Timestamp(d) not in old_d.index))
    if args.max_days:
        todo = todo[-args.max_days:]
    out(f"RTDREG: {len(rtd)} days listed, fetching {len(todo)}")
    rows = {}
    for d in todo:
        c = get(rtd[d][0])
        if c:
            try:
                rows[pd.Timestamp(d)] = demand_day(c)
            except Exception as e:  # noqa: BLE001  (one malformed file skips that day, not the run)
                out(f"  RTD {d}: {type(e).__name__}: {e}")
    dem = merge(old_d, pd.DataFrame.from_dict(rows, orient="index").round(1))

    # DIPC: the final file where IEMOP lists it (about 5-6 weeks after the day), else the raw one
    raw, fin = safe_listing("DIPCER"), safe_listing("DIPCEF")

    def source(d):
        if len(fin.get(d, [])) >= 24:
            return "DIPCEF"
        return "DIPCER" if len(raw.get(d, [])) >= 24 else None

    saved_src = old_m["Source"] if "Source" in old_m else pd.Series(dtype=str)
    saved_n = old_m["Intervals"] if "Intervals" in old_m else pd.Series(dtype=float)
    todo = []
    for d in sorted(set(raw) | set(fin)):
        src, ts = source(d), pd.Timestamp(d)
        if d > yesterday or src is None:
            continue
        need_price = d in revise or old_p.empty or ts not in old_p.index
        need_mix = (ts not in old_m.index or (src == "DIPCEF" and saved_src.get(ts) != "DIPCEF")
                    or (src == "DIPCER" and d in revise) or saved_n.get(ts, 0) < 280)
        if need_price or need_mix:
            todo.append(d)
    if args.max_days:
        todo = todo[-args.max_days:]
    out(f"DIPC: {len(raw)} raw / {len(fin)} final days listed, fetching {len(todo)} (24 hourly files each)")
    fuel_of = FuelMap()
    prow, mrow, energy = {}, {}, []
    for n, d in enumerate(todo):
        src = source(d)
        prices, mixes = [], []
        with ThreadPoolExecutor(8) as ex:   # the day's 24 hourly zips at once
            for c in ex.map(get, (fin if src == "DIPCEF" else raw)[d]):
                if not c:
                    continue
                try:
                    p, m, r = dipc_hour(c, fuel_of)
                    prices.append(p)
                    mixes.append(m)
                    energy.append(r)
                except Exception as e:  # noqa: BLE001  (one malformed zip skips that hour)
                    out(f"  {src} {d}: {type(e).__name__}: {e}")
        ts = pd.Timestamp(d)
        if mixes:
            mrow[ts] = dict(mix_row(mixes), Source=src)
        x = pd.concat(prices) if prices else pd.DataFrame()
        if not x.empty:
            row = {}
            for reg, y in x.groupby("region"):
                row[f"{reg}_SMP_PHP_per_MWh"] = y["SMP"].mean()
                row[f"{reg}_LMP_genweighted_PHP_per_MWh"] = y["LMP_w"].sum() / y["MW"].sum() if y["MW"].sum() else None
            row["Philippines_LMP_genweighted_PHP_per_MWh"] = x["LMP_w"].sum() / x["MW"].sum() if x["MW"].sum() else None
            row["Intervals"] = x["t"].nunique()
            prow[ts] = row
        if n % 10 == 0 or n == len(todo) - 1:
            out(f"  {d} ({src})")
            # checkpoint: a timeout keeps what is done
            write(args.out, dem, merge(old_p, pd.DataFrame.from_dict(prow, orient="index").round(2)))
            write_mix(args.mix_out, merge(old_m, pd.DataFrame.from_dict(mrow, orient="index").round(1)), dem)
    if not todo:
        write(args.out, dem, old_p)
        write_mix(args.mix_out, old_m, dem)
    if energy:   # what the fuel map leaves out, for keeping asia/philippines_resource_fuels.csv up to date
        e = pd.concat(energy).groupby(level=0).sum()
        un = e[[fuel_of(r) == "Unmapped" for r in e.index]].sort_values(ascending=False)
        out(f"Unmapped resources: {len(un)}, {100 * un.sum() / e.sum():.2f}% of scheduled energy fetched: "
            + ", ".join(f"{r} {v:,.0f} MWh" for r, v in un.head(15).items()))


if __name__ == "__main__":
    main()
