"""
National gas consumption from the gas TSOs' / market-area operators' own published series, for the large countries where
ENTSOG's interconnection-point classification misses most of the demand (Germany reports final consumers as one aggregate,
Spain has few demand points, France's and Italy's end-user offtake is not tagged in ENTSOG's country totals):

  output/Data and Chart Outputs/europe_tso_gas_demand_daily.xlsx
    sheet "Daily": date (gas day), GWh per day:
        DE_distribution, DE_industry_power, DE_total   Trading Hub Europe (THE) aggregated consumption, SLP + RLM, H- and L-gas
        FR_industrial, FR_distribution, FR_power, FR_total   ODRE (GRTgaz / Teréga / RTE open data): industrial offtake,
                                                              public distribution (GRD/ELD) and gas-fired power plants (CCCG)
        ES_total                                     Enagás GTS national demand (the "Demand history" page's data)
    sheet "Units": source and definitions

Sources (all free, no key)
  DE  https://www.tradinghub.eu/en-gb/Publications/Transparency/Aggregated-consumption-data  (kWh -> GWh)
  FR  https://odre.opendatasoft.com  datasets conso-journa-industriel-grtgazterega, courbe-de-charge-eldgrd-regional-grtgaz-terega,
      conso-horaire-cccg-nat  (MWh -> GWh)
  ES  https://www.enagas.es/en/technical-management-system/energy-data/demand/history/  (GWh)

The Great Britain (National Gas NTS) and Ireland (Gas Networks Ireland) consumption come from their own workbooks in this repo.

Incremental: the committed workbook is the history store; each run re-fetches the last 14 days (the operators restate recent
days) plus any gap; history from 2021-01-01. A country that fails is logged and left as it was.

Usage: python3 TSO_GAS_DEMAND_DAILY.py [--out-dir DIR] [--start 2021-01-01]
"""
import argparse
import os
import sys
import time
from datetime import date, datetime, timedelta, timezone

import pandas as pd
import requests

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
sys.path.insert(0, ROOT)
import xlsx_notes  # noqa: E402

OUT_DEFAULT = os.path.join(ROOT, "output", "Data and Chart Outputs")
FILE = "europe_tso_gas_demand_daily.xlsx"
REVISION_DAYS = 14
UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36"
COLUMNS = ["DE_distribution", "DE_industry_power", "DE_total", "FR_industrial", "FR_distribution", "FR_power", "FR_total", "ES_total"]


def get(url, tries=3, **kw):
    for i in range(tries):
        try:
            r = requests.get(url, timeout=(15, 60), **kw)
            if r.ok:
                return r
            if r.status_code in (429, 500, 502, 503, 504):
                time.sleep(4 * (i + 1))
                continue
            raise RuntimeError(f"HTTP {r.status_code} {r.text[:200]} for {r.url}")
        except requests.RequestException:
            time.sleep(4 * (i + 1))
    raise RuntimeError(f"gave up on {url}")


# ---- Germany: Trading Hub Europe -------------------------------------------------------------------------------------
THE_URL = "https://datenservice-api.tradinghub.eu/api/evoq/GetAggregierteVerbrauchsdatenTabelle"
THE_SLP = ["slPsyn_H_Gas", "slPana_H_Gas", "slPsyn_L_Gas", "slPana_L_Gas"]
THE_RLM = ["rlMmT_H_Gas", "rlMmT_L_Gas", "rlMoT_H_Gas", "rlMoT_L_Gas"]


def germany(d0, d1):
    h = {"User-Agent": UA, "Accept": "application/json", "Origin": "https://www.tradinghub.eu", "Referer": "https://www.tradinghub.eu/"}
    parts = []
    s = d0
    while s <= d1:                                     # yearly chunks keep each response modest
        e = min(date(s.year, 12, 31), d1)
        r = get(THE_URL, params={"DatumStart": s.strftime("%m-%d-%Y"), "DatumEnde": e.strftime("%m-%d-%Y"), "GasXType_Id": "all"}, headers=h)
        rows = r.json()
        if rows:
            parts.append(pd.DataFrame(rows))
        s = e + timedelta(days=1)
    if not parts:
        return pd.DataFrame()
    d = pd.concat(parts, ignore_index=True)
    d["date"] = pd.to_datetime(d["gastag"]).dt.normalize()
    for c in THE_SLP + THE_RLM:
        d[c] = pd.to_numeric(d.get(c), errors="coerce")
    dup = int(d["date"].duplicated().sum())
    if dup:
        print(f"  THE: {dup} repeated gas-day rows (several publication versions per day); the last one is kept", flush=True)
    d = d.drop_duplicates("date", keep="last").set_index("date").sort_index()
    out = pd.DataFrame(index=d.index)
    out["DE_distribution"] = d[THE_SLP].sum(axis=1, min_count=1) / 1e6
    out["DE_industry_power"] = d[THE_RLM].sum(axis=1, min_count=1) / 1e6
    out["DE_total"] = out["DE_distribution"] + out["DE_industry_power"]
    return out


# ---- France: ODRE ----------------------------------------------------------------------------------------------------
ODRE = "https://odre.opendatasoft.com/api/explore/v2.1/catalog/datasets/{ds}/records"
ODRE_SETS = {"FR_industrial": ("conso-journa-industriel-grtgazterega", "consommation_journaliere_mwh_pcs"),
             "FR_distribution": ("courbe-de-charge-eldgrd-regional-grtgaz-terega", "conso_journaliere_mwh_pcs_0degc"),
             "FR_power": ("conso-horaire-cccg-nat", "conso_journaliere_mwh_pcs_0degc")}


def france(d0, d1):
    cols = {}
    for name, (ds, field) in ODRE_SETS.items():
        rec, offset = {}, 0
        while True:
            r = get(ODRE.format(ds=ds), params={
                "select": f"date,sum({field}) as v", "group_by": "date", "order_by": "date", "limit": 100, "offset": offset,
                "where": f"date>=date'{d0.isoformat()}' and date<=date'{d1.isoformat()}'"})
            res = r.json().get("results", [])
            for row in res:
                if row.get("v") is not None:
                    rec[pd.Timestamp(row["date"][:10])] = float(row["v"]) / 1000.0     # MWh -> GWh
            if len(res) < 100:
                break
            offset += 100
        cols[name] = pd.Series(rec, dtype=float)
    out = pd.DataFrame(cols).sort_index()
    out["FR_total"] = out[list(ODRE_SETS)].sum(axis=1, min_count=len(ODRE_SETS))
    return out


# ---- Spain: Enagas ---------------------------------------------------------------------------------------------------
ENAGAS = ("https://www.enagas.es/content/enagas/en/gestion-tecnica-sistema/energy-data/demanda/historico/jcr:content/responsiveGrid/"
          "container_copy_19796/realdemand_copy_copy.realdemand.json")


def spain(d0, d1):
    h = {"Accept": "application/json, text/javascript, */*; q=0.01", "User-Agent": UA, "X-Requested-With": "XMLHttpRequest",
         "Referer": "https://www.enagas.es/en/technical-management-system/energy-data/demand/history/"}
    rec = {}
    q = min(d0 + timedelta(days=55), d1)
    while True:
        j = get(ENAGAS, params={"date": q.strftime("%d/%m/%Y")}, headers=h).json()
        for e in j.get("actual", []):
            try:
                day = date.fromisoformat(e["fecha_demanda"])
                if d0 <= day <= d1:
                    rec[pd.Timestamp(day)] = float(e["demanda"])
            except (KeyError, ValueError, TypeError):
                continue
        if q >= d1:
            break
        q = min(q + timedelta(days=55), d1)      # the endpoint returns a ~60-90 day window ending at the queried date
        time.sleep(0.3)
    return pd.DataFrame({"ES_total": pd.Series(rec, dtype=float)}).sort_index()


def read_existing(path):
    if not os.path.exists(path):
        return pd.DataFrame(columns=COLUMNS)
    try:
        d = pd.read_excel(path, sheet_name="Daily")
    except Exception as e:  # noqa: BLE001
        print(f"could not read {FILE} ({type(e).__name__}: {e}); starting over")
        return pd.DataFrame(columns=COLUMNS)
    d["date"] = pd.to_datetime(d["date"], errors="coerce")
    d = d.dropna(subset=["date"]).set_index("date").sort_index()
    for c in COLUMNS:
        if c not in d:
            d[c] = float("nan")
    return d[COLUMNS]


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out-dir", default=OUT_DEFAULT)
    ap.add_argument("--start", default="2021-01-01")
    ap.add_argument("--only", default="", help="comma-separated subset of DE,FR,ES (the workbook keeps the other countries' history)")
    args = ap.parse_args()
    only = {x.strip().upper() for x in args.only.split(",") if x.strip()}
    os.makedirs(args.out_dir, exist_ok=True)
    path = os.path.join(args.out_dir, FILE)
    start = datetime.strptime(args.start, "%Y-%m-%d").date()
    today = datetime.now(timezone.utc).date()
    old = read_existing(path)
    combined = old.copy()
    for code, label, fn, cols in (("DE", "Germany (THE)", germany, [c for c in COLUMNS if c.startswith("DE_")]),
                                  ("FR", "France (ODRE)", france, [c for c in COLUMNS if c.startswith("FR_")]),
                                  ("ES", "Spain (Enagas)", spain, ["ES_total"])):
        if only and code not in only:
            continue
        print(f"{label}: start", flush=True)
        have = old[cols].dropna(how="all")
        fs = start if have.empty else max(start, have.index.max().date() - timedelta(days=REVISION_DAYS))
        try:
            new = fn(fs, today)
        except Exception as e:  # noqa: BLE001
            print(f"{label}: FAILED {type(e).__name__}: {e}", flush=True)
            continue
        if new.empty:
            print(f"{label}: no rows from {fs}", flush=True)
            continue
        combined = combined.reindex(combined.index.union(new.index))
        for c in cols:
            if c in new:
                s = new[c].dropna()
                combined.loc[s.index, c] = s
        print(f"{label}: {len(new)} days {new.index.min():%Y-%m-%d} .. {new.index.max():%Y-%m-%d} (from {fs})", flush=True)
    combined = combined.sort_index().round(2).dropna(how="all")
    combined.index.name = "date"
    if combined.empty:
        print("no TSO gas demand data")
        return
    print("days per column:", combined.notna().sum().to_dict())
    print((combined.resample("YS").sum() / 1000).round(0).T.to_string())
    lines = ["Europe - national gas consumption from the gas TSOs' / market-area operators' own series", "",
             "Source", "Germany: Trading Hub Europe aggregated consumption (https://www.tradinghub.eu). France: ODRE open data of GRTgaz, "
             "Teréga and RTE (https://odre.opendatasoft.com). Spain: Enagás GTS demand history "
             "(https://www.enagas.es/en/technical-management-system/energy-data/demand/history/). Free, no key.",
             "", "Units and definitions",
             "GWh per gas day. DE_distribution = THE SLP (standard-profile consumers on distribution networks); DE_industry_power = THE RLM "
             "(metered large consumers: industry and gas-fired power, not split further); DE_total = both. FR_industrial = direct industrial "
             "connections, FR_distribution = public distribution (GRD/ELD), FR_power = gas-fired power plants (CCCG), FR_total = the three. "
             "ES_total = Enagás national demand. Recent days are preliminary and restated.",
             f"Re-fetches the last {REVISION_DAYS} days each run plus gaps; history from {args.start}.",
             "", "Last pull", f"{datetime.now(timezone.utc):%Y-%m-%d %H:%M} UTC; {len(combined)} days, "
             f"{combined.index.min():%Y-%m-%d} to {combined.index.max():%Y-%m-%d}"]
    xlsx_notes.write_workbook(path, {"Daily": combined}, lines, {"Source", "Units and definitions", "Last pull"})
    print(f"saved {FILE}")


if __name__ == "__main__":
    main()
