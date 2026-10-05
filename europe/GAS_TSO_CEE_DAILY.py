"""
National gas consumption (daily) for Austria, Czechia and Lithuania from the gas operators' own public data:

  output/Data and Chart Outputs/europe_tso_gas_demand_cee_daily.xlsx
    sheet "Daily": date (gas day), GWh per day:
        AT_total, AT_slp, AT_industry_gt300, AT_industry_lt300, AT_power_east
                       AGGM (Austrian Gas Grid Management) data monitor: end-customer consumption ("EKV"), determined from
                       allocations/metering. AT_total = "determined consumption Austria flow"; the rest are its classes
                       (standard-load-profile customers, load-profile customers above / below 300 MW... see Units), AT_power_east =
                       gas-fired power plants in market area East only (the series AGGM publishes).
        AT_net_entry, AT_net_exit, AT_storage_withdrawal, AT_storage_injection, AT_production
        AT_exit_baumgarten_mab, AT_entry_baumgarten_mab, AT_exit_baumgarten_bog, AT_entry_baumgarten_gesamt
                       AGGM's allocations at the Baumgarten points. BOG exit / GESAMT entry (BOG + GCA) are the Eustream border (Slovakia): they match Eustream's
                       and ENTSOG's Baumgarten allocations (5.8 / 0.9 TWh in 2025). MAB exit / entry (20.4 / 7.7 TWh in 2025) are a second Baumgarten boundary
                       that neither ENTSOG nor Eustream's border point carries; they make up the whole gap between AGGM's net border exit and the entries the
                       neighbours record from Austria (see EUROPE_MASTER.py, austria_slovakia_mab()).
                       AGGM's market-area balance components: net border entry and exit (all border points, net of flows through the
                       virtual points), storage withdrawal and injection (all storage in the market area) and domestic production
                       ("Production East": OMV, RAG). AT_total = net entry - net exit + storage withdrawal - injection + production.
        CZ_total       NET4GAS (CAMS public API) system balance: border entries - border exits + storage withdrawals - storage
                       injections + virtual production point, on allocated daily quantities. NET4GAS publishes no domestic-exit
                       series, so this is consumption plus losses/own use/line-pack change, labelled as a balance.
        CZ_border_entry, CZ_border_exit, CZ_storage_withdrawal, CZ_storage_injection, CZ_production   the balance components
        LT_total, LT_distribution, LT_direct
                       Amber Grid open data (public Power BI report "Lithuania consumption"): domestic consumption = gas
                       transmitted to distribution systems + to directly connected consumers.
    sheet "Units": source and definitions

Sources (all free, no key)
  AT  https://platform.aggm.at/vis-service/api/ts/values (POST, the data monitor's public API; kWh -> GWh)
  CZ  https://extranet.cams.net4gas.cz/usy-cams-balancingg01-public/.../publicApi/tsData/list (NET4GAS CAMS public API; kWh -> GWh)
  LT  https://wabi-west-europe-d-primary-api.analysis.windows.net/public/reports/querydata (Amber Grid's embedded public Power BI,
      resource key = report id; GWh)

Incremental: the committed workbook is the history store; each run re-fetches the last 14 days (the operators restate recent days)
plus any gap; history from 2021-01-01. A country that fails is logged and left as it was.

Usage: python3 GAS_TSO_CEE_DAILY.py [--out-dir DIR] [--start 2021-01-01] [--only AT,CZ,LT]
"""
import argparse
import os
import sys
import time
from datetime import date, datetime, timedelta, timezone
from zoneinfo import ZoneInfo

import pandas as pd
import requests

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
sys.path.insert(0, ROOT)
import xlsx_notes  # noqa: E402

OUT_DEFAULT = os.path.join(ROOT, "output", "Data and Chart Outputs")
FILE = "europe_tso_gas_demand_cee_daily.xlsx"
REVISION_DAYS = 14
UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36"
AT_COLS = ["AT_total", "AT_slp", "AT_industry_gt300", "AT_industry_lt300", "AT_power_east",
           "AT_net_entry", "AT_net_exit", "AT_storage_withdrawal", "AT_storage_injection", "AT_production",
           "AT_exit_baumgarten_mab", "AT_entry_baumgarten_mab", "AT_exit_baumgarten_bog", "AT_entry_baumgarten_gesamt"]
CZ_COLS = ["CZ_total", "CZ_border_entry", "CZ_border_exit", "CZ_storage_withdrawal", "CZ_storage_injection", "CZ_production"]
LT_COLS = ["LT_total", "LT_distribution", "LT_direct"]
COLUMNS = AT_COLS + CZ_COLS + LT_COLS


def post(url, tries=3, **kw):
    for i in range(tries):
        try:
            r = requests.post(url, timeout=(15, 120), **kw)
            if r.ok:
                return r
            if r.status_code in (429, 500, 502, 503, 504):
                time.sleep(4 * (i + 1))
                continue
            raise RuntimeError(f"HTTP {r.status_code} {r.text[:200]} for {url}")
        except requests.RequestException:
            time.sleep(4 * (i + 1))
    raise RuntimeError(f"gave up on {url}")


def get(url, tries=3, **kw):
    for i in range(tries):
        try:
            r = requests.get(url, timeout=(15, 120), **kw)
            if r.ok:
                return r
            time.sleep(4 * (i + 1))
        except requests.RequestException:
            time.sleep(4 * (i + 1))
    raise RuntimeError(f"gave up on {url}")


# ---- Austria: AGGM data monitor --------------------------------------------------------------------------------------
AT_URL = "https://platform.aggm.at/vis-service/api/ts/values"
AT_SERIES = {"ErmittelterEKVOesterreich": "AT_total", "SummeEKV_SLP_Oesterreich": "AT_slp", "SummeEKV_LPZGR300_Oesterreich": "AT_industry_gt300",
             "SummeEKV_LPZKL300_Oesterreich": "AT_industry_lt300", "SummeEKV_Kraftwerke_MGO": "AT_power_east",
             "NettoEntryOesterreich": "AT_net_entry", "NettoExitOesterreich": "AT_net_exit",
             "NettoEntrySpeicherOesterreich": "AT_storage_withdrawal", "NettoExitSpeicherOesterreich": "AT_storage_injection",
             "Production East": "AT_production",
             "ExitBaumgartenMAB_MGM-Allokationen": "AT_exit_baumgarten_mab", "EntryBaumgartenMAB_MGM-Allokationen": "AT_entry_baumgarten_mab",
             "ExitBaumgartenBOG_MGM-Allokationen": "AT_exit_baumgarten_bog", "EntryBaumgartenGESAMT_MGM-Allokationen": "AT_entry_baumgarten_gesamt"}


def austria(d0, d1):
    h = {"User-Agent": UA, "Content-Type": "application/json", "Accept": "application/json"}
    frames = []
    s = d0
    while s <= d1:                                       # half-year chunks keep each response modest
        e = min(s + timedelta(days=182), d1)
        body = {"rangeType": "individual", "from": f"{s:%Y-%m-%d}T06:00:00", "to": f"{e + timedelta(days=1):%Y-%m-%d}T06:00:00",
                "granularity": "day", "timeseries": list(AT_SERIES)}
        j = post(AT_URL, json=body, headers=h).json()
        for cd in j["timeSeriesData"]["chartData"]:
            col = AT_SERIES.get(cd["header"]["name"])
            if not col:
                continue
            for p in cd["dataSet"]:
                if p.get("y") is None:
                    continue
                t = pd.Timestamp(p["x"], unit="ms", tz="UTC").tz_convert("Europe/Vienna")
                frames.append((t.tz_localize(None).normalize(), col, p["y"] / 1e6))          # kWh -> GWh; x = 06:00 local gas-day start
        s = e + timedelta(days=1)
    if not frames:
        return pd.DataFrame()
    df = pd.DataFrame(frames, columns=["date", "col", "v"]).drop_duplicates(["date", "col"], keep="last")
    return df.pivot(index="date", columns="col", values="v").sort_index()


# ---- Czechia: NET4GAS CAMS public API --------------------------------------------------------------------------------
CAMS = "https://extranet.cams.net4gas.cz/usy-cams-balancingg01-public/11100000000000000000000000000010/publicApi/tsData/list"
# interconnection-point ids of the CAMS public API (allocated daily quantities, kWh/d)
CZ_GROUPS = {
    "CZ_border_entry": [(3, "entry"), (15, "entry"), (20, "entry"), (23, "entry")],              # Brandov, Waidhaus, Lanzhot, Cesky Tesin
    "CZ_border_exit": [(9, "exit"), (17, "exit"), (21, "exit"), (24, "exit")],
    "CZ_storage_withdrawal": [(26, "entry"), (41, "entry"), (44, "entry"), (76, "entry")],       # Gas Storage CZ, MND Energy Storage, MND Gas Storage, SPP Storage
    "CZ_storage_injection": [(33, "exit"), (42, "exit"), (45, "exit"), (77, "exit")],
    "CZ_production": [(66, "entry")],                                                            # virtual production point
}


def prague_utc(d):
    return datetime(d.year, d.month, d.day, 6, tzinfo=ZoneInfo("Europe/Prague")).astimezone(timezone.utc).strftime("%Y-%m-%dT%H:%M:%S.000Z")


def czechia(d0, d1):
    h = {"User-Agent": UA, "Content-Type": "application/json", "Accept": "application/json"}
    idx = pd.date_range(d0, d1, freq="D")
    n = len(idx)
    out = {}
    for grp, pts in CZ_GROUPS.items():
        tot = pd.Series(0.0, index=idx)
        seen = False
        for ipo, d in pts:
            body = {"tsList": [{"tsId": {"tsDefinitionCode": "allocation_KwhD_day_ipoD", "tsDimensionValueIdList": [str(ipo), d]},
                                "timeInterval": {"from": prague_utc(d0), "to": prague_utc(d1 + timedelta(days=1))},
                                "tsValueListSampleType": "fixed", "loadLastPreviousFilledValue": "false"}], "pageInfo": {"pageIndex": 0, "pageSize": 1000}}
            j = post(CAMS, json=body, headers=h).json()
            for ts in j.get("tsList", []):
                vals = [(v.get("value") if isinstance(v, dict) else None) for v in ts.get("tsValueList", [])]
                if len(vals) != n:
                    print(f"  CZ ipo {ipo} {d}: {len(vals)} values for {n} days", flush=True)
                s = pd.Series(vals[:n], index=idx[:len(vals)], dtype="float64") / 1e6                 # kWh -> GWh
                if s.notna().any():
                    seen = True
                tot = tot.add(s.fillna(0), fill_value=0)
                # a day with no value in any of the points stays NaN in the group below
                out.setdefault(grp + "_n", pd.Series(0, index=idx))
                out[grp + "_n"] = out[grp + "_n"].add(s.notna().astype(int), fill_value=0)
        if seen:
            out[grp] = tot
    df = pd.DataFrame({k: v for k, v in out.items() if not k.endswith("_n")})
    if df.empty:
        return df
    have = out.get("CZ_border_entry_n", pd.Series(0, index=idx)) > 0           # days where the main border series exist
    df = df[have.reindex(df.index).fillna(False)]
    for c in CZ_GROUPS:
        if c not in df:
            df[c] = 0.0
    df["CZ_total"] = (df["CZ_border_entry"] - df["CZ_border_exit"] + df["CZ_storage_withdrawal"] - df["CZ_storage_injection"] + df["CZ_production"])
    return df


# ---- Lithuania: Amber Grid public Power BI ---------------------------------------------------------------------------
PBI_URL = "https://wabi-west-europe-d-primary-api.analysis.windows.net/public/reports/querydata?synchronous=true"
PBI_RID = "305e13f3-56da-4cce-813c-92a5ac0181e1"
PBI_MODEL = 1732958
PBI_DATASET = "dbe6a55e-3ce8-42b8-ace8-7a1df7e0cf58"
LT_PAV = {"Domestic consumption": "LT_total", "Transmitted to distribution systems": "LT_distribution",
          "Transmitted to directly connected cons.": "LT_direct"}


def _col(src, prop):
    return {"Column": {"Expression": {"SourceRef": {"Source": src}}, "Property": prop}}


def pbi_rows(resp):
    ds = resp["results"][0]["result"]["data"]["dsr"]["DS"][0]
    dicts = ds.get("ValueDicts", {})
    rows, prev, sch = [], None, None
    for r in ds["PH"][0]["DM0"]:
        if "S" in r:
            sch = r["S"]
        R, N, C, vals = r.get("R", 0), r.get("Ø", 0), list(r.get("C", [])), []
        for i in range(len(sch)):
            if R >> i & 1:
                v = prev[i]
            elif N >> i & 1:
                v = None
            else:
                v = C.pop(0)
                dn = sch[i].get("DN")
                if dn is not None and isinstance(v, int) and dn in dicts:
                    v = dicts[dn][v]
            vals.append(v)
        prev = vals
        rows.append(vals)
    return rows


def lithuania(d0, d1):
    h = {"User-Agent": UA, "X-PowerBI-ResourceKey": PBI_RID, "Content-Type": "application/json;charset=UTF-8",
         "Accept": "application/json, text/plain, */*", "Origin": "https://app.powerbi.com", "Referer": "https://app.powerbi.com/"}
    frm = [{"Name": "d", "Entity": "Date", "Type": 0}, {"Name": "s", "Entity": "Σ Measures", "Type": 0},
           {"Name": "v", "Entity": "VARTOTOJU_DUJU_SUNAUDOJIMAS", "Type": 0}]
    where = [{"Condition": {"In": {"Expressions": [_col("v", "pav")], "Values": [[{"Literal": {"Value": f"'{x}'"}}] for x in LT_PAV]}}},
             {"Condition": {"Comparison": {"ComparisonKind": 2, "Left": _col("d", "Date"),
                                           "Right": {"Literal": {"Value": f"datetime'{d0:%Y-%m-%d}T00:00:00'"}}}}}]
    sel = [{**_col("d", "Date"), "Name": "Date.Date"}, {**_col("v", "pav"), "Name": "pav"},
           {"Measure": {"Expression": {"SourceRef": {"Source": "s"}}, "Property": "SuvartojimoKiekis, GWh"}, "Name": "GWh"}]
    body = {"version": "1.0.0", "queries": [{"Query": {"Commands": [{"SemanticQueryDataShapeCommand": {
        "Query": {"Version": 2, "From": frm, "Select": sel, "Where": where},
        "Binding": {"Primary": {"Groupings": [{"Projections": [0, 1, 2]}]}, "DataReduction": {"DataVolume": 4, "Primary": {"Window": {"Count": 30000}}}, "Version": 1},
        "ExecutionMetricsKind": 1}}]}, "QueryId": "", "ApplicationContext": {"DatasetId": PBI_DATASET, "Sources": [{"ReportId": "2182883"}]}}],
        "cancelQueries": [], "modelId": PBI_MODEL}
    try:
        rows = pbi_rows(post(PBI_URL, json=body, headers=h).json())
    except Exception as e:  # noqa: BLE001
        print(f"  LT date-filtered query failed ({type(e).__name__}: {str(e)[:150]}); pulling the full history", flush=True)
        body["queries"][0]["Query"]["Commands"][0]["SemanticQueryDataShapeCommand"]["Query"]["Where"] = where[:1]
        rows = pbi_rows(post(PBI_URL, json=body, headers=h).json())
    recs = [(pd.Timestamp(r[0], unit="ms").normalize(), LT_PAV[r[1]], r[2]) for r in rows if r[0] is not None and r[2] is not None and r[1] in LT_PAV]
    df = pd.DataFrame(recs, columns=["date", "col", "v"])
    df["v"] = pd.to_numeric(df["v"], errors="coerce")
    df = df.drop_duplicates(["date", "col"], keep="last").pivot(index="date", columns="col", values="v").sort_index()
    df = df[df.index >= pd.Timestamp(d0)]
    return df.iloc[:-3]                                                  # the latest days are still being filled in (direct consumers lag)


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
    ap.add_argument("--only", default="", help="comma-separated subset of AT,CZ,LT (the workbook keeps the other countries' history)")
    args = ap.parse_args()
    only = {x.strip().upper() for x in args.only.split(",") if x.strip()}
    os.makedirs(args.out_dir, exist_ok=True)
    path = os.path.join(args.out_dir, FILE)
    start = datetime.strptime(args.start, "%Y-%m-%d").date()
    today = datetime.now(timezone.utc).date()
    old = read_existing(path)
    combined = old.copy()
    for code, label, fn, cols in (("AT", "Austria (AGGM)", austria, AT_COLS), ("CZ", "Czechia (NET4GAS)", czechia, CZ_COLS),
                                  ("LT", "Lithuania (Amber Grid)", lithuania, LT_COLS)):
        if only and code not in only:
            continue
        print(f"{label}: start", flush=True)
        have = old[cols].dropna(how="all")
        fs = start if have.empty else max(start, have.index.max().date() - timedelta(days=REVISION_DAYS))
        if any(old[c].notna().sum() == 0 for c in cols):      # a column added later is back-filled from the start
            fs = start
        try:
            new = fn(fs, today - timedelta(days=1) if code != "LT" else today)
        except Exception as e:  # noqa: BLE001
            print(f"{label}: FAILED {type(e).__name__}: {e}", flush=True)
            continue
        if new.empty:
            print(f"{label}: no rows from {fs}", flush=True)
            continue
        new.index = pd.DatetimeIndex(new.index).astype("datetime64[ns]")
        new = new.astype("float64")
        combined = combined.reindex(combined.index.union(new.index))
        for c in cols:
            if c in new:
                s = new[c].dropna()
                combined.loc[s.index, c] = s
        print(f"{label}: {len(new)} days {new.index.min():%Y-%m-%d} .. {new.index.max():%Y-%m-%d} (from {fs})", flush=True)
    combined = combined.sort_index().round(3).dropna(how="all")
    combined.index.name = "date"
    if combined.empty:
        print("no CEE gas demand data")
        return
    print("days per column:", combined.notna().sum().to_dict())
    print((combined.resample("YS").sum(min_count=1) / 1000).round(2).T.to_string())
    lines = ["Central Europe - national gas consumption (Austria, Czechia, Lithuania) from the gas operators' own public series", "",
             "Source", "Austria: AGGM Austrian Gas Grid Management data monitor, end-customer consumption (https://platform.aggm.at). Czechia: NET4GAS "
             "CAMS public data (https://extranet.cams.net4gas.cz), allocated daily quantities at the interconnection points. Lithuania: Amber Grid open "
             "data, Lithuanian gas consumption report (https://ambergrid.lt/en/for-clients/open-data/650). Free, no key.",
             "", "Units and definitions",
             "GWh per gas day. AT_total = AGGM 'determined consumption Austria flow' (end-customer consumption, allocated/metered; AGGM publishes it from "
             "Oct 2022 in this form); AT_slp = standard-load-profile customers, AT_industry_gt300 / AT_industry_lt300 = load-profile-metered "
             "customers above / below 300 MW capacity classes (AGGM naming), AT_power_east = gas-fired power plants in market area East only; AT_net_entry / AT_net_exit / AT_storage_withdrawal / AT_storage_injection / AT_production = AGGM's market-area balance (border, storage and domestic production); AT_exit_baumgarten_mab / AT_entry_baumgarten_mab = allocations at AGGM's Baumgarten MAB boundary, AT_exit_baumgarten_bog / AT_entry_baumgarten_gesamt = the Baumgarten border to Eustream (Slovakia). "
             "CZ_total = NET4GAS system balance: CZ_border_entry (Brandov, Waidhaus, Lanzhot, Cesky Tesin) - CZ_border_exit + CZ_storage_withdrawal "
             "- CZ_storage_injection + CZ_production (virtual production point). NET4GAS publishes no domestic-exit series, so the balance is "
             "consumption plus own use, losses and line-pack change. LT_total = Amber Grid 'domestic consumption' = LT_distribution (gas "
             "transmitted to distribution systems) + LT_direct (to directly connected consumers); the latest 3 days published are dropped as incomplete. "
             "Recent days are preliminary and restated.",
             f"Re-fetches the last {REVISION_DAYS} days each run plus gaps; history from {args.start}.",
             "", "Last pull", f"{datetime.now(timezone.utc):%Y-%m-%d %H:%M} UTC; {len(combined)} days, "
             f"{combined.index.min():%Y-%m-%d} to {combined.index.max():%Y-%m-%d}"]
    xlsx_notes.write_workbook(path, {"Daily": combined}, lines, {"Source", "Units and definitions", "Last pull"})
    print(f"saved {FILE}")


if __name__ == "__main__":
    main()
