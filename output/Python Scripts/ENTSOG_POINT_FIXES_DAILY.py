"""
ENTSOG points that the main country-balance pull (ENTSOG_GAS_FLOWS_DAILY.py) classifies away, pulled one by one:

  output/Data and Chart Outputs/entsog_point_fixes_daily.xlsx
    sheet "Daily": date (gas day), GWh per day, columns
        GR_tap_imports        Nea Mesimvria (TAP -> DESFA): Azerbaijani gas entering Greece. ENTSOG lists TAP's operator (AL-TSO-0001) with
                              country GR, so the main pull sees it as a flow inside Greece and drops it (7-11 TWh a year, Greece's balance -15%).
        HU_production_exit    "Exit for Blending (HU)": imported gas leaving the transmission system to be blended with high-CO2 domestic gas
                              and re-entering at the "Aggregated Single Production" entry. The main pull counts that entry as production, so
                              this exit is subtracted from Hungarian production (13-14 TWh a year, Hungary's balance +14%).
        UK_moffat_exit        Moffat exit (GB NTS -> Ireland, Northern Ireland and the Isle of Man). ENTSOG gives the point's far side as country
                              UK, so the main pull drops it and Great Britain's exports omit Ireland (63 TWh in 2025).
        DE_greifswald_nel     Greifswald / NEL entry (Nord Stream 1 gas into the NEL pipeline, operator NEL Gastransport). The same gas is also
        DE_greifswald_opal    reported by GUD and Fluxys Deutschland (Greifswald / GUD, / Fluxys), so one operator is taken. Greifswald / OPAL entry
                              (OPAL Gastransport; also reported as Greifswald / LBTG). ENTSOG lists no far side for these points, so the main pull
                              finds no adjacent system and drops them: 620 TWh in 2021 and 314 TWh in 2022 of Russian gas, which left Germany's
                              (and the Germany + Netherlands) balance about 100 TWh a quarter short until Nord Stream stopped in Sept 2022.
        DE_emden_oge          Emden (EPT1) entries of OGE, GUD and GTS (Norwegian gas; Thyssengas reports the same flow as GUD and is left out).
        DE_emden_gud          Kept as a check on the Gassco-based Emden estimate in EUROPE_MASTER.emden_gap; not used by the master.
        NL_emden_gts
    sheet "Units": source and definitions

Incremental: reads the committed workbook, re-fetches the last 45 days plus any gap; history from 2021-10-04 (ENTSOG keeps ~5 years).

Usage: python3 ENTSOG_POINT_FIXES_DAILY.py [--out-dir DIR] [--start 2021-10-04]
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

API = "https://transparency.entsog.eu/api/v1"
H = {"User-Agent": "Mozilla/5.0", "Accept": "application/json"}
OUT_DEFAULT = os.path.join(ROOT, "output", "Data and Chart Outputs")
FILE = "entsog_point_fixes_daily.xlsx"
POINTS = {"GR_tap_imports": ("GR-TSO-0001", "ITP-00427", "entry"),
          "HU_production_exit": ("HU-TSO-0001", "PRD-00235", "exit"),
          "UK_moffat_exit": ("UK-TSO-0001", "ITP-00090", "exit"),
          "DE_greifswald_nel": ("DE-TSO-0017", "ITP-00247", "entry"),
          "DE_greifswald_opal": ("DE-TSO-0016", "ITP-00251", "entry"),
          "DE_emden_oge": ("DE-TSO-0009", "ITP-00080", "entry"),
          "DE_emden_gud": ("DE-TSO-0005", "ITP-00081", "entry"),
          "NL_emden_gts": ("NL-TSO-0001", "ITP-00160", "entry")}
RELOAD_DAYS = 45


def get(path, params, tries=4):
    for i in range(tries):
        try:
            r = requests.get(f"{API}/{path}", params=params, headers=H, timeout=(15, 300))
            if r.status_code == 404:
                return {}
            if r.ok:
                return r.json()
            time.sleep(10 * (i + 1))
        except requests.RequestException:
            time.sleep(10 * (i + 1))
    raise RuntimeError(f"ENTSOG {path} failed")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out-dir", default=OUT_DEFAULT)
    ap.add_argument("--start", default="2021-10-04")
    args = ap.parse_args()
    os.makedirs(args.out_dir, exist_ok=True)
    path = os.path.join(args.out_dir, FILE)
    start = date.fromisoformat(args.start)
    end = date.today() - timedelta(days=1)
    old = pd.DataFrame()
    if os.path.exists(path):
        try:
            old = pd.read_excel(path, sheet_name="Daily")
            old["date"] = pd.to_datetime(old["date"])
            old = old.dropna(subset=["date"]).set_index("date").sort_index()
        except Exception as e:  # noqa: BLE001
            print(f"could not read {FILE} ({type(e).__name__}: {e}); starting over")
    fs = start if old.empty else max(start, (old.index.max() - timedelta(days=RELOAD_DAYS)).date())
    if len(old) and not set(POINTS) <= set(old.columns):
        fs = start                      # a point added since the last run: backfill its whole history
    cols = {(o, p, d): n for n, (o, p, d) in POINTS.items()}
    print(f"{len(cols)} points; fetching {fs} -> {end}", flush=True)
    series = {}
    for (opk, pk, dr), name in sorted(cols.items()):
        rec, cur = {}, fs
        while cur <= end:                       # one call per year keeps each response small
            nxt = min(date(cur.year, 12, 31), end)
            rows = get("operationalData", {"indicator": "Physical Flow", "periodType": "day", "from": cur.isoformat(), "to": nxt.isoformat(),
                                           "pointDirection": f"{opk}{pk}{dr}", "limit": -1}).get("operationalData", [])
            for r in rows:
                try:
                    rec[pd.Timestamp(r["periodFrom"][:10])] = float(r["value"]) / 1e6      # kWh/d -> GWh/d
                except (TypeError, ValueError, KeyError):
                    pass
            cur = nxt + timedelta(days=1)
        if rec:
            s = pd.Series(rec, dtype=float)
            series[name] = series[name].add(s, fill_value=0) if name in series else s
            print(f"  {name}: {len(rec)} days, {s.sum() / 1000:.1f} TWh", flush=True)
    new = pd.DataFrame(series).sort_index()
    if new.empty:
        print("no data")
        return
    comb = new.combine_first(old) if len(old) else new
    comb.loc[new.index, new.columns] = new
    comb = comb[sorted(comb.columns)].sort_index().round(3)
    comb.index.name = "date"
    print("TWh per year:\n" + (comb.groupby(comb.index.year).sum() / 1000).round(1).T.to_string())
    lines = ["Europe - ENTSOG points dropped by the country-balance classification (physical flow)", "",
             "Source", "ENTSOG Transparency Platform, operational data: Physical Flow, daily (https://transparency.entsog.eu/).",
             "", "Units and definitions",
             "GWh per gas day. GR_tap_imports = Nea Mesimvria entry (TAP gas into Greece); HU_production_exit = Exit for Blending (HU), to be "
             "subtracted from the Aggregated Single Production entry; UK_moffat_exit = Moffat exit from the GB NTS to Ireland (ROI + Northern Ireland + Isle of Man); "
             "DE_greifswald_nel / DE_greifswald_opal = Nord Stream 1 gas entering Germany at Greifswald into the NEL and OPAL pipelines (Russian origin).",
             f"Re-fetches the last {RELOAD_DAYS} days each run; history from {args.start}.",
             "", "Last pull", f"{datetime.now(timezone.utc):%Y-%m-%d %H:%M} UTC; {len(comb)} days, {comb.index.min():%Y-%m-%d} to {comb.index.max():%Y-%m-%d}"]
    xlsx_notes.write_workbook(path, {"Daily": comb}, lines, {"Source", "Units and definitions", "Last pull"})
    print(f"saved {FILE}")


if __name__ == "__main__":
    main()
