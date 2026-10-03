"""
Daily power generation by fuel for Europe from the ENTSO-E Transparency Platform
(Actual Generation per Production Type [16.1.B&C], API document A75), one workbook per country:

  output/Data and Chart Outputs/<country>_power_generation_daily.xlsx
    sheet "Daily": date, Hydro_MWh, PumpedStorage_MWh, Gas_MWh, Coal_MWh, Oil_MWh, Nuclear_MWh, Wind_MWh,
                   Solar_MWh, Bioenergy_MWh, Other_MWh, Storage_MWh, Total_MWh,
                   PumpedStorageConsumption_MWh, StorageCharging_MWh, Load_MWh
    sheet "Units": source, definitions, zones, coverage rules

Same layout as the other raw country generation workbooks, so add_charts.py charts it with no registry entry
(monthly GWh stacked by fuel; pumped storage and batteries are kept in the sheet but not in the chart).

ENTSO-E is the TSOs' own statutory reporting platform. In a settled-week check (14-20 Sep 2026) it matched
SMARD (Germany) and RTE (France) to within 0.5% and Energy-Charts for 10 zones, so it is used as the primary
source for Europe, labelled as such on the master's Sources tab. National feeds replace it where they are better.

Definitions
  - A day is a UTC day. Energy = sum of (MW x slot hours) over the 15/30/60-minute slots ENTSO-E reports.
  - Fuel columns are NET: generation (inBiddingZone series) minus the same fuel's own consumption
    (outBiddingZone series, e.g. auxiliary load of idle plants). Pumped storage generation and battery
    discharge are separate columns; pumping and charging are kept as their own consumption columns.
  - Load_MWh is ENTSO-E's actual total load (A65) for the same zones and days (blank where a zone publishes none).
  - Countries with several bidding zones (Italy, Denmark, Sweden, Norway) are the sum of their zones; a day
    is kept only if every zone has it. Luxembourg is inside the DE-LU zone (Germany row).
  - A zone-day counts as complete when some generation series covers at least 21 of its 24 hours; partial days
    (late publication, today) are left out and picked up on a later run.

Incremental: reads the committed workbook and fetches only from (last saved day - 10 days) so recent revisions
are picked up, plus any gap in the last 120 days; older gaps are treated as permanent (the platform has no data
for them). A first run backfills from 2021-01-01 in 60-day windows, stopping gracefully at --max-minutes (the next
run continues, because each finished country is saved as it completes).

Usage: python3 ENTSOE_GENERATION_DAILY.py [--out-dir DIR] [--countries DE,FR] [--start 2021-01-01]
                                           [--max-minutes 320]
Requires ENTSOE_API_KEY (environment / GitHub secret, or api_keys.py).
"""
import argparse
import os
import sys
import time
from datetime import date, datetime, timedelta, timezone

import pandas as pd

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
sys.path.insert(0, HERE)
sys.path.insert(0, ROOT)
import entsoe_common as C  # noqa: E402
import xlsx_notes  # noqa: E402

OUT_DEFAULT = os.path.join(ROOT, "output", "Data and Chart Outputs")
START_DEFAULT = "2021-01-01"
REVISION_DAYS = 10
GAP_DAYS = 120
WINDOW_DAYS = 60
MIN_HOURS = 21.0
GEN_COLS = [f"{f}_MWh" for f in C.GEN_FUELS]
CONS_COLS = ["PumpedStorageConsumption_MWh", "StorageCharging_MWh"]
ALL_COLS = GEN_COLS + ["Total_MWh"] + CONS_COLS + ["Load_MWh"]


def out_path(out_dir, slug):
    return os.path.join(out_dir, f"{slug}_power_generation_daily.xlsx")


def read_existing(path):
    if not os.path.exists(path):
        return pd.DataFrame(columns=ALL_COLS)
    try:
        d = pd.read_excel(path, sheet_name="Daily")
    except Exception as e:  # noqa: BLE001
        print(f"  could not read {os.path.basename(path)} ({type(e).__name__}: {e}); starting over")
        return pd.DataFrame(columns=ALL_COLS)
    d["date"] = pd.to_datetime(d["date"], errors="coerce")
    d = d.dropna(subset=["date"]).set_index("date").sort_index()
    for c in ALL_COLS:
        if c not in d:
            d[c] = float("nan")
    return d[ALL_COLS]


def zone_days(eic, d0, d1, deadline, stats):
    """{date: {column: MWh}} for complete UTC days of one bidding zone in [d0, d1)."""
    rows, complete, load = {}, set(), {}
    for w0, w1 in C.windows(d0, d1, WINDOW_DAYS):
        if time.time() > deadline:
            raise TimeoutError
        res = C.fetch_split({"documentType": "A75", "processType": "A16", "in_Domain": eic}, w0, w1,
                            C.parse_generation, C.merge_generation)
        stats["windows"] += 1
        time.sleep(0.25)
        lres = C.fetch_split({"documentType": "A65", "processType": "A16", "outBiddingZone_Domain": eic}, w0, w1,
                             C.parse_energy, C.merge_energy)
        time.sleep(0.25)
        for d, mwh in ((lres or {}).get("e") or {}).items():
            if lres["h"].get(d, 0.0) >= MIN_HOURS:
                load[d] = mwh
        if res is None:
            stats["empty"] += 1
            continue
        e, h = res["e"], res["h"]
        tot = {"in": 0.0, "out": 0.0, "": 0.0}
        for (psr, dr), days in e.items():
            if C.GEN_COLUMN.get(psr) not in (None, "PumpedStorage", "Storage"):
                tot[dr] += sum(days.values())
        gen_dir = "out" if tot["out"] > tot["in"] * 1.5 and tot["out"] > 0 else "in"
        if gen_dir == "out":
            stats["flipped"] += 1
        for (psr, dr), days in e.items():
            col = C.GEN_COLUMN.get(psr)
            if col is None:
                continue
            is_gen = dr in (gen_dir, "")
            for d, mwh in days.items():
                r = rows.setdefault(d, {})
                if col in ("PumpedStorage", "Storage"):
                    key = col if is_gen else ("PumpedStorageConsumption" if col == "PumpedStorage" else "StorageCharging")
                    r[key] = r.get(key, 0.0) + mwh
                else:
                    r[col] = r.get(col, 0.0) + (mwh if is_gen else -mwh)
        # a day is complete when its best-covered generation series reaches MIN_HOURS
        best = {}
        for (psr, dr), days in h.items():
            if dr in (gen_dir, "") and C.GEN_COLUMN.get(psr) not in (None, "PumpedStorage", "Storage"):
                for d, hrs in days.items():
                    best[d] = max(best.get(d, 0.0), hrs)
        complete |= {d for d, hrs in best.items() if hrs >= MIN_HOURS}
    out = {}
    for d, r in rows.items():
        if d in complete and d0.date() <= d < d1.date():
            if d in load:
                r["Load"] = load[d]
            out[d] = r
    return out


def country_frame(zones, d0, d1, deadline, stats):
    per_zone = {}
    for label, eic in zones:
        zs = {"windows": 0, "empty": 0, "flipped": 0}
        per_zone[label] = zone_days(eic, d0, d1, deadline, zs)
        stats.setdefault("zones", {})[label] = {**zs, "days": len(per_zone[label])}
        stats["windows"] += zs["windows"]
        stats["empty"] += zs["empty"]
    common = set.intersection(*(set(z) for z in per_zone.values())) if per_zone else set()
    stats["days_dropped_missing_zone"] = len(set().union(*(set(z) for z in per_zone.values())) - common) if len(per_zone) > 1 else 0
    data = {}
    for d in sorted(common):
        row = {}
        for z in per_zone.values():
            for k, v in z[d].items():
                row[k] = row.get(k, 0.0) + v
        if not all("Load" in z[d] for z in per_zone.values()):
            row.pop("Load", None)
        data[pd.Timestamp(d)] = row
    df = pd.DataFrame.from_dict(data, orient="index")
    if df.empty:
        return pd.DataFrame(columns=ALL_COLS)
    for f in C.GEN_FUELS:
        if f not in df:
            df[f] = 0.0
    for c in ("PumpedStorageConsumption", "StorageCharging"):
        if c not in df:
            df[c] = 0.0
    load = df["Load"] if "Load" in df else pd.Series(float("nan"), index=df.index)
    df = df.drop(columns=["Load"], errors="ignore").fillna(0.0)
    df["Total"] = df[C.GEN_FUELS].sum(axis=1)
    df["Load"] = load
    df = df[df["Total"] > 0]
    df = df.rename(columns=lambda c: f"{c}_MWh")
    df.index.name = "date"
    return df[ALL_COLS].round(1).sort_index()


def notes(name, zones, df, start):
    lines = [f"{name} - daily power generation by fuel (ENTSO-E Transparency Platform)", "",
             "Source", "ENTSO-E Transparency Platform, Actual Generation per Production Type [16.1.B&C] "
             "(API document A75, process A16 realised). The TSOs' own statutory reporting. https://transparency.entsoe.eu/",
             "", "Zones", ", ".join(f"{label} ({eic})" for label, eic in zones),
             "", "Units and definitions",
             "MWh per UTC day: sum of MW x slot hours over the 15/30/60-minute slots ENTSO-E reports for the day.",
             "Fuel columns are net: generation minus the same fuel's own consumption (auxiliary load of idle plants).",
             "Hydro = run-of-river + reservoir (excludes pumped storage). Gas includes coal-derived gas. Coal = hard coal, "
             "lignite, oil shale, peat. Bioenergy = biomass. Other = geothermal, marine, other renewable, waste, other.",
             "PumpedStorage_MWh and Storage_MWh (batteries) are generation/discharge; pumping and charging are in "
             "PumpedStorageConsumption_MWh and StorageCharging_MWh. Total_MWh = all generation columns.",
             "", "Coverage rules",
             f"A zone-day is kept when some generation series covers >= {MIN_HOURS:.0f} of 24 hours; multi-zone countries "
             "need every zone. Partial days are left out and fetched again on a later run.",
             f"Each run re-fetches the last {REVISION_DAYS} days (ENTSO-E values are revised) and any gap within "
             f"{GAP_DAYS} days; older gaps are treated as permanent.",
             f"History is pulled from {start}.",
             "", "Last pull", f"{datetime.now(timezone.utc):%Y-%m-%d %H:%M} UTC; "
             f"{len(df)} days, {df.index.min():%Y-%m-%d} to {df.index.max():%Y-%m-%d}" if len(df) else "no data"]
    titles = {"Source", "Zones", "Units and definitions", "Coverage rules", "Last pull"}
    return lines, titles


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out-dir", default=OUT_DEFAULT)
    ap.add_argument("--countries", default="", help="comma-separated ISO codes (default: all)")
    ap.add_argument("--start", default=START_DEFAULT)
    ap.add_argument("--max-minutes", type=float, default=320.0)
    args = ap.parse_args()
    C.api_key()
    os.makedirs(args.out_dir, exist_ok=True)
    deadline = time.time() + args.max_minutes * 60
    start = datetime.strptime(args.start, "%Y-%m-%d").replace(tzinfo=timezone.utc)
    end = C.today_utc()   # exclusive: through yesterday (partial days are dropped by the coverage rule)
    want = [c.strip().upper() for c in args.countries.split(",") if c.strip()] or list(C.COUNTRIES)

    summary, failed = [], []
    for code in want:
        if code not in C.COUNTRIES:
            print(f"unknown country code {code}")
            continue
        name, slug, zones = C.COUNTRIES[code]
        path = out_path(args.out_dir, slug)
        old = read_existing(path)
        fs = start
        if len(old):
            last = old.index.max()
            fs = max(start, last.to_pydatetime().replace(tzinfo=timezone.utc) - timedelta(days=REVISION_DAYS))
            if old["Load_MWh"].notna().sum() == 0:
                fs = start   # workbook predates the Load column: backfill it
            have = set(old.index.date)
            cutoff = (end - timedelta(days=GAP_DAYS)).date()
            gaps = [d for d in (start + timedelta(days=i) for i in range((end - start).days))
                    if d.date() >= cutoff and d.date() not in have and d < fs]
            if gaps:
                fs = min(fs, gaps[0])
        have_txt = f"have {len(old)} days, last {old.index.max():%Y-%m-%d}" if len(old) else "new"
        print(f"{code} {name}: {len(zones)} zone(s), fetching {fs:%Y-%m-%d} -> {end:%Y-%m-%d} ({have_txt})", flush=True)
        stats = {"windows": 0, "empty": 0}
        t0 = time.time()
        try:
            new = country_frame(zones, fs, end, deadline, stats)
        except TimeoutError:
            print(f"  time budget reached during {name}; stopping (finished countries are saved)")
            summary.append((code, name, "stopped (time budget)", len(old)))
            break
        except Exception as e:  # noqa: BLE001
            print(f"  FAILED {name}: {type(e).__name__}: {e}")
            failed.append((code, f"{type(e).__name__}: {e}"))
            summary.append((code, name, "failed", len(old)))
            continue
        combined = pd.concat([old[~old.index.isin(new.index)], new]).sort_index()
        combined = combined[ALL_COLS]
        combined.index.name = "date"
        zinfo = "; ".join(f"{z}: {v['days']}d/{v['empty']} empty of {v['windows']}" for z, v in stats.get("zones", {}).items())
        print(f"  {len(new)} new/updated days in {time.time() - t0:.0f}s ({stats['windows']} requests, "
              f"{stats['empty']} empty); {zinfo}"
              + (f"; dropped {stats['days_dropped_missing_zone']} days missing a zone" if stats.get("days_dropped_missing_zone") else ""),
              flush=True)
        if combined.empty:
            summary.append((code, name, "NO DATA", 0))
            continue
        lines, titles = notes(name, zones, combined, args.start)
        xlsx_notes.write_workbook(path, {"Daily": combined}, lines, titles)
        summary.append((code, name, f"{combined.index.min():%Y-%m-%d}..{combined.index.max():%Y-%m-%d}", len(combined)))

    print("\nSUMMARY")
    for code, name, rng, n in summary:
        print(f"  {code} {name:28s} {n:5d} days  {rng}")
    if failed:
        print("FAILED:", failed)
        sys.exit(1 if len(failed) == len(want) else 0)


if __name__ == "__main__":
    main()
