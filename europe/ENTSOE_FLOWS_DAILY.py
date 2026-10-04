"""
Cross-border electricity flows for Europe from the ENTSO-E Transparency Platform (Cross-Border Physical Flow
[12.1.G], API document A11), one workbook:

  output/Data and Chart Outputs/europe_cross_border_flows_daily.xlsx
    sheet "Net imports": date, one column per country, GWh per day (positive = net import, negative = net export)
    sheet "Borders":     date, one column per directed border "<from>><to>", GWh per day of physical flow
    sheet "Units":       source and definitions

Imports and exports matter for the balance: a country's supply is its generation plus net imports, less what goes
into pumped storage and batteries. Great Britain and Turkey are not in the generation set, but their links to the
countries here are included so those countries' net imports are complete.

Definitions
  - A border is a pair of bidding zones in different countries (zones inside one country - Italy, Denmark, Sweden,
    Norway - are internal and not counted). Each border is fetched in both directions.
  - Physical flow, MWh per UTC day (MW x slot hours); a directed-border day is kept when it covers >= 21 of 24 hours.
  - A country's net import on a day = flows into it - flows out of it over ALL its (valid) borders; the day is
    blank if any of those borders is missing that day. A border with no ENTSO-E data at all (a candidate link that
    does not exist or is not reported) is dropped from the list and noted in the log.

Incremental: the committed workbook is the history store; each border-direction re-fetches the last 7 days plus any
gap in the last 120 days. First run backfills from 2021-01-01 in 180-day windows, stopping gracefully at --max-minutes.

Usage: python3 ENTSOE_FLOWS_DAILY.py [--out-dir DIR] [--start 2021-01-01] [--max-minutes 150]
Requires ENTSOE_API_KEY (environment / GitHub secret, or api_keys.py).
"""
import argparse
import os
import sys
import time
from datetime import datetime, timedelta, timezone

import pandas as pd

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
sys.path.insert(0, HERE)
sys.path.insert(0, ROOT)
import entsoe_common as C  # noqa: E402
import xlsx_notes  # noqa: E402

OUT_DEFAULT = os.path.join(ROOT, "output", "Data and Chart Outputs")
FILE = "europe_cross_border_flows_daily.xlsx"
REVISION_DAYS = 7
GAP_DAYS = 120
WINDOW_DAYS = 180
MIN_HOURS = 21.0

EXTERNAL = {"GB": ("Great Britain", "10YGB----------A"), "TR": ("Turkey", "10YTR-TEIAS----W"),
            # eastern neighbours: Ukraine (ENTSO-E synchronised in 2022; the Burshtyn "island" before), Moldova, Russia, Belarus.
            # Candidate links only - those ENTSO-E does not publish are dropped and logged.
            "UA": ("Ukraine", "10Y1001C--00003F"), "UA-IPS": ("Ukraine", "10Y1001C--000182"),
            "UA-BEI": ("Ukraine", "10YUA-WEPS-----0"), "MD": ("Moldova", "10Y1001A1001A990"),
            "RU": ("Russia", "10Y1001A1001A49F"), "RU-KGD": ("Russia", "10Y1001A1001A50U"), "BY": ("Belarus", "10Y1001A1001A51S")}

# zone -> neighbouring zones (candidate interconnectors; pairs inside one country are skipped, and a pair with no
# ENTSO-E data in either direction is dropped)
ADJACENCY = """
DE-LU: FR NL BE AT CH CZ PL DK1 DK2 SE4 NO2
FR: ES IT-North CH BE GB
ES: PT
IT-North: CH AT SI
IT-South: GR
IT-Centre-South: ME
NL: BE GB DK1 NO2
BE: GB
AT: CZ HU SI CH
CZ: PL SK
SK: PL HU UA UA-IPS
HU: RO RS HR SI UA UA-BEI UA-IPS
RO: RS BG UA UA-IPS MD
BG: GR RS MK TR
GR: MK AL TR
PL: LT SE4 UA UA-IPS BY
HR: SI RS BA
RS: BA ME MK XK
BA: ME
ME: AL XK
MK: XK
AL: XK
DK1: NL NO2 SE3 GB
DK2: SE4
SE1: NO4 FI
SE2: NO3 NO4
SE3: NO1 FI
SE4: LT
LT: BY RU-KGD
NO1: SE3
NO2: GB
FI: NO4 EE RU
EE: LV RU
LV: LT RU BY
IE(SEM): GB
"""


def zone_tables():
    """label -> EIC and label -> country name, for the covered countries plus the external nodes."""
    eic, country = {}, {}
    for code, (name, slug, zones) in C.COUNTRIES.items():
        for label, e in zones:
            eic[label], country[label] = e, name
    for label, (name, e) in EXTERNAL.items():
        eic[label], country[label] = e, name
    return eic, country


def borders():
    eic, country = zone_tables()
    pairs, seen = [], set()
    for line in ADJACENCY.strip().splitlines():
        a, rest = line.split(":")
        for b in rest.split():
            if a not in eic or b not in eic or country[a] == country[b]:
                continue
            key = frozenset((a, b))
            if key in seen:
                continue
            seen.add(key)
            pairs.append((a, b))
    return pairs, eic, country


def read_sheet(path, sheet):
    if not os.path.exists(path):
        return pd.DataFrame()
    try:
        d = pd.read_excel(path, sheet_name=sheet)
    except Exception:  # noqa: BLE001
        return pd.DataFrame()
    d["date"] = pd.to_datetime(d["date"], errors="coerce")
    return d.dropna(subset=["date"]).set_index("date").sort_index()


GAP_INTERPOLATE_DAYS = 14
ENDED_AFTER_DAYS = 30


def fill_for_net(borders_df):
    """Border table used for the country net imports (the Borders sheet itself stays as reported).
    ENTSO-E publishes nothing for a link before it starts or while it carries no flow, so a blank is read as follows:
    before a link's first reported day = 0 (link not yet in service); a gap of up to 14 days inside the series =
    interpolated (outage/late data); a longer gap = 0 (link out of service, e.g. Moyle/EWIC outages); days after
    the last reported day count as 0 if that was over 30 days before the end of the data (link discontinued, e.g. the Russian and
    Belarusian links in 2022), otherwise stay blank so a recent lag is not mistaken for zero flow."""
    out = borders_df.reindex(pd.date_range(borders_df.index.min(), borders_df.index.max()))
    out.index.name = "date"
    for c in out:
        s = out[c]
        first, last = s.first_valid_index(), s.last_valid_index()
        out.loc[:first, c] = out.loc[:first, c].fillna(0.0)
        if last < out.index.max() - pd.Timedelta(days=ENDED_AFTER_DAYS):   # link stopped reporting long ago (e.g. Russia/Belarus 2022)
            out.loc[last:, c] = out.loc[last:, c].fillna(0.0)
        inner = s.loc[first:last]
        gap = inner.isna()
        if not gap.any():
            continue
        run_id = (gap != gap.shift()).cumsum()
        run_len = gap.groupby(run_id).transform("sum")
        interp = inner.interpolate(limit_area="inside")
        fill = inner.copy()
        short = gap & (run_len <= GAP_INTERPOLATE_DAYS)
        fill[short] = interp[short]
        fill[gap & ~short] = 0.0
        out.loc[first:last, c] = fill
    return out


def net_imports(borders_df, pairs, country):
    """Country net imports: into - out over all valid borders touching the country; blank if any border is missing.
    ENTSO-E reports the same Ukrainian tie-lines under several Ukraine zones (UA, UA-IPS, UA-BEI) - the values are identical
    wherever they overlap - so for one neighbour and direction the zones are merged (largest value per day), not summed
    (summing doubled Slovakia's, Hungary's and Romania's Ukraine flows)."""
    covered = {name for name, _, _ in C.COUNTRIES.values()}
    valid = [(x, y) for a, b in pairs for x, y in ((a, b), (b, a)) if f"{x}>{y}" in borders_df]
    filled = fill_for_net(borders_df)
    groups = {}
    for x, y in valid:
        if country[x] == "Ukraine" or country[y] == "Ukraine":
            key = (x, "out") if country[y] == "Ukraine" else (y, "in")
            groups.setdefault(key, []).append((x, y))
    drop = set()
    for g in groups.values():
        if len(g) > 1:
            cols = [f"{x}>{y}" for x, y in g]
            filled[cols[0]] = filled[cols].max(axis=1)
            drop.update(g[1:])
    valid = [v for v in valid if v not in drop]
    net = {}
    for name in sorted(covered):
        ins = [f"{x}>{y}" for x, y in valid if country[y] == name]
        outs = [f"{x}>{y}" for x, y in valid if country[x] == name]
        if not ins and not outs:
            continue
        cols = ins + outs
        full = filled[cols].notna().all(axis=1)
        v = filled[ins].sum(axis=1, min_count=1) - filled[outs].sum(axis=1, min_count=1)
        if not ins:
            v = -filled[outs].sum(axis=1, min_count=1)
        elif not outs:
            v = filled[ins].sum(axis=1, min_count=1)
        net[name] = v.where(full)
    net_df = pd.DataFrame(net).dropna(how="all").round(3)
    net_df.index.name = "date"
    return net_df


def flow_days(out_eic, in_eic, d0, d1, deadline):
    """{date: GWh} physical flow out_eic -> in_eic for the complete UTC days in [d0, d1); None if no data at all."""
    got, any_data = {}, False
    for w0, w1 in C.windows(d0, d1, WINDOW_DAYS):
        if time.time() > deadline:
            raise TimeoutError
        res = C.fetch_split({"documentType": "A11", "out_Domain": out_eic, "in_Domain": in_eic}, w0, w1,
                            C.parse_energy, C.merge_energy)
        time.sleep(0.25)
        if res is None:
            continue
        any_data = True
        for d, mwh in res["e"].items():
            if res["h"].get(d, 0.0) >= MIN_HOURS and w0.date() <= d < w1.date():
                got[d] = mwh / 1000.0
    return got if any_data else None


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out-dir", default=OUT_DEFAULT)
    ap.add_argument("--start", default="2021-01-01")
    ap.add_argument("--max-minutes", type=float, default=150.0)
    ap.add_argument("--net-only", action="store_true", help="recompute Net imports from the saved Borders sheet, no API calls")
    args = ap.parse_args()
    if not args.net_only:
        C.api_key()
    os.makedirs(args.out_dir, exist_ok=True)
    path = os.path.join(args.out_dir, FILE)
    deadline = time.time() + args.max_minutes * 60
    start = datetime.strptime(args.start, "%Y-%m-%d").replace(tzinfo=timezone.utc)
    end = C.today_utc()
    pairs, eic, country = borders()
    old = read_sheet(path, "Borders")
    new_cols, dead, stopped = {}, [], False
    print(f"{len(pairs)} candidate borders ({2 * len(pairs)} directions)", flush=True)
    for a, b in ([] if args.net_only else pairs):
        for x, y in ((a, b), (b, a)):
            col = f"{x}>{y}"
            if stopped:
                break
            s = old[col].dropna() if col in old else pd.Series(dtype=float)
            fs = start
            if len(s):
                fs = max(start, s.index.max().to_pydatetime().replace(tzinfo=timezone.utc) - timedelta(days=REVISION_DAYS))
                cutoff = (end - timedelta(days=GAP_DAYS)).date()
                gaps = [start + timedelta(days=i) for i in range((end - start).days)]
                gaps = [g for g in gaps if g.date() >= cutoff and pd.Timestamp(g.date()) not in s.index and g < fs]
                if gaps:
                    fs = min(fs, gaps[0])
            try:
                got = flow_days(eic[x], eic[y], fs, end, deadline)
            except TimeoutError:
                print("time budget reached; stopping (borders done so far are saved)")
                stopped = True
                break
            except Exception as e:  # noqa: BLE001
                print(f"{col}: FAILED {type(e).__name__}: {e}")
                continue
            if got is None and not len(s):
                dead.append(col)
                print(f"{col}: no ENTSO-E data", flush=True)
                continue
            if got:
                new_cols[col] = pd.Series({pd.Timestamp(d): v for d, v in got.items()}, dtype=float)
            print(f"{col}: {len(got or {})} days from {fs:%Y-%m-%d}", flush=True)
        if stopped:
            break
    borders_df = old.copy() if len(old) else pd.DataFrame()
    if new_cols:
        new = pd.DataFrame(new_cols)
        borders_df = borders_df.reindex(borders_df.index.union(new.index)) if len(borders_df) else new.copy()
        for c in new:
            if c not in borders_df:
                borders_df[c] = float("nan")
            borders_df.loc[new[c].dropna().index, c] = new[c].dropna()
    borders_df = borders_df.dropna(how="all", axis=1).dropna(how="all").sort_index()
    if borders_df.empty:
        print("no flow data")
        return
    borders_df.index.name = "date"

    net_df = net_imports(borders_df, pairs, country)
    lines = ["Europe - cross-border electricity flows and net imports (ENTSO-E Transparency Platform)", "",
             "Source", "ENTSO-E Transparency Platform, Cross-Border Physical Flow [12.1.G] (API document A11). https://transparency.entsoe.eu/",
             "", "Units and definitions",
             "Borders: GWh per UTC day of physical flow from the first zone to the second. Net imports: per country, GWh per day, "
             "flows in minus flows out over all its borders (positive = net import, negative = net export); blank on a day when any "
             "of its borders is missing. Zones inside one country are internal and not counted.",
             "A border blank in the source counts as zero before the link's first reported day and in gaps longer than 14 days (link not in service), and gaps up to 14 days are interpolated, so a short outage does not blank the country's net import. The Borders sheet keeps the raw values. "
             "Great Britain and Turkey are not generation countries here, but their links are included so the neighbours' "
             "net imports are complete. Candidate borders with no ENTSO-E data are dropped: " + (", ".join(dead) or "none this run") + ".",
             f"Re-fetches the last {REVISION_DAYS} days each run plus gaps within {GAP_DAYS} days; history from {args.start}.",
             "", "Last pull", f"{datetime.now(timezone.utc):%Y-%m-%d %H:%M} UTC; {len(borders_df)} days, "
             f"{borders_df.index.min():%Y-%m-%d} to {borders_df.index.max():%Y-%m-%d}; {borders_df.shape[1]} directed borders"]
    xlsx_notes.write_workbook(path, {"Net imports": net_df, "Borders": borders_df.round(3)}, lines,
                              {"Source", "Units and definitions", "Last pull"})
    print(f"saved {FILE}: {len(borders_df)} days x {borders_df.shape[1]} directed borders; net imports for {net_df.shape[1]} countries")
    if dead:
        print("dead borders:", dead)


if __name__ == "__main__":
    main()
