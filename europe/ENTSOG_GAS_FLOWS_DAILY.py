"""
Gas flows by country for Europe from the ENTSOG Transparency Platform (physical flows at every interconnection
point, API operationalData / "Physical Flow", daily), one workbook:

  output/Data and Chart Outputs/europe_gas_flows_daily.xlsx
    sheet "Country balance": date, then per country <CC>_<category>_GWhd with categories
        production        gas entering the country's grid from production points
        imports           pipeline gas entering from another country's grid (any origin, EU or not)
        exports           pipeline gas leaving to another country's grid
        lng               gas entering from LNG terminals (ALSI's terminal send-out is the better LNG series; kept for checks)
        storage_out / storage_in   gas leaving / entering storage (AGSI+ is the better storage series)
        distribution      gas leaving to distribution networks (households, commerce, small industry)
        final_consumers   gas leaving directly to large final consumers (industry, power plants)
    sheet "Imports by origin": date, one column per origin country, GWh per day entering EU27 grids from outside the EU (the UK counts as outside)
    sheet "Exports by destination": date, one column per destination country, GWh per day leaving EU27 grids to outside the EU
    sheet "Border flows": date, then one column per directed border "<from>><to>" (e.g. "NL>DE"), GWh per day; the flow on the border
        after de-duplication (see below). Lets a combined balance (Germany + Netherlands) cancel the flows between its members.
    sheet "Border flows by side": date, then per directed border "<from>><to> exit" (the sender's own total) and "<from>><to> entry" (the
        receiver's own total), GWh per day, after the same de-duplication; the two differ by the measurement difference between the operators.
        Border flows is the larger of the two. Lets a combined balance cancel intra-block flows with one number per border.
    sheet "Units": source and definitions

Each flow row is (point, operator, entry|exit). It is classed by the system on the other side of the point, taken from
ENTSOG's interconnections list: the operator's country is the first two letters of its key, the adjacent system gives
the type (Transmission / Production / LNG Terminals / Storage / Distribution / Final Consumers) and country.
A sender's exports and a receiver's imports use their own side of the border; the other side is used only where the own country
publishes no row for those points (Baumgarten on the Austrian side). The Border flows sheet shows the larger of the two sides. Within a side, operators that report the same gas at one point are counted once, and a virtual point (VIP) and the physical
points it aggregates are taken as the larger of the two, not summed (ENTSOG reports both; VIP Brandov = EUGAL + OPAL + Hora Svate Katerina). ENTSOG reports kWh/d; converted to GWh/d. Empty (unreported) values are left blank.

Caveats: ENTSOG data are operational flows (allocations/nominations), restated for recent days (the last 45 days are
re-fetched each run) and incomplete for some points; the platform keeps about 5 years, so the committed workbook is the
history store. Germany reports consumption as aggregated final-consumer points; Spain has few demand points. A country's
balance (supply - uses) therefore does not always close - the master shows the gap rather than hiding it.

Usage: python3 ENTSOG_GAS_FLOWS_DAILY.py [--out-dir DIR] [--start 2021-10-04]
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
HEADERS = {"User-Agent": "Mozilla/5.0", "Accept": "application/json"}
OUT_DEFAULT = os.path.join(ROOT, "output", "Data and Chart Outputs")
FILE = "europe_gas_flows_daily.xlsx"
START_DEFAULT = "2021-10-04"
RELOAD_DAYS = 45
WINDOW_DAYS = 31
EU27 = {"AT", "BE", "BG", "HR", "CY", "CZ", "DK", "EE", "FI", "FR", "DE", "GR", "HU", "IE", "IT", "LV", "LT", "LU",
        "MT", "NL", "PL", "PT", "RO", "SK", "SI", "ES", "SE"}
CATEGORIES = ["production", "imports", "exports", "lng", "storage_out", "storage_in", "distribution", "final_consumers"]
# ENTSOG country codes shown in the master (UK = Great Britain / Northern Ireland)
NAMES = {"AT": "Austria", "BE": "Belgium", "BG": "Bulgaria", "HR": "Croatia", "CZ": "Czechia", "DK": "Denmark",
         "EE": "Estonia", "FI": "Finland", "FR": "France", "DE": "Germany", "GR": "Greece", "HU": "Hungary",
         "IE": "Ireland", "IT": "Italy", "LV": "Latvia", "LT": "Lithuania", "LU": "Luxembourg", "NL": "Netherlands",
         "PL": "Poland", "PT": "Portugal", "RO": "Romania", "SK": "Slovakia", "SI": "Slovenia", "ES": "Spain",
         "SE": "Sweden", "UK": "United Kingdom", "CH": "Switzerland"}


def get_json(path, params, tries=4):
    last = ""
    for attempt in range(tries):
        try:
            r = requests.get(f"{API}/{path}", params=params, headers=HEADERS, timeout=(15, 300))
        except requests.RequestException as e:
            last = f"{type(e).__name__}"
            time.sleep(10 * (attempt + 1))
            continue
        if r.status_code in (429,) or r.status_code >= 500:
            last = f"HTTP {r.status_code}"
            time.sleep(20 * (attempt + 1))
            continue
        if r.status_code == 404:   # ENTSOG answers "no result" with 404
            return {}
        if r.status_code != 200:
            raise RuntimeError(f"{path} {params}: HTTP {r.status_code} {r.text[:150]}")
        return r.json()
    raise RuntimeError(f"{path} {params}: failed after {tries} tries ({last})")


VIP_KEYS = set()   # pointKeys of virtual interconnection points (label starts with "VIP"), filled by adjacency()


def adjacency():
    """(pointKey, operatorKey, direction) -> (adjacent infrastructure type, adjacent country key)."""
    ics = get_json("interconnections", {"limit": -1}).get("interconnections", [])
    adj = {}
    for i in ics:
        pk = i["pointKey"]
        if str(i.get("pointLabel") or "").strip().upper().startswith("VIP"):
            VIP_KEYS.add(pk)
        if i.get("toOperatorKey"):
            adj[(pk, i["toOperatorKey"], "entry")] = (i.get("fromInfrastructureTypeLabel"), i.get("fromCountryKey"))
        if i.get("fromOperatorKey"):
            adj[(pk, i["fromOperatorKey"], "exit")] = (i.get("toInfrastructureTypeLabel"), i.get("toCountryKey"))
    return adj


def classify(row, adj):
    """-> (country, category, origin) or None. origin is the adjacent country for entries (imports)."""
    op, pk, d = row.get("operatorKey") or "", row["pointKey"], row["directionKey"]
    country = op[:2]
    a = adj.get((pk, op, d))
    if not country or a is None:
        return None
    typ, acountry = a
    if typ == "Transmission":
        if acountry and acountry != country:
            return country, ("imports" if d == "entry" else "exports"), acountry
        return None   # between two systems of the same country
    if typ == "Production" and d == "entry":
        return country, "production", None
    if typ == "LNG Terminals" and d == "entry":
        return country, "lng", None
    if typ == "Storage":
        return country, ("storage_out" if d == "entry" else "storage_in"), None
    if typ == "Distribution" and d == "exit":
        return country, "distribution", None
    if typ == "Final Consumers" and d == "exit":
        return country, "final_consumers", None
    return None


def _dedupe_operators(vals):
    """Several operators on one side of a point often report the same gas (e.g. both German TSOs at Ueberackern, 24.2 and 24.1 TWh):
    values within 1% of one already kept are dropped."""
    kept = []
    for v in sorted(vals, reverse=True):
        if not any(abs(v - k) <= max(0.01 * max(v, k), 1e-3) for k in kept):
            kept.append(v)
    return sum(kept)


def border_flows(rows):
    """rows: (day, reporting country, direction, other country, pointKey, operator, GWh). -> {(day, from, to): GWh/d}.
    Each side of a border is the sum of its points with (1) operators that duplicate each other dropped and (2) the virtual point
    (VIP) and the physical points it aggregates taken as the larger of the two, not their sum (ENTSOG reports both: e.g. VIP Brandov =
    EUGAL + OPAL + Hora Svate Katerina, so Czech imports from Germany were counted twice). Each country keeps its own side; the
    other side only where the own country publishes no row for those points (Baumgarten on the Austrian side)."""
    by_pt = {}
    for day, c, d, oc, pk, op, g in rows:
        by_pt.setdefault((day, c, d, oc, pk), []).append(g)
    side, pks, seen = {}, {}, set()
    for (day, c, d, oc, pk), vals in by_pt.items():
        k = (day, c, d, oc, pk in VIP_KEYS)
        side[k] = side.get(k, 0.0) + _dedupe_operators(vals)
        pks.setdefault((day, c, d, oc), set()).add(pk)
        seen.add((day, c, d, pk))
    own = {}
    for (day, c, d, oc, vip), v in side.items():
        k = (day, c, d, oc)
        own[k] = max(own.get(k, 0.0), v)      # larger of the VIP sum and the physical-points sum
    flows = {}
    for (day, c, d, oc), v in own.items():
        a, b = (c, oc) if d == "exit" else (oc, c)
        flows.setdefault((day, a, b), {})["exit" if d == "exit" else "entry"] = v
    # (exports credited to the sender, imports credited to the receiver, border flow). Each country keeps its own side. The other side
    # is used only where the own country publishes no row at all for those points that day (Baumgarten has no Austrian-side row), not
    # where it publishes a zero (Greece's Kulata exit shows ~30 GWh/d in winter where Bulgaria's entry reports 0) and not where it
    # books the same point against a different neighbour (Komotini IGB is Bulgaria's entry "from AL" but Greece's exit "to BG"; filling
    # it counted the gas twice and put Bulgaria 55 points out).
    out = {}
    for (day, a, b), v in flows.items():
        ex = v["exit"] if "exit" in v else (v["entry"] if not any((day, a, "exit", pk) in seen for pk in pks[(day, b, "entry", a)]) else 0.0)
        im = v["entry"] if "entry" in v else (v["exit"] if not any((day, b, "entry", pk) in seen for pk in pks[(day, a, "exit", b)]) else 0.0)
        out[(day, a, b)] = (ex, im, max(v.values()))
    return out


def fetch_window(d0, d1, adj):
    """Daily GWh/d per (country, category), per origin / destination and per border for [d0, d1]."""
    data = get_json("operationalData", {"indicator": "Physical Flow", "periodType": "day", "from": d0.isoformat(),
                                        "to": d1.isoformat(), "limit": -1}).get("operationalData", [])
    cat, origin, dest, border, unclassified = {}, {}, {}, {}, 0
    border_side = {}      # ("A>B exit" = the sender's own total, "A>B entry" = the receiver's own total), as booked in the country balances
    trans = []
    for r in data:
        v = r.get("value")
        if v in (None, ""):
            continue
        try:
            gwh = float(v) / 1e6 if r.get("unit", "kWh/d") == "kWh/d" else float(v) * {"MWh/d": 1e-3, "GWh/d": 1.0}.get(r["unit"], float("nan"))
        except (TypeError, ValueError):
            continue
        c = classify(r, adj)
        if c is None:
            unclassified += 1
            continue
        country, category, orig = c
        day = pd.Timestamp(r["periodFrom"][:10])
        if category in ("imports", "exports"):
            trans.append((day, country, "entry" if category == "imports" else "exit", orig, r["pointKey"], r.get("operatorKey"), gwh))
            continue
        cat[(day, f"{country}_{category}_GWhd")] = cat.get((day, f"{country}_{category}_GWhd"), 0.0) + gwh
    for (day, a, b), (ex, im, g) in border_flows(trans).items():
        border[(day, f"{a}>{b}")] = g
        border_side[(day, f"{a}>{b} exit")] = ex
        border_side[(day, f"{a}>{b} entry")] = im
        cat[(day, f"{a}_exports_GWhd")] = cat.get((day, f"{a}_exports_GWhd"), 0.0) + ex
        cat[(day, f"{b}_imports_GWhd")] = cat.get((day, f"{b}_imports_GWhd"), 0.0) + im
        if a not in EU27 and b in EU27:
            origin[(day, a)] = origin.get((day, a), 0.0) + im
        if a in EU27 and b not in EU27:
            dest[(day, b)] = dest.get((day, b), 0.0) + ex
    return cat, origin, dest, border, border_side, len(data), unclassified


def to_frame(d, name):
    if not d:
        return pd.DataFrame()
    s = pd.Series(d)
    s.index = pd.MultiIndex.from_tuples(s.index)
    return s.unstack().sort_index()


def read_sheet(path, sheet):
    if not os.path.exists(path):
        return pd.DataFrame()
    try:
        d = pd.read_excel(path, sheet_name=sheet)
    except Exception:  # noqa: BLE001
        return pd.DataFrame()
    d["date"] = pd.to_datetime(d["date"], errors="coerce")
    return d.dropna(subset=["date"]).set_index("date").sort_index()


def merge(old, new):
    if old.empty:
        return new.sort_index()
    out = old.reindex(old.index.union(new.index))
    for c in new.columns:
        if c not in out:
            out[c] = float("nan")
        s = new[c].dropna()
        out.loc[s.index, c] = s
    return out.sort_index()


def drop_spikes(bal, ratio=15.0, min_median=20.0, min_value=1500.0):
    """A single bad ENTSOG value (e.g. Germany final consumers 249,215 GWh on 2025-08-21 against a normal ~350) would swamp
    a monthly total. A value above `ratio` x the centred 61-day median (median above 20 GWh/d, value above 1,500 GWh/d) is blanked and logged;
    genuine cold-snap or storage peaks are well inside that."""
    med = bal.rolling(61, center=True, min_periods=20).median()
    bad = (bal > ratio * med) & (med > min_median) & (bal > min_value)
    for col in bal.columns[bad.any()]:
        for day in bal.index[bad[col]]:
            print(f"  dropped implausible value {col} {day:%Y-%m-%d}: {bal.at[day, col]:.0f} vs median {med.at[day, col]:.0f}")
    return bal.mask(bad)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out-dir", default=OUT_DEFAULT)
    ap.add_argument("--start", default=START_DEFAULT)
    ap.add_argument("--max-minutes", type=float, default=90.0)
    ap.add_argument("--rebuild", action="store_true", help="ignore the committed workbook and re-pull everything from --start")
    args = ap.parse_args()
    os.makedirs(args.out_dir, exist_ok=True)
    path = os.path.join(args.out_dir, FILE)
    deadline = time.time() + args.max_minutes * 60
    start = date.fromisoformat(args.start)
    end = date.today() - timedelta(days=1)
    bal, org, dst, brd, sid = (read_sheet(path, "Country balance"), read_sheet(path, "Imports by origin"),
                               read_sheet(path, "Exports by destination"), read_sheet(path, "Border flows"),
                               read_sheet(path, "Border flows by side"))
    if args.rebuild or brd.empty or sid.empty:     # workbooks written before the border sheets / VIP de-duplication hold double-counted flows: pull again
        bal, org, dst, brd, sid = (pd.DataFrame(),) * 5
    fs = start
    if len(bal):
        fs = max(start, (bal.index.max() - timedelta(days=RELOAD_DAYS)).date())
    print(f"fetching {fs} -> {end} (have {len(bal)} days)", flush=True)
    adj = adjacency()
    print(f"{len(adj)} (point, operator, direction) classifications", flush=True)
    cur = fs
    while cur <= end:
        if time.time() > deadline:
            print("time budget reached; stopping (windows done so far are saved)")
            break
        nxt = min(cur + timedelta(days=WINDOW_DAYS - 1), end)
        t0 = time.time()
        cat, origin, dest, border, bside, n, unc = fetch_window(cur, nxt, adj)
        print(f"  {cur} -> {nxt}: {n} rows, {unc} unclassified, {len(cat)} country-day values in {time.time() - t0:.0f}s", flush=True)
        bal = merge(bal, to_frame(cat, "balance"))
        org = merge(org, to_frame(origin, "origin"))
        dst = merge(dst, to_frame(dest, "dest"))
        brd = merge(brd, to_frame(border, "border"))
        sid = merge(sid, to_frame(bside, "side"))
        bal.index.name = org.index.name = dst.index.name = brd.index.name = sid.index.name = "date"
        cur = nxt + timedelta(days=1)
        time.sleep(1)
    if bal.empty:
        print("no data")
        return
    bal = bal[sorted(bal.columns)].round(3)
    org = org[sorted(org.columns)].round(3) if not org.empty else org
    dst = dst[sorted(dst.columns)].round(3) if not dst.empty else dst
    lines = ["Europe - gas flows by country (ENTSOG Transparency Platform)", "",
             "Source", "ENTSOG Transparency Platform, operational data: Physical Flow, daily, at every interconnection point "
             "(https://transparency.entsog.eu/). The TSOs' own reporting.",
             "", "Units and definitions",
             "GWh per gas day (ENTSOG kWh/d / 1e6). <CC>_<category>_GWhd: production, imports (pipeline gas entering from another country), "
             "exports (leaving to another country), lng (entering from LNG terminals), storage_out / storage_in, distribution "
             "(exits to distribution networks) and final_consumers (exits to large consumers: industry, power plants). UK = Great Britain "
             "and Northern Ireland.",
             "Each flow row is classed by the system on the other side of the point (ENTSOG interconnections list); cross-border flows "
             "are own side per country (other side only where the own side publishes no point; VIP and physical points not summed, duplicate operators counted once). Imports by origin / Exports by destination: pipeline gas entering EU27 grids from, or leaving "
             "to, countries outside the EU27 (the UK counts as outside), by country.",
             "Operational data: restated for recent days (last " + str(RELOAD_DAYS) + " days re-fetched each run) and unreported for "
             "some points; Germany reports aggregated final consumers, Spain has few demand points, so a country's supply and uses do "
             "not always balance. Single values above 1,500 GWh/d and 15x the local median are dropped as data errors. ENTSOG keeps about 5 years of history; this workbook is the history store.",
             "", "Last pull", f"{datetime.now(timezone.utc):%Y-%m-%d %H:%M} UTC; {len(bal)} days, {bal.index.min():%Y-%m-%d} to {bal.index.max():%Y-%m-%d}"]
    bal = drop_spikes(bal)
    sheets = {"Country balance": bal}
    if not brd.empty:
        sheets["Border flows"] = brd[sorted(brd.columns)].round(3)
    if not sid.empty:
        sheets["Border flows by side"] = drop_spikes(sid[sorted(sid.columns)]).round(3)
    if not org.empty:
        sheets["Imports by origin"] = org
    if not dst.empty:
        sheets["Exports by destination"] = dst
    xlsx_notes.write_workbook(path, sheets, lines, {"Source", "Units and definitions", "Last pull"})
    print(f"saved {FILE}: {len(bal)} days x {bal.shape[1]} columns; origins: {list(org.columns) if not org.empty else 'none'}")


if __name__ == "__main__":
    main()
