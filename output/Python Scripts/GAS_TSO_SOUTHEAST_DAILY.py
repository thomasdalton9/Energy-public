"""
National gas consumption from further gas TSOs' own published series (raw operator data, no Eurostat gap-filling).
One workbook, sheet "Daily" (GWh per gas day), europe/GAS_TSO_SOUTHEAST_DAILY.py -> europe_tso_gas_demand_extra_daily.xlsx.

  PL  Gaz-System Moduł Informacji Rynkowej, "Actual quantity of gas transmitted" (KspRealization, kWh per zone and gas day, billing
      data since 2018, operative data for the latest days):  https://swi.gaz-system.pl/mir/#/public/bil/ksp-realization
      PL_distribution = exit to DSO networks (H and L gas; PL_dso_return = gas entering the grid back from DSO networks, shown for
      information, not netted); PL_final_customers = aggregated
      final-customer exit points; PL_other = TSO own needs, nitrogen removal / mixing plants and aggregated PPG points;
      PL_total = the three. Exchange/OTC points, storage, interconnectors and compulsory stocks are not consumption and are left out.

Incremental: the committed workbook is the history store; each run re-fetches the last REVISION_DAYS days (operative values are
replaced by billing values) plus any gaps.

Usage: python3 GAS_TSO_SOUTHEAST_DAILY.py [--out-dir DIR] [--start 2021-01-01] [--only PL,RO,ES,FI]
"""
import argparse
import io
import os
import re
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
FILE = "europe_tso_gas_demand_extra_daily.xlsx"
REVISION_DAYS = 45
UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36"
COLUMNS = ["PL_distribution", "PL_final_customers", "PL_other", "PL_total", "PL_dso_return", "RO_final_customers", "RO_distribution", "RO_total", "HR_distribution", "HR_final_customers", "HR_total", "FI_total"]


def get(url, tries=3, **kw):
    for i in range(tries):
        try:
            r = requests.get(url, timeout=(15, 90), **kw)
            if r.ok:
                return r
            if r.status_code in (429, 500, 502, 503, 504):
                time.sleep(4 * (i + 1))
                continue
            raise RuntimeError(f"HTTP {r.status_code} {r.text[:200]} for {r.url}")
        except requests.RequestException:
            time.sleep(4 * (i + 1))
    raise RuntimeError(f"gave up on {url}")


# ---- Poland: Gaz-System MIR --------------------------------------------------------------------------------------------
GS = "https://swi.gaz-system.pl/mir/api/v1/en/KspRealization/data"
PL_DSO_EXIT = {"900557", "900558"}
PL_DSO_ENTRY = {"907555", "907556"}
PL_FINAL = {"909027", "909028"}
PL_OTHER = {"100003", "100007", "900008", "900009", "900131", "900133"}


def num(v):
    s = re.sub(r"[^\d,.\-]", "", str(v or ""))        # "646 036 768" (thin/no-break spaces as thousands separators)
    if not s or s == "-":
        return float("nan")
    return float(s.replace(",", ""))


def poland(d0, d1):
    rec = {}
    s = d0
    while s <= d1:
        e = min(s + timedelta(days=30), d1)
        start, got = 0, 0
        while True:
            body = {"start": start, "length": 2000, "globalFilter": {"value": ""}, "lang": "en",
                    "sort": [{"column": "gasDay", "order": "asc"}, {"column": "zoneCode", "order": "asc"}],
                    "customFilters": {"filters": [{"filterName": "gasDay$from", "values": s.isoformat()},
                                                  {"filterName": "gasDay$to", "values": e.isoformat()}]}}
            for i in range(3):
                r = requests.post(GS, json=body, headers={"User-Agent": UA}, timeout=(15, 120))
                if r.ok:
                    break
                time.sleep(4 * (i + 1))
            r.raise_for_status()
            j = r.json()
            rows = j.get("data", [])
            for x in rows:
                z = str(x.get("zoneCode"))
                if z in PL_DSO_EXIT | PL_DSO_ENTRY | PL_FINAL | PL_OTHER:
                    rec.setdefault(pd.Timestamp(x["gasDay"]), {})[z] = num(x.get("realization"))
            got += len(rows)
            if not rows or got >= int(j.get("total") or 0):
                break
            start += len(rows)
        s = e + timedelta(days=1)
    if not rec:
        return pd.DataFrame()
    d = pd.DataFrame.from_dict(rec, orient="index").sort_index()
    g = lambda ids: d[[c for c in ids if c in d]].sum(axis=1, min_count=1) / 1e6          # kWh -> GWh
    out = pd.DataFrame(index=d.index)
    out["PL_distribution"] = g(PL_DSO_EXIT)
    out["PL_dso_return"] = g(PL_DSO_ENTRY)
    out["PL_final_customers"] = g(PL_FINAL)
    out["PL_other"] = g(PL_OTHER).fillna(0)
    out["PL_total"] = out["PL_distribution"] + out["PL_final_customers"] + out["PL_other"]
    out = out[out["PL_distribution"].notna() & out["PL_final_customers"].notna()]
    return out


# ---- Finland: Gasgrid Finland ------------------------------------------------------------------------------------------
GG_PAGE = "https://gasgrid.fi/en/gas-business/transparency-and-market-information/"


def finland(d0, d1):
    """Gasgrid publishes one workbook 'Gas consumption in Finland <date>.xlsx' (a sheet per year, GWh/day, GCV) on its transparency
    page; the file name changes with each upload, so the link is read from the page."""
    page = get(GG_PAGE, headers={"User-Agent": UA}).text
    links = re.findall(r'href="([^"]*[Gg]as-[Cc]onsumption-in-[Ff]inland[^"]*\.xlsx)"', page)
    if not links:
        raise RuntimeError("no 'Gas consumption in Finland' xlsx link on the Gasgrid page")
    url = links[0]
    print(f"  Gasgrid file: {url}", flush=True)
    x = pd.ExcelFile(io.BytesIO(get(url, headers={"User-Agent": UA}).content))
    rec = {}
    for sh in x.sheet_names:
        raw = x.parse(sh, header=None)
        for _, row in raw.iterrows():
            vals = list(row)
            for i, v in enumerate(vals):
                if isinstance(v, (pd.Timestamp, datetime)) and not pd.isna(v):
                    nums = [w for w in vals[i + 1:] if isinstance(w, (int, float)) and not pd.isna(w)]
                    if nums:
                        rec[pd.Timestamp(v).normalize()] = float(nums[0])
                    break
    out = pd.DataFrame({"FI_total": pd.Series(rec, dtype=float)}).sort_index()
    return out[(out.index >= pd.Timestamp(d0)) & (out.index <= pd.Timestamp(d1))]


# ---- Romania: Transgaz -------------------------------------------------------------------------------------------------
TG = "https://www.transgaz.ro/new-tabel-transparenta-masuratori_en.php?poz=197"
TG_REF = "https://www.transgaz.ro/en/clients/operational-data/physical-flows"


def romania(d0, d1):
    """Transgaz 'Physical flows' table (commercial measurements per relevant point, MWh/day at 15C/15C): the Excel export of the
    page's grid for a date range. Exit 'SM-CF001' = final clients connected directly to the transmission system, 'SM-SD001' =
    distribution systems; RO_total is their sum (gas reaching consumers through the grid; production consumed outside it is not seen)."""
    rec = {}
    s = d0
    while s <= d1:
        e = min(s + timedelta(days=30), d1)
        ses = requests.Session()
        ses.headers.update({"User-Agent": UA, "Referer": TG_REF})
        page = ses.get(TG, timeout=(15, 60)).text
        vs = re.search(r"id='grid_viewstate'[^>]*value='([^']*)'", page)
        data = {"data_start": s.isoformat(), "data_stop": e.isoformat(), "puncte": "SM", "um": "MW", "Exportbtn": "Export",
                "IgnorePaging": "on", "grid_cmd": "", "grid_viewstate": vs.group(1) if vs else ""}
        content = None
        for i in range(3):
            try:
                r = ses.post(TG, data=data, timeout=(15, 120))
                if r.ok and r.content[:2] == b"PK":
                    content = r.content
                    break
            except requests.RequestException:
                pass
            time.sleep(4 * (i + 1))
        if content is None:
            raise RuntimeError(f"Transgaz export failed for {s}..{e}")
        x = pd.read_excel(io.BytesIO(content), header=None)
        for _, row in x.iterrows():
            try:
                day = pd.Timestamp(str(row[0])[:10])
            except Exception:  # noqa: BLE001
                continue
            if not re.match(r"\d{4}-\d\d-\d\d", str(row[0])[:10]):
                continue
            code = str(row[1]).strip()
            if code in ("SM-CF001", "SM-SD001"):
                rec.setdefault(day, {})[code] = num(row[3])
        s = e + timedelta(days=1)
    if not rec:
        return pd.DataFrame()
    d = pd.DataFrame.from_dict(rec, orient="index").sort_index() / 1000.0                 # MWh -> GWh
    out = pd.DataFrame(index=d.index)
    out["RO_final_customers"] = d.get("SM-CF001")
    out["RO_distribution"] = d.get("SM-SD001")
    out["RO_total"] = out["RO_final_customers"] + out["RO_distribution"]
    return out.dropna(subset=["RO_total"])


# ---- Spain 2021-22: Enagas monthly statistical bulletin ------------------------------------------------------------------
ENAGAS_PAGE = "https://www.enagas.es/en/technical-management-system/energy-data/publications/gas-statistical-bulletin/"
MONTHS = {"jan": 1, "ene": 1, "feb": 2, "mar": 3, "apr": 4, "abr": 4, "may": 5, "jun": 6, "jul": 7, "aug": 8, "ago": 8, "sep": 9,
          "oct": 10, "nov": 11, "dec": 12, "dic": 12}
ES_MONTHLY_COLS = ["ES_national", "ES_conventional", "ES_power", "source_file"]


def _gwh(tok):
    return float(re.sub(r"[.,]", "", tok))                     # whole GWh; '.' or ',' are thousands separators


FULL_MONTHS = ["january", "february", "march", "april", "may", "june", "july", "august", "september", "october", "november", "december"]


def parse_enagas_bulletin(content):
    """Month, national market demand, conventional and power-generation demand (GWh) from the 'Evolution of gas demand' table of an
    Enagas monthly bulletin. Two layouts exist: a clean text table ('National Market demand 26.953 ...', month in the 'GWh May-2021'
    header) and one whose PDF text comes out letter-spaced in table order; for the latter the month is taken from the title page and
    the first three numbers after the 'Month' column header are national, conventional and power (checked: national = conv + power)."""
    import pdfplumber
    with pdfplumber.open(io.BytesIO(content)) as pdf:
        pages = [pg.extract_text() or "" for pg in pdf.pages[:8]]
    for t in pages:
        m = re.search(r"National\s*Market\s*demand\s*([\d.,]+)", t, re.I)
        if not m:
            continue
        h = re.search(r"GWh\s*([A-Za-z]{3})[A-Za-z]*[-\s]*(20\d\d)", t)
        if not h or h.group(1).lower() not in MONTHS:
            break
        conv = re.search(r"^Conventional\s*([\d.,]+)", t, re.M | re.I)
        pw = re.search(r"^Power\s*generation\s*([\d.,]+)", t, re.M | re.I)
        return (pd.Timestamp(int(h.group(2)), MONTHS[h.group(1).lower()], 1), _gwh(m.group(1)),
                _gwh(conv.group(1)) if conv else float("nan"), _gwh(pw.group(1)) if pw else float("nan"))
    title = re.search(r"(" + "|".join(FULL_MONTHS) + r")\s*(20\d\d)", pages[0].lower()) if pages else None
    if not title:
        return None
    month = pd.Timestamp(int(title.group(2)), FULL_MONTHS.index(title.group(1)) + 1, 1)
    for t in pages[1:6]:
        lines = [re.sub(r"\s+", "", ln) for ln in t.split("\n")]
        idx = next((i for i, ln in enumerate(lines) if ln.startswith("Month")), None)
        if idx is None:
            continue
        nums = [ln for ln in lines[idx + 1:] if re.fullmatch(r"\d{1,3}(?:[.,]\d{3})*", ln)]
        if len(nums) >= 3:
            nat, conv, pw = (_gwh(x) for x in nums[:3])
            if abs(nat - conv - pw) <= 5:
                return month, nat, conv, pw
    return None


def spain_monthly(start, old):
    """Monthly national gas demand 2021-22 (and later, for checking against the daily JSON) from the Enagas bulletin archive. The
    Enagas demand-history JSON used by TSO_GAS_DEMAND_DAILY.py is empty before 2023; the monthly bulletins (PDF) go back to 2018.
    Bulletins already parsed (source_file) are skipped, except the newest three, which are re-read for revisions."""
    h = {"User-Agent": UA}
    files = {}
    today = datetime.now(timezone.utc).date()
    for y in range(start.year, today.year + 1):
        for m in range(1, 13):
            if (y, m) < (start.year, start.month) or (y, m) > (today.year, today.month):
                continue
            try:
                r = get(ENAGAS_PAGE, params={"category": "", "month": m, "year": y}, headers=h)
            except Exception as e:  # noqa: BLE001
                print(f"  Enagas listing {y}-{m:02d}: {type(e).__name__}", flush=True)
                continue
            for x in re.findall(r'href="(/content/dam[^"]+\.pdf)"', r.text):
                files.setdefault(x, (y, m))
    print(f"  Enagas bulletins listed: {len(files)}", flush=True)
    known = set(old["source_file"].dropna()) if "source_file" in old else set()
    newest = {f for f, _ in sorted(files.items(), key=lambda kv: kv[1])[-3:]}
    rows = {}
    for f, ym in sorted(files.items(), key=lambda kv: kv[1]):
        name = f.split("/")[-1]
        if name in known and f not in newest:
            continue
        try:
            res = parse_enagas_bulletin(get("https://www.enagas.es" + f, headers=h).content)
        except Exception as e:  # noqa: BLE001
            print(f"  {name}: {type(e).__name__} {str(e)[:80]}", flush=True)
            continue
        if res is None:
            print(f"  {name}: demand table not found", flush=True)
            continue
        month, nat, conv, pw = res
        if month < pd.Timestamp(start):
            continue
        rows[month] = {"ES_national": nat, "ES_conventional": conv, "ES_power": pw, "source_file": name}
    return pd.DataFrame.from_dict(rows, orient="index", columns=ES_MONTHLY_COLS).sort_index() if rows else pd.DataFrame(columns=ES_MONTHLY_COLS)


# ---- Croatia: Plinacro SUKAP -------------------------------------------------------------------------------------------
PLIN = "https://www.sukap.plinacro.hr/pub/protoci/search"


def croatia(d0, d1):
    """Plinacro public data portal (SUKAP), 'Realised physical flows' (kWh/d at GCV 25C/0C) aggregated by connection type:
    DISTRIBUTION = exits to distribution systems, END_BUYER = exits to final customers connected to the transmission system."""
    out = {}
    for ptype, col in (("DISTRIBUTION", "HR_distribution"), ("END_BUYER", "HR_final_customers")):
        rec = {}
        s = d0
        while s <= d1:
            e = min(s + timedelta(days=120), d1)
            start = 0
            while True:
                body = {"gasDayFrom": f"{(s - timedelta(days=1)).isoformat()}T06:00:00.000+00:00", "gasDayTo": f"{e.isoformat()}T06:00:00.000+00:00",
                        "pointType": ptype, "pointId": None, "sortFieldList": [{"property": "gasDay", "direction": "ASC"}],
                        "page": start // 500 + 1, "start": start, "limit": 500}
                for i in range(3):
                    r = requests.post(PLIN, json=body, headers={"User-Agent": UA}, timeout=(15, 90))
                    if r.ok:
                        break
                    time.sleep(4 * (i + 1))
                r.raise_for_status()
                j = r.json()
                rows = j.get("data", [])
                for x in rows:
                    v = x.get("measuredGcv")
                    if v is not None:
                        rec[pd.Timestamp(str(x["gasDay"])[:10])] = rec.get(pd.Timestamp(str(x["gasDay"])[:10]), 0.0) + float(v) / 1e6
                start += len(rows)
                if not rows or start >= int(j.get("total") or 0):
                    break
            s = e + timedelta(days=1)
        out[col] = pd.Series(rec, dtype=float)
    d = pd.DataFrame(out).sort_index()
    d = d[(d.index >= pd.Timestamp(d0)) & (d.index <= pd.Timestamp(d1))]
    d["HR_total"] = d["HR_distribution"] + d["HR_final_customers"]
    return d.dropna(subset=["HR_total"])


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


def read_monthly(path):
    if os.path.exists(path):
        try:
            d = pd.read_excel(path, sheet_name="Monthly")
            d["month"] = pd.to_datetime(d["month"], errors="coerce")
            return d.dropna(subset=["month"]).set_index("month").sort_index()
        except Exception:  # noqa: BLE001
            pass
    return pd.DataFrame(columns=ES_MONTHLY_COLS)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out-dir", default=OUT_DEFAULT)
    ap.add_argument("--start", default="2021-01-01")
    ap.add_argument("--only", default="", help="comma-separated subset of PL,FI,... (the workbook keeps the other countries' history)")
    args = ap.parse_args()
    only = {x.strip().upper() for x in args.only.split(",") if x.strip()}
    os.makedirs(args.out_dir, exist_ok=True)
    path = os.path.join(args.out_dir, FILE)
    start = datetime.strptime(args.start, "%Y-%m-%d").date()
    today = datetime.now(timezone.utc).date()
    old = read_existing(path)
    combined = old.copy()
    for code, label, fn, cols in (("PL", "Poland (Gaz-System)", poland, [c for c in COLUMNS if c.startswith("PL_")]),
                                  ("RO", "Romania (Transgaz)", romania, [c for c in COLUMNS if c.startswith("RO_")]),
                                  ("HR", "Croatia (Plinacro)", croatia, [c for c in COLUMNS if c.startswith("HR_")]),
                                  ("FI", "Finland (Gasgrid)", finland, ["FI_total"])):
        if only and code not in only:
            continue
        print(f"{label}: start", flush=True)
        have = old[cols].dropna(how="all")
        fs = start if (have.empty or any(old[c].dropna().empty for c in cols)) else max(start, have.index.max().date() - timedelta(days=REVISION_DAYS))
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
    monthly = read_monthly(path)
    if not only or "ES" in only:
        print("Spain (Enagas bulletin): start", flush=True)
        try:
            new_m = spain_monthly(start, monthly)
            if not new_m.empty:
                monthly = monthly.reindex(monthly.index.union(new_m.index))
                for c in ES_MONTHLY_COLS:
                    monthly.loc[new_m.index, c] = new_m[c]
                print(f"Spain (Enagas bulletin): {len(new_m)} months {new_m.index.min():%Y-%m} .. {new_m.index.max():%Y-%m}", flush=True)
        except Exception as e:  # noqa: BLE001
            print(f"Spain (Enagas bulletin): FAILED {type(e).__name__}: {e}", flush=True)
    monthly = monthly.sort_index()
    monthly.index.name = "month"
    combined = combined.sort_index().round(2).dropna(how="all")
    combined.index.name = "date"
    if combined.empty and monthly.empty:
        print("no TSO gas demand data")
        return
    print("days per column:", combined.notna().sum().to_dict())
    print((combined.resample("YS").sum() / 1000).round(1).T.to_string())
    if not monthly.empty:
        print((monthly[["ES_national", "ES_conventional", "ES_power"]].resample("YS").sum() / 1000).round(1).T.to_string())
    lines = ["Europe - national gas consumption from further gas operators' own series (raw operator data, no Eurostat gap-filling)", "",
             "Source", "Poland: Gaz-System Market Information Module, 'Actual quantity of gas transmitted' per zone "
             "(https://swi.gaz-system.pl/mir/#/public/bil/ksp-realization; billing data, operative values for the latest days). "
             "Romania: Transgaz 'Physical flows' table, exit points to consumers "
             "(https://www.transgaz.ro/en/clients/operational-data/physical-flows). Finland: Gasgrid Finland 'Gas consumption in Finland' "
             "workbook (https://gasgrid.fi/en/gas-business/transparency-and-market-information/). Croatia: Plinacro public data portal SUKAP, "
             "'Realised physical flows' (https://www.sukap.plinacro.hr/pub/flow). Spain 2021-22: Enagas monthly "
             "statistical bulletin PDFs (https://www.enagas.es/en/technical-management-system/energy-data/publications/gas-statistical-bulletin/); "
             "the Enagas daily demand-history JSON used in europe_tso_gas_demand_daily.xlsx starts only in 2023. Free, no key.",
             "", "Units and definitions",
             "Sheet Daily: GWh per gas day. PL_distribution = exit to DSO networks (H and L gas); PL_final_customers = aggregated final-customer "
             "exit points; PL_other = TSO own needs, nitrogen-removal / mixing plants and aggregated PPG points; PL_total = the three; "
             "PL_dso_return = gas entering the grid back from DSO networks (information only, not netted). Exchange/OTC points, storage, "
             "interconnectors and compulsory stocks are left out. RO_final_customers = final clients connected directly to the transmission "
             "system (SM-CF001), RO_distribution = distribution systems (SM-SD001), RO_total = both (MWh at 15C/15C converted to GWh). "
             "HR_distribution = exits from the Plinacro transmission system to distribution systems, HR_final_customers = exits to final customers "
             "connected to it, HR_total = both (kWh/d GCV 25C/0C from the Plinacro public data portal SUKAP, converted to GWh). "
             "FI_total = Gasgrid's gas consumption in Finland (GCV). These are what the transmission operators deliver to consumers: gas "
             "produced and consumed without entering the transmission grid is not included, so the totals sit about 8-10% (PL, RO) below "
             "Eurostat's gross inland consumption; Finland is 20-30% below Eurostat in 2023-25 (biomethane and distribution-connected gas).",
             "Sheet Monthly: Spain, GWh per month, from the Enagas bulletin table 'Evolution of gas demand' (month column). ES_national = "
             "national market demand = ES_conventional (conventional market, incl. industry) + ES_power (power generation); LNG vessel "
             "loading and international exports are excluded. source_file = the bulletin PDF the month was read from.",
             f"Daily series re-fetch the last {REVISION_DAYS} days each run plus gaps; history from {args.start}. Bulletins already read are not downloaded again.",
             "", "Last pull", f"{datetime.now(timezone.utc):%Y-%m-%d %H:%M} UTC; daily {len(combined)} days"
             + (f" {combined.index.min():%Y-%m-%d} to {combined.index.max():%Y-%m-%d}" if len(combined) else "")
             + (f"; Spain monthly {len(monthly)} months {monthly.index.min():%Y-%m} to {monthly.index.max():%Y-%m}" if len(monthly) else "")]
    sheets = {"Daily": combined}
    if not monthly.empty:
        sheets["Monthly"] = monthly
    xlsx_notes.write_workbook(path, sheets, lines, {"Source", "Units and definitions", "Last pull"})
    print(f"saved {FILE}")


if __name__ == "__main__":
    main()
