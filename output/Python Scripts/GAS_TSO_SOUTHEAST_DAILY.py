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
COLUMNS = ["PL_distribution", "PL_final_customers", "PL_other", "PL_total", "PL_dso_return", "RO_final_customers", "RO_distribution", "RO_total", "FI_total"]


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
             "(https://www.enagas.es/en/technical-management-system/energy-data/demand/history/). Denmark: Energinet Energi Data Service, "
             "dataset Gasflow (https://www.energidataservice.dk). Portugal: REN DataHub (https://datahub.ren.pt). Free, no key.",
             "", "Units and definitions",
             "GWh per gas day. DE_distribution = THE SLP (standard-profile consumers on distribution networks); DE_industry_power = THE RLM "
             "(metered large consumers: industry and gas-fired power, not split further); DE_total = both. FR_industrial = direct industrial "
             "connections, FR_distribution = public distribution (GRD/ELD), FR_power = gas-fired power plants (CCCG), FR_total = the three. "
             "ES_total = Enagás national demand. DK_total = gas delivered to Denmark from the transmission system plus biogas injected "
             "(Energinet's definition). PT_total = REN total consumption (whole GWh), split into conventional market, electricity market "
             "(gas-fired power), distribution (GRMS) and high-pressure clients. Recent days are preliminary and restated.",
             f"Re-fetches the last {REVISION_DAYS} days each run plus gaps; history from {args.start}.",
             "", "Last pull", f"{datetime.now(timezone.utc):%Y-%m-%d %H:%M} UTC; {len(combined)} days, "
             f"{combined.index.min():%Y-%m-%d} to {combined.index.max():%Y-%m-%d}"]
    xlsx_notes.write_workbook(path, {"Daily": combined}, lines, {"Source", "Units and definitions", "Last pull"})
    print(f"saved {FILE}")


if __name__ == "__main__":
    main()
