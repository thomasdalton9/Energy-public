"""
Belgium (monthly), Netherlands (annual) and Luxembourg (annual) installed generating capacity from the national sources.

BELGIUM - Elia open data (opendata.elia.be, ODS v2.1 API, free, no key) -> belgium_power_capacity_monthly.xlsx, month-end MW:
  Wind  ods031 'Wind power production estimation and forecast on Belgian grid (Historical)': monitoredcapacity (MW) of the wind fleet
        Elia monitors, by Offshore/Onshore, region and Elia/DSO connection, summed. COMPLETE fleet (this is the series ENTSO-E reports).
  Solar ods032 'Photovoltaic power production estimation and forecast ...': monitoredcapacity for region 'Belgium' (total fleet, from
        Sep 2020; the series ENTSO-E reports).
  Other ods036 'Actual installed power - by unit and by fuel type' (1 January of 2015-2023; fuel codes NG gas, NU nuclear, LF liquid fuel,
        CP coal, WA water, Other) and, from 21 May 2024, ods179 'Individual data of contracted technical units' (daily, sum of technicalpmax
        by fuel). PARTIAL: units Elia contracts (CIPU, roughly 25 MW and over); nuclear and the large gas units are complete,
        small CHP/biomass/hydro/storage are not. 'Water' is mainly Coo-Trois-Ponts pumped storage and is booked to PumpedStorage_MW;
        run-of-river hydro, oil and the smaller bioenergy have no monthly source (blank, never filled). Months between the two
        conventional sources (Feb 2023 - Apr 2024) are blank.
  Incremental: stored months older than two are kept; months without a value and the last two are queried again.

NETHERLANDS - CBS StatLine (opendata.cbs.nl OData, free): 86266NED 'Elektriciteit en warmte; productie, inzet en vermogen naar energiedrager'
  (electric capacity MW by carrier, end of year, 2015-) and 82610NED 'Hernieuwbare elektriciteit; productie en vermogen' (wind on land /
  at sea, biomass). ANNUAL ONLY - no monthly or quarterly capacity source was found (CBS solar table 85005NED is annual too) ->
  netherlands_power_capacity_annual.xlsx, one row per year dated 1 January holding END-OF-YEAR capacity.

LUXEMBOURG - ILR (Institut luxembourgeois de regulation) 'Donnees statistiques du secteur d'electricite', data.public.lu, sheet '2. Production':
  installed power by source in kW, ANNUAL ONLY 2022- -> luxembourg_power_capacity_annual.xlsx (end of year, dated 1 January of that year).
  The Vianden pumped-storage plant (1.3 GW) is not in ILR's production table.

Cross-checks: end of December vs the 1 January ENTSO-E annual values already in the repo (<country>_power_capacity.xlsx) and, as a reference
only, Energy-Charts yearly capacity (Fraunhofer ISE, FALLBACK - not used in any series).

Usage: python3 BENELUX_CAPACITY.py [--country be|nl|lu|all] [--out-dir DIR]
"""
import argparse
import io
import os
import re
import sys
import time

import pandas as pd
import requests

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.dirname(os.path.dirname(HERE)))
import xlsx_notes  # noqa: E402

OUTDIR = os.path.join("output", "Data and Chart Outputs")
UA = {"User-Agent": "Mozilla/5.0 (research; energy data)"}
ELIA = "https://opendata.elia.be/api/explore/v2.1/catalog/datasets/"
CBS = "https://opendata.cbs.nl/ODataApi/odata/"
EC = "https://api.energy-charts.info/installed_power?country={}&time_step=yearly&installation_decommission=false"
STD = ["Hydro", "Gas", "Wind", "Solar", "Coal", "Nuclear", "Oil", "Bioenergy", "Other"]
EXTRA = ["PumpedStorage_MW", "Storage_MW", "Wind_Onshore_MW", "Wind_Offshore_MW"]
REVISE_MONTHS = 2


def log(*a):
    print(*a, flush=True)


def rget(url, tries=6, **kw):
    for k in range(tries):
        try:
            r = requests.get(url, timeout=(30, 120), headers=UA, **kw)
            if r.status_code < 500 and r.status_code != 429:
                return r
        except Exception as e:  # noqa: BLE001
            log("   retry", k, repr(e)[:90])
        time.sleep(3 * (k + 1))
    raise RuntimeError("giving up on " + url)


def load(path, sheet):
    try:
        d = pd.read_excel(path, sheet_name=sheet, index_col=0)
        d.index = pd.to_datetime(d.index, errors="coerce")
        return d[d.index.notna()]
    except Exception:  # noqa: BLE001
        return pd.DataFrame()


def energy_charts_yearly(code):
    """Energy-Charts yearly installed power (GW) by type: {year: {type: GW}}. Reference only."""
    try:
        d = rget(EC.format(code)).json()
        return {int(t): {s["name"]: s["data"][i] for s in d["production_types"]} for i, t in enumerate(d["time"])}
    except Exception as e:  # noqa: BLE001
        log("Energy-Charts yearly", code, "not read:", repr(e)[:80])
        return {}


def entsoe_crosscheck(monthly, entsoe_path, ours_end_of_prior_period, annual):
    """1 January ENTSO-E value vs our value at the end of the previous month (monthly) / previous year (annual rows dated 1 Jan Y = end of Y)."""
    e = load(entsoe_path, "Monthly")
    rows = []
    for ts in e.index:
        prev = (ts - pd.DateOffset(years=1)) if annual else (ts - pd.DateOffset(months=1))
        if prev not in monthly.index:
            continue
        for f in STD + ["Total", "PumpedStorage", "Storage"]:
            c = f"{f}_MW"
            if c not in e.columns or c not in monthly.columns:
                continue
            o, t = monthly.loc[prev, c], float(e.loc[ts, c])
            if pd.isna(o):
                rows.append({"1 Jan": ts.date(), "Fuel": f, "Ours (MW)": None, "ENTSO-E A68 (MW)": round(t), "Difference (MW)": None, "Difference (%)": None})
                continue
            row = {"1 Jan": ts.date(), "Fuel": f, "Ours (MW)": round(float(o)), "ENTSO-E A68 (MW)": round(t), "Difference (MW)": round(float(o) - t),
                   "Difference (%)": round((float(o) - t) / t * 100, 1) if t else None}
            if annual:   # ENTSO-E's annual value is not always the 1 January stock: also show our end of the SAME year
                so = monthly.loc[ts, c] if ts in monthly.index else None
                row["Ours, end of same year (MW)"] = round(float(so)) if so is not None and pd.notna(so) else None
                row["Difference same year (%)"] = round((float(so) - t) / t * 100, 1) if so is not None and pd.notna(so) and t else None
            rows.append(row)
    return pd.DataFrame(rows)


def finish(monthly):
    for c in [f"{f}_MW" for f in STD] + EXTRA:
        if c not in monthly:
            monthly[c] = pd.NA
    monthly["Total_MW"] = monthly[[f"{f}_MW" for f in STD]].apply(pd.to_numeric, errors="coerce").sum(axis=1, min_count=1)
    cols = ["date"] if False else []
    order = [f"{f}_MW" for f in STD] + ["Total_MW"] + EXTRA + [c for c in monthly.columns if c not in [f"{f}_MW" for f in STD] + ["Total_MW"] + EXTRA]
    monthly = monthly[order]
    monthly.index.name = "date"
    return monthly


# ----------------------------------------------------------------------------------------------------- Belgium
def elia(ds, **p):
    r = rget(ELIA + ds + "/records", params=p)
    if r.status_code != 200:
        raise RuntimeError(f"{ds} {r.status_code} {r.text[:200]}")
    return r.json()["results"]


def with_day(res, known):
    """Elia returns an aliased date_format() under the expression text, not the alias: rename that key to 'd'."""
    out = []
    for r in res:
        k = next((k for k in r if k not in known and k != "d"), None)
        r = dict(r)
        if "d" not in r or r["d"] is None:
            r["d"] = r.pop(k) if k else None
        out.append(r)
    return out


def month_end_window(ts):
    last = (ts + pd.offsets.MonthEnd(0)).normalize()
    return last, last + pd.Timedelta(days=1)


def elia_monitored(ds, month, group):
    """monitoredcapacity at the last day of `month` with data in the last 3 days: dict {group tuple: MW} for the latest day that has rows."""
    last, _ = month_end_window(month)
    lo = last - pd.Timedelta(days=2)
    sel = "date_format(datetime,'yyyy-MM-dd') as d, " + ", ".join(group) + ", max(monitoredcapacity) as mx"
    res = elia(ds, select=sel, where=f"datetime>='{lo:%Y-%m-%d}T00:00:00Z' and datetime<'{last + pd.Timedelta(days=1):%Y-%m-%d}T00:00:00Z'",
               group_by="date_format(datetime,'yyyy-MM-dd'), " + ", ".join(group), limit=100)
    res = with_day(res, set(group) | {"mx"})
    if not res:
        return None, None
    day = max(r["d"] for r in res)
    return day, {tuple(r[g] for g in group): r["mx"] for r in res if r["d"] == day and r["mx"] is not None}


def elia_months_with_data(ds):
    res = elia(ds, select="date_format(datetime,'yyyy-MM') as m, count(*) as n", group_by="date_format(datetime,'yyyy-MM')", limit=100,
               where="datetime>='2020-09-01T00:00:00Z'")
    key = [k for k in res[0] if k != "n"][0] if res else None
    return sorted(pd.Timestamp(r[key] + "-01") for r in res) if key else []


CIPU_MAP = {"Natural Gas": "Gas_MW", "Nuclear": "Nuclear_MW", "Water": "PumpedStorage_MW", "Energy Storage": "Storage_MW", "Biofuels": "Bioenergy_MW",
            "Other": "Other_MW", "Other Fossil Fuels": "Other_MW"}
ODS036_MAP = {"NG": "Gas_MW", "NU": "Nuclear_MW", "WA": "PumpedStorage_MW", "LF": "Oil_MW", "CP": "Coal_MW", "Other": "Other_MW"}


def belgium(outdir):
    path = os.path.join(outdir, "belgium_power_capacity_monthly.xlsx")
    old = load(path, "Monthly")
    log("Belgium: stored", len(old), "months")
    now = pd.Timestamp.now().normalize()
    last_done = (now.replace(day=1) - pd.DateOffset(months=1))              # last complete month
    rows = {}
    log("wind/solar months with data ...")
    wind_m, pv_m = elia_months_with_data("ods031"), elia_months_with_data("ods032")
    keep_until = last_done - pd.DateOffset(months=REVISE_MONTHS - 1)
    def have(ts, col):
        return len(old) and ts in old.index and pd.notna(old.loc[ts].get(col)) and ts < keep_until
    for ts in sorted(set(wind_m) | set(pv_m)):
        if ts > last_done:
            continue
        r = rows.setdefault(ts, {})
        if ts in set(wind_m) and not have(ts, "Wind_MW"):
            day, vals = elia_monitored("ods031", ts, ["offshoreonshore", "region", "gridconnectiontype"])
            if vals:
                r["Wind_Offshore_MW"] = sum(v for k, v in vals.items() if k[0] == "Offshore")
                r["Wind_Onshore_MW"] = sum(v for k, v in vals.items() if k[0] == "Onshore")
                r["Wind_MW"] = r["Wind_Offshore_MW"] + r["Wind_Onshore_MW"]
                r["Wind_day"] = day
        if ts in set(pv_m) and not have(ts, "Solar_MW"):
            day, vals = elia_monitored("ods032", ts, ["region"])
            if vals and ("Belgium",) in vals:
                r["Solar_MW"] = vals[("Belgium",)]
                r["Solar_day"] = day
    # conventional: ods036 (1 January rows)
    for r in elia("ods036", select="date, fuelcode, sum(productionunitnominalpower) as s", group_by="date, fuelcode", order_by="date", limit=100):
        ts = pd.Timestamp(r["date"][:7] + "-01")
        # Elia's date is the start of the year (local 1 Jan = 31 Dec 23:00Z): the row holds the fleet at that date
        if ts.month == 12:
            ts = ts + pd.DateOffset(months=1)
        if r["fuelcode"] in ODS036_MAP and r["s"] is not None and ts.year >= 2015:
            # a 1 January row is the stock at the END of the previous December
            end_month = ts - pd.DateOffset(months=1)
            rows.setdefault(end_month, {})[ODS036_MAP[r["fuelcode"]]] = r["s"]
            rows[end_month]["Conventional_source"] = "ods036 (units, 1 Jan)"
    # conventional: ods179 from May 2024, one request per month (sum of technicalpmax on the last day of the month with data)
    cipu_months = pd.date_range("2024-05-01", last_done, freq="MS")
    for ts in cipu_months:
        if have(ts, "Gas_MW") and ts < keep_until:
            continue
        last, _ = month_end_window(ts)
        res = elia("ods179", select="date_format(date,'yyyy-MM-dd') as d, fueltypepublication, sum(technicalpmax) as s",
                   where=f"date>='{(last - pd.Timedelta(days=2)):%Y-%m-%d}T00:00:00Z' and date<'{last + pd.Timedelta(days=2):%Y-%m-%d}T00:00:00Z'",
                   group_by="date_format(date,'yyyy-MM-dd'), fueltypepublication", limit=100)
        res = with_day(res, {"fueltypepublication", "s"})
        if not res:
            continue
        # Elia's date stamp is 22:00Z/23:00Z the evening before the local day: the day label shifts by one
        days = sorted({r["d"] for r in res})
        target = max(d for d in days if d <= f"{last + pd.Timedelta(days=1):%Y-%m-%d}") if any(d <= f"{last + pd.Timedelta(days=1):%Y-%m-%d}" for d in days) else days[-1]
        r = rows.setdefault(ts, {})
        for x in (y for y in res if y["d"] == target):
            c = CIPU_MAP.get(x["fueltypepublication"])
            if c and x["s"] is not None:
                r[c] = r.get(c, 0.0) + x["s"]
        r["Conventional_source"] = "ods179 (contracted units >= ~25 MW)"
    new = pd.DataFrame.from_dict(rows, orient="index")
    new.index.name = "date"
    if len(old):
        keep = old[[c for c in old.columns if c in new.columns or c == "Total_MW"]].drop(columns=["Total_MW"], errors="ignore")
        merged = new.combine_first(keep) if len(new) else keep
        # fresh values replace stored ones for the months just queried
        for ts in new.index:
            for c in new.columns:
                if pd.notna(new.loc[ts, c]):
                    merged.loc[ts, c] = new.loc[ts, c]
    else:
        merged = new
    merged = merged.sort_index()
    merged = merged[merged.index >= "2015-01-01"]
    for c in merged.columns:
        if c not in ("Wind_day", "Solar_day", "Conventional_source"):
            merged[c] = pd.to_numeric(merged[c], errors="coerce").round(1)
    monthly = finish(merged)
    xc = entsoe_crosscheck(monthly, os.path.join(outdir, "belgium_power_capacity.xlsx"), None, annual=False)
    ec = energy_charts_yearly("be")
    ecrows = [{"Year": y, "Type": k, "Energy-Charts (MW), FALLBACK reference": round(v * 1000)} for y, d in ec.items() for k, v in d.items() if y >= 2021]
    rel = pd.DataFrame([{"Run (UTC)": pd.Timestamp.utcnow().strftime("%Y-%m-%d %H:%M"), "Months stored": int(monthly.shape[0]), "Last month with wind": str(max(wind_m).date()) if wind_m else "",
                         "Last month with solar": str(max(pv_m).date()) if pv_m else ""}]).set_index("Run (UTC)")
    notes = ["UNITS", "MW installed at the END of each month (row date = 1st of the month it closes). Total_MW = sum of the fuel columns that have a value; it is NOT a full national total (see COVERAGE). PumpedStorage_MW and Storage_MW are outside Total_MW, as in the ENTSO-E annual workbook.",
             "", "SOURCE", "Elia open data (opendata.elia.be): ods031 wind monitored capacity, ods032 photovoltaic monitored capacity, ods036 actual installed power by unit and fuel type (1 January 2015-2023), ods179 individual data of contracted technical units (daily from 21 May 2024). Licence: Elia open data (CC BY 4.0).",
             "", "COVERAGE", "COMPLETE fleet: Wind_MW (offshore + onshore), Solar_MW, Nuclear_MW. PARTIAL (units Elia contracts, roughly 25 MW and over, CIPU): Gas_MW, Bioenergy_MW, Other_MW, Storage_MW and PumpedStorage_MW ('Water' = mainly Coo-Trois-Ponts). Run-of-river Hydro_MW and Oil_MW have no monthly source (blank). Conventional months Feb 2023 - Apr 2024 are blank: ods036 stops at 1 Jan 2023 and ods179 starts on 21 May 2024. Column 'Conventional_source' says which dataset a month comes from.",
             "Elia stamps daily data at 22:00/23:00 UTC the evening before the local day; the last day of the month is taken from the last three stamped days.",
             "", "VALIDATION", "'Cross-check' compares end-December with the 1 January ENTSO-E annual values. ENTSO-E's Belgian solar and wind are the same Elia monitored-capacity series. 'Cross-check Energy-Charts' lists Fraunhofer ISE's yearly figures, reference only."]
    sheets = {"Monthly": monthly, "Cross-check": xc.set_index("1 Jan") if len(xc) else xc,
              "Cross-check Energy-Charts": pd.DataFrame(ecrows), "Releases": rel}
    xlsx_notes.write_workbook(path, sheets, notes, {"UNITS", "SOURCE", "COVERAGE", "VALIDATION"})
    log("Belgium wrote", path, "rows", len(monthly), "last", monthly.index.max().date())
    log(xc.to_string() if len(xc) else "no ENTSO-E cross-check")
    return path


# ----------------------------------------------------------------------------------------------------- Netherlands
def cbs_table(tb, flt, select):
    r = rget(CBS + tb + "/TypedDataSet", params={"$format": "json", "$filter": flt, "$select": select, "$top": 5000})
    return r.json()["value"]


def netherlands(outdir):
    path = os.path.join(outdir, "netherlands_power_capacity_annual.xlsx")
    cat = rget("https://opendata.cbs.nl/ODataCatalog/Tables", params={"$format": "json", "$select": "Identifier,Modified,Period",
                                                                     "$filter": "Identifier eq '86266NED' or Identifier eq '82610NED'"}).json()["value"]
    a = cbs_table("86266NED", "CentraleDecentraleProductie eq 'E007022'", "Energiedragers,Perioden,ElektrischVermogenMW_7")
    b = cbs_table("82610NED", "BronTechniek eq 'E006587' or BronTechniek eq 'E006637' or BronTechniek eq 'E006638' or BronTechniek eq 'E006590' or BronTechniek eq 'E006566'",
                  "BronTechniek,Perioden,OpgesteldElektrischVermogenEindeJaar_8")
    A = {}
    for x in a:
        if x["ElektrischVermogenMW_7"] is not None:
            A.setdefault(int(x["Perioden"][:4]), {})[x["Energiedragers"].strip()] = x["ElektrischVermogenMW_7"]
    B = {}
    for x in b:
        if x["OpgesteldElektrischVermogenEindeJaar_8"] is not None:
            B.setdefault(int(x["Perioden"][:4]), {})[x["BronTechniek"].strip()] = x["OpgesteldElektrischVermogenEindeJaar_8"]
    rows = {}
    for y in sorted(A):
        if y < 2015:
            continue
        r = {"Coal_MW": A[y].get("E006461"), "Gas_MW": A[y].get("E006560"), "Hydro_MW": A[y].get("E006587"), "Solar_MW": A[y].get("E006589"),
             "Wind_MW": A[y].get("E006588"), "Nuclear_MW": A[y].get("E006602"), "Storage_MW": A[y].get("E006607"), "Other_MW": A[y].get("E007262"),
             "Wind_Onshore_MW": B.get(y, {}).get("E006637"), "Wind_Offshore_MW": B.get(y, {}).get("E006638"), "Bioenergy_MW": B.get(y, {}).get("E006566"),
             "Solar_MW_82610": B.get(y, {}).get("E006590")}
        rows[pd.Timestamp(year=y, month=1, day=1)] = r      # year-end y, dated 1 January y
    monthly = finish(pd.DataFrame.from_dict(rows, orient="index"))
    monthly["Oil_MW"] = pd.NA
    xc = entsoe_crosscheck(monthly, os.path.join(outdir, "netherlands_power_capacity.xlsx"), None, annual=True)
    ec = energy_charts_yearly("nl")
    ecrows = [{"Year": y, "Type": k, "Energy-Charts (MW), FALLBACK reference": round(v * 1000)} for y, d in ec.items() for k, v in d.items() if y >= 2021]
    rel = pd.DataFrame([{"Table": c["Identifier"], "CBS Modified": c["Modified"], "Periods": c["Period"]} for c in cat])
    notes = ["UNITS", "MW installed at the END of each year; the row dated 1 January Y holds end-of-year Y capacity (so ENTSO-E's 1 January Y+1 value is compared with row Y). ANNUAL ONLY: the Netherlands has no monthly or quarterly capacity source (CBS solar table 85005NED is annual too).",
             "", "SOURCE", "CBS StatLine 86266NED 'Elektriciteit en warmte; productie, inzet en vermogen naar energiedrager' (Elektrisch vermogen MW, central + decentral) and 82610NED 'Hernieuwbare elektriciteit; productie en vermogen' (wind on land / at sea, biomass, end of year). Statistics Netherlands, CC BY 4.0.",
             "", "COVERAGE", "Gas = natural gas, Coal = steenkool, Other = steam and gas-expansion energy, Storage_MW = batteries (from 2025 in CBS), Oil has no CBS line. Bioenergy_MW from 82610NED (the 86266NED biomass line is empty).",
             "", "VALIDATION", "'Cross-check' vs ENTSO-E annual; 'Cross-check Energy-Charts' lists Fraunhofer ISE yearly figures (FALLBACK reference only)."]
    sheets = {"Monthly": monthly, "Cross-check": xc.set_index("1 Jan") if len(xc) else xc, "Cross-check Energy-Charts": pd.DataFrame(ecrows), "Releases": rel}
    xlsx_notes.write_workbook(path, sheets, notes, {"UNITS", "SOURCE", "COVERAGE", "VALIDATION"})
    log("Netherlands wrote", path, len(monthly))
    log(xc.to_string() if len(xc) else "no cross-check")
    return path


# ----------------------------------------------------------------------------------------------------- Luxembourg
ILR_SLUG = "donnees-statistiques-du-secteur-delectricite-ilr"
LU_MAP = {"biogaz": "Bioenergy_MW", "biomasse": "Bioenergy_MW", "cogeneration a gaz naturel": "Gas_MW", "cogeneration": "Gas_MW", "eolienne": "Wind_MW",
          "hydroelectrique": "Hydro_MW", "photovoltaique": "Solar_MW", "usine d'incineration": "Other_MW"}


def strip_acc(s):
    import unicodedata
    return unicodedata.normalize("NFKD", str(s)).encode("ascii", "ignore").decode().lower().strip()


def luxembourg(outdir):
    path = os.path.join(outdir, "luxembourg_power_capacity_annual.xlsx")
    ds = rget("https://data.public.lu/api/1/datasets/" + ILR_SLUG + "/").json()
    res = [r for r in ds["resources"] if str(r.get("format")).lower() == "xlsx"][0]
    import openpyxl
    wb = openpyxl.load_workbook(io.BytesIO(rget(res["url"]).content), read_only=True, data_only=True)
    ws = [w for w in wb.worksheets if strip_acc(w.title).startswith("2. production")][0]
    grid = [list(r) for r in ws.iter_rows(values_only=True)]
    top = next(i for i, r in enumerate(grid) if any("puissance install" in strip_acc(c) for c in r if c))
    hdr = next(i for i in range(top, top + 6) if sum(str(c).strip().isdigit() for c in grid[i] if c is not None) >= 2)
    years = {j: int(str(c).strip()) for j, c in enumerate(grid[hdr]) if c is not None and str(c).strip().isdigit()}
    rows = {}
    for r in grid[hdr + 1:hdr + 14]:
        lab = next((strip_acc(c) for c in r if isinstance(c, str)), "")
        if not lab or lab == "total":
            if lab == "total":
                tot = {years[j]: r[j] for j in years if r[j] is not None}
            continue
        col = LU_MAP.get(lab)
        if col is None:
            continue
        for j, y in years.items():
            if isinstance(r[j], (int, float)):
                rows.setdefault(pd.Timestamp(year=y, month=1, day=1), {})
                rows[pd.Timestamp(year=y, month=1, day=1)][col] = rows[pd.Timestamp(year=y, month=1, day=1)].get(col, 0.0) + r[j] / 1000.0
    monthly = finish(pd.DataFrame.from_dict(rows, orient="index").sort_index())
    ilr_total = {pd.Timestamp(year=y, month=1, day=1): v / 1000.0 for y, v in tot.items()} if "tot" in dir() else {}
    chk = pd.DataFrame({"Sum of sources (MW)": monthly["Total_MW"], "ILR 'Total' row (MW)": pd.Series(ilr_total)}).round(1)
    ec = energy_charts_yearly("lu")
    ecrows = [{"Year": y, "Type": k, "Energy-Charts (MW), FALLBACK reference": round(v * 1000)} for y, d in ec.items() for k, v in d.items() if y >= 2021 and v]
    rel = pd.DataFrame([{"File": res["title"], "Last modified (data.public.lu)": res.get("last_modified"), "URL": res["url"]}])
    notes = ["UNITS", "MW installed at the END of each year (ILR 'puissance installee par source d'energie', kW / 1000); the row dated 1 January Y holds end-of-year Y capacity. ANNUAL ONLY, 2022 onward: no monthly source exists.",
             "", "SOURCE", "ILR, 'Donnees statistiques du secteur d'electricite', data.public.lu (" + ILR_SLUG + "), sheet '2. Production'. Licence CC BY 4.0.",
             "", "COVERAGE", "Hydro, Photovoltaic -> Solar, Wind, Biogas + Biomass -> Bioenergy, natural-gas cogeneration -> Gas, waste incineration -> Other. The Vianden pumped-storage plant (about 1.3 GW) is NOT in ILR's production table; there is no Luxembourg-only ENTSO-E series (Luxembourg is in the DE-LU zone).",
             "", "VALIDATION", "'Check' compares the sum of the mapped sources with ILR's own total row; 'Cross-check Energy-Charts' lists Fraunhofer ISE's yearly Luxembourg figures (FALLBACK reference only)."]
    sheets = {"Monthly": monthly, "Check": chk, "Cross-check Energy-Charts": pd.DataFrame(ecrows), "Releases": rel}
    xlsx_notes.write_workbook(path, sheets, notes, {"UNITS", "SOURCE", "COVERAGE", "VALIDATION"})
    log("Luxembourg wrote", path, len(monthly))
    log(monthly.to_string())
    return path


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--country", default="all")
    ap.add_argument("--out-dir", default=OUTDIR)
    a = ap.parse_args()
    os.makedirs(a.out_dir, exist_ok=True)
    for code, fn in (("be", belgium), ("nl", netherlands), ("lu", luxembourg)):
        if a.country in ("all", code):
            try:
                fn(a.out_dir)
            except Exception as e:  # noqa: BLE001
                import traceback
                traceback.print_exc()
                log(f"{code} FAILED: {e!r}")


if __name__ == "__main__":
    main()
