"""
South Korea power pulls from EPSIS (Electric Power Statistics Information System, https://epsis.kpx.or.kr), the statistics
portal of KPX (Korea Power Exchange, the system and market operator). Keyless; the portal's own grid pages call
POST /epsisnew/<name>.ajax and answer with JavaScript that fills `gridData` (parsed here, never executed). Probes:
discovery_archive/south_korea/KOREA_SOURCES_PROBE*.py, raw samples in discovery_archive/results/south_korea/.

Workbooks written to "output/Data and Chart Outputs/" (history store = the committed workbook; each run keeps the stored
rows and re-reads only a revision window):
  south_korea_power_generation_monthly.xlsx  Data: monthly electricity traded on the KPX market + PPA by fuel (GWh,
                                             'Electricity trading volume by fuel', menu 040501, from Jan 2002), Annual: total
                                             generation incl. self-generators (menu 060101, GWh, from 1961)
  south_korea_power_capacity.xlsx            Monthly: installed capacity by fuel (MW, menu 020100; monthly from Jan 2018,
                                             December of 2012-2017 before that)
  south_korea_smp.xlsx                       Monthly: system marginal price (KRW/kWh, menu 040201, from Jan 2015),
                                             Marginal fuel: hours per month each fuel set the land SMP (040203),
                                             Daily: land SMP by hour, daily max/min/weighted average (040202, from 2021)
  south_korea_demand_daily.xlsx              Daily: installed capacity, supply capability, peak and minimum demand, reserve
                                             (menu 030100, from 2015)
  south_korea_gas_power_use.xlsx             Annual: fuel burned by power generators incl. gas (1,000 t) (menu 060200)

    python3 south_korea/SOUTH_KOREA_EPSIS.py [--out-dir "output/Data and Chart Outputs"] [--only smp,demand,...]
"""
import argparse
import datetime as dt
import os
import re
import sys
import time

import numpy as np
import pandas as pd
import requests

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))  # repo root, for xlsx_notes
import xlsx_notes  # noqa: E402

B = "https://epsis.kpx.or.kr"
UA = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36",
      "Accept-Language": "ko-KR,ko;q=0.9,en;q=0.8", "X-Requested-With": "XMLHttpRequest"}
OUT_DIR = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "output", "Data and Chart Outputs")
PAGES = {
    "ptd": "/epsisnew/selectEkmaPtdBftChart.do?menuId=040501",
    "bft": "/epsisnew/selectEkpoBftChart.do?menuId=020100",
    "smp": "/epsisnew/selectEkmaSmpSmpChart.do?menuId=040201",
    "nsm": "/epsisnew/selectEkmaSmpNsmChart.do?menuId=040203",
    "shd": "/epsisnew/selectEkmaSmpShdChart.do?menuId=040202",
    "mep": "/epsisnew/selectEkgeEpsMepChart.do?menuId=030100",
    "gep": "/epsisnew/selectEkgeGepTotChart.do?menuId=060101",
    "ffu": "/epsisnew/selectEkgeFfuChart.do?menuId=060200",
}
S = requests.Session()
S.headers.update(UA)
REVISION_MONTHS = 3     # monthly series: stored rows older than this are kept, the rest re-read
REVISION_DAYS = 10      # daily series
SHD_START = pd.Timestamp("2021-01-01")
MEP_START = pd.Timestamp("2015-01-01")


# ------------------------------------------------------------------------------------------------------- transport
def _retry(fn, what):
    last = None
    for i in range(4):
        try:
            return fn()
        except requests.RequestException as e:
            last = e
            print(f"  {what}: attempt {i + 1}/4 {type(e).__name__}", file=sys.stderr)
            time.sleep(4 * (i + 1))
    raise last


def page(key):
    """GET the portal page (sets the session cookie, and for the embedded annual tables holds the data itself)."""
    return _retry(lambda: S.get(B + PAGES[key], timeout=(15, 120)), f"GET {key}").text


def ajax(key, name, data):
    def go():
        r = S.post(B + f"/epsisnew/{name}.ajax", data=data, headers={"Referer": B + PAGES[key]}, timeout=(15, 180))
        r.raise_for_status()
        return r.text
    return _retry(go, f"POST {name}")


_ASSIGN = re.compile(r'\b(\w+)\s*=\s*(?:textFormmat\(|\()?"([^"]*)"')
_PUSH = re.compile(r"gridData\.push\(\{(.*?)\}\);", re.S)


def parse_grid(js):
    """The ajax answer is JS: `c1 = textFormmat("12.3",count); ... gridData.push({"Period":year, "c1":c1, ...});`.
    Returns one dict per push (variables resolved to the quoted values assigned before it)."""
    rows, env, pos = [], {}, 0
    for m in _PUSH.finditer(js):
        for a in _ASSIGN.finditer(js[pos:m.start()]):
            env[a.group(1)] = a.group(2)
        row = {}
        for k, v in re.findall(r'"(\w+)"\s*:\s*("[^"]*"|\w+)', m.group(1)):
            row[k] = v[1:-1] if v.startswith('"') else env.get(v)
        rows.append(row)
        pos = m.end()
    return rows


def num(x):
    try:
        return float(x)
    except (TypeError, ValueError):
        return np.nan


def frame(rows, mapping, date_of):
    """rows -> DataFrame indexed by date_of(row), columns renamed by mapping {source key: column}."""
    recs = {}
    for r in rows:
        d = date_of(r)
        if d is None:
            continue
        recs[d] = {col: num(r.get(src)) for src, col in mapping.items()}
    df = pd.DataFrame.from_dict(recs, orient="index", columns=list(mapping.values())).sort_index()
    df.index = pd.to_datetime(df.index)
    return df


def month_of(r):
    p = r.get("Period") or ""
    m = re.match(r"(\d{4})/(\d{2})$", p)
    return f"{m.group(1)}-{m.group(2)}-01" if m else None


def read_old(path, sheet):
    if not os.path.exists(path):
        return None
    try:
        d = pd.read_excel(path, sheet_name=sheet, index_col=0)
        d.index = pd.to_datetime(d.index)
        return d.sort_index()
    except Exception as e:  # noqa: BLE001
        print(f"  stored {os.path.basename(path)}[{sheet}] unreadable ({type(e).__name__}); starting again", file=sys.stderr)
        return None


def merge_monthly(old, new):
    """Keep stored months older than the revision window, take the source's months from there on (and any month the
    store lacks)."""
    if old is None or old.empty:
        return new
    cut = new.index.max() - pd.DateOffset(months=REVISION_MONTHS)
    keep = old[old.index < cut]
    fresh = new[(new.index >= cut) | ~new.index.isin(keep.index)]
    return pd.concat([keep, fresh]).sort_index()


def merge_daily(old, new, since):
    if old is None or old.empty:
        return new
    keep = old[old.index < since]
    return pd.concat([keep, new[new.index >= since]]).sort_index()


def date_range_chunks(start, end, days):
    s = start
    while s <= end:
        e = min(s + pd.Timedelta(days=days - 1), end)
        yield s, e
        s = e + pd.Timedelta(days=1)


def write(path, sheets, notes):
    titles = {n for n in notes if n.isupper() and n}
    xlsx_notes.write_workbook(path, {k: v.rename_axis(v.index.name or "date") for k, v in sheets.items()}, notes, titles)
    print(f"Saved {path}: " + ", ".join(f"{k} {len(v)} rows" for k, v in sheets.items()), flush=True)


# ------------------------------------------------------------------------------------------------------ the pulls
PTD_COLS = {"c1": "Nuclear_GWh", "c2": "Coal_Bituminous_GWh", "c3": "Coal_Anthracite_GWh", "c23": "Coal_GWh",
            "c4": "Oil_GWh", "c5": "Gas_GWh", "c6": "Pumped_Storage_GWh", "c91": "Fuel_Cell_GWh", "c92": "IGCC_GWh",
            "c93": "Solar_GWh", "c94": "Wind_GWh", "c95": "Hydro_GWh", "c96": "Ocean_GWh", "c97": "Bioenergy_GWh",
            "c98": "Waste_GWh", "c912": "Renewables_Total_GWh", "c7": "Other_GWh", "c8": "Market_Total_GWh",
            "c9": "PPA_Total_GWh", "c10": "Total_GWh"}


def annual_blocks(html):
    """The annual tables (generation 060101, fuel use 060200) are embedded in the page as one
    `if("YYYY" >= beginDate ...) { c1 = textFormmat("..."); ...}` block per year."""
    parts = re.split(r'if\("(\d{4})" >= beginDate', html)
    out = {}
    for i in range(1, len(parts), 2):
        out[int(parts[i])] = dict(re.findall(r'\b(c\d+) = textFormmat\("([^"]*)"', parts[i + 1][:6000]))
    return out


GEP_COLS = {"c1": "Hydro_General_MWh", "c2": "Pumped_Storage_MWh", "c3": "Hydro_Small_MWh", "c5": "Steam_Anthracite_MWh",
            "c6": "Steam_Bituminous_MWh", "c7": "Steam_Heavy_Oil_MWh", "c8": "Steam_Gas_MWh", "c24": "Combined_Cycle_LNG_MWh",
            "c25": "Combined_Cycle_Oil_MWh", "c12": "Combined_Cycle_Total_MWh", "c14": "Nuclear_MWh", "c16": "Renewables_MWh",
            "c15": "District_Energy_MWh", "c13": "Internal_Combustion_MWh", "c23": "Other_MWh", "c17": "Total_Generators_MWh",
            "c18": "Self_Gen_Sold_to_KEPCO_MWh", "c19": "Self_Gen_Own_Use_MWh", "c20": "Self_Gen_Total_MWh",
            "c21": "Total_Generators_plus_KEPCO_purchases_MWh", "c22": "Total_Generators_plus_Self_Gen_MWh"}
FFU_COLS = {"c1": "Anthracite_kt", "c3": "Bituminous_kt", "c5": "Heavy_Oil_1000kl", "c7": "Diesel_1000kl", "c9": "Gas_kt",
            "c11": "Fuel_Heat_10e9_kcal"}


def pull_generation(out_dir):
    path = os.path.join(out_dir, "south_korea_power_generation_monthly.xlsx")
    js = ajax("ptd", "selectEkmaPtdBft", {"selYear": "N", "selRegion": "1"})
    new = frame(parse_grid(js), PTD_COLS, month_of)
    if new.empty:
        raise RuntimeError("no monthly rows parsed from selectEkmaPtdBft")
    data = merge_monthly(read_old(path, "Data"), new)
    data.index.name = "month"
    ann = {y: {GEP_COLS[c]: num(v) for c, v in b.items() if c in GEP_COLS} for y, b in annual_blocks(page("gep")).items()}
    annual = pd.DataFrame.from_dict(ann, orient="index").sort_index()
    annual.index = pd.to_datetime(annual.index.astype(str) + "-01-01")
    annual.index.name = "year"
    annual = annual[annual.index >= "1990-01-01"]
    write(path, {"Data": data, "Annual": annual}, [
        "UNITS",
        "Data: GWh per month of electricity traded through the KPX electricity market and PPAs, by fuel (EPSIS menu 'Electricity "
        "trading volume (power market and PPA)', all regions). Months are dated the 1st. Total_GWh = market + PPA total; "
        "Coal_GWh = bituminous + anthracite; Renewables_Total_GWh = fuel cell + IGCC + solar + wind + hydro + ocean + bio + waste "
        "(the grid-connected renewable plants that trade on the market; small behind-the-meter systems and self-generation are "
        "not in it). Pumped_Storage_GWh is the output of pumped-storage plants.",
        "Annual: EPSIS 'Generation' table in MWh per year, all generators plus the commercial self-generators' purchases and own use "
        "(Total_Generators_MWh = business operators only; Total_Generators_plus_Self_Gen_MWh adds the self-generators). Used to "
        "check the monthly market series.",
        "",
        "SOURCE",
        "Korea Power Exchange (KPX), Electric Power Statistics Information System (EPSIS): https://epsis.kpx.or.kr/epsisnew/"
        "selectEkmaPtdBftChart.do?menuId=040501 (monthly) and selectEkgeGepTotChart.do?menuId=060101 (annual). Probe: "
        "discovery_archive/south_korea/KOREA_SOURCES_PROBE*.py.",
        "",
        "UPDATES",
        f"Incremental: the source answers with its whole history in one request; stored months older than {REVISION_MONTHS} months "
        "are kept, the latest months are re-read (EPSIS revises recent months). Latest month is usually the previous calendar month.",
    ])


def pull_capacity(out_dir):
    path = os.path.join(out_dir, "south_korea_power_capacity.xlsx")
    cols = {"c1": "Nuclear_MW", "c2": "Coal_Bituminous_MW", "c3": "Coal_Anthracite_MW", "c4": "Oil_MW", "c5": "Gas_MW",
            "c6": "Pumped_Storage_MW", "c91": "Fuel_Cell_MW", "c92": "IGCC_MW", "c93": "Solar_MW", "c94": "Wind_MW",
            "c95": "Hydro_Renewable_MW", "c96": "Ocean_MW", "c97": "Bio_Plants_MW", "c98": "Waste_MW", "c7": "Other_Source_MW",
            "c8": "Total_MW"}
    js = ajax("bft", "selectEkpoBft", {"selYear": "N", "selRegion": "1", "selMemgubun": "", "selTelgramform": "",
                                       "selBusiType": ""})
    new = frame(parse_grid(js), cols, month_of)
    if new.empty:
        raise RuntimeError("no rows parsed from selectEkpoBft")
    d = merge_monthly(read_old(path, "Monthly"), new)
    d["Hydro_MW"] = d["Hydro_Renewable_MW"] + d["Pumped_Storage_MW"]
    d["Coal_MW"] = d["Coal_Bituminous_MW"] + d["Coal_Anthracite_MW"]
    d["Bioenergy_MW"] = d["Bio_Plants_MW"] + d["Waste_MW"]
    d["Other_MW"] = d["Other_Source_MW"] + d["Fuel_Cell_MW"] + d["IGCC_MW"] + d["Ocean_MW"]
    d.index.name = "date"
    write(path, {"Monthly": d}, [
        "UNITS",
        "Monthly: installed capacity at the end of the month in MW by fuel (EPSIS 'Installed capacity by fuel', all regions, all "
        "members, all dispatch types and business types). Monthly from Jan 2018; before that the source gives December of each "
        "year (2012-2017). Derived columns used by the charts: Hydro_MW = renewable hydro + pumped storage; Coal_MW = bituminous "
        "+ anthracite; Gas_MW = LNG; Bioenergy_MW = bio plants + waste; Other_MW = other + fuel cell + IGCC + ocean. Total_MW is "
        "the source's total.",
        "",
        "SOURCE",
        "Korea Power Exchange (KPX), EPSIS: https://epsis.kpx.or.kr/epsisnew/selectEkpoBftChart.do?menuId=020100",
        "",
        "UPDATES",
        f"Incremental: whole history in one request; stored months older than {REVISION_MONTHS} months kept, the rest re-read.",
    ])


def pull_smp(out_dir):
    path = os.path.join(out_dir, "south_korea_smp.xlsx")
    end_m = dt.date.today().strftime("%Y%m")
    js = ajax("smp", "selectEkmaSmpSmp", {"beginDate": "201501", "endDate": end_m, "selYear": "N"})
    monthly = frame(parse_grid(js), {"c1": "SMP_Land_KRW_per_kWh", "c2": "SMP_Jeju_KRW_per_kWh",
                                     "c3": "SMP_Integrated_KRW_per_kWh", "c4": "BLMP_KRW_per_kWh"}, month_of)
    if monthly.empty:
        raise RuntimeError("no rows parsed from selectEkmaSmpSmp")
    monthly = monthly.mask(monthly == 0)       # 0 = not yet published (BLMP, Jeju before 2015), not a price
    monthly = merge_monthly(read_old(path, "Monthly"), monthly)
    monthly.index.name = "month"
    js = ajax("nsm", "selectEkmaSmpNsm", {"beginDate": "201501", "endDate": end_m, "selYear": "N"})
    marg = frame(parse_grid(js), {"c1": "LNG_hours", "c2": "Oil_hours", "c3": "Anthracite_hours", "c4": "Bituminous_hours",
                                  "c5": "Nuclear_hours", "c6": "None_hours", "c7": "Total_hours"}, month_of)
    marg = merge_monthly(read_old(path, "Marginal fuel"), marg)
    marg.index.name = "month"
    # daily land SMP by hour
    old = read_old(path, "Daily")
    today = pd.Timestamp(dt.date.today())
    since = SHD_START if old is None or old.empty else max(SHD_START, old.index.max() - pd.Timedelta(days=REVISION_DAYS))
    hrs = {f"c{i}": f"H{i:02d}" for i in range(1, 25)}
    hrs.update({"c25": "Max", "c26": "Min", "c27": "Weighted_avg"})
    parts = []
    for s, e in date_range_chunks(since, today - pd.Timedelta(days=1), 100):
        js = ajax("shd", "selectEkmaSmpShd", {"beginDate": s.strftime("%Y%m%d"), "endDate": e.strftime("%Y%m%d"),
                                              "selYear": "N", "selMonth": "N", "selKind": "land", "locale": ""})
        rows = parse_grid(js)
        if rows and not parts:
            print("  first SMP daily row:", rows[0], flush=True)

        def dkey(r):
            d = r.get("Date") or ""
            m = re.match(r"(\d{4})[/-](\d{2})[/-](\d{2})$", d)
            return f"{m.group(1)}-{m.group(2)}-{m.group(3)}" if m else None
        parts.append(frame(rows, hrs, dkey))
    new = pd.concat(parts).sort_index() if parts else pd.DataFrame(columns=list(hrs.values()))
    new = new[~new.index.duplicated(keep="last")]
    daily = merge_daily(old, new, since) if len(new) or old is not None else new
    daily = daily[daily.index < today]
    daily.index.name = "date"
    write(path, {"Monthly": monthly, "Marginal fuel": marg, "Daily": daily}, [
        "UNITS",
        "Monthly: system marginal price (SMP) in KRW per kWh (EPSIS 'SMP'): mainland (land), Jeju, integrated; BLMP is the "
        "boundary load marginal price (blank until published). A 0 in the source is stored as blank.",
        "Marginal fuel: hours per month in which each fuel set the mainland SMP (EPSIS 'SMP setting fuel'); None = no marginal unit.",
        "Daily: mainland SMP by hour (H01 = 00:00-01:00 ... H24), KRW per kWh, and the day's maximum, minimum and weighted average "
        "(EPSIS 'SMP by hour', land). From 2021. The current day is not stored.",
        "The master converts KRW/kWh to US$/MWh with the Federal Reserve H.10 won per US$ rate (south_korea_fx_usd_daily.xlsx).",
        "",
        "SOURCE",
        "Korea Power Exchange (KPX), EPSIS: https://epsis.kpx.or.kr/epsisnew/selectEkmaSmpSmpChart.do?menuId=040201 (monthly), "
        "selectEkmaSmpNsmChart.do?menuId=040203 (marginal fuel), selectEkmaSmpShdChart.do?menuId=040202 (hourly).",
        "",
        "UPDATES",
        f"Monthly sheets: stored months older than {REVISION_MONTHS} months kept, the rest re-read. Daily: only days after the "
        f"last stored day minus {REVISION_DAYS} days are requested (first run backfills from 2021-01-01).",
    ])


def pull_demand(out_dir):
    path = os.path.join(out_dir, "south_korea_demand_daily.xlsx")
    old = read_old(path, "Daily")
    today = pd.Timestamp(dt.date.today())
    since = MEP_START if old is None or old.empty else max(MEP_START, old.index.max() - pd.Timedelta(days=REVISION_DAYS))
    cols = {"c1": "Installed_Capacity_MW", "c3": "Supply_Capability_MW", "c4": "Peak_Demand_MW", "c8": "Min_Demand_MW",
            "c5": "Supply_Reserve_MW", "c6": "Supply_Reserve_pct"}
    parts = []
    for s, e in date_range_chunks(since, today - pd.Timedelta(days=1), 366):
        js = ajax("mep", "selectEkgeEpsMep", {"beginDate": s.strftime("%Y%m%d"), "endDate": e.strftime("%Y%m%d"),
                                              "selYear": "N", "selMonth": "N"})
        rows = parse_grid(js)
        parts.append(frame(rows, cols, lambda r: f"{r['year']}-{r['month']}-{r['day']}" if r.get("day") else None))
    new = pd.concat(parts).sort_index() if parts else pd.DataFrame(columns=list(cols.values()))
    new = new[~new.index.duplicated(keep="last")]
    new["Min_Demand_MW"] = new["Min_Demand_MW"].mask(new["Min_Demand_MW"] == 0)   # 0 = not recorded (early years)
    daily = merge_daily(old, new, since)
    daily = daily[daily.index < today]
    daily.index.name = "date"
    write(path, {"Daily": daily}, [
        "UNITS",
        "Daily: installed capacity and supply capability in MW, the day's peak and minimum demand in MW, supply reserve in MW and "
        "% (EPSIS 'Power supply and demand results', daily). Peak demand is the highest hourly demand of the day on the KPX "
        "system; a minimum of 0 in the source (early years) is stored as blank. The current day is not stored.",
        "",
        "SOURCE",
        "Korea Power Exchange (KPX), EPSIS: https://epsis.kpx.or.kr/epsisnew/selectEkgeEpsMepChart.do?menuId=030100",
        "",
        "UPDATES",
        f"Incremental: only days after the last stored day minus {REVISION_DAYS} days are requested (first run from 2015-01-01).",
    ])


def pull_gas_use(out_dir):
    path = os.path.join(out_dir, "south_korea_gas_power_use.xlsx")
    ann = {y: {FFU_COLS[c]: num(v) for c, v in b.items() if c in FFU_COLS} for y, b in annual_blocks(page("ffu")).items()}
    d = pd.DataFrame.from_dict(ann, orient="index").sort_index()
    d.index = pd.to_datetime(d.index.astype(str) + "-01-01")
    d = d[d.index >= "1990-01-01"].mask(lambda x: x == 0)
    d.index.name = "year"
    old = read_old(path, "Annual")
    if old is not None:      # annual table: keep stored years, take the source's latest two
        cut = d.index.max() - pd.DateOffset(years=1)
        d = pd.concat([old[old.index < cut], d[(d.index >= cut) | ~d.index.isin(old.index)]]).sort_index()
        d.index.name = "year"
    write(path, {"Annual": d}, [
        "UNITS",
        "Annual: fuel burned by electricity generators (EPSIS 'Generation fuel consumption'): anthracite and bituminous coal in "
        "1,000 tonnes, heavy oil and diesel in 1,000 kilolitres, gas (LNG) in 1,000 tonnes, total fuel heat in 10^9 kcal. A 0 in "
        "the source is stored as blank (fuel not used). Business generators only; no volume conversion of LNG is made.",
        "",
        "SOURCE",
        "Korea Power Exchange (KPX), EPSIS: https://epsis.kpx.or.kr/epsisnew/selectEkgeFfuChart.do?menuId=060200",
        "",
        "UPDATES",
        "Annual table embedded in the page; stored years kept, the latest two re-read.",
    ])


PULLS = {"generation": pull_generation, "capacity": pull_capacity, "smp": pull_smp, "demand": pull_demand,
         "gas": pull_gas_use}


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--out-dir", default=OUT_DIR)
    ap.add_argument("--only", default="", help="comma list of: " + ",".join(PULLS))
    args = ap.parse_args()
    os.makedirs(args.out_dir, exist_ok=True)
    want = [k for k in PULLS if not args.only or k in args.only.split(",")]
    failed = []
    for k in want:
        try:
            PULLS[k](args.out_dir)
        except Exception as e:  # noqa: BLE001
            failed.append(k)
            print(f"{k} FAILED: {type(e).__name__}: {e}", file=sys.stderr, flush=True)
    if failed:
        sys.exit(f"failed: {failed}")


if __name__ == "__main__":
    main()
