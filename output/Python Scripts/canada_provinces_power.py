"""
Canada provincial grid-operator pulls (raw sources, no key), one daily workbook with a sheet per province:

  Quebec           Hydro-Quebec open data (donnees.hydroquebec.com + the hydroquebec.com JSON feeds):
                   generation by source (hydraulic, wind, solar, thermal, other) and demand.
                   History table 2019-01-01 to the last annual update (hourly), then the rolling two-day
                   production.json / demande.json feeds (hourly / 15-minute), saved day by day.
  Alberta          AESO Current Supply and Demand report (ets.aeso.ca CSDReportServlet): net generation (TNG) and
                   maximum capability (MC) by generation group, sampled every run and averaged per Alberta day;
                   plus the hourly pool price and Alberta Internal Load (SMPriceReportServlet).
                   AESO's own history files stop in 2017 and its API needs a key, so history starts when this pull
                   first ran.
  British Columbia BC Hydro balancing-authority hourly control-area load (BCHA), yearly .xls files (2024 on) and the
                   current-month file. BC Hydro publishes load only, not generation by fuel.
  Quebec reservoirs daily mean water level (m) at eight large Hydro-Quebec reservoirs, saved day by day from the
                   rolling ~10-day hydrometeorological open-data table.
  New Brunswick    NB Power TSO system information archive (hourly NB load / demand and net scheduled interchange,
                   monthly CSV, 2019 on). Load and interchange only.

Incremental: the committed workbook is the history store. Each run reads it and fetches only what is not saved yet
(Quebec: the history table is pulled once per annual update; New Brunswick: only months after the saved ones plus the
last two; BC Hydro: completed years once, the current year and month every run). Daily energy is mean hourly MW x 24.

Usage: python3 canada_provinces_power.py [--out workbook.xlsx]
"""
import argparse
import io
import os
import re
import sys
from datetime import datetime, timezone
from zoneinfo import ZoneInfo

import pandas as pd
import requests

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)  # repo root, for xlsx_notes
import xlsx_notes  # noqa: E402

DEFAULT_OUT = os.path.join(ROOT, "output", "Data and Chart Outputs", "canada_provincial_power_daily.xlsx")
HEADERS = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) "
                         "Chrome/124.0 Safari/537.36"}
TIMEOUT = (10, 120)
START = "2019-01-01"

QC_TZ, AB_TZ = ZoneInfo("America/Toronto"), ZoneInfo("America/Edmonton")

HQ_ODS = "https://donnees.hydroquebec.com/api/explore/v2.1/catalog/datasets/{ds}"
HQ_JSON = "https://www.hydroquebec.com/data/documents-donnees/donnees-ouvertes/json/{name}.json"
AESO_CSD = "http://ets.aeso.ca/ets_web/ip/Market/Reports/CSDReportServlet"
AESO_PRICE = "http://ets.aeso.ca/ets_web/ip/Market/Reports/SMPriceReportServlet"
BCH_BASE = ("https://www.bchydro.com/content/dam/BCHydro/customer-portal/documents/corporate/suppliers/"
            "transmission-system/balancing_authority_load_data/")
NB_ARCHIVE = "https://tso.nbpower.com/Public/en/system_information_archive.aspx"


def get(url, retries=3, **kw):
    last = None
    for i in range(retries):
        try:
            r = requests.get(url, headers=HEADERS, timeout=TIMEOUT, **kw)
            r.raise_for_status()
            return r
        except requests.RequestException as e:
            last = e
            print(f"    attempt {i + 1}/{retries} failed: {type(e).__name__}: {str(e)[:150]}", flush=True)
    raise last


def load(path, sheet):
    try:
        d = pd.read_excel(path, sheet_name=sheet, index_col=0)
    except (FileNotFoundError, ValueError, KeyError, OSError):
        return pd.DataFrame()
    d.index = pd.to_datetime(pd.Index(d.index).astype(str), errors="coerce")
    return d[d.index.notna()].sort_index()


def upsert(old, new):
    """New rows win; saved rows the new data does not cover are kept."""
    if new is None or new.empty:
        return old
    if old.empty:
        return new.sort_index()
    return new.combine_first(old)[list(new.columns) + [c for c in old.columns if c not in new.columns]].sort_index()


def daily_energy(s, min_points, per_day_points=24):
    """Series of MW readings indexed by local timestamp -> daily MWh (mean MW x 24), days with enough readings."""
    g = s.groupby(s.index.normalize())
    mean, n = g.mean(), g.count()
    return (mean * 24.0).where(n >= min_points)


# ---------------------------------------------------------------- Quebec
HQ_MAP = {"hydraulique": "Hydro_MWh", "eolien": "Wind_MWh", "solaire": "Solar_MWh", "thermique": "Thermal_MWh",
          "autres": "Other_MWh", "total": "Total_MWh"}


def hq_history(since):
    """Hourly generation by source from the ODS history table, as daily MWh (Quebec local day)."""
    ds = HQ_ODS.format(ds="historique-production-electricite-quebec")
    top = get(ds + "/records", params={"select": "max(date) as m", "limit": 1}).json()["results"][0]["m"]
    print(f"  history table runs to {top}", flush=True)
    r = get(ds + "/exports/csv", params={"select": "date," + ",".join(HQ_MAP), "where": f'date >= "{since}"',
                                          "order_by": "date", "delimiter": ",", "timezone": "America/Toronto"})
    d = pd.read_csv(io.StringIO(r.text))
    d["date"] = pd.to_datetime(d["date"].astype(str).str[:19])
    d = d.set_index("date").sort_index()
    d = d[~d.index.duplicated()]
    out = pd.DataFrame({new: daily_energy(d[old], 20) for old, new in HQ_MAP.items() if old in d})
    out.index.name = "date"
    return out.dropna(how="all"), pd.Timestamp(top)


def hq_live():
    """Rolling two-day JSON feeds: generation (hourly) and demand (15-minute) -> daily MWh."""
    prod = get(HQ_JSON.format(name="production")).json()
    rows = [{"date": x["date"], **x["valeurs"]} for x in prod.get("details", [])]
    p = pd.DataFrame(rows)
    out = pd.DataFrame()
    if not p.empty:
        p["date"] = pd.to_datetime(p["date"])
        p = p.set_index("date").apply(pd.to_numeric, errors="coerce").dropna(how="all")
        out = pd.DataFrame({new: daily_energy(p[old], 23) for old, new in HQ_MAP.items() if old in p})
    dem = get(HQ_JSON.format(name="demande")).json()
    q = pd.DataFrame([{"date": x["date"], "v": x["valeurs"].get("demandeTotal")} for x in dem.get("details", [])])
    if not q.empty:
        q["date"] = pd.to_datetime(q["date"])
        q = q.set_index("date")["v"].astype(float).dropna()
        out["Demand_MWh"] = daily_energy(q, 92)
    out.index.name = "date"
    return out.dropna(how="all")


def hq_demand_history(since):
    ds = HQ_ODS.format(ds="historique-demande-electricite-quebec")
    r = get(ds + "/exports/csv", params={"select": "date,moyenne_mw", "where": f'date >= "{since}"', "order_by": "date",
                                         "delimiter": ",", "timezone": "America/Toronto"})
    d = pd.read_csv(io.StringIO(r.text))
    d["date"] = pd.to_datetime(d["date"].astype(str).str[:19])
    s = d.set_index("date")["moyenne_mw"].sort_index()
    return daily_energy(s[~s.index.duplicated()], 20).rename("Demand_MWh").to_frame().dropna()


def pull_quebec(saved, state):
    new = pd.DataFrame()
    # History table: pulled when it has grown past what we hold (it is updated about once a year).
    try:
        ds = HQ_ODS.format(ds="historique-production-electricite-quebec")
        top = pd.Timestamp(get(ds + "/records", params={"select": "max(date) as m", "limit": 1}).json()["results"][0]["m"])
        top = top.tz_convert(QC_TZ).tz_localize(None) if top.tzinfo else top
        if saved.empty or state.get("qc_hist_end") != str(top.date()):
            hist, top = hq_history(START)
            new = upsert(new, hist)
            dem = hq_demand_history(START)
            new = upsert(new, dem)
            state["qc_hist_end"] = str(hist.index.max().date()) if not hist.empty else ""
            print(f"  Quebec history: {len(hist)} days to {state['qc_hist_end']}", flush=True)
        else:
            print(f"  Quebec history unchanged (to {state['qc_hist_end']})", flush=True)
    except Exception as e:  # noqa: BLE001
        print(f"  Quebec history FAILED: {type(e).__name__}: {e}", flush=True)
    try:
        live = hq_live()
        # the rolling feed's last day is partial until the day ends: daily_energy() already needs >= 20 hourly points
        new = upsert(new, live)
        print(f"  Quebec live feed: {len(live)} days, last {live.index.max():%Y-%m-%d}" if not live.empty else
              "  Quebec live feed empty", flush=True)
    except Exception as e:  # noqa: BLE001
        print(f"  Quebec live FAILED: {type(e).__name__}: {e}", flush=True)
    return upsert(saved, new)


# Hydro-Quebec hydrometeorological dataset: hourly water levels (m) at its facilities, a rolling ~10-day window.
RESERVOIRS = {"Caniapiscau": "Caniapiscau Sud (CMCS)", "Gouin": "Gouin Barrage amont",
              "Eastmain": "Eastmain Barrage amont", "Manic-5 (Daniel-Johnson)": "Manic-5 Est",
              "La Grande-4": "La Grande-4 Centrale amont", "La Grande-3": "La Grande-3 Centrale amont",
              "La Grande-2-A (Robert-Bourassa)": "La Grande-2-A Centrale amont", "Bersimis-1": "Bersimis-1 Centrale amont"}


def pull_quebec_levels(saved):
    since = (pd.Timestamp.now(tz="UTC") - pd.Timedelta(days=14)).strftime("%Y-%m-%d")
    ds = HQ_ODS.format(ds="donnees-hydrometeorologiques")
    r = get(ds + "/exports/csv", params={"select": "nom,date,valeur", "delimiter": ",", "timezone": "America/Toronto",
                                         "where": f'composition_depil_type_point_donnee = "Niveau" and date >= "{since}"'})
    d = pd.read_csv(io.StringIO(r.text))
    d["date"] = pd.to_datetime(d["date"].astype(str).str[:19])
    d["valeur"] = pd.to_numeric(d["valeur"], errors="coerce")
    out = {}
    for label, nom in RESERVOIRS.items():
        x = d[d["nom"] == nom].dropna(subset=["valeur"])
        if x.empty:
            continue
        h = x.groupby("date")["valeur"].mean()
        g = h.groupby(h.index.normalize())
        out[f"{label} level_m"] = g.mean().where(g.count() >= 20)
    new = pd.DataFrame(out).dropna(how="all")
    new.index.name = "date"
    print(f"  Quebec reservoir levels: {len(new)} days, {list(new.columns)}", flush=True)
    return upsert(saved, new)


# ---------------------------------------------------------------- Alberta
GROUP_COL = {"COGENERATION": "Cogeneration", "COMBINED CYCLE": "Combined cycle", "GAS FIRED STEAM": "Gas fired steam",
             "SIMPLE CYCLE": "Simple cycle", "HYDRO": "Hydro", "WIND": "Wind", "SOLAR": "Solar",
             "ENERGY STORAGE": "Energy storage", "OTHER": "Other", "COAL": "Coal", "BIOMASS": "Biomass",
             "DUAL FUEL": "Dual fuel"}


def text_of(html):
    t = re.sub(r"<script.*?</script>|<style.*?</style>", " ", html, flags=re.S | re.I)
    return re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", t)).replace("&nbsp;", " ")


def aeso_snapshot():
    """One row from the Current Supply and Demand report: TNG / MC by group, total net generation, AIL, interchange."""
    t = text_of(get(AESO_CSD).text)
    m = re.search(r"Last Update\s*:\s*([A-Za-z]{3} \d{1,2}, \d{4} \d{1,2}:\d{2})", t)
    if not m:
        raise ValueError("no 'Last Update' stamp in the AESO CSD report")
    stamp = datetime.strptime(m.group(1), "%b %d, %Y %H:%M").replace(tzinfo=AB_TZ).astimezone(timezone.utc)
    row = {}
    for key, lab in (("TNG_total", r"Alberta Total Net Generation"), ("Interchange", r"Net Actual Interchange"),
                     ("AIL", r"Alberta Internal Load \(AIL\)"), ("NetToGrid", r"Net-To-Grid Generation")):
        mm = re.search(lab + r"\s+(-?\d+)", t)
        row[key] = float(mm.group(1)) if mm else None
    seg = t[t.find("GENERATION GROUP MC TNG DCR") + len("GENERATION GROUP MC TNG DCR"):t.find("INTERCHANGE PATH")]
    for name, mc, tng, _dcr in re.findall(r"([A-Z][A-Z ]*?[A-Z])\s+(-?\d+)\s+(-?\d+)\s+(-?\d+)", seg):
        name = name.strip()
        if name == "TOTAL":
            continue
        col = GROUP_COL.get(name, name.title())
        row[f"{col} TNG_MW"], row[f"{col} MC_MW"] = float(tng), float(mc)
    if len(row) < 8:
        raise ValueError(f"AESO CSD parse found too little: {row}")
    return pd.Timestamp(stamp).tz_localize(None), row


def aeso_prices():
    """Hourly pool price and AIL demand from the report's recent rows -> DataFrame by Alberta hour-ending timestamp."""
    t = text_of(get(AESO_PRICE).text)
    rows = []
    for d, he, price, avg, ail in re.findall(r"(\d{2}/\d{2}/\d{4})\s+(\d{1,2})\s+(-|-?[\d.,]+)\s+(-|-?[\d.,]+)\s+(-|-?[\d.,]+)", t):
        if price == "-":
            continue
        rows.append({"date": pd.to_datetime(d, format="%m/%d/%Y"), "he": int(he), "price": float(price.replace(",", "")),
                     "ail": float(ail.replace(",", "")) if ail != "-" else None})
    return pd.DataFrame(rows)


def pull_alberta(snaps, hourly):
    try:
        ts, row = aeso_snapshot()
        snaps = upsert(snaps, pd.DataFrame([row], index=pd.DatetimeIndex([ts], name="utc")))
        print(f"  AESO snapshot {ts} UTC: TNG {row.get('TNG_total')} MW, AIL {row.get('AIL')} MW", flush=True)
    except Exception as e:  # noqa: BLE001
        print(f"  AESO snapshot FAILED: {type(e).__name__}: {e}", flush=True)
    try:
        p = aeso_prices()
        if not p.empty:
            p["key"] = p["date"] + pd.to_timedelta(p["he"], unit="h")
            p = p.set_index("key")[["price", "ail"]].rename(columns={"price": "Pool price (CAD per MWh)", "ail": "AIL_MW"})
            p.index.name = "hour_ending"
            hourly = upsert(hourly, p)
            print(f"  AESO price table: {len(p)} hours, last {p.index.max()}", flush=True)
    except Exception as e:  # noqa: BLE001
        print(f"  AESO price FAILED: {type(e).__name__}: {e}", flush=True)
    return snaps, hourly


GAS_GROUPS = ("Cogeneration", "Combined cycle", "Gas fired steam", "Simple cycle")


def alberta_daily(snaps, hourly):
    """Daily table. Generation: the LAST AESO snapshot of each Alberta day, in MW (one instantaneous reading a day, not
    a daily mean - the report has no history and the pull runs once a day). Pool price and AIL: daily means of the hourly
    rows (complete days only)."""
    out = pd.DataFrame()
    if not snaps.empty:
        loc = snaps.copy()
        loc.index = loc.index.tz_localize("UTC").tz_convert(AB_TZ).tz_localize(None)
        last = loc.groupby(loc.index.normalize()).tail(1)
        last.index = last.index.normalize()
        tng = lambda grp: last[[f"{x} TNG_MW" for x in grp if f"{x} TNG_MW" in last]].sum(axis=1, min_count=1)  # noqa: E731
        out = pd.DataFrame({
            "Gas_MW": tng(GAS_GROUPS), "Coal_MW": tng(["Coal"]), "Hydro_MW": tng(["Hydro"]), "Wind_MW": tng(["Wind"]),
            "Solar_MW": tng(["Solar"]), "Storage_MW": tng(["Energy storage"]),
            "Other_MW": tng(["Other", "Biomass", "Dual fuel"]), "Total_MW": last["TNG_total"]})
        if "Coal_MW" in out and (out["Coal_MW"].fillna(0) == 0).all():
            out = out.drop(columns="Coal_MW")
    if not hourly.empty:
        h = hourly.copy()
        day = (h.index - pd.Timedelta(hours=1)).normalize()   # hour ending 24 belongs to the same day
        g = h.groupby(day)
        full = g["AIL_MW"].count() >= 23
        h2 = pd.DataFrame({"Pool price avg (CAD per MWh)": g["Pool price (CAD per MWh)"].mean().where(full),
                           "AIL_MWh": (g["AIL_MW"].mean() * 24).where(full)})
        out = out.join(h2, how="outer") if not out.empty else h2
    out.index.name = "date"
    return out.dropna(how="all").sort_index()


# ---------------------------------------------------------------- British Columbia
def bch_parse(content):
    x = pd.read_excel(io.BytesIO(content), sheet_name=0, header=None)
    hdr = x.index[x[0].astype(str).str.strip().eq("Date")]
    if len(hdr) == 0:
        raise ValueError("no Date header row in BC Hydro file")
    d = x.iloc[hdr[0] + 1:, :3].copy()
    d.columns = ["date", "he", "load"]
    d["date"] = pd.to_datetime(d["date"], format="%m/%d/%Y", errors="coerce")
    d["load"] = pd.to_numeric(d["load"], errors="coerce")
    d = d.dropna(subset=["date", "load"])
    g = d.groupby("date")["load"]
    return (g.sum().where(g.count() >= 23)).rename("Load_MWh").to_frame().dropna()


def pull_bc(saved):
    """Yearly BalancingAuthorityLoad files (completed years once, the current year every run) and the current month."""
    this_year = datetime.now(timezone.utc).year
    have_years = set(saved.index.year) if not saved.empty else set()
    new = pd.DataFrame()
    for y in range(2024, this_year + 1):
        if y < this_year and y in have_years and (saved.index.year == y).sum() >= 360:
            continue
        for path in (f"Historical Transmission Data/BalancingAuthorityLoad {y}.xls", f"BalancingAuthorityLoad {y}.xls"):
            try:
                r = get(BCH_BASE + path, retries=2)
                new = upsert(new, bch_parse(r.content))
                print(f"  BC Hydro {path}: ok", flush=True)
                break
            except Exception as e:  # noqa: BLE001
                print(f"  BC Hydro {path}: {type(e).__name__}: {str(e)[:100]}", flush=True)
    try:
        new = upsert(new, bch_parse(get(BCH_BASE + "CurrentHourlyBALoad.xls").content))
        print("  BC Hydro CurrentHourlyBALoad.xls: ok", flush=True)
    except Exception as e:  # noqa: BLE001
        print(f"  BC Hydro current month FAILED: {type(e).__name__}: {e}", flush=True)
    return upsert(saved, new)


# ---------------------------------------------------------------- New Brunswick
def nb_month(s, base_html, year, month):
    data = {k: (re.search(rf'id="{k}" value="([^"]*)"', base_html) or [None, ""])[1]
            for k in ("__VIEWSTATE", "__VIEWSTATEGENERATOR", "__EVENTVALIDATION")}
    data.update({"__EVENTTARGET": "ctl00$cphMainContent$lbGetData", "__EVENTARGUMENT": "",
                 "ctl00$cphMainContent$ddlMonth": str(month), "ctl00$cphMainContent$ddlYear": str(year)})
    r = s.post(NB_ARCHIVE, data=data, timeout=TIMEOUT)
    r.raise_for_status()
    if "csv" not in r.headers.get("content-type", ""):
        return pd.DataFrame()
    d = pd.read_csv(io.StringIO(r.text))
    d["HOUR"] = pd.to_datetime(d["HOUR"], errors="coerce")
    d = d.dropna(subset=["HOUR"]).set_index("HOUR").apply(pd.to_numeric, errors="coerce")
    d = d[~d.index.duplicated()]
    names = {"NB_LOAD": "NB load", "NB_DEMAND": "NB demand", "ISO_NE": "ISO-NE", "NMISA": "Northern Maine",
             "QUEBEC": "Quebec", "HYDRO_QUEBEC": "Quebec", "NOVA_SCOTIA": "Nova Scotia", "PEI": "PEI"}
    out = pd.DataFrame({f"{names.get(c, c)}_MWh": daily_energy(d[c], 23) for c in d.columns})
    return out.dropna(how="all")


def pull_nb(saved):
    s = requests.Session()
    s.headers.update(HEADERS)
    page = s.get(NB_ARCHIVE, timeout=TIMEOUT)
    page.raise_for_status()
    today = pd.Timestamp.now(tz="UTC").tz_localize(None)
    months = pd.period_range(START, today.to_period("M"), freq="M")
    if not saved.empty:
        last = saved.index.max().to_period("M")
        have = saved.index.to_period("M").value_counts()
        # the last saved month and the one before (revisions), and any earlier month that is short of days (a failed
        # or missing pull) so gaps heal
        months = [m for m in months if m >= last - 1 or have.get(m, 0) < min(27, m.days_in_month - 2)]
    new = pd.DataFrame()
    for m in months:
        try:
            new = upsert(new, nb_month(s, page.text, m.year, m.month))
        except Exception as e:  # noqa: BLE001
            print(f"  NB Power {m}: {type(e).__name__}: {str(e)[:100]}", flush=True)
    print(f"  NB Power: {len(months)} month(s) pulled, {len(new)} days", flush=True)
    return upsert(saved, new)


# ---------------------------------------------------------------- main
NOTES = [
    "UNITS",
    "Daily energy in MWh = mean hourly MW x 24 (a day needs most of its readings: 20 of 24 hours for Quebec generation, "
    "80 of 96 for Quebec demand, 23 of 24 for New Brunswick and BC Hydro). Alberta generation is NOT a daily "
    "mean: it is one instantaneous AESO snapshot per day in MW (the last snapshot of the Alberta day; the earlier "
    "3-hourly averaging was dropped, the pull runs once a day). Pool price: CAD per MWh, daily mean of hourly prices; "
    "AIL_MWh = daily energy from hourly AIL.",
    "Quebec: Hydro_MWh (hydraulique), Wind_MWh (eolien), Solar_MWh, Thermal_MWh, Other_MWh (autres), Total_MWh, "
    "Demand_MWh. Local (Eastern) day.",
    "Alberta: Gas_MW = cogeneration + combined cycle + gas fired steam + simple cycle (AESO net generation, TNG). "
    "AESO's report has no coal group any more (Alberta's coal units are gas-fired or retired). Storage_MW is net "
    "battery output. AIL = Alberta Internal Load (demand). Alberta (Mountain) day.",
    "Quebec reservoirs: daily mean water level in metres above sea level (Hydro-Quebec hydrometeorological dataset; the "
    "open-data table holds only about the last 10 days, so history builds from the first run).",
    "British Columbia: Load_MWh = BC Hydro balancing-authority (BCHA) control-area load, not generation: BC Hydro does "
    "not publish generation by fuel.",
    "New Brunswick: NB load / NB demand are load, not generation. Interchange columns: positive = export from New "
    "Brunswick, negative = import (ISO-NE, Northern Maine, Quebec, Nova Scotia, PEI).",
    "",
    "COVERAGE",
    "Quebec: hourly history table 2019-01-01 to its last annual update (about the end of the previous year), then the "
    "rolling two-day open-data feeds saved day by day from the first run, so the months between the two are missing "
    "until Hydro-Quebec's next annual update fills them. Alberta: from the first run (AESO publishes no hourly "
    "history by fuel after 2017; its API needs a key). British Columbia: 2024 on. New Brunswick: 2019 on.",
    "Not available as structured public data: Saskatchewan (SaskPower), Nova Scotia (NS Power: OASIS reports need "
    "an account), Manitoba Hydro (live hydrological application only), Newfoundland and Labrador.",
    "",
    "SOURCE",
    "Hydro-Quebec open data: https://donnees.hydroquebec.com/explore/dataset/historique-production-electricite-quebec/ , "
    "https://www.hydroquebec.com/data/documents-donnees/donnees-ouvertes/json/production.json",
    "AESO Current Supply and Demand report: http://ets.aeso.ca/ets_web/ip/Market/Reports/CSDReportServlet ; "
    "pool price report: http://ets.aeso.ca/ets_web/ip/Market/Reports/SMPriceReportServlet",
    "BC Hydro balancing authority load data: https://www.bchydro.com/energy-in-bc/operations/transmission/"
    "transmission-system/balancing-authority-load-data.html",
    "NB Power Transmission & System Operator: https://tso.nbpower.com/Public/en/system_information_archive.aspx",
    "",
    "STATE",
]


def main():
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--out", default=DEFAULT_OUT)
    args = ap.parse_args()
    out = args.out
    state = {}
    try:
        st = pd.read_excel(out, sheet_name="Units", header=0)
        for line in st["Notes"].astype(str):
            m = re.match(r"qc_hist_end=(\S*)", line)
            if m:
                state["qc_hist_end"] = m.group(1)
    except Exception:  # noqa: BLE001
        pass

    qc, ab_d = load(out, "Quebec"), load(out, "Alberta")
    snaps, hourly = load(out, "Alberta snapshots"), load(out, "Alberta hourly")
    bc, nb = load(out, "British Columbia"), load(out, "New Brunswick")
    qlev = load(out, "Quebec reservoirs")

    print("Quebec (Hydro-Quebec open data)", flush=True)
    qc = pull_quebec(qc, state)
    try:
        qlev = pull_quebec_levels(qlev)
    except Exception as e:  # noqa: BLE001
        print(f"  Quebec reservoir levels FAILED: {type(e).__name__}: {e}", flush=True)
    print("Alberta (AESO)", flush=True)
    snaps, hourly = pull_alberta(snaps, hourly)
    ab_d = alberta_daily(snaps, hourly) if not snaps.empty or not hourly.empty else ab_d   # rebuilt from the stores
    print("British Columbia (BC Hydro)", flush=True)
    try:
        bc = pull_bc(bc)
    except Exception as e:  # noqa: BLE001
        print(f"  BC FAILED: {type(e).__name__}: {e}", flush=True)
    print("New Brunswick (NB Power)", flush=True)
    try:
        nb = pull_nb(nb)
    except Exception as e:  # noqa: BLE001
        print(f"  NB FAILED: {type(e).__name__}: {e}", flush=True)

    sheets = {"Quebec": qc, "Alberta": ab_d, "British Columbia": bc, "New Brunswick": nb,
              "Quebec reservoirs": qlev, "Alberta snapshots": snaps, "Alberta hourly": hourly}
    sheets = {k: v for k, v in sheets.items() if not v.empty}
    if not sheets:
        raise SystemExit("Nothing pulled")
    for k, v in sheets.items():
        v.index.name = v.index.name or "date"
        print(f"  {k}: {len(v)} rows, {v.index.min()} to {v.index.max()}", flush=True)
    for k in ("Quebec", "Alberta", "British Columbia", "New Brunswick", "Quebec reservoirs"):
        if k in sheets:
            sheets[k].index = pd.to_datetime(sheets[k].index).strftime("%Y-%m-%d")
            sheets[k].index.name = "date"
    for k in ("Alberta snapshots", "Alberta hourly"):
        if k in sheets:
            sheets[k].index = pd.to_datetime(sheets[k].index).strftime("%Y-%m-%d %H:%M")
    notes = NOTES + [f"qc_hist_end={state.get('qc_hist_end', '')}"]
    os.makedirs(os.path.dirname(os.path.abspath(out)), exist_ok=True)
    xlsx_notes.write_workbook(out, sheets, notes, {"UNITS", "COVERAGE", "SOURCE", "STATE"})
    print(f"Saved {out}", flush=True)


if __name__ == "__main__":
    main()
