"""
Americas installed capacity, ONE gap-free monthly dataset per country and fuel (Jan 2021 - latest month, MW), every filled cell
flagged.   output/Data and Chart Outputs/south_central_america_capacity_monthly_complete.xlsx
           output/Data and Chart Outputs/north_america_capacity_monthly_complete.xlsx

No downloads: it reads the capacity workbooks the other pulls have already committed (`<country>_power_capacity.xlsx`) and the Ember
yearly capacity workbook (`ember_capacity_yearly.xlsx`, EMBER_CAPACITY_YEARLY.py, one whole-file download) and combines them with the
generic rules in the repo-root `capacity_fill.py` (docstring: fill rules and flags). This file is only the Americas CONFIG (which
source for which country and fuel, and why) and the output; the layout is Europe's (europe/capacity/EUROPE_CAPACITY_COMPLETE.py).

Regions (the two masters' country sets)
  South & Central America and the Caribbean (SOUTH_AMERICA_MASTER):
    monthly national sources: Argentina (CAMMESA), Brazil (ANEEL), Chile (CNE), Colombia (XM), Ecuador (ARCONEL: annual December
      stocks 2021-23, monthly from Apr 2024, May 2025 missing);
    annual national sources: Bolivia (CNDC), Peru (COES), Uruguay (MIEM, 2026 = the half-year value), Paraguay (Itaipu and Yacyreta
      halves + Ember for ANDE);
    no capacity workbook: Venezuela, Guyana, Suriname, Guatemala, El Salvador, Honduras, Nicaragua, Costa Rica, Panama, Belize,
      Trinidad & Tobago, Puerto Rico, Jamaica, Dominican Republic - Ember yearly capacity, every cell FALLBACK_EMBER.
  North America (NORTH_AMERICA_MASTER): United States (EIA-860M monthly, plus Lower 48 and the nine ISO / utility regions as own sheets),
    Canada (Statistics Canada year-end 2019-24, annual), Mexico (Ember yearly, FALLBACK_EMBER: no raw Mexican capacity feed yet).

Date convention (capacity_fill.py): a row dated the 1st is the stock at the END of the previous month. The national monthly workbooks
date a row by the month it closes ('date' = 1st of the month, value = end of that month), so they are shifted one month on reading; an
annual row dated 1 January Y that holds the 31 December Y stock is shifted twelve months; Ember year Y = stock at the end of Y =
1 January Y+1. Where a monthly source starts in Jan 2021 its first row therefore lands on 1 Feb 2021 and the Jan 2021 cell is HELD_BACK
(the first stock held one month back; Argentina instead has CAMMESA's own December 2020 stock as ANCHOR_A). Annual national sources that
start with the Dec 2021 stock are back-spliced for 2021 from Ember (SPLICED, ratio at the first 1 January both hold, only within
0.7-1.3, else SPLICED_UNSCALED and listed on 'Breaks').

Sheets (per workbook): Units, Basis, Breaks, Coverage, Coverage by month, '<Region> total', one sheet per country (date, Total_MW =
all fuels except PumpedStorage and Storage, <Fuel>_MW, memo) and '<country> flags'.
Usage: python3 AMERICAS_CAPACITY_COMPLETE.py [--region sa|na|all] [--data-dir DIR] [--end YYYY-MM] [--no-charts]
"""
import argparse
import datetime
import os
import sys

import numpy as np
import pandas as pd

ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
sys.path.insert(0, ROOT)
import capacity_fill as cf  # noqa: E402
import xlsx_notes  # noqa: E402

DATA_DIR = os.path.join(ROOT, "output", "Data and Chart Outputs")
SA_OUT = "south_central_america_capacity_monthly_complete.xlsx"
NA_OUT = "north_america_capacity_monthly_complete.xlsx"
EMBER_FILE = "ember_capacity_yearly.xlsx"
START = "2021-01-01"
FUELS = cf.FUEL_ORDER
SA_TOTAL = "South & Central America total"
NA_TOTAL = "North America total"
EMBER_SRC = "Ember yearly electricity data, installed capacity by fuel (year-end stock)"
EMBER_URL = "https://ember-energy.org/data/yearly-electricity-data/"

_cache = {}
ZERO_DROPPED = []      # (country, date, fuel or 'all fuels') Ember zeros treated as 'not reported'


# ------------------------------------------------------------------ readers
def sheet(data_dir, file, name, shift=0):
    """A workbook sheet indexed by its first column (a date), shifted by `shift` months and snapped to month starts."""
    key = (data_dir, file, name, shift)
    if key not in _cache:
        p = os.path.join(data_dir, file)
        d = None
        if os.path.exists(p):
            try:
                d = pd.read_excel(p, sheet_name=name)
                dc = d.columns[0]
                d[dc] = pd.to_datetime(d[dc], errors="coerce")
                d = d.dropna(subset=[dc]).set_index(dc)
                d.index = (d.index + pd.DateOffset(months=shift)).to_period("M").to_timestamp()
            except Exception as e:  # noqa: BLE001
                print(f"  could not read {file}/{name}: {type(e).__name__}: {e}")
                d = None
        _cache[key] = d
    return _cache[key]


def col(d, c, res, kind=None, source="", url="", note="", native=None):
    if d is None or c not in d:
        return None
    return cf.Series(pd.to_numeric(d[c], errors="coerce"), res, kind, source, url, note, native)


_FUEL_COLS = [f"{f}_MW" for f in ("Hydro", "Gas", "Wind", "Solar", "Coal", "Nuclear", "Oil", "Bioenergy", "Other")]


def ember_clean(data_dir, iso3, label):
    """Ember yearly rows of one country, indexed by 1 Jan Y+1 (year-end stock), with not-reported values removed: a whole year whose
    fuels add to under 20% of the country's median year, or a fuel at 0 between two years above 0, is dropped so the series
    interpolates / holds across it (the same rule as the ENTSO-E zeros in the Europe build)."""
    key = ("ember", data_dir, iso3)
    if key in _cache:
        return _cache[key]
    p = os.path.join(data_dir, EMBER_FILE)
    out = None
    if os.path.exists(p):
        if ("emberall", data_dir) not in _cache:
            _cache[("emberall", data_dir)] = pd.read_excel(p, sheet_name="Yearly")
        d = _cache[("emberall", data_dir)]
        d = d[d["iso3"] == iso3].copy()
        if not d.empty:
            d["date_1_Jan"] = pd.to_datetime(d["date_1_Jan"])
            d = d.set_index("date_1_Jan").sort_index()[_FUEL_COLS].apply(pd.to_numeric, errors="coerce")
            tot = d.sum(axis=1)
            med = tot[tot > 0].median() if (tot > 0).any() else 0
            for i in d.index[(tot < 0.2 * med).values]:
                d.loc[i, :] = np.nan
                ZERO_DROPPED.append((label, i, "all fuels"))
            for c in d.columns:
                s = d[c]
                for j, i in enumerate(s.index):
                    if s[i] == 0 and 0 < j < len(s) - 1 and s.iloc[j - 1] > 0 and s.iloc[j + 1] > 0:
                        d.loc[i, c] = np.nan
                        ZERO_DROPPED.append((label, i, c[:-3]))
            out = d
    _cache[key] = out
    return out


def ember(data_dir, iso3, fuel, label, res="A"):
    """res 'A' = splice reference; res 'EMBER' = the primary of a country with no national source (every cell FALLBACK_EMBER).
    A fuel Ember does not list for the country is None (no source)."""
    d = ember_clean(data_dir, iso3, label)
    s = col(d, f"{fuel}_MW", res, None, EMBER_SRC, EMBER_URL)
    return s if s is not None and not s.clean().empty else None


# ------------------------------------------------------------------ the national sources
# file, sheet, shift (months; 1 = row dated by the month it closes, 12 = annual row dated 1 Jan Y holding the 31 Dec Y stock), res, kind
SA_NAT = {
    "Argentina": dict(iso="ARG", file="argentina_power_capacity.xlsx", shift=1, res="M", kind="REBUILT", lead="anchor",
                      source="CAMMESA Potencia Instalada: month-end machine list rebuilt by enabling / retirement date, Decembers = CAMMESA's published stock",
                      url="https://cammesaweb.cammesa.com/download/potencia-instalada/"),
    "Brazil": dict(iso="BRA", file="brazil_power_capacity.xlsx", shift=1, res="M", kind="REBUILT", lead="hold",
                   source="ANEEL SIGA plant register + MMGD distributed generation, month-end rebuild",
                   url="https://dadosabertos.aneel.gov.br/"),
    "Chile": dict(iso="CHL", file="chile_power_capacity.xlsx", shift=1, res="M", kind="REBUILT", lead="hold",
                  source="CNE Capacidad Instalada de Generacion (net capacity in service), month-end rebuild",
                  url="https://www.cne.cl/estadisticas/electricidad/"),
    "Colombia": dict(iso="COL", file="colombia_power_capacity.xlsx", shift=1, res="M", kind="RAW", lead="hold",
                     source="XM net effective capacity (CapEfecNeta) at month end, SIN plants",
                     url="https://www.xm.com.co/"),
    "Ecuador": dict(iso="ECU", file="ecuador_power_capacity.xlsx", shift=1, res="M", kind="RAW", lead=None,
                    anchors=["2021-12-01", "2022-12-01", "2023-12-01"],
                    source="ARCONEL Balance Nacional de Energia Electrica (nominal capacity): December stocks 2021-23 (annual report), monthly from Apr 2024",
                    url="https://arconel.gob.ec/balance-nacional-de-energia-electrica/"),
    "Bolivia": dict(iso="BOL", file="bolivia_power_capacity.xlsx", shift=12, res="A", kind=None, lead=None,
                    source="CNDC Bolivia, capacity by technology at December (SIN)", url="https://www.cndc.bo/"),
    "Peru": dict(iso="PER", file="peru_power_capacity.xlsx", shift=12, res="A", kind=None, lead=None,
                 source="COES annual statistics, effective capacity at 31 December (SEIN)",
                 url="https://www.coes.org.pe/Portal/publicaciones/estadisticas/"),
    "Uruguay": dict(iso="URY", file="uruguay_power_capacity.xlsx", shift=12, res="A", kind=None, lead=None, half_year_last=True,
                    source="MIEM / DNE Uruguay, potencia instalada (31 December; 2026 = latest half-year value)",
                    url="https://www.gub.uy/ministerio-industria-energia-mineria/datos-y-estadisticas/datos/series-estadisticas-energia-electrica"),
    "Paraguay": dict(iso="PRY", file="paraguay_power_capacity.xlsx", shift=12, res="A", kind=None, lead=None,
                     source="Itaipu and Yacyreta (EBY) Paraguayan halves + ANDE's own plants from Ember (annual)",
                     url="https://www.eby.gov.py/datos-tecnicos/"),
}
SA_EMBER_ONLY = [("Venezuela", "VEN"), ("Guyana", "GUY"), ("Suriname", "SUR"), ("Belize", "BLZ"), ("Costa Rica", "CRI"),
                 ("El Salvador", "SLV"), ("Guatemala", "GTM"), ("Honduras", "HND"), ("Nicaragua", "NIC"), ("Panama", "PAN"),
                 ("Dominican Republic", "DOM"), ("Jamaica", "JAM"), ("Puerto Rico", "PRI"), ("Trinidad & Tobago", "TTO")]
SA_ORDER = ["Argentina", "Bolivia", "Brazil", "Chile", "Colombia", "Ecuador", "Paraguay", "Peru", "Uruguay", "Venezuela", "Guyana",
            "Suriname", "Belize", "Costa Rica", "El Salvador", "Guatemala", "Honduras", "Nicaragua", "Panama",
            "Dominican Republic", "Jamaica", "Puerto Rico", "Trinidad & Tobago"]

US_SRC = ("EIA-860M monthly generator inventory (net summer capacity, operating + standby + out-of-service), EIA API v2",
          "https://www.eia.gov/electricity/data/eia860m/")
CAN_SRC = ("Statistics Canada tables 25-10-0022-01 (capacity by type) and 25-10-0023-01 (thermal by principal fuel), year-end, all producers",
           "https://www150.statcan.gov.ca/t1/tbl1/en/tv.action?pid=2510002201")
US_REGIONS = [("United States", "Monthly"), ("US Lower 48", "US_Total"), ("US ERCOT", "ERCOT"), ("US PJM", "PJM"), ("US MISO", "MISO"),
              ("US SPP", "SPP"), ("US CAISO", "CAISO"), ("US NYISO", "NYISO"), ("US ISONE", "ISONE"), ("US Southern", "Southern"),
              ("US TVA", "TVA")]
US_COLMAP = {"PumpedStorage": "Pumped_storage_MW", "Storage": "Battery_storage_MW"}
NA_COUNTRIES = ["United States", "Canada", "Mexico"]        # the regional total
NA_ORDER = NA_COUNTRIES + [n for n, _ in US_REGIONS[1:]]

NOTES = {
    ("Argentina", "Hydro"): "CAMMESA large + small hydro; includes the Rio Grande and Los Reyunos pumped-storage plants (no separate PumpedStorage).",
    ("Argentina", "Gas"): "CAMMESA publishes thermal capacity by machine type, not fuel: combined cycles + gas turbines + steam turbines (most can burn gasoil or fuel oil).",
    ("Argentina", "Oil"): "CAMMESA 'Motor Diesel' engines.",
    ("Argentina", "Coal"): "0 by the source's mapping (Rio Turbio is not in CAMMESA's list; coal-capable steam units are in Gas).",
    ("Brazil", "Solar"): "Includes ~54 GW of distributed (MMGD) solar declared to ANEEL.",
    ("Brazil", "Oil"): "ANEEL register rebuilt from today's SIGA plus a top-up for retired thermal plant (taper to the latest month); neither register has retirement dates.",
    ("Brazil", "Hydro"): "ANEEL registers pumped-storage-capable plants as hydro; no separate PumpedStorage.",
    ("Chile", "Coal"): "CNE snapshot rebuild; months before Jan 2026 carry the part of Ember's change since end-2025 that CNE's start dates do not explain (retirements).",
    ("Chile", "Storage"): "The CNE capacity table does not list battery storage; no source in this build.",
    ("Colombia", "Solar"): "XM net effective capacity; behind-the-meter rooftop solar without an XM record is not included.",
    ("Colombia", "Wind"): "XM lists a single small wind plant (Jemeton); the other projects are not yet in the metric.",
    ("Ecuador", "Gas"): "ARCONEL technology 'Turbogas' (a technology proxy: includes diesel-fired turbines).",
    ("Ecuador", "Oil"): "ARCONEL MCI engines + Turbovapor (fuel-oil steam); some engines burn associated gas.",
    ("Ecuador", "Solar"): "ARCONEL 'Fotovoltaica' (Ember shows 4x this for 2025: ARCONEL's table lists large plants).",
    ("Bolivia", "Bioenergy"): "CNDC 'Biomasa' (sugar-mill bagasse plants); Ember shows about twice this.",
    ("Peru", "Bioenergy"): "COES bagasse + biogas in the SEIN; Ember shows about twice this.",
    ("Uruguay", "Gas"): "MIEM lists all fossil plant as 'Total Fosil' (shown as Oil); Ember counts the gas turbines / combined cycle as Gas. The fossil total matches.",
    ("Uruguay", "Oil"): "All MIEM 'Total Fosil' (steam, gas turbines, engines): the fossil total matches Ember's Gas + Other Fossil.",
    ("Paraguay", "Hydro"): "Paraguay's 50% of Itaipu (7,000 MW) and Yacyreta (1,600 MW) + ANDE's Acaray (Ember residual).",
    ("Canada", "Other"): "StatCan 'other principal fuels' (biomass and other combustibles, not split) + geothermal / other non-combustible; bioenergy is inside it.",
    ("Canada", "Gas"): "StatCan conventional thermal capacity by principal fuel (natural gas).",
    ("Canada", "Coal"): "StatCan conventional thermal (coal): falls with the Ontario / Alberta phase-outs.",
    ("United States", "Hydro"): "Conventional hydro only; pumped storage is its own column.",
    ("United States", "Other"): "Geothermal, other gases, waste heat, hydrogen and others in EIA-860M.",
    ("Mexico", "Oil"): "Ember 'Other Fossil' (mostly oil-fired plant, incl. fuel-oil steam).",
    ("Mexico", "Other"): "Ember lists no 'Other Renewables' for Mexico.",
}
for _n in dict(SA_EMBER_ONLY):
    NOTES[(_n, "Oil")] = "Ember 'Other Fossil' (oil-fired plant, mostly diesel / fuel oil)."
COMMON_NOTE = {"Storage": "batteries; excluded from Total", "PumpedStorage": "excluded from Total"}


# ------------------------------------------------------------------ builders
def _fuel_map(df_cols, fuel, colmap=None):
    return (colmap or {}).get(fuel, f"{fuel}_MW")


def argentina_dec2020(data_dir):
    """CAMMESA's own December 2020 stock by our fuels (year-end table, MW) as {fuel: MW}, or {}. Coal and Other are 0 by the monthly
    workbook's mapping (no CAMMESA machine type maps to them)."""
    p = os.path.join(data_dir, "argentina_power_capacity.xlsx")
    if not os.path.exists(p):
        return {}
    raw = pd.read_excel(p, sheet_name="CAMMESA annual by type")
    raw = raw.set_index(raw.columns[0])
    if 2020 not in raw.index:
        return {}
    r = raw.loc[2020]

    def g(*names):
        return float(sum(pd.to_numeric(r.get(n), errors="coerce") for n in names if pd.notna(pd.to_numeric(r.get(n), errors="coerce"))))
    return {"Hydro": g("Hidráulica > 50 MW", "Hidráulica < 50 MW"), "Gas": g("Ciclos Combinados", "Turbina a gas", "Turbovapor"),
            "Nuclear": g("Nuclear"), "Oil": g("Motor Diesel"), "Wind": g("Eólica"), "Solar": g("Solar"),
            "Bioenergy": g("Biogas", "Biomasa"), "Coal": 0.0, "Other": 0.0}


def national_series(data_dir, name, spec, fuel):
    """The national Series for one fuel (None if the workbook lacks the column)."""
    d = sheet(data_dir, spec["file"], "Monthly", spec["shift"])
    if d is None:
        return None
    c = _fuel_map(d.columns, fuel, spec.get("colmap"))
    if c not in d:
        return None
    s = pd.to_numeric(d[c], errors="coerce").copy()
    native = {}
    if spec.get("half_year_last") and len(s):
        # the last annual row is MIEM's latest HALF-YEAR value (30 June): date it 1 July of that year
        last = s.index[-1]
        s = s.rename(index={last: pd.Timestamp(year=last.year - 1, month=7, day=1)})
    if name == "Ecuador":
        native = {pd.Timestamp(a) + pd.DateOffset(months=1): "ANCHOR_A" for a in spec["anchors"]}
    if name == "Argentina":
        a = argentina_dec2020(data_dir).get(fuel)
        if a is not None:
            s = pd.concat([pd.Series([a], index=[pd.Timestamp("2021-01-01")]), s]).sort_index()
            native = {"2021-01-01": "ANCHOR_A"}
    if name == "Uruguay":
        u = uruguay_2020(data_dir).get(fuel)
        if u is not None:
            s = pd.concat([pd.Series([u], index=[pd.Timestamp("2021-01-01")]), s]).sort_index()
    return cf.Series(s, spec["res"], spec["kind"], spec["source"], spec["url"], "", native or None)


def uruguay_2020(data_dir):
    """MIEM's own 31 December 2020 stock (its 'Pot Inst Fuente' table kept on the 'MIEM by source' sheet), by our fuels."""
    p = os.path.join(data_dir, "uruguay_power_capacity.xlsx")
    if not os.path.exists(p):
        return {}
    d = pd.read_excel(p, sheet_name="MIEM by source")
    d = d.set_index(d.columns[0])
    d.columns = [str(c).replace(".0", "") for c in d.columns]
    if "2020" not in d:
        return {}
    pick = {"Hydro": "Total Hidráulica", "Wind": "Total Eólica", "Solar": "Total Solar", "Oil": "Total Fósil", "Bioenergy": "Total Biomasa"}
    out = {f: float(d.loc[r, "2020"]) for f, r in pick.items() if r in d.index and pd.notna(d.loc[r, "2020"])}
    if len(out) == len(pick):
        out.update({"Gas": 0.0, "Coal": 0.0, "Nuclear": 0.0, "Other": 0.0})
    return out


def sa_config(data_dir, name):
    if name in SA_NAT:
        spec = SA_NAT[name]
        iso = spec["iso"]
        cfg = {}
        for fuel in FUELS:
            prim = national_series(data_dir, name, spec, fuel)
            cfg[fuel] = cf.FuelCfg(primary=[prim] if prim is not None else [], ref=ember(data_dir, iso, fuel, name),
                                   note=NOTES.get((name, fuel), ""),
                                   lead_hold=1 if spec.get("lead") == "hold" else 0)
        if name == "Argentina":     # CAMMESA's hydro includes the two pumped-storage plants
            cfg["PumpedStorage"] = cf.FuelCfg(included_in="Hydro", note=NOTES[(name, "Hydro")])
        return cfg
    iso = dict(SA_EMBER_ONLY)[name]
    return {fuel: cf.FuelCfg(primary=[ember(data_dir, iso, fuel, name, "EMBER")], note=NOTES.get((name, fuel), ""))
            for fuel in FUELS}


def us_series(data_dir, sheetname, fuel):
    """EIA-860M register: a fuel absent from a month (or from a region's sheet) has no operating capacity there = 0, not a gap."""
    d = sheet(data_dir, "us_power_capacity.xlsx", sheetname, 1)
    if d is None:
        return None
    c = US_COLMAP.get(fuel, f"{fuel}_MW")
    s = pd.to_numeric(d[c], errors="coerce").fillna(0.0) if c in d else pd.Series(0.0, index=d.index)
    return cf.Series(s, "M", "RAW", US_SRC[0], US_SRC[1])


def na_config(data_dir, name):
    cfg = {}
    if name in dict(US_REGIONS):
        sh = dict(US_REGIONS)[name]
        for fuel in FUELS:
            ref = ember(data_dir, "USA", fuel, name) if name == "United States" else None
            s = us_series(data_dir, sh, fuel)
            cfg[fuel] = cf.FuelCfg(primary=[s] if s is not None else [], ref=ref, note=NOTES.get((name, fuel), ""),
                                   lead_hold=1)
    elif name == "Canada":
        d = sheet(data_dir, "canada_power_capacity.xlsx", "Monthly", 12)
        for fuel in FUELS:
            s = col(d, f"{fuel}_MW", "A", None, *CAN_SRC) if fuel != "Bioenergy" else None
            cfg[fuel] = cf.FuelCfg(primary=[s] if s is not None else [], ref=ember(data_dir, "CAN", fuel, name),
                                   note=NOTES.get((name, fuel), ""),
                                   included_in="Other" if fuel == "Bioenergy" and s is None and d is not None else None)
    else:   # Mexico
        for fuel in FUELS:
            cfg[fuel] = cf.FuelCfg(primary=[ember(data_dir, "MEX", fuel, name, "EMBER")], note=NOTES.get((name, fuel), ""))
    return cfg


def build_region(data_dir, region, end=None):
    names = SA_ORDER if region == "sa" else NA_ORDER
    make = sa_config if region == "sa" else na_config
    cfgs = {n: make(data_dir, n) for n in names}
    cap = pd.Timestamp(datetime.date.today().replace(day=1))
    end = pd.Timestamp(end) if end else cap     # held flat (HELD_LATEST) from each source's last point to the current month
    results = {}
    for n in names:
        if not any((s is not None and not s.clean().empty) for c in cfgs[n].values() for s in c.primary + [c.ref]):
            print(f"  {n}: no source at all, skipped")
            continue
        results[n] = cf.fill_country(n, FUELS, START, end, cfgs[n])
    return results, end


# ------------------------------------------------------------------ output
def country_sheet(r):
    df = r.values.copy()
    df.columns = [f"{c}_MW" for c in df.columns]
    df.insert(0, "Total_MW", cf.total(r.values))
    df.index.name = "date"
    return df.round(1)


EXTRA_BREAKS = {
    "sa": [
        dict(country="Argentina", fuel="Hydro / Gas", month="all", kind="BASIS (classification)", ratio="",
             detail="CAMMESA lists machine types, not fuels: Gas = combined cycle + gas + steam turbines (dual-fuel); hydro includes the two pumped-storage plants"),
        dict(country="Brazil", fuel="Solar", month="all", kind="BASIS (distributed solar)", ratio="",
             detail="ANEEL includes ~54 GW of distributed (MMGD) solar declared to the regulator; Ember's solar is on a similar but not identical basis"),
        dict(country="Chile", fuel="Coal / Gas / Oil", month="2021-2025", kind="BASIS (retirements)", ratio="",
             detail="CNE publishes only the current snapshot; earlier months count plants still in service plus Ember's fossil change (see the Chile workbook)"),
        dict(country="Colombia", fuel="Solar", month="2025-03", kind="SERIES NOTE (test operation)", ratio="",
             detail="From 2 Mar 2025 XM stopped reporting plants in test operation; the Colombia workbook carries them forward (about 0.6 GW)"),
        dict(country="Ecuador", fuel="all", month="2024-01..2024-04", kind="SOURCE CHANGE (annual -> monthly)", ratio="",
             detail="December stocks 2021-23 come from the annual statistics, monthly BNEE workbooks start Apr 2024 (stock end-Mar 2024 on the 1 Apr 2024 row); the four months between are INTERP_GAP / HELD_GAP"),
        dict(country="Ecuador", fuel="all", month="2025-06", kind="MISSING SOURCE MONTH", ratio="",
             detail="the BNEE workbook for May 2025 is missing: INTERP_GAP / HELD_GAP"),
        dict(country="Uruguay", fuel="Gas / Oil", month="all", kind="BASIS (classification)", ratio="",
             detail="MIEM does not split fossil plant by fuel: all shown as Oil, Gas is 0 (Ember: Gas 1.1 GW, Other Fossil 0.07 GW). The fossil total matches Ember."),
        dict(country="Uruguay", fuel="all", month="2026-07", kind="HALF-YEAR VALUE", ratio="",
             detail="MIEM's 2026 column is the latest half-year value, placed at 1 July 2026 (stock at 30 June); the earlier columns are 31 December stocks"),
        dict(country="Paraguay", fuel="all", month="all", kind="BASIS (partly Ember)", ratio="",
             detail="Itaipu / Yacyreta halves are raw; ANDE's own plants are Ember's Paraguay figures less the two halves (ande.gov.py is behind a captcha). Constant in 2021-25."),
    ],
    "na": [
        dict(country="United States", fuel="all", month="all", kind="BASIS (net summer capacity)", ratio="",
             detail="EIA-860M net summer capacity; Ember shows the same stock on a nameplate-leaning basis (hence differences of a few %); 50 states, Alaska and Hawaii included, territories not"),
        dict(country="US regions", fuel="all", month="all", kind="BASIS (balancing authority)", ratio="",
             detail="the ISO / utility sheets sum EIA-860M generators by the balancing-authority code EIA assigns (ERCO, PJM, MISO, SWPP, CISO, NYIS, ISNE, SOCO, TVA); they are not part of the North America total (the United States sheet is)"),
        dict(country="Canada", fuel="Other / Bioenergy", month="all", kind="BASIS (not split)", ratio="",
             detail="StatCan 'other principal fuels' (biomass etc.) is not split from geothermal / other: all in Other, Bioenergy INCLUDED_IN_OTHER"),
        dict(country="Canada", fuel="all", month="2025-01", kind="HELD AFTER LAST YEAR-END", ratio="",
             detail="StatCan publishes the year-end stock about a year later: the 31 Dec 2024 stock is held flat from 1 Jan 2025 (HELD_LATEST); Ember's 2025 stock is not spliced on"),
        dict(country="Mexico", fuel="all", month="all", kind="FALLBACK_EMBER", ratio="",
             detail="no raw Mexican capacity feed in the repo (SENER / CENACE publish capacity only in PDF reports); Ember yearly capacity, interpolated like ENTSO-E annual data"),
    ],
}


def drop_zero_splice_breaks(breaks_df, results):
    """The engine lists a splice as a level break when it cannot form a ratio; where the fuel is 0 on both sides (national source 0 and
    Ember 0 / not listed) nothing differs, so the row is dropped."""
    if breaks_df.empty:
        return breaks_df
    keep = []
    for _, b in breaks_df.iterrows():
        if b["kind"] == "SPLICE LEVEL BREAK" and (b["ratio"] == "" or pd.isna(b["ratio"])) and b["country"] in results:
            v = results[b["country"]].values[b["fuel"]]
            before = v[v.index < pd.Timestamp(b["month"] + "-01")]
            if len(before) and (before.fillna(0) == 0).all():
                continue
        keep.append(b)
    return pd.DataFrame(keep, columns=breaks_df.columns) if keep else breaks_df.iloc[0:0]


def large_steps(results, share=0.05, min_mw=100.0):
    """Month-on-month changes of a fuel above 5% of the country total (and 100 MW): real commissionings and closures, annual stocks
    stepping at a 1 January, and source / splice boundaries all show up - listed so none is mistaken for a smooth series."""
    rows = []
    for name, r in results.items():
        tot = cf.total(r.values).shift()
        for fuel in r.values.columns:
            dl = r.values[fuel].diff()
            for m in dl.index[(dl.abs() > np.maximum(min_mw, share * tot.fillna(0))).fillna(False).to_numpy()]:
                rows.append(dict(country=name, fuel=fuel, month=f"{m:%Y-%m}", kind="LARGE STEP", ratio=round(float(dl[m] / tot[m]), 3),
                                 detail=f"{dl[m]:+.0f} MW ({dl[m] / tot[m] * 100:+.1f}% of the previous month's country total), cell flag {r.flags.loc[m, fuel]}"))
    return rows


def write_region(data_dir, region, out, end_arg=None, with_charts=True):
    results, end = build_region(data_dir, region, end_arg)
    total_name = SA_TOTAL if region == "sa" else NA_TOTAL
    label = "South & Central America and the Caribbean" if region == "sa" else "North America"
    members = [n for n in results if region == "sa" or n in NA_COUNTRIES]
    sheets, basis, breaks, cov_rows, flag_sheets = {}, [], [], [], {}
    for name, r in results.items():
        sheets[name[:31]] = country_sheet(r)
        fl = r.flags.copy()
        fl.index.name = "date"
        flag_sheets[f"{name} flags"[:31]] = fl
        basis.extend(r.basis)
        breaks.extend(r.breaks)
        c = cf.coverage(r.flags)
        c.name = name
        cov_rows.append(c)
    coverage = pd.DataFrame(cov_rows).fillna(0.0).round(1)
    coverage.index.name = "country"
    tot = sum(results[n].values.fillna(0) for n in members)
    anyv = sum(results[n].values.notna().astype(int) for n in members)
    tot = tot.where(anyv > 0)
    reg = tot.copy()
    reg.columns = [f"{c}_MW" for c in reg.columns]
    reg.insert(0, "Total_MW", cf.total(tot))
    reg.index.name = "date"
    reg = reg.round(1)
    allflags = pd.concat([results[n].flags for n in members], axis=1)

    def share(group):
        return allflags.isin(group).sum(axis=1) / allflags.shape[1] * 100

    obs = share(cf.OBSERVED)
    none = share(["NONE"]) + allflags.apply(lambda c: c.astype(str).str.startswith("INCLUDED_IN_")).sum(axis=1) / allflags.shape[1] * 100
    held = share(["HELD_LATEST"])
    fb = share(["FALLBACK_EMBER"])
    cbm = pd.DataFrame({"Observed (RAW / REBUILT / ANCHOR)": obs, "Interpolated / held between points": 100 - obs - none - held - fb,
                        "Held after last point": held, "Ember yearly fallback": fb,
                        "No source (NONE) or included in another fuel": none}).round(1)
    cbm.index.name = "date"
    # the Ember fallback cells hold both interpolated and held values; 'Interpolated' above is what is left after the others
    basis_df = pd.DataFrame(basis)
    basis_df["caveat"] = [(b["caveat"] + (" " if b["caveat"] else "") + COMMON_NOTE.get(b["fuel"], "")).strip() for b in basis]
    breaks_df = pd.DataFrame(breaks) if breaks else pd.DataFrame(columns=["country", "fuel", "month", "kind", "ratio", "detail"])
    breaks_df = drop_zero_splice_breaks(breaks_df, results)
    extra = [e for e in EXTRA_BREAKS[region] if e["country"] in results or e["country"] == "US regions"]
    extra += [dict(country=sl, fuel=fu, month=f"{dt:%Y-%m}", kind="EMBER ZERO TREATED AS NOT REPORTED", ratio="",
                   detail="Ember returns 0 (or a collapse) for a year / fuel not reported; the point is dropped and the series interpolates / holds across it")
              for sl, dt, fu in sorted(set(ZERO_DROPPED), key=lambda t: (t[0], str(t[1]), t[2])) if sl in results]
    extra += large_steps(results)
    breaks_df = pd.concat([breaks_df, pd.DataFrame(extra)], ignore_index=True)
    none_list = basis_df[basis_df["rule"].astype(str).str.startswith("NONE")][["country", "fuel"]]
    n_ember = sum(1 for n in results if (results[n].flags == "FALLBACK_EMBER").to_numpy().mean() > 0.5)
    notes = [
        f"{label} installed generating capacity - one gap-free monthly dataset per country and fuel (MW)", "",
        "What this is",
        "Built by americas/capacity/AMERICAS_CAPACITY_COMPLETE.py with the generic rules in capacity_fill.py from the capacity workbooks already in this folder "
        "(national monthly and annual sources, <country>_power_capacity.xlsx) and the Ember yearly capacity workbook (ember_capacity_yearly.xlsx, fuel split).",
        f"Span: {START[:7]} to {end:%Y-%m}, monthly, MW. A row dated the 1st of a month is the stock at the END of the previous month (the repo-wide convention): "
        "national monthly rows dated by the month they close are shifted one month, annual 31 December stocks sit on 1 January of the next year, "
        "Ember year Y on 1 January Y+1. Total_MW = all fuels except PumpedStorage and Storage (batteries).",
        "", "Fill rules (per country and fuel, ONE source basis for the whole span)",
        "RAW / REBUILT: a monthly source used as is (REBUILT = a plant register or machine list rebuilt to month-end). ANCHOR_A: an annual official stock at its own date. "
        "INTERP_A: wind, solar, batteries and bioenergy linear between annual stocks. HELD_A: hydro, pumped storage, gas, coal, nuclear, oil and other held at the last stock until the next one.",
        "INTERP_GAP / HELD_GAP: a hole of up to 6 months inside a monthly source (Ecuador May 2025, Jan-Apr 2024) filled the same way. HELD_LATEST: after the last official point the stock is held flat to the current month, no trend. "
        "SPLICED / SPLICED_UNSCALED: months before an annual source starts are back-filled from Ember scaled by the ratio at the first 1 January both hold (only within 0.7-1.3; else unscaled and listed on Breaks). "
        "HELD_BACK: the first stock of a monthly source held one month back (its first row is Jan 2021, which lands on 1 Feb 2021). FALLBACK_EMBER: Ember yearly capacity (interpolated or held like annual official data) where a country has no national source. "
        "INCLUDED_IN_OTHER: Canada's bioenergy is inside StatCan's 'other principal fuels'. NONE: no source holds the fuel (blank).",
        "", "Sheets",
        f"Basis (source, rule, ratio, caveat per country and fuel), Breaks (level and series breaks), Coverage (share of country-fuel-months per flag), Coverage by month, "
        f"'{total_name}' ({'all countries' if region == 'sa' else 'United States, Canada, Mexico; the US regional sheets are not added'}), one sheet and one '<name> flags' sheet per country"
        + (" and per US region (Lower 48, ERCOT, PJM, MISO, SPP, CAISO, NYISO, ISONE, Southern, TVA)." if region == "na" else "."),
        "", "Caveats",
        "A held or interpolated value is not an observation: use Coverage before drawing conclusions about the latest months or about step fuels inside a year. "
        "Annual stocks are 31 December values: closures and commissionings inside a year appear at the next 1 January for held fuels. "
        "Ember is a compilation (national ministries and grid operators, IRENA) on its own fuel classification; a national source and Ember can differ by more than the 0.7-1.3 band for small fuels (see Breaks).",
        f"{n_ember} of {len(results)} sheets are mostly Ember fallback.",
    ]
    titles = ["What this is", "Fill rules (per country and fuel, ONE source basis for the whole span)", "Sheets", "Caveats"]
    all_sheets = {"Basis": basis_df.set_index("country"), "Breaks": breaks_df.set_index("country"), "Coverage": coverage,
                  "Coverage by month": cbm, total_name: reg}
    all_sheets.update(sheets)
    all_sheets.update(flag_sheets)
    os.makedirs(os.path.dirname(out), exist_ok=True)
    xlsx_notes.write_workbook(out, all_sheets, notes, titles)
    print(f"{out}: {len(sheets)} sheets, {START[:7]} - {end:%Y-%m}")
    show = [c for c in ["OBSERVED", "FILLED", "FALLBACK_EMBER", "HELD_LATEST", "NONE", "INCLUDED"] if c in coverage]
    print(coverage[show].to_string())
    print("NONE:", none_list.to_dict("records"))
    if with_charts:
        import add_charts
        add_charts.add_charts(out)
    return results


# ---------------------------------------------------------------- charts (registered in add_charts.REGISTRY)
CHART_GROUPS = ["Hydro", "Gas", "Wind", "Solar", "Coal", "Nuclear", "Oil", "Bioenergy", "Other"]


def chart_specs(path):
    import add_charts
    names = pd.ExcelFile(path).sheet_names
    total = next(n for n in names if n.endswith(" total"))
    specs = []
    t = add_charts.by_date(add_charts.read(path, total), "date")
    g = pd.DataFrame({f: t.get(f"{f}_MW") for f in CHART_GROUPS}, index=t.index) / 1000.0
    specs.append(add_charts.spec(total, g, f"{total.replace(' total', '')} installed capacity, complete monthly series",
                                 "GW installed", "stacked_bar"))
    cbm = add_charts.by_date(add_charts.read(path, "Coverage by month"), "date")
    specs.append(add_charts.spec("Coverage", cbm, "Share of country-fuel-months observed vs filled", "% of country-fuel cells", "stacked_bar"))
    for n in names:
        if n in ("Units", "Basis", "Breaks", "Coverage", "Coverage by month", total) or n.endswith(" flags") or n.startswith("Chart"):
            continue
        d = add_charts.by_date(add_charts.read(path, n), "date")
        gg = pd.DataFrame({f: d.get(f"{f}_MW") for f in CHART_GROUPS}, index=d.index) / 1000.0
        gg = gg.dropna(how="all", axis=1)
        if gg.empty:
            continue
        specs.append(add_charts.spec(n, gg, f"{n} installed capacity, complete monthly series", "GW installed", "stacked_bar"))
    return specs


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--region", default="all", choices=["sa", "na", "all"])
    ap.add_argument("--data-dir", default=DATA_DIR)
    ap.add_argument("--end", default=None)
    ap.add_argument("--out-dir", default=None)
    ap.add_argument("--no-charts", action="store_true")
    args = ap.parse_args()
    od = args.out_dir or args.data_dir
    if args.region in ("sa", "all"):
        write_region(args.data_dir, "sa", os.path.join(od, SA_OUT), args.end, not args.no_charts)
    if args.region in ("na", "all"):
        write_region(args.data_dir, "na", os.path.join(od, NA_OUT), args.end, not args.no_charts)


if __name__ == "__main__":
    main()
