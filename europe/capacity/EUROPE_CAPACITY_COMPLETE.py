"""
Europe installed capacity, ONE gap-free monthly dataset per country and fuel (Jan 2021 - latest month, MW), every filled cell
flagged.   output/Data and Chart Outputs/europe_capacity_monthly_complete.xlsx

No downloads: it reads the capacity workbooks the other europe/capacity pulls have already committed (national monthly /
quarterly / annual sources and the ENTSO-E A68 annual workbooks `<slug>_power_capacity.xlsx`) and combines them with the generic
rules in the repo-root `capacity_fill.py` (see its docstring for the fill rules and flags). This file is only the Europe CONFIG
(which source for which country and fuel, and why) and the output.

Basis: for every country and fuel ONE source is used for the whole span (no change of level between sources, except the
documented splice). Fuels a national monthly source lacks come from ENTSO-E annual (linear for wind / solar / storage /
bioenergy, held for the others); after the last official point the stock is held flat (HELD_LATEST), no trend. Nothing is guessed:
a fuel no source holds is blank and flagged NONE.

Date convention: a row dated the 1st is the stock at the end of the previous month. Germany, Belgium, Ireland and the United
Kingdom date their rows by the month they close, so they are shifted one month on reading (`shift=1`).

Countries: the Europe master's ENTSO-E set (europe/europe_countries.py, Ireland = Republic from EirGrid), the United Kingdom
(not in that file; the master shows GB power from Elexon; figures here are UK incl. Northern Ireland) and, if their workbooks
exist, Turkey and any country with a `<slug>_power_capacity_monthly.xlsx` that is not configured below (used as a RAW monthly
source for the fuels it holds, ENTSO-E annual for the rest, the other agents' Greece / Balkans / Turkey pulls land that way).

Sheets: Units, Basis, Breaks, Coverage, Coverage by month, Europe total, one sheet per country (date, <Fuel>_MW, Total_MW =
all fuels except pumped storage and batteries, memo columns) and '<country> flags'.
Usage: python3 EUROPE_CAPACITY_COMPLETE.py [--data-dir DIR] [--out FILE] [--end YYYY-MM]
"""
import argparse
import datetime
import os
import sys

import pandas as pd

ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
sys.path.insert(0, ROOT)
sys.path.insert(0, os.path.join(ROOT, "europe"))
import capacity_fill as cf  # noqa: E402
import xlsx_notes  # noqa: E402
from europe_countries import COUNTRIES  # noqa: E402

DATA_DIR = os.path.join(ROOT, "output", "Data and Chart Outputs")
OUT_FILE = "europe_capacity_monthly_complete.xlsx"
START = "2021-01-01"
FUELS = cf.FUEL_ORDER
ENTSOE_SRC = "ENTSO-E A68 installed capacity per type (1 January)"
ENTSOE_URL = "https://transparency.entsoe.eu/"
ALL_FUELS = " ".join(FUELS).split()

_cache = {}


def sheet(data_dir, file, name, shift=0):
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


def col(d, c, res, kind=None, source="", url="", note=""):
    if d is None or c not in d:
        return None
    s = pd.to_numeric(d[c], errors="coerce")
    return cf.Series(s, res, kind, source, url, note)


ZERO_DROPPED = []   # (slug, date, fuel or 'all row') ENTSO-E zeros treated as "not reported"
_FUEL_COLS = [f"{f}_MW" for f in ("Hydro", "Gas", "Wind", "Solar", "Coal", "Nuclear", "Oil", "Bioenergy", "Other")]


def entsoe_clean(data_dir, slug):
    """ENTSO-E annual workbook with its not-reported years removed: ENTSO-E A68 returns 0 (not blank) for a year or a fuel a
    country did not report. A whole row whose fuels add to under 20% of the country's median row (Slovakia 2026, Denmark 2024),
    or a fuel at 0 between two years that are both above 0 (Albania oil 2025), is dropped, so the series interpolates across it."""
    key = ("entsoe_clean", data_dir, slug)
    if key in _cache:
        return _cache[key]
    d = sheet(data_dir, f"{slug}_power_capacity.xlsx", "Monthly")
    if d is None:
        _cache[key] = None
        return None
    d = d.copy().apply(pd.to_numeric, errors="coerce")
    cols = [c for c in _FUEL_COLS if c in d]
    tot = d[cols].sum(axis=1)
    med = tot[tot > 0].median() if (tot > 0).any() else 0
    for i in d.index[(tot < 0.2 * med).values]:
        d.loc[i, :] = float("nan")
        ZERO_DROPPED.append((slug, i, "all fuels"))
    for c in d.columns:
        if not c.endswith("_MW"):
            continue
        s = d[c]
        for j, i in enumerate(s.index):
            if s[i] == 0 and 0 < j < len(s) - 1 and s.iloc[j - 1] > 0 and s.iloc[j + 1] > 0:
                d.loc[i, c] = float("nan")
                ZERO_DROPPED.append((slug, i, c[:-3]))
    _cache[key] = d
    return d


def entsoe(data_dir, slug, fuel):
    if not slug:
        return None
    return col(entsoe_clean(data_dir, slug), f"{fuel}_MW", "A", None, ENTSOE_SRC, ENTSOE_URL)


# ---------------------------------------------------------------- the config
# name -> dict(code, slug (ENTSO-E workbook), mon = national monthly/quarterly source for SOME fuels)
#   mon: file, sheet, shift, res, kind, fuels (taken from it), colmap {fuel: column}, kinds {fuel: kind}, source, url
#   memo: columns of the mon sheet passed through unflagged
MON = {
    "Germany": dict(file="germany_power_capacity_monthly.xlsx", sheet="Monthly", shift=1, res="M", kind="REBUILT",
                    fuels="Hydro PumpedStorage Gas Coal Nuclear Oil Wind Solar Bioenergy Storage Other",
                    source="Bundesnetzagentur MaStR (Gesamtdatenexport), month-end rebuild",
                    url="https://www.marktstammdatenregister.de/",
                    memo=["Lignite_MW", "Wind_Onshore_MW", "Wind_Offshore_MW"]),
    "France": dict(file="france_power_capacity_monthly.xlsx", sheet="Monthly", shift=0, res="M", kind="REBUILT",
                   fuels="Hydro PumpedStorage Gas Coal Nuclear Oil Wind Solar Bioenergy Storage Other",
                   source="ODRE national register of production and storage installations, month-end rebuild",
                   url="https://odre.opendatasoft.com/", memo=["Wind_Onshore_MW", "Wind_Offshore_MW"]),
    "Spain": dict(file="spain_power_capacity_monthly.xlsx", sheet="Monthly", shift=0, res="M", kind="RAW",
                  fuels="Hydro Gas Coal Nuclear Oil Wind Solar Bioenergy Other",
                  source="Red Electrica REData potencia instalada (generacion)", url="https://www.ree.es/en/datos/generation"),
    "Portugal": dict(file="portugal_power_capacity_monthly.xlsx", sheet="Monthly", shift=0, res="M", kind="RAW",
                     fuels="Hydro PumpedStorage Nuclear Wind Solar Bioenergy Other",
                     source="DGEG monthly renewables report (31 Dec rows to 2023, monthly from 2024)",
                     url="https://www.dgeg.gov.pt/"),
    "Switzerland": dict(file="switzerland_power_capacity_monthly.xlsx", sheet="Monthly", shift=0, res="M", kind="REBUILT",
                        fuels="Hydro PumpedStorage Gas Coal Nuclear Oil Wind Solar Bioenergy Other",
                        source="BFE / Pronovo plant register, month-end rebuild",
                        url="https://www.bfe.admin.ch/"),
    "Czechia": dict(file="czechia_power_capacity_monthly.xlsx", sheet="Monthly", shift=0, res="M", kind="RAW",
                    fuels="Hydro PumpedStorage Gas Coal Nuclear Wind Solar",
                    source="ERU quarterly electricity system report (monthly rows)", url="https://eru.gov.cz/"),
    "Hungary": dict(file="hungary_power_capacity_monthly.xlsx", sheet="Monthly", shift=0, res="M", kind="RAW",
                    fuels="Hydro PumpedStorage Gas Coal Nuclear Oil Wind Solar Bioenergy Storage Other",
                    source="MAVIR gross installed capacity by production type (monthly PDF)", url="https://www.mavir.hu/"),
    "Poland": dict(file="poland_power_capacity_quarterly.xlsx", sheet="Quarterly", shift=0, res="Q", kind="RAW",
                   fuels="Wind", source="URE renewable installations register, quarterly", url="https://www.ure.gov.pl/"),
    "Belgium": dict(file="belgium_power_capacity_monthly.xlsx", sheet="Monthly", shift=1, res="M", kind="RAW",
                    fuels="Wind Solar", source="Elia open data ods031 / ods032 (complete monitored wind and solar fleet)",
                    url="https://opendata.elia.be/", memo=["Wind_Onshore_MW", "Wind_Offshore_MW"]),
    "Denmark": dict(file="denmark_power_capacity_monthly.xlsx", sheet="Monthly", shift=0, res="M", kind="RAW",
                    fuels="Wind Solar", source="Energinet Energi Data Service CapacityPerMunicipality",
                    url="https://www.energidataservice.dk/", memo=["Wind_Onshore_MW", "Wind_Offshore_MW", "Thermal_unsplit_MW"]),
    "Norway": dict(file="norway_power_capacity_monthly.xlsx", sheet="Monthly", shift=0, res="M", kind="REBUILT",
                   fuels="Hydro Wind", colmap={"Hydro": "Hydro_incl_pumped_MW"},
                   source="NVE hydropower and wind databases, month-end rebuild", url="https://www.nve.no/"),
    "Ireland": dict(file="ireland_power_capacity_monthly.xlsx", sheet="Monthly", shift=1, res="M", kind="REBUILT",
                    fuels="Hydro PumpedStorage Wind Solar Bioenergy Storage", kinds={"Wind": "RAW"},
                    source="EirGrid (wind installed capacity, raw) and EirGrid / ESB Networks register rebuild (Republic of Ireland)",
                    url="https://www.eirgrid.ie/"),
}
SLUG_FOR = {code: slug for code, (_, slug, _z) in COUNTRIES.items()}

NOTES = {
    ("Germany", "Hydro"): "MaStR Germany-registered units only; ENTSO-E's DE-LU zone also holds Vianden (Luxembourg) and Austrian-registered units. MaStR kept; Luxembourg is not added on top.",
    ("Germany", "PumpedStorage"): "MaStR Germany-registered units only (no Vianden); see Hydro.",
    ("Germany", "Oil"): "Includes small CHP / emergency units (Oil_ge10MW is a memo on the source sheet; closer to ENTSO-E).",
    ("Germany", "Coal"): "Hard coal + lignite (Lignite_MW memo).",
    ("Spain", "PumpedStorage"): "REData has no pumped storage; ENTSO-E annual.",
    ("Spain", "Solar"): "Includes self-consumption (ENTSO-E annual does not).",
    ("Spain", "Other"): "Includes cogeneration and other non-renewables as REData classifies them.",
    ("Portugal", "Solar"): "DGEG includes self-consumption / small units (about twice ENTSO-E). Kept as published; level difference listed on Breaks.",
    ("Portugal", "Hydro"): "Excludes pumped storage (DGEG Hydro_incl_pumped is the sum).",
    ("Czechia", "Solar"): "SERIES BREAK Jan 2024: ERU's coverage widened (+1.2 GW), not new build; not smoothed.",
    ("Czechia", "Coal"): "ERU: all steam plant (hard coal, lignite and steam plant on other fuels).",
    ("Czechia", "Gas"): "ERU: CCGT + gas turbines / engines.",
    ("Belgium", "Wind"): "Elia monitored fleet (= ENTSO-E series).",
    ("Belgium", "Nuclear"): "ENTSO-E 1 January step used: Elia's contracted-unit list still shows closed units for months after closure.",
    ("Belgium", "Gas"): "ENTSO-E annual: Elia's list holds contracted units only.",
    ("Belgium", "Bioenergy"): "ENTSO-E annual: Elia's list holds contracted units only.",
    ("Belgium", "Hydro"): "ENTSO-E annual: Elia's list holds contracted units only (no run-of-river).",
    ("Belgium", "Oil"): "ENTSO-E annual: Elia's list holds contracted units only.",
    ("Denmark", "Gas"): "Energinet's thermal capacity is split by size only (Thermal_unsplit_MW, memo column); fuels taken from ENTSO-E annual rather than splitting the unsplit figure by shares.",
    ("Denmark", "Coal"): "ENTSO-E annual (Energinet thermal not split by fuel; memo Thermal_unsplit_MW).",
    ("Denmark", "Oil"): "ENTSO-E annual (Energinet thermal not split by fuel).",
    ("Denmark", "Bioenergy"): "ENTSO-E annual (Energinet thermal not split by fuel).",
    ("Norway", "Hydro"): "NVE conventional + pumped-storage plants (matches ENTSO-E Hydro within ~1%); pumped storage is inside Hydro.",
    ("Poland", "Solar"): "URE register excludes micro-installations (about a quarter of ENTSO-E in 2021); ENTSO-E annual used.",
    ("Poland", "Hydro"): "URE and ENTSO-E differ by more than 30% (different plant scope); ENTSO-E annual used for one basis across the Polish thermal / hydro fuels.",
    ("Poland", "Bioenergy"): "URE includes biomass co-firing units (1.8 x ENTSO-E); ENTSO-E annual used.",
    ("Ireland", "Gas"): "ENTSO-E SEM is all-island (Republic + Northern Ireland); EirGrid's thermal register has a connection year only, so no Republic-only history exists. Fuels from the register (hydro, wind, solar, bio, storage) are Republic only.",
    ("Ireland", "Coal"): "ENTSO-E SEM all-island; see Gas.",
    ("Ireland", "Oil"): "ENTSO-E SEM all-island; see Gas.",
    ("Ireland", "Other"): "ENTSO-E SEM all-island (peat, waste, other); see Gas.",
    ("Ireland", "Wind"): "EirGrid TSO + DSO (Republic only); DSO figures post later, so months after the last complete one are held.",
    ("United Kingdom", "Solar"): "DESNZ Energy Trends 6.3, UK incl. Northern Ireland and small / domestic.",
    ("United Kingdom", "Wind"): "REPD rebuild (>= 150 kW, operational today only, so earlier months understated) scaled to the ET 6.1 quarterly stock at every quarter end.",
    ("United Kingdom", "Bioenergy"): "REPD rebuild scaled to the ET 6.1 quarterly stock.",
    ("United Kingdom", "PumpedStorage"): "REPD rebuild scaled to DUKES 5.12 year-end stocks.",
    ("United Kingdom", "Storage"): "REPD batteries are ~40% below DUKES 5.16; DUKES annual (linear) used instead.",
    ("United Kingdom", "Hydro"): "ET 6.1 quarterly (REPD lacks the older large hydro).",
    ("United Kingdom", "Gas"): "DUKES 5.12 annual, held between year ends (no monthly or dated-closure source we hold).",
    ("United Kingdom", "Coal"): "DUKES 5.12 annual, held between year ends: a closure inside a year shows at the next 1 January (the last station closed in 2024).",
    ("United Kingdom", "Nuclear"): "DUKES 5.12 annual, held between year ends.",
    ("United Kingdom", "Oil"): "DUKES 5.12 annual, held between year ends.",
    ("United Kingdom", "Other"): "DUKES 5.12 annual, held between year ends.",
    ("Netherlands", "Solar"): "ENTSO-E annual; CBS (86266NED / 82610NED) ends at year-end 2024 and was kept as a cross-check only (CBS / ENTSO-E 0.88 at 1 Jan 2025).",
}
COMMON_NOTE = {"Storage": "batteries; excluded from Total", "PumpedStorage": "excluded from Total"}


def uk_config(data_dir):
    m = sheet(data_dir, "united_kingdom_power_capacity_monthly.xlsx", "Monthly", 1)
    q = sheet(data_dir, "united_kingdom_power_capacity_monthly.xlsx", "Quarterly_ET61", 1)
    a = sheet(data_dir, "united_kingdom_power_capacity_monthly.xlsx", "Annual_DUKES")
    if a is not None:
        a = a.copy()
        a.index = pd.to_datetime(a["date_1_Jan"]).dt.to_period("M").dt.to_timestamp() if "date_1_Jan" in a else a.index
    f = "United Kingdom"
    et63 = ("DESNZ Energy Trends 6.3 (solar, monthly)", "https://www.gov.uk/government/statistics/solar-photovoltaics-deployment")
    repd = ("DESNZ REPD, month-end rebuild by operational date", "https://www.gov.uk/government/publications/renewable-energy-planning-database-monthly-extract")
    et61 = ("DESNZ Energy Trends 6.1 (quarterly renewable capacity)", "https://www.gov.uk/government/statistics/energy-trends-section-6-renewables")
    dukes = ("DESNZ DUKES 5.12 / 5.16 (year-end capacity)", "https://www.gov.uk/government/statistics/electricity-chapter-5-digest-of-united-kingdom-energy-statistics-dukes")

    def S(d, c, res, kind, lab):
        return col(d, c, res, kind, lab[0], lab[1])

    def D(c):
        return S(a, f"{c}_MW", "A", None, dukes)
    cfg = {}
    for fuel in ["Gas", "Coal", "Nuclear", "Oil", "Other"]:
        cfg[fuel] = cf.FuelCfg(primary=[D(fuel)], note=NOTES.get((f, fuel), ""))
    cfg["Solar"] = cf.FuelCfg(primary=[S(m, "Solar_MW", "M", "RAW", et63)], ref=D("Solar"), note=NOTES[(f, "Solar")])
    cfg["Hydro"] = cf.FuelCfg(primary=[S(q, "Hydro_MW", "Q", None, et61)], ref=D("Hydro"), note=NOTES[(f, "Hydro")])
    cfg["Wind"] = cf.FuelCfg(primary=[S(m, "Wind_MW", "M", "REBUILT", repd)], scale_to=S(q, "Wind_MW", "Q", None, et61),
                             ref=D("Wind"), note=NOTES[(f, "Wind")])
    cfg["Bioenergy"] = cf.FuelCfg(primary=[S(m, "Bioenergy_MW", "M", "REBUILT", repd)], scale_to=S(q, "Bioenergy_MW", "Q", None, et61),
                                  ref=D("Bioenergy"), note=NOTES[(f, "Bioenergy")])
    cfg["PumpedStorage"] = cf.FuelCfg(primary=[S(m, "PumpedStorage_MW", "M", "REBUILT", repd)], scale_to=D("PumpedStorage"),
                                      ref=D("PumpedStorage"), note=NOTES[(f, "PumpedStorage")])
    cfg["Storage"] = cf.FuelCfg(primary=[S(m, "Storage_MW", "M", "REBUILT", repd)], scale_to=D("Storage"), ref=D("Storage"),
                                note=NOTES[(f, "Storage")])
    return cfg


def country_list(data_dir):
    """[(name, code, slug)] in a fixed order: the master's countries (Ireland = Republic), the UK, then extras."""
    out = []
    for code, (name, slug, _z) in COUNTRIES.items():
        if code == "IE":
            out.append(("Ireland", "IE", slug))
        else:
            out.append((name, code, slug))
    out.append(("United Kingdom", "GB", None))
    have = {s for _n, _c, s in out}
    for fn in sorted(os.listdir(data_dir)) if os.path.isdir(data_dir) else []:
        if fn.endswith("_power_capacity_monthly.xlsx"):
            slug = fn[: -len("_power_capacity_monthly.xlsx")]
            if slug in have or slug in ("united_kingdom", "ireland", "luxembourg") or any(
                    v["file"] == fn for v in MON.values()):
                continue
            out.append((slug.replace("_", " ").title(), slug.upper()[:2], slug))
    return out


def build_country(data_dir, name, code, slug):
    if name == "United Kingdom":
        return uk_config(data_dir), []
    cfg, memo = {}, []
    mon = MON.get(name)
    md = None
    if mon is None:
        fn = f"{slug}_power_capacity_monthly.xlsx"
        d = sheet(data_dir, fn, "Monthly")
        if d is not None:
            # a national file may split coal and lignite or carry geothermal (Turkey's EPDK): fold them into the standard
            # Coal / Other columns and keep the parts as memo columns
            d, colmap = d.copy(), {}
            if "Lignite_MW" in d and "Coal_MW" in d:
                d["Coal_incl_lignite_MW"] = d[["Coal_MW", "Lignite_MW"]].apply(pd.to_numeric, errors="coerce").sum(
                    axis=1, min_count=1)
                colmap["Coal"] = "Coal_incl_lignite_MW"
            if "Geothermal_MW" in d:
                parts = [c for c in ("Other_MW", "Geothermal_MW") if c in d]
                d["Other_incl_geothermal_MW"] = d[parts].apply(pd.to_numeric, errors="coerce").sum(axis=1, min_count=1)
                colmap["Other"] = "Other_incl_geothermal_MW"
            fuels = [f for f in FUELS if colmap.get(f, f"{f}_MW") in d and d[colmap.get(f, f"{f}_MW")].notna().any()]
            mon = dict(file=fn, sheet="Monthly", shift=0, res="M", kind="RAW", fuels=" ".join(fuels), colmap=colmap,
                       memo=[c for c in ("Lignite_MW", "Geothermal_MW") if c in d],
                       source=f"{name}: national monthly source ({fn})", url="")
            md = d
    if mon and md is None:
        md = sheet(data_dir, mon["file"], mon["sheet"], mon.get("shift", 0))
    for fuel in FUELS:
        ref = entsoe(data_dir, slug, fuel)
        prim = []
        if mon and fuel in mon["fuels"].split():
            c = (mon.get("colmap") or {}).get(fuel, f"{fuel}_MW")
            prim = [col(md, c, mon["res"], (mon.get("kinds") or {}).get(fuel, mon["kind"]), mon["source"], mon["url"])]
        cfg[fuel] = cf.FuelCfg(primary=prim, ref=ref, note=NOTES.get((name, fuel), ""))
    if name == "Norway":
        cfg["PumpedStorage"] = cf.FuelCfg(included_in="Hydro", note=NOTES[(name, "Hydro")])
    if mon:
        for c in mon.get("memo", []):
            if md is not None and c in md:
                cfg.setdefault("_memo", {})[c] = cf.Series(md[c], "M", "RAW")
    return cfg, mon.get("memo", []) if mon else []


def build_all(data_dir, end=None):
    countries = country_list(data_dir)
    # END = latest month any configured source reaches (capped at the current month)
    last = pd.Timestamp(START)
    built = {}
    for name, code, slug in countries:
        cfg, _memo = build_country(data_dir, name, code, slug)
        built[name] = (code, slug, cfg)
        for k, c in cfg.items():
            if k == "_memo":
                continue
            for s in c.primary:
                if s is not None and not s.clean().empty:
                    last = max(last, s.clean().index.max())
    cap = pd.Timestamp(datetime.date.today().replace(day=1))
    end = pd.Timestamp(end) if end else min(last, cap)
    results = {}
    for name, (code, slug, cfg) in built.items():
        memo = cfg.pop("_memo", {})
        if name != "United Kingdom" and not any(
                (s is not None and not s.clean().empty) for c in cfg.values() for s in c.primary + [c.ref]):
            continue    # no source at all for this country (e.g. a Turkey placeholder)
        if name == "Turkey" and not any(c.primary and c.primary[0] is not None for c in cfg.values()):
            continue
        for k, s in memo.items():
            cfg.setdefault(FUELS[0], cf.FuelCfg()).memo = {**(cfg[FUELS[0]].memo or {}), k: s}
        r = cf.fill_country(name, FUELS, START, end, cfg)
        results[name] = (code, r)
    return results, end


def country_sheet(r):
    df = r.values.copy()
    df.columns = [f"{c}_MW" for c in df.columns]
    df.insert(0, "Total_MW", cf.total(r.values))
    df = df[[f"{c}_MW" for c in FUELS[:0]] + [c for c in df.columns]]
    for c in r.memo.columns:
        df[c] = r.memo[c]
    df.index.name = "date"
    return df.round(1)


def europe_set(results):
    """The master's Europe total set: its ENTSO-E countries (Ireland = Republic) + the United Kingdom, not Turkey."""
    names = {"Ireland" if c == "IE" else n for c, (n, _s, _z) in COUNTRIES.items()} | {"United Kingdom"}
    return [n for n in results if n in names]


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--data-dir", default=DATA_DIR)
    ap.add_argument("--out", default=None)
    ap.add_argument("--end", default=None)
    ap.add_argument("--no-charts", action="store_true")
    args = ap.parse_args()
    out = args.out or os.path.join(args.data_dir, OUT_FILE)
    results, end = build_all(args.data_dir, args.end)

    sheets, basis, breaks, cov_rows = {}, [], [], []
    flag_sheets = {}
    for name, (code, r) in results.items():
        sheets[name[:31]] = country_sheet(r)
        fl = r.flags.copy()
        fl.index.name = "date"
        flag_sheets[f"{name} flags"[:31]] = fl
        basis.extend(r.basis)
        breaks.extend(r.breaks)
        c = cf.coverage(r.flags)
        c.name = name
        cov_rows.append(c)
    coverage = pd.DataFrame(cov_rows).round(1)
    coverage.index.name = "country"
    # Europe total
    members = europe_set(results)
    tot = sum(results[n][1].values.fillna(0) for n in members)
    allnone = sum(results[n][1].values.notna().astype(int) for n in members)
    tot = tot.where(allnone > 0)
    eu = tot.copy()
    eu.columns = [f"{c}_MW" for c in eu.columns]
    eu.insert(0, "Total_MW", cf.total(tot))
    eu.index.name = "date"
    eu = eu.round(1)
    # coverage by month, all countries and fuels (excl. INCLUDED_IN, NONE shown)
    allflags = pd.concat([results[n][1].flags for n in members], axis=1)

    def share(group):
        return allflags.isin(group).sum(axis=1) / allflags.shape[1] * 100

    obs = share(cf.OBSERVED)
    none = share(["NONE"]) + allflags.apply(lambda c: c.astype(str).str.startswith("INCLUDED_IN_")).sum(axis=1) / allflags.shape[1] * 100
    held = share(["HELD_LATEST"])
    cbm = pd.DataFrame({"Observed (RAW / REBUILT / ANCHOR)": obs, "Interpolated / held between points": 100 - obs - none - held,
                        "Held after last point": held, "No source (NONE) or included in another fuel": none}).round(1)
    cbm.index.name = "date"
    # Basis / Breaks
    basis_df = pd.DataFrame(basis)
    basis_df["caveat"] = [(b["caveat"] + (" " if b["caveat"] else "") + COMMON_NOTE.get(b["fuel"], "")).strip() for b in basis]
    if breaks:
        breaks_df = pd.DataFrame(breaks)
    else:
        breaks_df = pd.DataFrame(columns=["country", "fuel", "month", "kind", "ratio", "detail"])
    extra = [  # level / series breaks documented from the source workbooks (not detectable from the numbers alone)
        dict(country="Czechia", fuel="Solar", month="2024-02", kind="SERIES BREAK (source coverage)", ratio="",
             detail="ERU solar jumps +1.2 GW between Dec 2023 and Jan 2024 because its coverage widened, not new build; kept as published"),
        dict(country="Germany", fuel="Hydro", month="all", kind="BASIS (zone)", ratio="",
             detail="MaStR Germany-registered units; ENTSO-E DE-LU also holds Vianden (LU) and Austrian-registered plant. Capacity factors on DE generation (ENTSO-E zone) are slightly overstated for hydro / pumped storage."),
        dict(country="Spain", fuel="Solar", month="all", kind="BASIS (self-consumption)", ratio="",
             detail="REData includes self-consumption; ENTSO-E annual does not (+27% at 1 Jan 2026)"),
        dict(country="Portugal", fuel="Solar", month="all", kind="BASIS (self-consumption)", ratio="",
             detail="DGEG includes self-consumption / small units, about twice ENTSO-E"),
        dict(country="Ireland", fuel="Gas / Coal / Oil / Other", month="all", kind="BASIS (all-island)", ratio="",
             detail="thermal fuels are ENTSO-E SEM all-island; the register fuels are Republic only"),
        dict(country="United Kingdom", fuel="all", month="all", kind="BASIS (geography)", ratio="",
             detail="UK incl. Northern Ireland; the master's GB generation excludes Northern Ireland"),
    ]
    extra += [dict(country=sl, fuel=fu, month=f"{dt:%Y-%m}", kind="ENTSO-E ZERO TREATED AS NOT REPORTED", ratio="",
                   detail="ENTSO-E A68 returns 0 for a year / fuel not reported; the point is dropped and the series interpolates / holds across it")
              for sl, dt, fu in sorted(set(ZERO_DROPPED), key=lambda t: (t[0], str(t[1]), t[2]))]
    breaks_df = pd.concat([breaks_df, pd.DataFrame(extra)], ignore_index=True)
    none_list = basis_df[basis_df["rule"].astype(str).str.startswith("NONE")][["country", "fuel", "caveat"]]
    notes = [
        "Europe installed generating capacity - one gap-free monthly dataset per country and fuel (MW)", "",
        "What this is",
        "Built by europe/capacity/EUROPE_CAPACITY_COMPLETE.py with the generic rules in capacity_fill.py from the capacity workbooks already in this folder (national monthly / quarterly / annual sources and the ENTSO-E A68 annual workbooks). No downloads.",
        f"Span: {START[:7]} to {end:%Y-%m}, monthly, MW. A row dated the 1st of a month is the stock at the END of the previous month (the repo-wide convention). Total_MW = all fuels except PumpedStorage and Storage (batteries).",
        "", "Fill rules (per country and fuel, ONE source basis for the whole span)",
        "RAW / REBUILT: a monthly source used as is (REBUILT = a plant register rebuilt to month-end). REBUILT_SCALED = a register rebuild multiplied by the interpolated ratio official anchors / rebuild (UK).",
        "ANCHOR_Q / ANCHOR_A: a quarterly or annual official stock at its own date. INTERP_Q: linear between quarter-end stocks. INTERP_A: wind, solar, batteries, bioenergy linear between annual stocks. HELD_A: hydro, pumped storage, gas, coal, lignite, nuclear, oil, other held at the last annual stock until the next one.",
        "INTERP_GAP / HELD_GAP: a hole of up to 6 months inside a monthly source, filled the same way. HELD_LATEST: after the last official point the stock is held flat to the latest month, no trend. SPLICED / SPLICED_UNSCALED: months before a late-starting monthly source back-filled from the annual source (scaled by the ratio at the first overlapping 1 January when it lies within 0.7-1.3). FALLBACK_EMBER: Ember yearly (not used in Europe). NONE: no source holds the fuel, left blank.",
        "INCLUDED_IN_HYDRO: Norway's pumped storage is inside its Hydro figure and is not added.",
        "", "Sheets",
        "Basis (source, rule, ratio, caveat per country and fuel), Breaks (level and series breaks), Coverage (share of country-fuel-months per flag), Coverage by month, Europe total (the master's country set + the United Kingdom; Turkey and extras excluded), one sheet per country and '<country> flags'.",
        "", "Caveats",
        "A held or interpolated value is not an observation: use Coverage before drawing conclusions about the latest months or about step fuels inside a year. ENTSO-E annual stocks are 1 January; closures and commissionings inside a year are not dated unless a monthly source holds them.",
        "Luxembourg is inside ENTSO-E's DE-LU zone and is not a Europe-master country; it is not added to Germany.",
    ]
    titles = ["What this is", "Fill rules (per country and fuel, ONE source basis for the whole span)", "Sheets", "Caveats"]
    all_sheets = {"Basis": basis_df.set_index("country"), "Breaks": breaks_df.set_index("country"), "Coverage": coverage,
                  "Coverage by month": cbm, "Europe total": eu}
    all_sheets.update(sheets)
    all_sheets.update(flag_sheets)
    os.makedirs(os.path.dirname(out), exist_ok=True)
    xlsx_notes.write_workbook(out, all_sheets, notes, titles)
    print(f"{out}: {len(sheets)} countries, {START[:7]} - {end:%Y-%m}")
    print(coverage[["OBSERVED", "FILLED", "NONE", "INCLUDED"]].to_string())
    print("NONE:", none_list.to_dict("records"))
    if not args.no_charts:
        import add_charts
        add_charts.add_charts(out)


# ---------------------------------------------------------------- charts (registered in add_charts.REGISTRY)
CHART_GROUPS = ["Hydro", "Gas", "Wind", "Solar", "Coal", "Nuclear", "Oil", "Bioenergy", "Other"]


def chart_specs(path):
    import add_charts
    specs = []
    eu = add_charts.by_date(add_charts.read(path, "Europe total"), "date")
    g = pd.DataFrame({f: eu.get(f"{f}_MW") for f in CHART_GROUPS}, index=eu.index) / 1000.0
    specs.append(add_charts.spec("Europe total", g, "Europe installed capacity, complete monthly series (national sources, ENTSO-E annual where no monthly)",
                                 "GW installed", "stacked_bar"))
    cbm = add_charts.by_date(add_charts.read(path, "Coverage by month"), "date")
    specs.append(add_charts.spec("Coverage", cbm, "Share of country-fuel-months observed vs filled", "% of country-fuel cells", "stacked_bar"))
    for n in pd.ExcelFile(path).sheet_names:
        if n in ("Units", "Basis", "Breaks", "Coverage", "Coverage by month", "Europe total") or n.endswith(" flags") or n.startswith("Chart"):
            continue
        d = add_charts.by_date(add_charts.read(path, n), "date")
        gg = pd.DataFrame({f: d.get(f"{f}_MW") for f in CHART_GROUPS}, index=d.index) / 1000.0
        gg = gg.dropna(how="all", axis=1)
        if gg.empty:
            continue
        specs.append(add_charts.spec(n, gg, f"{n} installed capacity, complete monthly series", "GW installed", "stacked_bar"))
    return specs


if __name__ == "__main__":
    main()
