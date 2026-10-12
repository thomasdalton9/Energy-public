"""
Installed generation capacity by fuel for the United Kingdom and Ireland, from the official registers and statistics
(no model estimates, nothing interpolated).

  output/Data and Chart Outputs/united_kingdom_power_capacity_monthly.xlsx
  output/Data and Chart Outputs/ireland_power_capacity_monthly.xlsx

UNITED KINGDOM (Great Britain + Northern Ireland). No single source gives all fuels monthly, so each fuel family takes
the best raw source and says so (sheet 'Units'):
  * Solar  - DESNZ Energy Trends ET 6.3 (monthly cumulative solar PV capacity, GB / NI / UK, from Jan 2010): RAW, monthly.
  * Wind (onshore / offshore), hydro, pumped storage, batteries, bioenergy and waste - REPD, DESNZ Renewable Energy
    Planning Database (quarterly release): month-end capacity REBUILT from each project's 'Operational' date. REPD lists
    projects of 150 kW and over that are operational today, so it understates earlier months (decommissioned and repowered
    plant are not in it) and small plant is missing; the 'Check' sheet sets it against the official quarterly ET 6.1.
  * ET 6.1 (quarterly renewable capacity, UK and NI, from Q1 2010): RAW, QUARTERLY (sheet 'Quarterly_ET61').
  * DUKES 5.12 (installed capacity by fuel incl. gas, coal, nuclear, oil; UK, end of year) + 5.16 (batteries): RAW, ANNUAL
    (sheet 'Annual_DUKES'). Gas, coal, nuclear and oil have NO monthly source: the Elexon BMU register has no dates or history
    and still lists closed plant, the NESO TEC register lists only live agreements, so those columns of 'Monthly' stay blank.

IRELAND (Republic of Ireland; Northern Ireland on its own sheet) from EirGrid's 'System and Renewable Data Summary Report':
  * Wind - EirGrid monthly installed wind capacity (TSO + DSO), RAW, monthly, from 1992.
  * Solar, hydro, bioenergy / waste, batteries and pumped storage - EirGrid / ESB Networks connected-generator registers
    (energisation dates) rebuilt to month-end capacity; transmission- and distribution-connected projects in the registers
    only (no rooftop / small-scale solar).
  * Thermal plant (gas, coal, oil, peat) carries a connection YEAR only, so no monthly history: each release's snapshot is
    appended to sheet 'Thermal_snapshots'.

Incremental: every source is a whole file. Each run looks up the file's current address (gov.uk content API / EirGrid
page), compares it and its Last-Modified / length with sheet 'Releases', and downloads and re-parses it only when it
changed; otherwise the stored sheets are kept. The workbook is the history store.

Usage: python3 GB_IE_CAPACITY_MONTHLY.py [--out-dir DIR] [--only uk|ie] [--force] [--samples DIR]
  --samples DIR parses local copies of the source files instead of downloading (development / tests).
"""
import argparse
import calendar
import hashlib
import io
import os
import re
import sys
import time
from datetime import datetime, timezone

import numpy as np
import pandas as pd
import requests

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(os.path.dirname(HERE))
sys.path.insert(0, ROOT)
import xlsx_notes  # noqa: E402

OUT_DEFAULT = os.path.join(ROOT, "output", "Data and Chart Outputs")
UK_FILE = "united_kingdom_power_capacity_monthly.xlsx"
IE_FILE = "ireland_power_capacity_monthly.xlsx"

ET_API = "https://www.gov.uk/api/content/government/statistics/energy-trends-section-6-renewables"
DUKES_API = ("https://www.gov.uk/api/content/government/statistics/"
             "electricity-chapter-5-digest-of-united-kingdom-energy-statistics-dukes")
REPD_PAGE = "https://www.gov.uk/government/publications/renewable-energy-planning-database-monthly-extract"
EIRGRID_PAGE = "https://www.eirgrid.ie/grid/system-and-renewable-data-reports"

S = requests.Session()
S.headers["User-Agent"] = "Mozilla/5.0 (X11; Linux x86_64) energy-research"


# ----------------------------------------------------------------------------------------------- generic helpers
def http(method, url, tries=4, **kw):
    last = None
    for i in range(tries):
        try:
            r = S.request(method, url, timeout=240, **kw)
            if r.status_code < 400:
                return r
            last = f"HTTP {r.status_code}"
        except requests.RequestException as e:  # noqa: PERF203
            last = f"{type(e).__name__}: {e}"
        time.sleep(3 * (i + 1))
    raise RuntimeError(f"{method} {url}: {last}")


def month_start(ts):
    return pd.Timestamp(ts).to_period("M").to_timestamp()


def clean_label(s):
    s = re.sub(r"\[[^\]]*\]", "", str(s))
    return re.sub(r"\s+", " ", s).strip()


def col_name(label, prefix=""):
    return prefix + re.sub(r"[^A-Za-z0-9]+", "_", clean_label(label)).strip("_") + "_MW"


def num(v):
    try:
        x = float(str(v).replace(",", ""))
        return x
    except ValueError:
        return np.nan


def merge_replace(old, new):
    """New rows replace stored rows of the same index value; stored rows the new file no longer has are kept."""
    if old is None or old.empty:
        return new.sort_index()
    if new is None or new.empty:
        return old
    keep = old.loc[~old.index.isin(new.index)]
    return pd.concat([keep, new]).sort_index()


def read_store(path):
    """All data sheets of the existing workbook (the history store), {name: DataFrame}; {} if absent / unreadable."""
    if not os.path.exists(path):
        return {}
    out = {}
    try:
        x = pd.ExcelFile(path)
        for name in x.sheet_names:
            if name in ("Units", "Chart") or name.startswith("Chart"):
                continue
            d = x.parse(name)
            if name == "Releases":
                out[name] = d
                continue
            if d.empty:
                continue
            idx = d.columns[0]
            if idx == "date":
                d[idx] = pd.to_datetime(d[idx], errors="coerce")
                d = d.dropna(subset=[idx])
            d = d.set_index(idx).sort_index()
            out[name] = d
    except Exception as e:  # noqa: BLE001
        print(f"  could not read {os.path.basename(path)} ({type(e).__name__}: {e}); starting over")
        return {}
    return out


class Releases:
    """Sheet 'Releases': one row per source file (key, address, Last-Modified, length, sha256, fetched)."""

    def __init__(self, store):
        d = store.get("Releases")
        self.rows = {} if d is None or d.empty else {r["source"]: dict(r) for _, r in d.iterrows()}

    def same(self, source, url, lm, length):
        r = self.rows.get(source)
        return bool(r) and str(r.get("url")) == str(url) and str(r.get("last_modified")) == str(lm) \
            and str(r.get("length")) == str(length)

    def set(self, source, url, lm, length, sha, note=""):
        self.rows[source] = {"source": source, "url": url, "last_modified": lm, "length": length, "sha256": sha,
                             "fetched_utc": datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M"), "note": note}

    def frame(self):
        cols = ["source", "url", "last_modified", "length", "sha256", "fetched_utc", "note"]
        return pd.DataFrame(list(self.rows.values()), columns=cols)


def fetch_if_new(rel, source, url, force=False, samples=None, sample_name=None):
    """Return the file bytes when the source changed (or force), else None. In --samples mode read the local copy."""
    if samples:
        b = open(os.path.join(samples, sample_name), "rb").read()
        rel.set(source, f"sample:{sample_name}", "", len(b), hashlib.sha256(b).hexdigest()[:16])
        return b
    lm, length = "", ""
    try:
        h = http("HEAD", url, tries=2, allow_redirects=True)
        lm, length = h.headers.get("Last-Modified", ""), h.headers.get("Content-Length", "")
    except RuntimeError as e:
        print(f"  HEAD failed for {source} ({e}); will download")
    if not force and lm and rel.same(source, url, lm, length):
        print(f"  {source}: unchanged ({lm})")
        return None
    if not force and not lm and source in rel.rows and str(rel.rows[source].get("url")) == str(url):
        print(f"  {source}: same address, no Last-Modified header - kept")
        return None
    r = http("GET", url)
    b = r.content
    rel.set(source, url, lm or r.headers.get("Last-Modified", ""), length or str(len(b)), hashlib.sha256(b).hexdigest()[:16])
    print(f"  {source}: downloaded {len(b):,} bytes ({url})")
    return b


def gov_attachments(api_url):
    """{title: url} of the attachments of a gov.uk publication (content API)."""
    j = http("GET", api_url).json()
    out = {}

    def walk(o):
        if isinstance(o, dict):
            u = o.get("url")
            if isinstance(u, str) and re.search(r"\.(xlsx|csv|ods|xls)$", u, re.I) and o.get("title"):
                out.setdefault(str(o["title"]), u)
            for v in o.values():
                walk(v)
        elif isinstance(o, list):
            for v in o:
                walk(v)
    walk(j)
    return out


# ----------------------------------------------------------------------------------------------- UK parsers
def parse_et61(b, sheet, prefix=""):
    """ET 6.1 'Quarter' (UK) / 'Northern Ireland - Qtr': cumulative installed capacity table -> frame indexed by the
    first day of the quarter's LAST month (2026-03-01 = end of March 2026, Q1), MW, `[x]` -> blank."""
    d = pd.read_excel(io.BytesIO(b), sheet_name=sheet, header=None)
    r0 = next(i for i in range(len(d)) if str(d.iat[i, 0]).strip().upper().startswith("CUMULATIVE INSTALLED CAPACITY"))
    cols = {}
    for j in range(1, d.shape[1]):
        m = re.search(r"(\d{4})\s+(\d)\w*\s+quarter", str(d.iat[r0, j]), re.I)
        if m:
            cols[j] = pd.Timestamp(int(m.group(1)), 3 * int(m.group(2)), 1)
    rows = {}
    for i in range(r0 + 1, len(d)):
        lab = d.iat[i, 0]
        if pd.isna(lab):
            break
        lab = clean_label(lab)
        if lab.lower().startswith("co-firing"):
            continue
        rows[col_name(lab, prefix)] = [num(d.iat[i, j]) for j in cols]
        if lab.upper() == "TOTAL":
            break
    out = pd.DataFrame(rows, index=pd.DatetimeIndex(list(cols.values()), name="date")).sort_index()
    return out


def et61_standard(q, prefix=""):
    """Standard fuel columns from the ET 6.1 technology columns."""
    g = lambda *names: q[[prefix + n for n in names if prefix + n in q]].sum(axis=1, min_count=1)  # noqa: E731
    s = pd.DataFrame(index=q.index)
    s["Wind_Onshore_MW"] = g("Onshore_wind_MW")
    s["Wind_Offshore_MW"] = g("Offshore_wind_seabed_MW", "Offshore_wind_floating_MW")
    s["Wind_MW"] = s["Wind_Onshore_MW"].add(s["Wind_Offshore_MW"], fill_value=0)
    s["Solar_MW"] = g("Solar_photovoltaics_MW")
    s["Hydro_MW"] = g("Small_scale_hydro_MW", "Large_scale_hydro_MW")
    s["Bioenergy_MW"] = g("Landfill_gas_MW", "Sewage_sludge_digestion_MW", "Energy_from_waste_MW",
                          "Animal_Biomass_non_AD_MW", "Anaerobic_digestion_MW", "Plant_biomass_MW", "Liquid_biofuels_MW")
    s["Other_MW"] = g("Shoreline_wave_tidal_MW")
    s["Renewables_total_MW"] = q[prefix + "TOTAL_MW"] if prefix + "TOTAL_MW" in q else np.nan
    s.index.name = "date"
    return s


def parse_et63(b):
    """ET 6.3 Table_1_by_Capacity: cumulative solar PV capacity (MW) monthly, GB / NI / UK totals + UK domestic."""
    d = pd.read_excel(io.BytesIO(b), sheet_name="Table_1_by_Capacity", header=None)
    r0 = next(i for i in range(len(d)) if str(d.iat[i, 0]).strip().upper().startswith("CUMULATIVE CAPACITY"))
    cols = {}
    for j in range(1, d.shape[1]):
        m = re.search(r"([A-Za-z]{3})\s+(\d{4})", str(d.iat[r0, j]).replace("\n", " "))
        if m:
            cols[j] = pd.Timestamp(datetime.strptime(f"{m.group(1)} {m.group(2)}", "%b %Y"))
    end = next((i for i in range(r0 + 1, len(d)) if str(d.iat[i, 0]).strip().upper().startswith("CUMULATIVE COUNT")), len(d))
    out = {}
    block = None
    for i in range(r0 + 1, end):
        lab = clean_label(d.iat[i, 0]) if not pd.isna(d.iat[i, 0]) else ""
        if lab in ("GB", "NI", "UK"):
            block = lab
        elif block and lab.upper() == "TOTAL":
            out[f"Solar_{block}_MW"] = [num(d.iat[i, j]) for j in cols]
        elif block == "UK" and lab.lower().startswith("of which: domestic"):
            out["Solar_UK_domestic_MW"] = [num(d.iat[i, j]) for j in cols]
    return pd.DataFrame(out, index=pd.DatetimeIndex(list(cols.values()), name="date")).sort_index()


def parse_dukes_512(b):
    """DUKES 5.12.C (UK, 'All' rows): installed capacity by fuel at the end of each year -> frame indexed by year."""
    d = pd.read_excel(io.BytesIO(b), sheet_name="5.12", header=None)
    r = next(i for i in range(len(d)) if str(d.iat[i, 0]).startswith("Table 5.12.C"))
    hdr = d.iloc[r + 1]
    years = {j: int(float(hdr.iat[j])) for j in range(2, d.shape[1]) if not pd.isna(hdr.iat[j])}
    rows, in_all = {}, False
    for i in range(r + 2, len(d)):
        net, lab = d.iat[i, 0], d.iat[i, 1]
        if pd.isna(lab):
            continue
        if str(net).strip() == "All":
            in_all = True
        if in_all and (str(net).strip() == "All" or pd.isna(net)):
            rows[col_name(lab, "UK_")] = [num(d.iat[i, j]) for j in years]
    out = pd.DataFrame(rows, index=pd.Index(list(years.values()), name="year")).sort_index()
    g = lambda *n: out[[f"UK_{x}_MW" for x in n if f"UK_{x}_MW" in out]].sum(axis=1, min_count=1)  # noqa: E731
    s = pd.DataFrame(index=out.index)
    s["Coal_MW"] = g("Coal")
    s["Oil_MW"] = g("Oil")
    s["Gas_MW"] = g("OCGT_and_conventional_thermal_gas", "CCGT_gas")
    s["Nuclear_MW"] = g("Nuclear_Magnox", "Nuclear_PWR", "Nuclear_AGR")
    s["Hydro_MW"] = g("Hydro_natural_flow")
    s["Wind_Onshore_MW"] = g("Wind_onshore")
    s["Wind_Offshore_MW"] = g("Wind_offshore")
    s["Wind_MW"] = s["Wind_Onshore_MW"] + s["Wind_Offshore_MW"]
    s["Solar_MW"] = g("Solar")
    s["Bioenergy_MW"] = g("Bioenergy")
    s["PumpedStorage_MW"] = g("Pumped_storage")
    s["Other_MW"] = g("Other_fuels", "Wave_and_tidal")
    s["Total_MW"] = g("Total")
    return s


def parse_dukes_516(b):
    d = pd.read_excel(io.BytesIO(b), sheet_name="5.16", header=None)
    r = next(i for i in range(len(d)) if str(d.iat[i, 0]).startswith("Table 5.16.C"))
    hdr = d.iloc[r + 1]
    years = {j: int(float(hdr.iat[j])) for j in range(1, d.shape[1]) if not pd.isna(hdr.iat[j])}
    out = {}
    for i in range(r + 2, len(d)):
        lab = clean_label(d.iat[i, 0]) if not pd.isna(d.iat[i, 0]) else ""
        if lab.lower().startswith("battery storage"):
            out["Storage_MW"] = [num(d.iat[i, j]) for j in years]
        elif lab.lower().startswith("pumped storage"):
            out["PumpedStorage_5_16_MW"] = [num(d.iat[i, j]) for j in years]
    return pd.DataFrame(out, index=pd.Index(list(years.values()), name="year")).sort_index()


REPD_GROUP = {
    "Wind Onshore": "Wind_Onshore", "Wind Offshore": "Wind_Offshore", "Solar Photovoltaics": "Solar",
    "Battery": "Storage", "Pumped Storage Hydroelectricity": "PumpedStorage",
    "Large Hydro": "Hydro", "Small Hydro": "Hydro",
    "Biomass (dedicated)": "Bioenergy", "Biomass (co-firing)": "Bioenergy", "Anaerobic Digestion": "Bioenergy",
    "Landfill Gas": "Bioenergy", "EfW Incineration": "Bioenergy", "Sewage Sludge Digestion": "Bioenergy",
    "Advanced Conversion Technologies": "Bioenergy",
}


def parse_repd(b, release_name):
    """REPD: month-end cumulative capacity (MW) of projects with status Operational, by Operational date.
    Months after the release's quarter end are left blank."""
    d = pd.read_csv(io.BytesIO(b), encoding="cp1252", low_memory=False)
    d.columns = [str(c).strip() for c in d.columns]
    op = d[d["Development Status (short)"].astype(str).str.strip() == "Operational"].copy()
    op["mw"] = op["Installed Capacity (MWelec)"].map(num)
    op["dt"] = pd.to_datetime(op["Operational"], dayfirst=True, errors="coerce")
    undated = op["dt"].isna() & op["mw"].notna()
    op = op.dropna(subset=["dt", "mw"])
    op["grp"] = op["Technology Type"].map(REPD_GROUP).fillna("Other")
    m = re.search(r"Q([1-4])[_ ]?(\d{4})", release_name)
    cutoff = pd.Timestamp(int(m.group(2)), 3 * int(m.group(1)), 1) if m else op["dt"].max().to_period("M").to_timestamp()
    first = pd.Timestamp("2010-01-01")
    idx = pd.date_range(first, cutoff, freq="MS", name="date")
    op["m"] = op["dt"].dt.to_period("M").dt.to_timestamp().clip(lower=first)
    op = op[op["m"] <= cutoff]
    piv = op.pivot_table(index="m", columns="grp", values="mw", aggfunc="sum").reindex(idx).fillna(0.0).cumsum()
    piv.columns = [f"{c}_MW" for c in piv.columns]
    piv["Wind_MW"] = piv.get("Wind_Onshore_MW", 0) + piv.get("Wind_Offshore_MW", 0)
    note = (f"cutoff {cutoff:%Y-%m}; {len(op):,} operational rows with date; {int(undated.sum())} operational rows "
            f"without an Operational date left out ({d['Installed Capacity (MWelec)'].map(num)[undated.reindex(d.index, fill_value=False)].sum():.0f} MW)"
            if False else f"cutoff {cutoff:%Y-%m}; {len(op):,} dated operational rows; {int(undated.sum())} without a date left out")
    return piv.round(3), note


def uk_monthly(store, rel):
    """Assemble the UK 'Monthly' sheet from the stored source sheets."""
    s63 = store.get("ET63_solar_monthly")
    rp = store.get("REPD_monthly")
    idx = pd.date_range("2010-01-01", max([x.index.max() for x in (s63, rp) if x is not None] or [pd.Timestamp("2010-01-01")]),
                        freq="MS", name="date")
    m = pd.DataFrame(index=idx)
    for c in ["Hydro", "Gas", "Wind_Onshore", "Wind_Offshore", "Wind", "Solar", "Coal", "Nuclear", "Oil", "Bioenergy", "Other",
              "PumpedStorage", "Storage"]:
        m[f"{c}_MW"] = np.nan
    if rp is not None:
        # REPD hydro (35% of ET 6.1: the database lacks the old large hydro) and 'other' stay on REPD_monthly only
        for c in ["Wind_Onshore", "Wind_Offshore", "Wind", "Bioenergy", "PumpedStorage", "Storage"]:
            if f"{c}_MW" in rp:
                m[f"{c}_MW"] = rp[f"{c}_MW"].reindex(idx)
    if s63 is not None and "Solar_UK_MW" in s63:
        m["Solar_MW"] = s63["Solar_UK_MW"].reindex(idx)
    return m.dropna(how="all", axis=0).round(3)


def uk_check(store):
    """REPD rebuild and ET 6.3 against the official quarterly ET 6.1 at each quarter end (2021 on); DUKES vs ET at year end."""
    q = store.get("Quarterly_ET61")
    m = store.get("Monthly")
    rows = []
    if q is not None and m is not None:
        for d in q.index[q.index >= "2021-01-01"]:
            if d not in m.index:
                continue
            for fuel in ["Wind_Onshore", "Wind_Offshore", "Solar", "Hydro", "Bioenergy"]:
                a, b = m.at[d, f"{fuel}_MW"], q.at[d, f"{fuel}_MW"] if f"{fuel}_MW" in q else np.nan
                rows.append({"date": d, "fuel": fuel, "monthly_sheet_MW": a, "ET61_quarterly_MW": b,
                             "difference_MW": a - b if pd.notna(a) and pd.notna(b) else np.nan,
                             "ratio": a / b if pd.notna(a) and pd.notna(b) and b else np.nan,
                             "basis": "ET 6.3 monthly" if fuel == "Solar" else "REPD rebuild"})
    c1 = pd.DataFrame(rows)
    rows = []
    du = store.get("Annual_DUKES")
    if du is not None and q is not None:
        for y in du.index[du.index >= 2021]:
            d = pd.Timestamp(int(y), 12, 1)
            for fuel in ["Wind_Onshore", "Wind_Offshore", "Solar", "Hydro", "Bioenergy"]:
                a, b = du.at[y, f"{fuel}_MW"], q.at[d, f"{fuel}_MW"] if d in q.index and f"{fuel}_MW" in q else np.nan
                rows.append({"year_end": int(y), "fuel": fuel, "DUKES_5_12_MW": a, "ET61_quarterly_MW": b,
                             "difference_MW": a - b if pd.notna(a) and pd.notna(b) else np.nan})
            if m is not None and d in m.index:
                for fuel in ["Wind_Onshore", "Wind_Offshore", "Solar", "Hydro", "Bioenergy", "PumpedStorage", "Storage"]:
                    if fuel in ("PumpedStorage", "Storage"):
                        a = du.at[y, "PumpedStorage_MW"] if fuel == "PumpedStorage" else \
                            (du.at[y, "Storage_MW"] if "Storage_MW" in du else np.nan)
                    else:
                        a = du.at[y, f"{fuel}_MW"]
                    b = m.at[d, f"{fuel}_MW"]
                    rows.append({"year_end": int(y), "fuel": fuel + " (Monthly sheet)", "DUKES_5_12_MW": a,
                                 "ET61_quarterly_MW": b, "difference_MW": b - a if pd.notna(a) and pd.notna(b) else np.nan})
    c2 = pd.DataFrame(rows)
    return c1, c2


def build_uk(out_dir, rel_in, force, samples):
    path = os.path.join(out_dir, UK_FILE)
    store = read_store(path)
    rel = Releases(store)
    changed = False
    # ---- ET 6.1 / 6.3
    try:
        if samples:
            u61 = u63 = None
        else:
            att = gov_attachments(ET_API)
            u61 = next(u for t, u in att.items() if "ET 6.1" in t and u.endswith(".xlsx"))
            u63 = next(u for t, u in att.items() if "ET 6.3" in t and u.endswith(".xlsx"))
        b = fetch_if_new(rel, "ET 6.1 quarterly renewables", u61, force, samples, "ET_6.1.xlsx")
        if b:
            q = parse_et61(b, "Quarter")
            qni = parse_et61(b, "Northern Ireland - Qtr", "NI_")
            std = et61_standard(q)
            stdni = et61_standard(qni, "NI_")
            store["Quarterly_ET61"] = merge_replace(store.get("Quarterly_ET61"), pd.concat([std, q], axis=1))
            store["Quarterly_ET61_NI"] = merge_replace(store.get("Quarterly_ET61_NI"), pd.concat([stdni, qni], axis=1))
            changed = True
        b = fetch_if_new(rel, "ET 6.3 solar PV monthly", u63, force, samples, "ET_6.3_solar.xlsx")
        if b:
            store["ET63_solar_monthly"] = merge_replace(store.get("ET63_solar_monthly"), parse_et63(b))
            changed = True
    except Exception as e:  # noqa: BLE001
        print(f"  ET 6.x failed: {type(e).__name__}: {e}")
    # ---- REPD
    try:
        if samples:
            url, name = None, "REPD_Publication_Q2_2026.csv"
        else:
            html = http("GET", REPD_PAGE).text
            links = re.findall(r'href="(https://assets\.publishing\.service\.gov\.uk/[^"]+\.csv)"', html)
            if not links:
                raise RuntimeError("no REPD csv link on the page")
            url = links[0]
            name = os.path.basename(url)
        b = fetch_if_new(rel, "REPD operational projects", url, force, samples, "REPD.csv")
        if b:
            piv, note = parse_repd(b, name)
            store["REPD_monthly"] = piv
            rel.rows["REPD operational projects"]["note"] = note
            changed = True
    except Exception as e:  # noqa: BLE001
        print(f"  REPD failed: {type(e).__name__}: {e}")
    # ---- DUKES
    try:
        if samples:
            u512 = u516 = None
        else:
            att = gov_attachments(DUKES_API)
            u512 = next(u for t, u in att.items() if "DUKES 5.12" in t)
            u516 = next(u for t, u in att.items() if "DUKES 5.16" in t)
        b = fetch_if_new(rel, "DUKES 5.12 installed capacity", u512, force, samples, "DUKES_5.12.xlsx")
        b2 = fetch_if_new(rel, "DUKES 5.16 energy storage", u516, force, samples, "DUKES_5.16.xlsx")
        if b or b2:
            old = store.get("Annual_DUKES")
            parts = []
            if b:
                parts.append(parse_dukes_512(b))
            elif old is not None:
                parts.append(old[[c for c in old.columns if c not in ("Storage_MW",)]])
            if b2:
                parts.append(parse_dukes_516(b2)[["Storage_MW"]])
            elif old is not None and "Storage_MW" in old:
                parts.append(old[["Storage_MW"]])
            du = pd.concat(parts, axis=1)
            du = du[du["Total_MW"].notna()] if "Total_MW" in du else du
            store["Annual_DUKES"] = merge_replace(old, du)
            changed = True
    except Exception as e:  # noqa: BLE001
        print(f"  DUKES failed: {type(e).__name__}: {e}")
    if "Annual_DUKES" in store:
        store["Annual_DUKES"] = store["Annual_DUKES"].drop(columns=["date_1_Jan"], errors="ignore")
    if "Annual_DUKES" in store:   # 1 January layout used by the ENTSO-E workbooks: the stock at the end of the previous year
        du = store["Annual_DUKES"].copy()
        du.insert(0, "date_1_Jan", pd.to_datetime((du.index + 1).astype(str) + "-01-01"))
        store["Annual_DUKES"] = du
    # ---- assemble
    store["Monthly"] = uk_monthly(store, rel)
    c1, c2 = uk_check(store)
    sheets = {"Monthly": store["Monthly"]}
    for k in ["Quarterly_ET61", "Quarterly_ET61_NI", "ET63_solar_monthly", "REPD_monthly", "Annual_DUKES"]:
        if k in store:
            sheets[k] = store[k]
    if not c1.empty:
        sheets["Check"] = c1.set_index("date")
    if not c2.empty:
        sheets["Check_annual"] = c2.set_index("year_end")
    sheets["Releases"] = rel.frame().set_index("source")
    notes, titles = uk_notes()
    xlsx_notes.write_workbook(path, sheets, notes, titles)
    print(f"  wrote {path} ({'updated' if changed else 'no source changed'}): " +
          ", ".join(f"{k} {len(v)}" for k, v in sheets.items()))
    return path


def uk_notes():
    T = ["UNITED KINGDOM INSTALLED CAPACITY (MONTHLY WHERE A SOURCE EXISTS)", "SHEETS", "SOURCES AND BASIS", "WHAT IS NOT AVAILABLE"]
    L = [T[0], "Great Britain plus Northern Ireland, MW. A row dated 2026-03-01 is the capacity at the END of March 2026.", "",
         T[1],
         "Monthly - month-end MW by fuel: Solar = DESNZ ET 6.3 (official, includes small and domestic); Wind (onshore / offshore),"
         " Bioenergy, PumpedStorage, Storage (batteries) = REPD rebuild (see below). Hydro, Gas, Coal, Nuclear, Oil, Other are blank "
         "(hydro: REPD lacks the older large hydro, use the quarterly ET 6.1 sheet).",
         "Quarterly_ET61 / Quarterly_ET61_NI - DESNZ Energy Trends 6.1, official cumulative renewable capacity by technology at each "
         "quarter end, UK and Northern Ireland (the quarterly cross-check and the best official wind / hydro / bioenergy series).",
         "ET63_solar_monthly - DESNZ ET 6.3, solar PV GB / NI / UK, UK domestic.",
         "REPD_monthly - month-end capacity of operational REPD projects by commissioning month, by technology group.",
         "Annual_DUKES - DUKES 5.12.C installed capacity by fuel at the END of each year (UK, all fuels incl. gas, coal, nuclear, oil) "
         "and 5.16.C grid-scale batteries; date_1_Jan = the 1-January date used by the ENTSO-E workbooks.",
         "Check / Check_annual - Monthly sheet against ET 6.1 at quarter ends and DUKES at year ends. Releases - source file address, "
         "Last-Modified, length, hash and the date it was fetched.", "",
         T[2],
         "ET 6.3 'Solar photovoltaics deployment' (gov.uk, monthly, about 6 weeks lag): RAW, official; cumulative capacity from FiT, RO, "
         "CfD, MCS, REPD and the Major Power Producers survey; revised as DESNZ receives data.",
         "REPD (Renewable Energy Planning Database, DESNZ, quarterly): projects of 150 kW and over. Capacity of every project with "
         "status Operational is added in its 'Operational' month. It is therefore the operational fleet TODAY by commissioning date: "
         "plant that was decommissioned or repowered is not in it (earlier months are understated), a repowered site carries its new "
         "capacity from its original date, and plant under 150 kW is missing. Compare the Check sheet: onshore wind runs about 6-8% "
         "below the official figure, offshore wind about 4% below, batteries about 40% below DUKES (small and distribution-connected sites), hydro 65% below (so hydro is not taken from REPD). Use as the monthly SHAPE, and ET 6.1 / "
         "DUKES for the level.",
         "ET 6.1 (quarterly, about a quarter in arrears): RAW, official; the last quarter is provisional.",
         "DUKES 5.12 / 5.16 (annual, July): RAW, official. Gas = OCGT and conventional thermal + CCGT; nuclear = Magnox + PWR + AGR.", "",
         T[3],
         "Gas, coal, nuclear and oil have no monthly or quarterly source. The Elexon BMU register (data.elexon.co.uk reference/bmunits) holds "
         "only today's registered capacity with no dates and still lists closed units (3.3 GW of 'COAL' in Oct 2026, the last coal "
         "station closed in 2024); the NESO TEC register lists live agreements only; Elexon dataset IGCA returned no rows. "
         "Energy-Charts has monthly installed power for Germany only. So thermal is annual (Annual_DUKES) and nothing is interpolated.",
         "Pre-2010 months are not built (REPD months start Jan 2010; earlier projects are in the Jan 2010 opening stock)."]
    return L, set(T)


# ----------------------------------------------------------------------------------------------- Ireland parsers
def norm(s):
    return re.sub(r"\s+", " ", str(s)).strip()


def parse_irish_date(v):
    if isinstance(v, (pd.Timestamp, datetime)):
        return pd.Timestamp(v)
    if v is None or (isinstance(v, float) and np.isnan(v)):
        return pd.NaT
    s = str(v).strip()
    m = re.match(r"^Pre[\s-]*(\d{4})$", s, re.I)
    if m:
        return pd.Timestamp(int(m.group(1)) - 1, 12, 31)
    if re.match(r"^\d{4}$", s):
        return pd.Timestamp(int(s), 1, 1)
    return pd.to_datetime(s, dayfirst=True, errors="coerce")


def eirgrid_wind_monthly(b):
    d = pd.read_excel(io.BytesIO(b), sheet_name="Wind Installed Capacities", header=None)
    sub = d.iloc[4:, 16:24].copy()
    sub.columns = ["date", "Wind_IE_TSO_MW", "Wind_IE_DSO_MW", "Wind_IE_MW", "Wind_NI_TSO_MW", "Wind_NI_DSO_MW",
                   "Wind_NI_MW", "Wind_AllIsland_MW"]
    sub["date"] = pd.to_datetime(sub["date"], errors="coerce")
    sub = sub.dropna(subset=["date"]).set_index("date").apply(pd.to_numeric, errors="coerce")
    sub.index = [month_start(x) for x in sub.index]
    sub.index.name = "date"
    return sub[sub.index >= "2010-01-01"]


def eirgrid_registers(b):
    """Connected-generator registers of 'IE Connected Renewables' and storage / thermal of 'IE Connected Non-Renewables'
    -> long frame of (group, name, mw, date, source) for month-end rebuild, plus the as-of month and thermal snapshot."""
    x = pd.ExcelFile(io.BytesIO(b))
    s = x.parse("IE Connected Renewables", header=None)
    titles = [(j, str(s.iat[0, j])) for j in range(s.shape[1]) if isinstance(s.iat[0, j], str) and "Generation in Ireland" in s.iat[0, j]]
    asof = None
    m = re.search(r"end of ([A-Za-z]{3})-(\d{4})", str(s.iat[2, 0]))
    if m:
        asof = pd.Timestamp(datetime.strptime(f"{m.group(1)} {m.group(2)}", "%b %Y"))
    recs = []
    for k, (j0, title) in enumerate(titles):
        j1 = titles[k + 1][0] if k + 1 < len(titles) else s.shape[1]
        blk = s.iloc[3:, j0:j1].copy()
        blk.columns = [norm(c) for c in s.iloc[3, j0:j1]]
        blk = blk.iloc[1:].reset_index(drop=True)
        dist = title.startswith("Distribution")
        kind = "Wind" if "Wind" in title else "Solar" if "Solar" in title else "Other renewable"
        cap = next((c for c in blk.columns if c.lower() in ("mec (mw)", "mec_mw")), None) if dist else \
            next((c for c in blk.columns if c.lower().startswith("installed capacity")), None)
        dcol = next((c for c in blk.columns if c.lower() == "energisation date"), None) or \
            next((c for c in blk.columns if c.lower().startswith("full / latest connection")), None)
        tcol = next((c for c in blk.columns if c.lower() == "type"), None)
        scol = next((c for c in blk.columns if c.lower() == "status"), None)
        for _, r in blk.iterrows():
            mw = num(r[cap]) if cap else np.nan
            dt = parse_irish_date(r[dcol]) if dcol else pd.NaT
            if np.isnan(mw) or pd.isna(dt):
                continue
            if dist and scol and norm(r[scol]).lower() != "energised":
                continue
            typ = norm(r[tcol]) if tcol and not pd.isna(r[tcol]) else kind
            recs.append({"kind": kind, "connection": "DSO" if dist else "TSO", "type": typ, "mw": mw, "date": dt})
    n = x.parse("IE Connected Non-Renewables", header=None)
    # transmission storage table: header row with 'Code'
    sr = next(i for i in range(len(n)) if norm(n.iat[i, 0]) == "Generator" and norm(n.iat[i, 1]) == "Code")
    for i in range(sr + 1, len(n)):
        if pd.isna(n.iat[i, 0]):
            break
        typ, mw, dt = norm(n.iat[i, 2]), num(n.iat[i, 3]), parse_irish_date(n.iat[i, 6])
        if not np.isnan(mw) and pd.notna(dt):
            recs.append({"kind": "Pumped storage" if "pump" in typ.lower() else "Battery", "connection": "TSO",
                         "type": typ, "mw": mw, "date": dt})
    # distribution table (right block): batteries and thermal
    hr = next(i for i in range(len(n)) if norm(n.iat[i, 12]) == "ESBNG Reference")
    hdr = [norm(c) for c in n.iloc[hr, 12:22]]
    for i in range(hr + 1, len(n)):
        if pd.isna(n.iat[i, 12]) and pd.isna(n.iat[i, 13]):
            break
        row = dict(zip(hdr, n.iloc[i, 12:22].tolist()))
        typ, mw = norm(row.get("Type", "")), num(row.get("MEC_MW"))
        if norm(row.get("Status", "")).lower() != "energised" or np.isnan(mw):
            continue
        dt = parse_irish_date(row.get("Energisation Date"))
        if "battery" in typ.lower() and pd.notna(dt):
            recs.append({"kind": "Battery", "connection": "DSO", "type": typ, "mw": mw, "date": dt})
        elif "flywheel" in typ.lower() and pd.notna(dt):
            recs.append({"kind": "Other storage", "connection": "DSO", "type": typ, "mw": mw, "date": dt})
    reg = pd.DataFrame(recs)
    # thermal snapshot (transmission list + distribution list), MW by fuel family at the release date
    th = []
    for i in range(4, sr):
        if pd.isna(n.iat[i, 0]):
            if th:
                break
            continue
        if norm(n.iat[i, 0]) in ("Interconnection", "Generator"):
            break
        th.append((norm(n.iat[i, 1]), num(n.iat[i, 2])))
    for i in range(hr + 1, len(n)):
        if pd.isna(n.iat[i, 12]) and pd.isna(n.iat[i, 13]):
            break
        t = norm(n.iat[i, 14])
        if norm(n.iat[i, 19]).lower() == "energised" and not any(w in t.lower() for w in ("battery", "flywheel")):
            th.append((t, num(n.iat[i, 15])))
    fam = {"gas": "Gas", "ccgt": "Gas", "ocgt": "Gas", "chp": "Gas", "nat. gas": "Gas", "rice": "Gas", "gas/do": "Gas",
           "coal": "Coal", "hfo": "Oil", "distillate": "Oil", "diesel": "Oil", "do": "Oil", "peat": "Peat"}
    tot = {}
    for t, mw in th:
        if np.isnan(mw):
            continue
        f = fam.get(t.lower(), "Other thermal")
        tot[f] = tot.get(f, 0.0) + mw
    snap = pd.DataFrame([{"date": asof, **{f"{k}_MW": round(v, 2) for k, v in sorted(tot.items())}}]).set_index("date") \
        if asof is not None else pd.DataFrame()
    return reg, asof, snap


def rebuild_register(reg, asof, groups):
    """Month-end cumulative MW per group from the register (month of date; a date before 2010 is in the opening stock)."""
    first = pd.Timestamp("2010-01-01")
    idx = pd.date_range(first, asof, freq="MS", name="date")
    out = pd.DataFrame(index=idx)
    for name, sel in groups.items():
        r = reg[sel(reg)]
        mo = r["date"].dt.to_period("M").dt.to_timestamp().clip(lower=first)
        mo = mo[mo <= asof]
        out[name] = r.loc[mo.index, "mw"].groupby(mo).sum().reindex(idx).fillna(0.0).cumsum()
    return out.round(3)


IE_GROUPS = {
    "Solar_MW": lambda r: r["kind"] == "Solar",
    "Solar_TSO_MW": lambda r: (r["kind"] == "Solar") & (r["connection"] == "TSO"),
    "Hydro_MW": lambda r: (r["kind"] == "Other renewable") & r["type"].str.lower().str.contains("hydro"),
    "Bioenergy_MW": lambda r: (r["kind"] == "Other renewable") & ~r["type"].str.lower().str.contains("hydro"),
    "Storage_MW": lambda r: r["kind"] == "Battery",
    "PumpedStorage_MW": lambda r: r["kind"] == "Pumped storage",
    "Wind_register_TSO_MW": lambda r: (r["kind"] == "Wind") & (r["connection"] == "TSO"),
    "Wind_register_DSO_MW": lambda r: (r["kind"] == "Wind") & (r["connection"] == "DSO"),
}


def build_ie(out_dir, force, samples):
    path = os.path.join(out_dir, IE_FILE)
    store = read_store(path)
    rel = Releases(store)
    changed = False
    try:
        if samples:
            url = None
        else:
            html = http("GET", EIRGRID_PAGE).text
            links = re.findall(r'href="(https://cms\.eirgrid\.ie/sites/default/files/publications/'
                               r'System-and-Renewable-Data-Summary-Report[^"]*\.xlsx)"', html)
            if not links:
                raise RuntimeError("no EirGrid summary report link")
            url = links[0]
        b = fetch_if_new(rel, "EirGrid System and Renewable Data Summary Report", url, force, samples, "eirgrid_summary.xlsx")
        if b:
            wind = eirgrid_wind_monthly(b)
            reg, asof, snap = eirgrid_registers(b)
            store["EirGrid_wind_monthly"] = merge_replace(store.get("EirGrid_wind_monthly"), wind)
            rb = rebuild_register(reg, asof, IE_GROUPS)
            store["Register_monthly"] = rb
            if not snap.empty:
                store["Thermal_snapshots"] = merge_replace(store.get("Thermal_snapshots"), snap)
            rel.rows["EirGrid System and Renewable Data Summary Report"]["note"] = \
                f"registers as of {asof:%Y-%m}; {len(reg)} dated register rows"
            changed = True
    except Exception as e:  # noqa: BLE001
        print(f"  EirGrid failed: {type(e).__name__}: {e}")
    w, rb = store.get("EirGrid_wind_monthly"), store.get("Register_monthly")
    if w is not None and rb is not None:
        idx = pd.date_range("2010-01-01", max(w.index.max(), rb.index.max()), freq="MS", name="date")
        m = pd.DataFrame(index=idx)
        for c in ["Hydro", "Gas", "Wind", "Solar", "Coal", "Nuclear", "Oil", "Bioenergy", "Other", "PumpedStorage", "Storage"]:
            m[f"{c}_MW"] = np.nan
        m["Wind_MW"] = w["Wind_IE_MW"].reindex(idx)   # TSO + DSO; blank where EirGrid has not yet posted the DSO part
        m["Wind_TSO_MW"] = w["Wind_IE_TSO_MW"].reindex(idx)
        m["Wind_DSO_MW"] = w["Wind_IE_DSO_MW"].reindex(idx)
        for c in ["Solar_MW", "Hydro_MW", "Bioenergy_MW", "Storage_MW", "PumpedStorage_MW"]:
            m[c] = rb[c].reindex(idx)
        store["Monthly"] = m.dropna(how="all").round(3)
        ni = pd.DataFrame(index=idx)
        ni["Wind_NI_TSO_MW"] = w["Wind_NI_TSO_MW"].reindex(idx)
        ni["Wind_NI_DSO_MW"] = w["Wind_NI_DSO_MW"].reindex(idx)
        ni["Wind_NI_MW"] = w["Wind_NI_MW"].reindex(idx)
        s63 = None
        try:   # NI solar monthly lives in the UK workbook (ET 6.3)
            s63 = pd.read_excel(os.path.join(out_dir, UK_FILE), sheet_name="ET63_solar_monthly")
            s63["date"] = pd.to_datetime(s63["date"])
            ni["Solar_NI_MW"] = s63.set_index("date")["Solar_NI_MW"].reindex(idx)
        except Exception as e:  # noqa: BLE001
            print(f"  NI solar not added ({type(e).__name__})")
        store["Monthly_NI"] = ni.dropna(how="all").round(3)
        # cross-checks
        chk = pd.DataFrame(index=idx)
        chk["Wind_TSO_series_MW"] = w["Wind_IE_TSO_MW"].reindex(idx)
        chk["Wind_TSO_register_MW"] = rb["Wind_register_TSO_MW"].reindex(idx)
        chk["TSO_series_minus_register_MW"] = chk["Wind_TSO_series_MW"] - chk["Wind_TSO_register_MW"]
        chk["Wind_DSO_series_MW"] = w["Wind_IE_DSO_MW"].reindex(idx)
        chk["Wind_DSO_register_MW"] = rb["Wind_register_DSO_MW"].reindex(idx)
        store["Check_wind"] = chk[chk.index >= "2015-01-01"].dropna(how="all").round(2)
        store["Check_ENTSOE"] = ie_vs_entsoe(out_dir, w, m, ni)
    sheets = {k: store[k] for k in ["Monthly", "Monthly_NI", "EirGrid_wind_monthly", "Register_monthly", "Thermal_snapshots",
                                    "Check_wind", "Check_ENTSOE"] if k in store and store[k] is not None and len(store[k])}
    sheets["Releases"] = rel.frame().set_index("source")
    notes, titles = ie_notes()
    xlsx_notes.write_workbook(path, sheets, notes, titles)
    print(f"  wrote {path} ({'updated' if changed else 'no source changed'}): " +
          ", ".join(f"{k} {len(v)}" for k, v in sheets.items()))
    return path


def ie_vs_entsoe(out_dir, w, m, ni):
    """1 January values of the ENTSO-E SEM workbook (all-island) against EirGrid at the previous December."""
    p = os.path.join(out_dir, "ireland_sem_power_capacity.xlsx")
    if not os.path.exists(p):
        return pd.DataFrame()
    e = pd.read_excel(p, sheet_name="Monthly")
    e["date"] = pd.to_datetime(e["date"])
    rows = []
    for _, r in e.iterrows():
        dec = (r["date"] - pd.DateOffset(months=1)).to_period("M").to_timestamp()
        if dec not in w.index:
            continue
        wi = w.at[dec, "Wind_AllIsland_MW"]
        wie = w.at[dec, "Wind_IE_MW"]
        rows.append({"entsoe_date": r["date"], "ENTSOE_SEM_Wind_MW": r.get("Wind_MW"), "EirGrid_allisland_wind_Dec_MW": wi,
                     "EirGrid_IE_wind_Dec_MW": wie, "difference_allisland_MW": (wi - r.get("Wind_MW")) if pd.notna(wi) else np.nan,
                     "ENTSOE_SEM_over_EirGrid_IE_wind": r.get("Wind_MW") / wie if pd.notna(wie) and wie else np.nan,
                     "ENTSOE_SEM_Solar_MW": r.get("Solar_MW"), "register_IE_solar_Dec_MW": m.at[dec, "Solar_MW"] if dec in m.index else np.nan,
                     "ENTSOE_SEM_Gas_MW": r.get("Gas_MW"), "ENTSOE_SEM_Coal_MW": r.get("Coal_MW"), "ENTSOE_SEM_Oil_MW": r.get("Oil_MW")})
    return pd.DataFrame(rows).set_index("entsoe_date")


def ie_notes():
    T = ["IRELAND INSTALLED CAPACITY (MONTHLY WHERE A SOURCE EXISTS)", "SHEETS", "SOURCES AND BASIS", "WHAT IS NOT AVAILABLE"]
    L = [T[0], "Republic of Ireland, MW at the END of the month shown (2026-03-01 = end of March 2026). Northern Ireland on 'Monthly_NI'.", "",
         T[1],
         "Monthly - Wind = EirGrid monthly installed wind capacity, TSO + DSO (blank where EirGrid has posted the transmission part "
         "but not yet the distribution part); Solar, Hydro, Bioenergy (biomass, waste to energy, landfill gas), Storage (batteries) and "
         "PumpedStorage (Turlough Hill) rebuilt from the registers. Gas, Coal, Oil, Nuclear are blank (see below).",
         "Monthly_NI - Northern Ireland: wind TSO / DSO / total as EirGrid and SONI publish them (the DSO part and total only at year ends), "
         "solar monthly from DESNZ ET 6.3 (United Kingdom workbook).",
         "EirGrid_wind_monthly - the source series (IE, NI, all-island). Register_monthly - month-end rebuild of every register group.",
         "Thermal_snapshots - transmission- plus distribution-connected thermal MEC by fuel at each EirGrid release (appended; the "
         "registers carry a connection year or 'Pre 2000' only).",
         "Check_wind - EirGrid's monthly TSO / DSO wind against the sum of the register by energisation date. Check_ENTSOE - the ENTSO-E "
         "SEM workbook's 1-January values (all-island) against EirGrid at the previous December. Releases - file address, "
         "Last-Modified, length, hash.", "",
         T[2],
         "EirGrid 'System and Renewable Data Summary Report' (xlsx on eirgrid.ie, updated about quarterly; tabs 'Wind Installed Capacities', "
         "'IE Connected Renewables', 'IE Connected Non-Renewables'). Distribution-connected projects are ESB Networks' quarterly generator "
         "statistics, copied into the same workbook. Energisation dates are EirGrid's; ESBN marks some as best estimates.",
         "The register rebuilds list projects connected TODAY: a plant that has closed is not in them, and rooftop / small-scale solar "
         "(thousands of sites under the MEC cut-off) is not listed, so Solar is utility-scale and DSO-connected projects only. "
         "Wind is the official series and is not affected.", "",
         T[3],
         "Gas, coal, oil and peat plant: connection YEAR only, so no month-end history; see Thermal_snapshots. Northern Ireland "
         "thermal and other fuels: only annual (DUKES 5.12.B) and quarterly renewables (ET 6.1, United Kingdom workbook)."]
    return L, set(T)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out-dir", default=OUT_DEFAULT)
    ap.add_argument("--only", choices=["uk", "ie"])
    ap.add_argument("--force", action="store_true")
    ap.add_argument("--samples")
    a = ap.parse_args()
    os.makedirs(a.out_dir, exist_ok=True)
    if a.only in (None, "uk"):
        build_uk(a.out_dir, None, a.force, a.samples)
    if a.only in (None, "ie"):
        build_ie(a.out_dir, a.force, a.samples)


if __name__ == "__main__":
    main()
