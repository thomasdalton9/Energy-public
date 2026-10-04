"""
Biomethane (upgraded biogas) injected into gas grids - raw operator / statistical data, one workbook:

  output/Data and Chart Outputs/europe_biomethane_statistics.xlsx
    sheet "Monthly"    : month, GWh per month, one column per country where a monthly source exists
        GB_biomethane_GWh  DESNZ Energy Trends table 4.2 "Biomethane to grid" (official statistics, monthly, Jan 2019 on; from 2019 on the
                           revised methodology)
        AT_biomethane_GWh  AGGM data monitor (Austrian Gas Grid Management): biomethane allocated at the entry points of all three market
                           areas (series SummeBioOesterreich, else East + Tyrol + Vorarlberg entry series)
        SE_biomethane_GWh  Swedish Energy Agency, biogas upgraded and delivered to the gas grid (see SE section)
    sheet "Annual"     : year, TWh per year, one column per country: Eurostat energy balance nrg_bal_c, biogases (R5300) transformation
                         input "for blending with natural gas" (TI_BNG_E) - the biomethane blended into / injected into the natural gas
                         grid, ANNUAL ONLY, national statistics offices' submissions
    sheet "Crosscheck" : annual sum of the monthly national series next to the Eurostat figure (TWh)
    sheet "Units"      : source, unit and definition of every column (biomethane vs raw biogas vs biogas-to-power)

Incremental: the committed workbook is the history store. DESNZ ET 4.2 is a whole file, downloaded only when a new release
(new asset URL, recorded on the Units sheet) appears; AGGM is re-fetched from two months before the last month saved; Eurostat is a
small API call and is refreshed in full.

Usage: python3 BIOMETHANE_STATS.py [--out-dir DIR] [--only gb,at,se,eurostat] [--force]
"""
import argparse
import io
import os
import re
import sys
from datetime import date, datetime, timezone

import numpy as np
import pandas as pd
import requests

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
sys.path.insert(0, ROOT)
import xlsx_notes  # noqa: E402

OUT_DEFAULT = os.path.join(ROOT, "output", "Data and Chart Outputs")
FILE = "europe_biomethane_statistics.xlsx"
UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36"
H = {"User-Agent": UA}

ET_PAGE = "https://www.gov.uk/government/statistics/gas-section-4-energy-trends"
AT_URL = "https://platform.aggm.at/vis-service/api/ts/values"
AT_SERIES = ["SummeBioOesterreich", "EntryBiogasOst_MGM-Allokationen", "EntryBiogasTirol_MGM-Allokationen", "EntryBiogasVorarlberg_MGM-Allokationen"]
ESTAT = "https://ec.europa.eu/eurostat/api/dissemination/statistics/1.0/data/nrg_bal_c"
EU27 = ["AT", "BE", "BG", "HR", "CY", "CZ", "DK", "EE", "FI", "FR", "DE", "EL", "HU", "IE", "IT", "LV", "LT", "LU", "MT", "NL", "PL", "PT", "RO", "SK", "SI", "ES", "SE"]
ISO = {"EL": "GR"}     # Eurostat's code for Greece
ANNUAL_EXTRA = ["UK", "NO", "CH"]


def get(url, tries=3, **kw):
    for i in range(tries):
        try:
            r = requests.get(url, headers=H, timeout=kw.pop("timeout", 90), **kw)
            if r.ok:
                return r
        except requests.RequestException as e:
            print(f"  {type(e).__name__} on {url[:90]}")
    return None


# ---- Great Britain: DESNZ Energy Trends 4.2 -------------------------------------------------------------------------------
def gb(state):
    r = get(ET_PAGE)
    if r is None:
        raise RuntimeError("gov.uk page not reachable")
    m = re.search(r'href="(https://assets\.publishing\.service\.gov\.uk/media/[^"]+ET_4\.2[^"]*\.xlsx)"', r.text)
    if not m:
        raise RuntimeError("ET 4.2 link not found")
    url = m.group(1)
    if state.get("gb_url") == url and not state["force"] and "GB_biomethane_GWh" in state["monthly"]:
        print("GB: release unchanged, skipped:", url)
        return
    import openpyxl
    wb = openpyxl.load_workbook(io.BytesIO(get(url, timeout=120).content), data_only=True)
    ws = wb["Month (GWh)"]
    hdr_row = ci = None
    for i, row in enumerate(ws.iter_rows(min_row=1, max_row=20, values_only=True), start=1):
        for j, c in enumerate(row):
            if c and "biomethane" in str(c).lower():
                hdr_row, ci = i, j
        if hdr_row:
            break
    if not hdr_row:
        raise RuntimeError("'Biomethane to grid' column not found")
    recs = []
    for row in ws.iter_rows(min_row=hdr_row + 1, values_only=True):
        d, v = row[0], row[ci]
        if d is None:
            continue
        ts = pd.to_datetime(re.sub(r"^\[[^\]]*\]\s*", "", str(d).strip()), format="%B %Y", errors="coerce")
        if pd.isna(ts) or not isinstance(v, (int, float)):
            continue
        recs.append((ts.to_period("M").to_timestamp(), float(v)))
    s = pd.Series(dict(recs)).sort_index()
    s = s[s.index >= "2014-01-01"]
    print(f"GB: {len(s)} months {s.index.min():%Y-%m} .. {s.index.max():%Y-%m}; annual:",
          {int(y): round(v / 1000, 2) for y, v in s.groupby(s.index.year).sum().items() if y >= 2019})
    state["monthly"]["GB_biomethane_GWh"] = s
    state["gb_url"] = url


# ---- Austria: AGGM -------------------------------------------------------------------------------------------------------
def at(state):
    old = state["monthly"].get("AT_biomethane_GWh")
    start = pd.Timestamp("2019-01-01")
    if old is not None and len(old) and not state["force"]:
        start = max(start, old.index.max() - pd.DateOffset(months=2))
    today = pd.Timestamp(date.today())
    cols = {}
    y0 = start
    while y0 <= today:
        y1 = min(y0 + pd.DateOffset(years=1), today + pd.DateOffset(days=1))
        body = {"rangeType": "individual", "from": f"{y0:%Y-%m-%d}T06:00:00", "to": f"{y1:%Y-%m-%d}T06:00:00", "granularity": "month", "timeseries": AT_SERIES}
        j = requests.post(AT_URL, json=body, timeout=120, headers={**H, "Content-Type": "application/json", "Accept": "application/json"}).json()
        for cd in j["timeSeriesData"]["chartData"]:
            for p in cd["dataSet"]:
                if p.get("y") is None:
                    continue
                t = pd.Timestamp(p["x"], unit="ms", tz="UTC").tz_convert("Europe/Vienna").tz_localize(None)
                t = (t - pd.Timedelta(hours=6)).to_period("M").to_timestamp()      # gas day starts 06:00 local
                cols.setdefault(cd["header"]["name"], {})[t] = p["y"] / 1e6        # kWh -> GWh
        y0 = y1
    df = pd.DataFrame(cols).sort_index()
    if df.empty:
        raise RuntimeError("no AGGM data")
    parts = [c for c in AT_SERIES[1:] if c in df]
    total = df["SummeBioOesterreich"] if "SummeBioOesterreich" in df else pd.Series(dtype=float)
    entry = df[parts].sum(axis=1, min_count=1)
    s = total.combine_first(entry).dropna()
    # the current month is incomplete: drop it
    s = s[s.index < pd.Timestamp(date.today()).to_period("M").to_timestamp()]
    comb = s if old is None or state["force"] else s.combine_first(old)
    comb.update(s)
    state["monthly"]["AT_biomethane_GWh"] = comb.sort_index()
    print(f"AT: {len(comb)} months {comb.index.min():%Y-%m} .. {comb.index.max():%Y-%m}; annual GWh:",
          {int(y): round(v, 1) for y, v in comb.groupby(comb.index.year).sum().items()})


# ---- Eurostat annual ----------------------------------------------------------------------------------------------------
def eurostat(state):
    geos = EU27 + ANNUAL_EXTRA + ["EU27_2020"]
    r = get(ESTAT, params={"format": "JSON", "lang": "EN", "geo": geos, "siec": "R5300", "nrg_bal": "TI_BNG_E", "unit": "GWH", "freq": "A",
                           "sinceTimePeriod": "2010"})
    if r is None:
        raise RuntimeError("Eurostat not reachable")
    j = r.json()
    dims = j["id"]
    idx = {k: list(j["dimension"][k]["category"]["index"]) for k in dims}
    sizes = j["size"]
    out = {}
    for pos, v in j["value"].items():
        pos = int(pos)
        co = {}
        for s_, k in reversed(list(zip(sizes, dims))):
            co[k] = idx[k][pos % s_]
            pos //= s_
        out.setdefault(ISO.get(co["geo"], co["geo"]), {})[int(co["time"])] = v / 1000.0     # GWh -> TWh
    ann = pd.DataFrame(out).sort_index()
    ann.index.name = "year"
    ann = ann.rename(columns=lambda c: ("EU27" if c == "EU27_2020" else c) + "_biomethane_TWh")
    # EU27 total from the member states when Eurostat's aggregate is missing a year
    mem = [f"{ISO.get(c, c)}_biomethane_TWh" for c in EU27 if f"{ISO.get(c, c)}_biomethane_TWh" in ann]
    ann["EU27_sum_of_countries_TWh"] = ann[mem].sum(axis=1, min_count=1)
    state["annual"] = ann.round(4)
    print("Eurostat annual TWh (EU27):", ann["EU27_biomethane_TWh"].round(2).to_dict() if "EU27_biomethane_TWh" in ann else "n/a")
    for c in ("GB", "UK", "FR", "DE", "DK", "IT", "SE", "AT", "NL"):
        col = f"{c}_biomethane_TWh"
        if col in ann:
            print("  ", c, ann[col].dropna().round(2).to_dict())


# ---- Sweden ---------------------------------------------------------------------------------------------------------------
SE_PX = "https://pxexternal.energimyndigheten.se/api/v1/en/Energimyndighetens_statistikdatabas/Officiell_energistatistik/Produktion_av_biogas_och_rotrester/"


def px(table, pick):
    """Query a PX-Web table: pick = {variable code or text: [value texts]}; returns {(key texts...): value}."""
    meta = requests.get(SE_PX + table, headers=H, timeout=60).json()
    query = []
    for v in meta["variables"]:
        want = pick.get(v["code"]) or pick.get(v["text"])
        if want:
            codes = [c for c, t in zip(v["values"], v["valueTexts"]) if t in want]
        else:
            codes = list(v["values"])
        query.append({"code": v["code"], "selection": {"filter": "item", "values": codes}})
    r = requests.post(SE_PX + table, json={"query": query, "response": {"format": "json"}}, headers=H, timeout=60)
    r.raise_for_status()
    j = r.json()
    out = {}
    names = [dict(zip(v["values"], v["valueTexts"])) for v in meta["variables"]]
    for row in j["data"]:
        out[tuple(names[i].get(k, k) for i, k in enumerate(row["key"]))] = row["values"][0]
    return meta, j, out


def se(state):
    """Sweden, annual: biogas upgraded to biomethane (all uses, EN0124_2) and biogas fed into the gas networks by county (EN0124_4, 2023+)."""
    meta, j, o = px("EN0124_2.px", {"Användningsområde": ["Upgate"]})
    up = {}
    for k, v in o.items():
        try:
            up[int(k[0])] = float(v) / 1000.0
        except (ValueError, TypeError):
            pass
    meta4, j4, o4 = px("EN0124_4.px", {})
    feed = {}
    for k, v in o4.items():
        try:
            feed[int(k[0])] = feed.get(int(k[0]), 0.0) + float(v) / 1000.0
        except (ValueError, TypeError):
            pass
    df = pd.DataFrame({"SE_upgraded_biogas_all_uses_TWh": pd.Series(up), "SE_fed_into_gas_networks_TWh": pd.Series(feed)}).sort_index()
    df.index.name = "year"
    state["annual_nat"] = state["annual_nat"].combine_first(df) if len(state["annual_nat"]) else df
    state["annual_nat"].update(df)
    print("SE annual TWh:\n", df.tail(6).round(3).to_string())


# ---- workbook -------------------------------------------------------------------------------------------------------------
def read_state(path, force):
    st = {"monthly": {}, "annual": pd.DataFrame(), "annual_nat": pd.DataFrame(), "force": force}
    if not os.path.exists(path):
        return st
    def sheet(name, idx):
        try:
            d = pd.read_excel(path, sheet_name=name)
            return d.set_index(idx) if idx in d.columns else None
        except Exception as e:  # noqa: BLE001
            print(f"could not read sheet {name} ({type(e).__name__})")
    try:
        m = pd.read_excel(path, sheet_name="Monthly")
        m["month"] = pd.to_datetime(m["month"])
        for c, v in m.set_index("month").items():
            st["monthly"][c] = v.dropna()
    except Exception as e:  # noqa: BLE001
        print(f"could not read Monthly ({type(e).__name__})")
    a = sheet("Annual", "year")
    if a is not None:
        st["annual"] = a
    a = sheet("Annual national", "year")
    if a is not None:
        st["annual_nat"] = a
    try:
        u = pd.read_excel(path, sheet_name="Units")
        for line in u["Notes"].astype(str):
            mm = re.match(r"ET 4\.2 release used: (\S+)", line)
            if mm:
                st["gb_url"] = mm.group(1)
    except Exception as e:  # noqa: BLE001
        print(f"could not read Units ({type(e).__name__})")
    return st


NOTES = {
    "GB_biomethane_GWh": "Great Britain (UK national statistics): DESNZ Energy Trends table 4.2 'Biomethane to grid', GWh gross calorific value per month. Biomethane "
                         "injected into the gas grid (transmission and distribution), NOT raw biogas and NOT biogas burnt for power. Estimated for the latest "
                         "months; revised methodology from 2019.",
    "AT_biomethane_GWh": "Austria: AGGM data monitor, biomethane production allocated at the entry points of the three market areas (East, Tyrol, Vorarlberg), "
                         "GWh per month. Biomethane injected into the gas grid only; biogas used for power is not included.",
    "SE_biomethane_GWh": "Sweden: see Units text.",
}


def write(path, st, errors):
    monthly = pd.DataFrame(st["monthly"]).sort_index()
    monthly.index.name = "month"
    monthly = monthly.round(3)
    ann = st["annual"].copy()
    # crosscheck: annual sum of full-year monthly national series vs Eurostat
    cc = {}
    for c in monthly.columns:
        cc_ = c.split("_")[0]
        s = monthly[c].dropna()
        full = s.groupby(s.index.year).agg(["sum", "count"])
        full = full[full["count"] == 12]["sum"] / 1000.0
        cc[f"{cc_}_national_monthly_sum_TWh"] = full
        es = "UK" if cc_ == "GB" else cc_
        col = f"{es}_biomethane_TWh"
        if col in ann:
            cc[f"{cc_}_Eurostat_TWh"] = ann[col]
    cross = pd.DataFrame(cc).sort_index()
    cross.index.name = "year"
    cross = cross.dropna(how="all").round(3)
    lines = ["Europe - biomethane (upgraded biogas) injected into gas grids: raw operator and statistical data", "",
             "Source", "Great Britain: DESNZ Energy Trends 4.2 (https://www.gov.uk/government/statistics/gas-section-4-energy-trends). Austria: AGGM data monitor "
             "(https://platform.aggm.at). Annual all-country table: Eurostat energy balances nrg_bal_c (https://ec.europa.eu/eurostat/databrowser/view/nrg_bal_c).",
             "", "Units and definitions"]
    lines += [f"{k}: {v}" for k, v in NOTES.items() if k in monthly.columns]
    lines += ["SE_upgraded_biogas_all_uses_TWh / SE_fed_into_gas_networks_TWh (sheet Annual national): Swedish Energy Agency official statistics 'Production and use of "
              "biogas' (EN0124_2 'Upgate' = biogas upgraded to biomethane, whatever its use - mostly vehicle fuel; EN0124_4 = fed into the biogas/gas "
              "networks, counties summed, 2023 on). Sweden's gas grid is small: only the second series is grid injection.",
              "<CC>_biomethane_TWh (sheet Annual): Eurostat nrg_bal_c, product 'Biogases' (R5300), flow TI_BNG_E 'transformation input - for blending with "
              "natural gas', unit GWh converted to TWh, per year. This is the biogas upgraded to biomethane and fed into the natural gas grid as reported by the "
              "national statistics offices (annual only; latest years provisional or missing). It is not total biogas production and excludes biogas used "
              "for electricity/heat. UK = United Kingdom (Eurostat has UK data to 2019 only). EU27 = Eurostat aggregate; EU27_sum_of_countries_TWh is the sum of the "
              "member-state columns.",
              "Crosscheck sheet: annual (12-month) sums of the monthly national series next to the Eurostat figure.",
              "", "Caveats and validation",
              "Great Britain: the DESNZ monthly values are smooth (a constant daily rate per year, i.e. modelled/estimated from the quarterly and annual "
              "returns, note 6/7 on the sheet), not metered monthly data. National Gas's data portal carries biomethane for one NTS-connected plant only (Glentham), "
              "most GB biomethane enters distribution networks, so no operator daily series exists. DESNZ 2019 4.71 TWh vs Eurostat UK 4.89 TWh (gross vs net CV / revisions).",
              "Austria: AGGM series SummeBioOesterreich exists from 2023; before that the East market area only (Tyrol and Vorarlberg entry series start 2023), so "
              "2019-2022 are slightly understated. AGGM 2023 = 155 GWh vs Eurostat 121 GWh vs EBA Statistical Report 2024 (131 GWh injected by 14 plants). A zero/low month "
              "(Nov 2025) is a source gap, not a plant stop.",
              "EBA Statistical Report (full country tables) is members-only; only a 17-page preview is public, so EBA is not tabulated here. GIE/EBA European Biomethane Map "
              "2026 (PDF, no data download): installed capacity 8.2 bcm/year at end Q2 2026, 1,974 plants (86% grid-connected) - capacity, not injected volume. Eurostat EU27 "
              "was 33.2 TWh in 2023 and 37.1 TWh in 2024; Germany starts in 2021 (earlier years not reported, so the 2020-21 EU27 jump is a reporting break).",
              "Italy: Eurostat annual only (3.2 TWh in 2024) - GSE blocks GitHub (HTTP 403) and Snam's portal is JavaScript-only.",
              "", "Method", "Incremental: DESNZ whole file is downloaded only when a new release URL appears (recorded below); AGGM re-fetched from two months "
              "before the last month saved; Eurostat refreshed in full.",
              "", "Last pull", f"{datetime.now(timezone.utc):%Y-%m-%d %H:%M} UTC"]
    if st.get("gb_url"):
        lines.append(f"ET 4.2 release used: {st['gb_url']}")
    for e in errors:
        lines.append(f"Source problem this run: {e}")
    sheets = {"Monthly": monthly}
    if len(ann):
        sheets["Annual"] = ann
    if len(st["annual_nat"]):
        sheets["Annual national"] = st["annual_nat"].round(4)
    if not cross.empty:
        sheets["Crosscheck"] = cross
    xlsx_notes.write_workbook(path, sheets, lines, {"Source", "Units and definitions", "Method", "Last pull"})
    print(f"saved {FILE}: monthly {monthly.shape}, annual {ann.shape}")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out-dir", default=OUT_DEFAULT)
    ap.add_argument("--only", default="gb,at,se,eurostat")
    ap.add_argument("--force", action="store_true")
    args = ap.parse_args()
    os.makedirs(args.out_dir, exist_ok=True)
    path = os.path.join(args.out_dir, FILE)
    st = read_state(path, args.force)
    errors = []
    for name, fn in (("gb", gb), ("at", at), ("se", se), ("eurostat", eurostat)):
        if name not in args.only.split(","):
            continue
        try:
            fn(st)
        except Exception as e:  # noqa: BLE001
            print(f"{name} FAILED: {type(e).__name__}: {e}")
            errors.append(f"{name}: {type(e).__name__}")
    write(path, st, errors)


if __name__ == "__main__":
    main()
