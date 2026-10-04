"""
Long annual energy history for every country, for the long-term fundamentals pages of the master workbooks.

  Power  - Ember yearly full release (2000 onward): generation by fuel, demand, net imports, capacity, power CO2.
  Fuels  - Energy Institute Statistical Review of World Energy (1965 onward): gas, oil, coal, primary energy.
           If EI cannot be reached from GitHub, Our World in Data's energy dataset is used instead (labelled).

Output: output/Data and Chart Outputs/long_term_energy.xlsx, long format (one row per year, iso3, variable):
  Power / Fuels: year, iso3, country, variable, value.   Countries: iso3, name.

Both sources publish whole files, so each run sends a HEAD request first and only downloads and re-parses a
file whose Last-Modified / ETag differs from the one recorded on the Units sheet; otherwise the saved sheet is kept.

Usage: python3 LONG_TERM_ENERGY.py [--out PATH] [--force]
"""
import argparse
import io
import os
import re
import sys

import pandas as pd
import requests

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import xlsx_notes  # noqa: E402

OUT = "output/Data and Chart Outputs/long_term_energy.xlsx"
EMBER_URL = ("https://storage.googleapis.com/emb-prod-bkt-publicdata/public-downloads/"
             "yearly_full_release_long_format.csv")
EI_PAGE = "https://www.energyinst.org/statistical-review/resources-and-data-downloads"
OWID_URL = "https://raw.githubusercontent.com/owid/energy-data/master/owid-energy-data.csv"
TWH_PER_BCM = 10.0           # OWID fallback only: gas TWh -> bcm
EJ_PER_TWH = 0.0036
POWER_START = 2000
UA = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) "
                    "Chrome/126.0 Safari/537.36",
      "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
      "Accept-Language": "en-GB,en;q=0.9"}
COLS = ["year", "iso3", "country", "variable", "value"]

# Ember fuel -> our generation / capacity bucket
GEN_MAP = {"Hydro": "Hydro", "Gas": "Gas", "Wind": "Wind", "Solar": "Solar", "Coal": "Coal", "Nuclear": "Nuclear",
           "Other Fossil": "Oil", "Bioenergy": "Bio", "Other Renewables": "Bio"}


# ------------------------------------------------------------------ helpers

def stamp_of(resp):
    """Version stamp of a published file: Last-Modified and ETag (whatever the server gives)."""
    lm, et = resp.headers.get("Last-Modified", ""), resp.headers.get("ETag", "")
    s = " | ".join(x for x in (f"Last-Modified {lm}" if lm else "", f"ETag {et}" if et else "") if x)
    return s


def load_saved(path):
    """Units text and saved data sheets of the previous run (empty when there is none)."""
    try:
        xl = pd.ExcelFile(path)
        units = [str(x) for x in pd.read_excel(xl, sheet_name=0, header=None).iloc[:, 0].dropna()]
        sheets = {}
        for name in ("Power", "Fuels", "Countries"):
            if name in xl.sheet_names:
                df = pd.read_excel(xl, sheet_name=name)
                sheets[name] = df.drop(columns=[c for c in df.columns if str(c).startswith("Unnamed")])
        return units, sheets
    except (FileNotFoundError, ValueError, KeyError, OSError):
        return [], {}


def saved_line(units, prefix):
    return next((u for u in units if u.startswith(prefix)), "")


def head(url):
    try:
        r = requests.head(url, headers=UA, timeout=(10, 60), allow_redirects=True)
        print(f"HEAD {url}: {r.status_code} {stamp_of(r)}", flush=True)
        return r if r.ok else None
    except requests.RequestException as e:
        print(f"HEAD {url}: {e}", flush=True)
        return None


def get(url, timeout=600):
    r = requests.get(url, headers=UA, timeout=(15, timeout))
    r.raise_for_status()
    print(f"GET {url}: {len(r.content) / 1e6:.1f} MB", flush=True)
    return r


# ------------------------------------------------------------------ Ember (Power)

def parse_ember(content):
    d = pd.read_csv(io.BytesIO(content), low_memory=False)
    print("Ember columns:", list(d.columns), flush=True)
    print("Ember area types:", d["Area type"].dropna().unique().tolist(), flush=True)
    iso_col = next(c for c in d.columns if c.lower().startswith("iso"))
    is_country = d["Area type"].astype(str).str.contains("Country", case=False)
    d = d[is_country | (d["Area"] == "World")].copy()
    d["iso3"] = d[iso_col].where(d["Area"] != "World", "WLD")
    d = d[d["iso3"].astype(str).str.fullmatch(r"[A-Z]{3}")]
    d = d[d["Year"] >= POWER_START]
    d["Value"] = pd.to_numeric(d["Value"], errors="coerce")
    parts = []

    def add(rows, var):
        g = rows.groupby(["Year", "iso3", "Area"], as_index=False)["Value"].sum(min_count=1)
        g["variable"] = var
        parts.append(g)

    gen = d[(d["Category"] == "Electricity generation") & (d["Unit"] == "TWh")]
    fuel = gen[gen["Subcategory"] == "Fuel"]
    for bucket in ("Hydro", "Gas", "Wind", "Solar", "Coal", "Nuclear", "Oil", "Bio"):
        add(fuel[fuel["Variable"].map(GEN_MAP) == bucket], f"Gen_{bucket}_TWh")
    add(fuel[fuel["Variable"].isin(["Other Fossil", "Bioenergy", "Other Renewables"])], "Gen_Other_TWh")
    tot = gen[gen["Variable"] == "Total Generation"]
    add(tot if not tot.empty else fuel, "Gen_Total_TWh")
    add(d[(d["Category"] == "Electricity demand") & (d["Variable"] == "Demand") & (d["Unit"] == "TWh")], "Demand_TWh")
    add(d[(d["Category"] == "Electricity imports") & (d["Variable"] == "Net Imports") & (d["Unit"] == "TWh")],
        "Net_Imports_TWh")
    cap = d[(d["Category"] == "Capacity") & (d["Unit"] == "GW") & (d["Subcategory"] == "Fuel")]
    for bucket in ("Hydro", "Gas", "Wind", "Solar", "Coal", "Nuclear"):
        add(cap[cap["Variable"] == bucket], f"Cap_{bucket}_GW")
    add(cap[cap["Variable"].isin(["Other Fossil", "Bioenergy", "Other Renewables"])], "Cap_Other_GW")
    add(cap, "Cap_Total_GW")
    em = d[(d["Category"] == "Power sector emissions") & (d["Unit"].astype(str).str.lower() == "mtco2")]
    em_tot = em[em["Subcategory"] == "Total"]
    add(em_tot if not em_tot.empty else em[em["Subcategory"] == "Fuel"], "CO2_Mt")
    out = pd.concat(parts, ignore_index=True).rename(columns={"Year": "year", "Area": "country", "Value": "value"})
    out = out.dropna(subset=["value"])
    out["year"] = out["year"].astype(int)
    out["value"] = out["value"].round(3)
    return out[COLS].sort_values(["iso3", "variable", "year"]).reset_index(drop=True)


# ------------------------------------------------------------------ Energy Institute (Fuels)

# EI panel/narrow variable -> our variable (first match wins; matched case-insensitively on the full name)
EI_VARS = {
    "gascons_bcm": "Gas_Consumption_bcm", "gasprod_bcm": "Gas_Production_bcm",
    "oilcons_kbd": "Oil_Consumption_kbd", "oilprod_kbd": "Oil_Production_kbd",
    "coalcons_ej": "Coal_Consumption_EJ", "coalprod_ej": "Coal_Production_EJ",
    "primary_ej": "Primary_Energy_EJ",
    "lngimp_bcm": "LNG_Imports_bcm", "lngexp_bcm": "LNG_Exports_bcm", "lng_imp_bcm": "LNG_Imports_bcm",
    "lng_exp_bcm": "LNG_Exports_bcm", "pipeimp_bcm": "Pipeline_Imports_bcm", "pipe_imp_bcm": "Pipeline_Imports_bcm",
    "gas_lng_imports_bcm": "LNG_Imports_bcm", "gas_lng_exports_bcm": "LNG_Exports_bcm",
    "gas_pipeline_imports_bcm": "Pipeline_Imports_bcm",
}
WORLD_NAMES = {"total world", "world"}


def find_ei_link():
    """The EI downloads page lists the data files; pick the panel-format CSV, else the narrow file."""
    r = requests.get(EI_PAGE, headers=UA, timeout=(15, 60))
    print(f"EI page: HTTP {r.status_code}, {len(r.text)} chars", flush=True)
    r.raise_for_status()
    links = re.findall(r'href="([^"]+\.(?:csv|xlsx)[^"]*)"', r.text, flags=re.I)
    links = [requests.compat.urljoin(EI_PAGE, l.replace("&amp;", "&")) for l in links]
    for l in links:
        print("  EI link:", l, flush=True)
    for pat in (r"panel.*\.csv", r"narrow.*\.csv", r"panel", r"narrow"):
        hit = [l for l in links if re.search(pat, l, re.I)]
        if hit:
            return hit[0]
    raise RuntimeError("no panel/narrow data file linked on the EI downloads page")


def iso_lookup(name):
    try:
        import pycountry
        return pycountry.countries.lookup(name).alpha_3
    except Exception:
        return None


def parse_ei(content, url):
    if url.lower().split("?")[0].endswith(".csv"):
        d = pd.read_csv(io.BytesIO(content), low_memory=False, encoding_errors="replace")
    else:
        d = pd.read_excel(io.BytesIO(content), sheet_name=0)
    print("EI columns:", list(d.columns)[:40], f"... ({d.shape[1]} columns)", flush=True)
    lc = {c.lower(): c for c in d.columns}
    country = lc.get("country")
    year = lc.get("year")
    iso = next((lc[c] for c in lc if c.startswith("iso")), None)
    if "var" in lc and "value" in lc:            # narrow: one row per variable
        d = d.rename(columns={lc["var"]: "var", lc["value"]: "value"})
    else:                                         # panel: one column per variable
        idv = [c for c in (country, year, iso) if c]
        valvars = [c for c in d.columns if c.lower() in EI_VARS]
        d = d.melt(id_vars=idv, value_vars=valvars, var_name="var", value_name="value")
    allvars = sorted(d["var"].astype(str).str.lower().unique())
    print("EI variables (first 200):", allvars[:200], flush=True)
    d["variable"] = d["var"].astype(str).str.lower().map(EI_VARS)
    d = d.dropna(subset=["variable"])
    d = d.rename(columns={country: "country", year: "year"})
    d["iso3"] = d[iso] if iso else None
    world = d["country"].astype(str).str.strip().str.lower().isin(WORLD_NAMES)
    d.loc[world, "iso3"] = "WLD"
    d.loc[world, "country"] = "World"
    miss = d["iso3"].isna() | ~d["iso3"].astype(str).str.fullmatch(r"[A-Z]{3}")
    if miss.any():
        names = d.loc[miss, "country"].astype(str).unique()
        m = {n: iso_lookup(n) for n in names if not re.match(r"(total|other|of which|non-|rest of)", n, re.I)}
        d.loc[miss, "iso3"] = d.loc[miss, "country"].map(m)
        print("EI names left without ISO3 (aggregates dropped):", [n for n in names if not m.get(n)][:80], flush=True)
    d = d[d["iso3"].astype(str).str.fullmatch(r"[A-Z]{3}")]
    d["value"] = pd.to_numeric(d["value"], errors="coerce")
    d = d.dropna(subset=["value", "year"])
    d["year"] = d["year"].astype(int)
    d["value"] = d["value"].round(4)
    return d[COLS].drop_duplicates(["year", "iso3", "variable"]).sort_values(["iso3", "variable", "year"]) \
        .reset_index(drop=True)


def parse_owid(content):
    d = pd.read_csv(io.BytesIO(content), low_memory=False)
    d = d[d["iso_code"].astype(str).str.fullmatch(r"[A-Z]{3}") | (d["country"] == "World")].copy()
    d["iso3"] = d["iso_code"].where(d["country"] != "World", "WLD")
    conv = {"gas_production": ("Gas_Production_bcm", 1 / TWH_PER_BCM),
            "gas_consumption": ("Gas_Consumption_bcm", 1 / TWH_PER_BCM),
            "coal_production": ("Coal_Production_EJ", EJ_PER_TWH),
            "coal_consumption": ("Coal_Consumption_EJ", EJ_PER_TWH),
            "primary_energy_consumption": ("Primary_Energy_EJ", EJ_PER_TWH)}
    parts = []
    for col, (var, f) in conv.items():
        if col in d:
            p = d[["year", "iso3", "country"]].copy()
            p["variable"], p["value"] = var, (pd.to_numeric(d[col], errors="coerce") * f).round(4)
            parts.append(p.dropna(subset=["value"]))
    out = pd.concat(parts, ignore_index=True)
    out["year"] = out["year"].astype(int)
    return out[COLS].sort_values(["iso3", "variable", "year"]).reset_index(drop=True)


# ------------------------------------------------------------------ main

def coverage(df):
    if df is None or df.empty:
        return "none"
    return "; ".join(f"{v} {g['year'].min()}-{g['year'].max()} ({g['iso3'].nunique()} areas)"
                     for v, g in df.groupby("variable"))


def checks(power, fuels):
    def val(df, iso, var, yr):
        s = df[(df["iso3"] == iso) & (df["variable"] == var) & (df["year"] == yr)]["value"]
        return float(s.iloc[0]) if len(s) else float("nan")
    for df, iso, var, yr in [(power, "WLD", "Gen_Total_TWh", 2024), (power, "IND", "Demand_TWh", 2024),
                             (power, "WLD", "Cap_Total_GW", 2024), (power, "WLD", "CO2_Mt", 2024),
                             (fuels, "WLD", "Gas_Consumption_bcm", 2024), (fuels, "WLD", "Gas_Production_bcm", 2024),
                             (fuels, "QAT", "LNG_Exports_bcm", 2024), (fuels, "JPN", "LNG_Imports_bcm", 2024),
                             (fuels, "WLD", "Oil_Consumption_kbd", 2024), (fuels, "WLD", "Primary_Energy_EJ", 2024),
                             (fuels, "USA", "Gas_Production_bcm", 1970)]:
        if df is not None and not df.empty:
            print(f"CHECK {iso} {var} {yr}: {val(df, iso, var, yr):,.1f}", flush=True)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", default=OUT)
    ap.add_argument("--force", action="store_true", help="re-download both files even if unchanged")
    args = ap.parse_args()
    units, saved = load_saved(args.out)

    # ---- Ember
    power, ember_stamp = saved.get("Power"), saved_line(units, "Ember file: ")
    h = head(EMBER_URL)
    new_stamp = "Ember file: " + (stamp_of(h) if h is not None else "")
    if power is not None and not args.force and h is not None and stamp_of(h) and new_stamp == ember_stamp:
        print("Ember file unchanged - keeping the saved Power sheet", flush=True)
    else:
        try:
            r = get(EMBER_URL)
            power = parse_ember(r.content)
            ember_stamp = "Ember file: " + (stamp_of(r) or (stamp_of(h) if h is not None else ""))
        except Exception as e:  # keep the saved sheet if the download fails
            print(f"Ember download/parse failed: {e}", flush=True)
            if power is None:
                raise

    # ---- Energy Institute (fallback: OWID)
    fuels, fuels_stamp = saved.get("Fuels"), saved_line(units, "Fuels file: ")
    fuels_src = saved_line(units, "Fuels source: ") or ""
    try:
        url = find_ei_link()
        src = f"Fuels source: Energy Institute Statistical Review of World Energy ({url})"
    except Exception as e:
        print(f"EI not reachable ({e}) - falling back to Our World in Data", flush=True)
        url = OWID_URL
        src = (f"Fuels source: Our World in Data energy dataset ({OWID_URL}), which compiles the Energy Institute "
               "Statistical Review (fallback: EI could not be reached from GitHub)")
    h = head(url)
    new_stamp = f"Fuels file: {url} | " + (stamp_of(h) if h is not None else "")
    if fuels is not None and not args.force and h is not None and stamp_of(h) and new_stamp == fuels_stamp:
        print("Fuels file unchanged - keeping the saved Fuels sheet", flush=True)
    else:
        try:
            r = get(url)
            fuels = parse_owid(r.content) if url == OWID_URL else parse_ei(r.content, url)
            fuels_stamp = f"Fuels file: {url} | " + (stamp_of(r) or (stamp_of(h) if h is not None else ""))
            fuels_src = src
        except Exception as e:
            print(f"Fuels download/parse failed from {url}: {e}", flush=True)
            if url != OWID_URL:
                r = get(OWID_URL)
                fuels = parse_owid(r.content)
                fuels_stamp = f"Fuels file: {OWID_URL} | {stamp_of(r)}"
                fuels_src = (f"Fuels source: Our World in Data energy dataset ({OWID_URL}), which compiles the "
                             "Energy Institute Statistical Review (fallback: the EI file could not be read)")
            elif fuels is None:
                raise
    if fuels is None:
        fuels = pd.DataFrame(columns=COLS)

    names = pd.concat([power[["iso3", "country"]], fuels[["iso3", "country"]]]).drop_duplicates("iso3")
    names = names.rename(columns={"country": "name"}).sort_values("iso3").reset_index(drop=True)
    checks(power, fuels)
    owid = "Our World in Data" in fuels_src
    notes = [
        "UNITS",
        "Long format: one row per (year, iso3, variable). iso3 = ISO 3166 alpha-3; WLD = World. Other regional "
        "aggregates are dropped.",
        "Power sheet (Ember): Gen_*_TWh = electricity generation in TWh per year; Gen_Oil_TWh = Ember 'Other Fossil' "
        "(oil and other fossil); Gen_Bio_TWh = Bioenergy + Other Renewables; Gen_Other_TWh = Gen_Oil + Gen_Bio. "
        "Demand_TWh, Net_Imports_TWh (imports minus exports) in TWh. Cap_*_GW = installed capacity in GW "
        "(Cap_Other_GW = other fossil + bioenergy + other renewables). CO2_Mt = power-sector CO2 emissions, Mt.",
        ("Fuels sheet (Our World in Data fallback): gas converted from TWh to bcm at 10.0 TWh/bcm; coal and primary "
         "energy from TWh to EJ (x 0.0036). OWID gives no oil volumes or LNG/pipeline trade, so those are absent."
         if owid else
         "Fuels sheet (Energy Institute): Gas_*_bcm in billion cubic metres per year; Oil_*_kbd in thousand barrels "
         "per day; Coal_*_EJ and Primary_Energy_EJ in exajoules. EI's own units, unconverted."),
        "Countries sheet: iso3 and the name used by the source.",
        "",
        "COVERAGE",
        f"Power (from {POWER_START}): {coverage(power)}",
        f"Fuels: {coverage(fuels)}",
        "",
        "SOURCE",
        f"Power: Ember yearly electricity data, full release (CC-BY-4.0): {EMBER_URL}",
        ember_stamp,
        fuels_src,
        fuels_stamp,
        "Both sources publish whole files: each run sends a HEAD request and re-downloads only when the recorded "
        "Last-Modified / ETag above changes.",
    ]
    os.makedirs(os.path.dirname(os.path.abspath(args.out)), exist_ok=True)
    sheets = {"Power": power.set_index("year"), "Fuels": fuels.set_index("year"), "Countries": names.set_index("iso3")}
    xlsx_notes.write_workbook(args.out, sheets, notes, {"UNITS", "COVERAGE", "SOURCE"})
    print(f"Saved {args.out}: Power {len(power)} rows, Fuels {len(fuels)} rows, {len(names)} countries", flush=True)


if __name__ == "__main__":
    main()
