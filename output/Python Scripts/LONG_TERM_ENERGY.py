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
FUELS_START = 1965
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

# EI all-data workbook: sheet -> our variable (EI's own units). The 2025 edition renamed primary energy to
# "Total Energy Supply (TES)"; the old name is kept as a fallback.
EI_SHEETS = [
    (r"^Gas Production - Bcm", "Gas_Production_bcm"),
    (r"^Gas Consumption - Bcm", "Gas_Consumption_bcm"),
    (r"^Gas - LNG imports bcm", "LNG_Imports_bcm"),
    (r"^Gas - LNG exports bcm", "LNG_Exports_bcm"),
    (r"^Oil Production - barrels", "Oil_Production_kbd"),
    (r"^Oil Consumption - barrels", "Oil_Consumption_kbd"),
    (r"^Coal Production - EJ", "Coal_Production_EJ"),
    (r"^Coal Consumption - EJ", "Coal_Consumption_EJ"),
    (r"^(Total Energy Supply \(TES\)|Primary Energy\s*-?\s*Cons)", "Primary_Energy_EJ"),
]
# EI country names pycountry does not resolve
EI_ISO = {"US": "USA", "Russian Federation": "RUS", "Iran": "IRN", "South Korea": "KOR", "Vietnam": "VNM",
          "Trinidad & Tobago": "TTO", "China Hong Kong SAR": "HKG", "Turkey": "TUR", "Türkiye": "TUR",
          "Turkiye": "TUR", "Venezuela": "VEN", "Bolivia": "BOL", "Syria": "SYR", "Taiwan": "TWN",
          "Czech Republic": "CZE", "Czechia": "CZE", "Moldova": "MDA", "Tanzania": "TZA", "Brunei": "BRN",
          "Laos": "LAO", "North Macedonia": "MKD", "Republic of Congo": "COG", "Congo": "COG",
          "Democratic Republic of Congo": "COD", "Ivory Coast": "CIV", "Cote d'Ivoire": "CIV", "Kosovo": "XKX",
          "Slovakia": "SVK", "United Kingdom": "GBR", "UK": "GBR", "United Arab Emirates": "ARE",
          "Netherlands": "NLD", "Philippines": "PHL", "China": "CHN", "Ukraine": "UKR", "Egypt": "EGY",
          "Total World": "WLD", "World": "WLD"}
NOT_COUNTRY = re.compile(r"^(total|other|of which|non-|rest of|central america|eastern africa|middle africa|"
                         r"western africa|european union|oecd|cis|ussr|opec|africa|europe|asia|middle east|"
                         r"north america|s\. & cent|south america|world$)", re.I)


def iso_lookup(name):
    name = re.sub(r"[*#^]+$", "", str(name)).strip()
    if name in EI_ISO:
        return EI_ISO[name]
    if NOT_COUNTRY.match(name):
        return None
    try:
        import pycountry
        return pycountry.countries.lookup(name).alpha_3
    except Exception:
        try:
            import pycountry
            hits = pycountry.countries.search_fuzzy(name)
            return hits[0].alpha_3 if len(hits) == 1 else None
        except Exception:
            return None


def ei_session():
    """The EI site sits behind Cloudflare, which refuses GitHub runners unless the TLS handshake looks like a
    browser's: curl_cffi impersonates Chrome."""
    from curl_cffi import requests as cr
    return cr


def find_ei_link():
    """Current all-data workbook linked on the EI downloads page (the unversioned 'ALL-data' file)."""
    cr = ei_session()
    r = cr.get(EI_PAGE, impersonate="chrome", timeout=60)
    print(f"EI page: HTTP {r.status_code}, {len(r.text)} chars", flush=True)
    if r.status_code != 200:
        raise RuntimeError(f"EI page HTTP {r.status_code}")
    links = sorted(set(re.findall(r'href="([^"]+\.xlsx[^"]*)"', r.text, flags=re.I)))
    links = [requests.compat.urljoin(EI_PAGE, l.replace("&amp;", "&")) for l in links]
    for l in links:
        print("  EI link:", l, flush=True)
    data = [l for l in links if re.search(r"all.data", l, re.I)]
    # the current edition is the unversioned file; past editions carry a year in the name
    current = [l for l in data if not re.search(r"20\d\d\.xlsx", l)] or data
    if not current:
        raise RuntimeError("no all-data workbook linked on the EI downloads page")
    return current[0]


def ei_head(url):
    try:
        r = ei_session().head(url, impersonate="chrome", timeout=60, allow_redirects=True)
        print(f"HEAD {url}: {r.status_code} {dict((k, v) for k, v in r.headers.items() if k.lower() in ('last-modified', 'etag', 'content-length'))}",
              flush=True)
        return r if r.status_code == 200 else None
    except Exception as e:
        print(f"HEAD {url}: {e}", flush=True)
        return None


def ei_stamp(resp):
    """EI sends no Last-Modified/ETag: Content-Length stands in (the URL itself changes with each new asset)."""
    s = stamp_of(resp)
    cl = resp.headers.get("Content-Length", "")
    return " | ".join(x for x in (s, f"Content-Length {cl}" if cl else "") if x)


def ei_sheet(xl, sheet, var):
    raw = pd.read_excel(xl, sheet_name=sheet, header=None)
    hdr = None
    for i in range(min(10, len(raw))):
        yrs = pd.to_numeric(raw.iloc[i, 1:], errors="coerce")
        if yrs.between(1900, 2100).sum() >= 5:
            hdr = i
            break
    if hdr is None:
        print(f"EI {sheet}: no year header found", flush=True)
        return None
    cols = []
    for j in range(1, raw.shape[1]):        # contiguous run of years only (growth/share columns repeat the year)
        y = pd.to_numeric(raw.iat[hdr, j], errors="coerce")
        if pd.isna(y) or not 1900 <= y <= 2100 or (cols and int(y) != cols[-1][1] + 1):
            if cols:
                break
            continue
        cols.append((j, int(y)))
    body = raw.iloc[hdr + 1:, [0] + [j for j, _ in cols]]
    body.columns = ["country"] + [y for _, y in cols]
    body = body.dropna(subset=["country"])
    body["country"] = body["country"].astype(str).str.strip()
    body["iso3"] = body["country"].map(iso_lookup)
    dropped = sorted(set(body.loc[body["iso3"].isna(), "country"]))
    body = body.dropna(subset=["iso3"]).drop_duplicates("iso3")
    long = body.melt(id_vars=["country", "iso3"], var_name="year", value_name="value")
    long["value"] = pd.to_numeric(long["value"], errors="coerce")
    long = long.dropna(subset=["value"])
    long["variable"] = var
    long.loc[long["iso3"] == "WLD", "country"] = "World"
    long["country"] = long["country"].str.replace(r"[*#^]+$", "", regex=True).str.strip()
    print(f"EI {sheet} -> {var}: {long['iso3'].nunique()} areas, {cols[0][1]}-{cols[-1][1]}; "
          f"rows not mapped: {[d for d in dropped if not NOT_COUNTRY.match(d)][:40]}", flush=True)
    return long


def parse_ei(content):
    xl = pd.ExcelFile(io.BytesIO(content))
    parts = []
    for pat, var in EI_SHEETS:
        sheet = next((s for s in xl.sheet_names if re.search(pat, s.strip(), re.I)), None)
        if sheet is None:
            print(f"EI: no sheet for {var} ({pat})", flush=True)
            continue
        if var in {p["variable"].iat[0] for p in parts if len(p)}:
            continue
        df = ei_sheet(xl, sheet, var)
        if df is not None and len(df):
            parts.append(df)
    if not parts:
        raise RuntimeError("no EI sheets parsed")
    d = pd.concat(parts, ignore_index=True)
    d["year"] = d["year"].astype(int)
    d = d[d["year"] >= FUELS_START]
    d["value"] = d["value"].astype(float).round(4)
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
    out = out[out["year"] >= FUELS_START]
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
    owid_src = (f"Fuels source: Our World in Data energy dataset ({OWID_URL}), which compiles the Energy Institute "
                "Statistical Review (fallback: the EI file could not be read from GitHub)")
    try:
        url = find_ei_link()
        h = ei_head(url)
        new_stamp = f"Fuels file: {url} | " + (ei_stamp(h) if h is not None else "")
        if fuels is not None and not args.force and h is not None and ei_stamp(h) and new_stamp == fuels_stamp:
            print("EI file unchanged - keeping the saved Fuels sheet", flush=True)
        else:
            r = ei_session().get(url, impersonate="chrome", timeout=600)
            if r.status_code != 200:
                raise RuntimeError(f"EI file HTTP {r.status_code}")
            print(f"GET {url}: {len(r.content) / 1e6:.1f} MB", flush=True)
            fuels = parse_ei(r.content)
            fuels_stamp = f"Fuels file: {url} | " + (ei_stamp(h) if h is not None else ei_stamp(r))
            fuels_src = (f"Fuels source: Energy Institute, Statistical Review of World Energy, all-data workbook "
                         f"({url}; downloads page {EI_PAGE})")
    except Exception as e:
        print(f"EI failed ({e}) - falling back to Our World in Data", flush=True)
        if fuels is None or "Our World in Data" in fuels_src:
            h = head(OWID_URL)
            new_stamp = f"Fuels file: {OWID_URL} | " + (stamp_of(h) if h is not None else "")
            if fuels is not None and not args.force and h is not None and stamp_of(h) and new_stamp == fuels_stamp:
                print("OWID file unchanged - keeping the saved Fuels sheet", flush=True)
            else:
                r = get(OWID_URL)
                fuels = parse_owid(r.content)
                fuels_stamp = f"Fuels file: {OWID_URL} | {stamp_of(r)}"
                fuels_src = owid_src
        else:
            print("keeping the saved EI Fuels sheet", flush=True)
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
         "Fuels sheet (Energy Institute): Gas_Production_bcm, Gas_Consumption_bcm, LNG_Imports_bcm, LNG_Exports_bcm in "
         "billion cubic metres per year; Oil_Production_kbd (crude, condensate and NGLs) and Oil_Consumption_kbd in "
         "thousand barrels per day; Coal_Production_EJ, Coal_Consumption_EJ and Primary_Energy_EJ (EI's Total Energy "
         "Supply) in exajoules. EI's own units, unconverted. Pipeline imports are not included: EI gives them by "
         "country only for the latest year (a from/to matrix)."),
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
