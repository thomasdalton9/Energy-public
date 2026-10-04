"""
India's full daily generation mix: CEA conventional generation (National Power Portal, via INDIA_NPP.py) plus daily
wind, solar and other-renewable generation, and daily peak demand / energy met, from NITI Aayog's India Climate &
Energy Dashboard (ICED, https://iced.niti.gov.in/energy/electricity/generation#daily).

Why ICED: Grid-India's daily PSP report (grid-india.in, report.grid-india.in, webapi/webcdn) and the RLDC, MERIT
and Vidyut Pravah sites reset or time out from GitHub runners; CEA's daily RE generation report stopped on
2025-11-18 (only the monthly RE report continues) and NPP publishes no RE report (discovery_archive/asia/
INDIA_DISCOVERY1-8.py). ICED's public API serves the whole daily series in one request:
  GET https://icedapi.niti.gov.in/energy/electricity/generation/daily?source=all
      -> [dates dd-mm-yyyy, [Coal, Wind, Solar, Nuclear, Oil & Gas, Hydro, Other Res]] in MU/day, from 2015-04-01
         (wind / solar / other RE from 2019-06-12); its conventional series equal CEA's daily report (sub-report
         17: coal + lignite, natural gas, nuclear, hydro excl. Bhutan) to the MU, and RE runs a day or two ahead
         of CEA.
  GET https://icedapi.niti.gov.in/v1/dailyPeakDemand/energyMet
      -> [dates, [peak demand met MW, energy met MU]] from 2017-04-01 (Grid-India power supply position).
The responses are CryptoJS AES strings ('U2FsdGVkX1...', OpenSSL salted); the dashboard decrypts them in the
browser with the passphrase in its public JS bundle (environment KEY), which this script reads from the bundle.

Writes output/Data and Chart Outputs/india_power_generation_daily.xlsx:
  Daily        date, Coal / Gas / Oil / Nuclear / Hydro (CEA daily report, from india_npp_generation_daily.xlsx;
               a day missing there is filled from ICED's copy of the same CEA figures) + Wind / Solar / Other
               (ICED), Total, MWh/day
  Demand       date, Demand_peak_MW (peak demand met), Energy_met_MWh, Demand_avg_MW (= energy met / 24)
  ICED daily   all seven ICED series as published, MWh/day, from 2015 (the RE history store)

Incremental: ICED returns its whole history per request (two requests per run, ~0.5 MB), so each run re-reads it
and merges it over the saved sheets (saved days are kept if ICED drops or zeroes them). Runs on the 1st and 15th,
after INDIA_NPP.py.

    python3 asia/INDIA_RE_DAILY.py [--out PATH] [--npp PATH]
"""
import argparse
import base64
import hashlib
import json
import os
import re
import sys
import time

import pandas as pd
import requests

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)
import xlsx_notes  # noqa: E402

OUT_DIR = os.path.join(ROOT, "output", "Data and Chart Outputs")
SITE, API = "https://iced.niti.gov.in", "https://icedapi.niti.gov.in"
GEN_URL = f"{API}/energy/electricity/generation/daily?source=all"
DEMAND_URL = f"{API}/v1/dailyPeakDemand/energyMet"
H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 "
                   "Safari/537.36", "Accept": "application/json, text/plain, */*", "Origin": SITE, "Referer": SITE + "/"}
T = (20, 300)   # the API sometimes takes minutes, or answers 504 after 240 s; a retry usually comes back quickly
ICED_SERIES = ["Coal", "Wind", "Solar", "Nuclear", "Oil & Gas", "Hydro", "Other Res"]   # order in the bundle's chart
START = "2021-01-01"   # CEA conventional daily data (india_npp_generation_daily.xlsx) starts here
CONV = ["Coal", "Gas", "Oil", "Nuclear", "Hydro"]
RE = ["Wind", "Solar", "Other"]
DAY_RANGE = (2.5e6, 1.2e7)   # plausible all-India total, MWh/day


def out(*a):
    print(*a, flush=True)


def get(url, tries=4):
    for i in range(tries):
        t0 = time.time()
        try:
            r = requests.get(url, headers=H, timeout=T)
            if r.ok:
                out(f"  {url}: {len(r.content):,} bytes in {time.time() - t0:.0f}s")
                return r
            out(f"  {url}: HTTP {r.status_code} after {time.time() - t0:.0f}s")
        except requests.RequestException as e:
            out(f"  {url}: {type(e).__name__} after {time.time() - t0:.0f}s")
        if i < tries - 1:
            time.sleep(60)
    return None


def passphrase():
    r = requests.get(SITE + "/", headers=H, timeout=T)
    js = re.findall(r'src=["\']([^"\']*main[^"\']*\.js)["\']', r.text)[0]
    src = requests.get(f"{SITE}/{js.lstrip('/')}", headers=H, timeout=T).text
    return re.search(r'KEY:"([^"]+)"', src).group(1)


def decrypt(token, key):
    """CryptoJS.AES.decrypt(token, passphrase): OpenSSL 'Salted__' format, EVP_BytesToKey(MD5), AES-256-CBC."""
    from cryptography.hazmat.primitives import padding
    from cryptography.hazmat.primitives.ciphers import Cipher, algorithms, modes
    raw = base64.b64decode(token)
    if raw[:8] != b"Salted__":
        raise ValueError("not an OpenSSL salted string")
    salt, ct = raw[8:16], raw[16:]
    d, prev = b"", b""
    while len(d) < 48:
        prev = hashlib.md5(prev + key.encode() + salt).digest()
        d += prev
    dec = Cipher(algorithms.AES(d[:32]), modes.CBC(d[32:48])).decryptor()
    un = padding.PKCS7(128).unpadder()
    return json.loads((un.update(dec.update(ct) + dec.finalize()) + un.finalize()).decode("utf-8"))


def body(r, key):
    j = r.json()
    return decrypt(j, key) if isinstance(j, str) else j


def series_frame(dates, columns, names):
    df = pd.DataFrame({n: pd.to_numeric(pd.Series(v), errors="coerce") for n, v in zip(names, columns)})
    df.index = pd.to_datetime(pd.Series(dates), format="%d-%m-%Y")
    df.index.name = "date"
    return df[~df.index.duplicated(keep="last")].sort_index()


def read_sheet(path, sheet):
    try:
        df = pd.read_excel(path, sheet_name=sheet, index_col=0)
    except (FileNotFoundError, ValueError):
        return pd.DataFrame()
    df.index = pd.to_datetime(df.index)
    df.index.name = "date"
    return df.sort_index()


def merge(old, new):
    """New values over old; a day the new pull lacks (or has blank) keeps its saved value."""
    if old.empty or new.empty:
        return (new if old.empty else old).sort_index()
    return new.combine_first(old).sort_index()


def fetch_iced(key):
    gen = dem = pd.DataFrame()
    r = get(GEN_URL)
    if r is not None:
        dates, cols = body(r, key)["data"]
        if len(cols) != len(ICED_SERIES):
            raise ValueError(f"ICED daily generation has {len(cols)} series, expected {len(ICED_SERIES)}")
        gen = series_frame(dates, cols, ICED_SERIES) * 1000   # MU -> MWh
        gen = gen.mask(gen <= 0)   # 0 = not reported yet (the latest days, RE before mid-2019)
        gen.columns = [c.replace(" & ", "_and_").replace(" ", "_") + "_MWh" for c in gen.columns]
    time.sleep(5)
    r = get(DEMAND_URL)
    if r is not None:
        dates, (peak, energy) = body(r, key)
        dem = series_frame(dates, [peak, energy], ["Demand_peak_MW", "Energy_met_MWh"])
        dem["Energy_met_MWh"] *= 1000   # MU -> MWh
        dem = dem.mask(dem <= 0).dropna(how="all")
        dem["Demand_avg_MW"] = (dem["Energy_met_MWh"] / 24).round(0)
    return gen, dem


def build_daily(npp, iced):
    """CEA conventional (NPP workbook; gaps from ICED's identical CEA series) + ICED wind / solar / other RE."""
    conv = npp.reindex(columns=[f"{c}_MWh" for c in CONV]).copy()
    alt = pd.DataFrame({"Coal_MWh": iced.get("Coal_MWh"), "Gas_MWh": iced.get("Oil_and_Gas_MWh"),
                        "Nuclear_MWh": iced.get("Nuclear_MWh"), "Hydro_MWh": iced.get("Hydro_MWh")})
    alt = alt.dropna(subset=["Coal_MWh"])
    fill = alt.index.difference(conv.dropna(subset=["Coal_MWh"]).index)
    conv = pd.concat([conv[~conv.index.isin(fill)], alt.loc[fill]]).sort_index()
    re_ = pd.DataFrame({"Wind_MWh": iced.get("Wind_MWh"), "Solar_MWh": iced.get("Solar_MWh"),
                        "Other_MWh": iced.get("Other_Res_MWh")})
    d = conv.join(re_, how="inner")
    d = d[d.index >= START].dropna(subset=["Coal_MWh", "Wind_MWh", "Solar_MWh"])
    d["Oil_MWh"] = d["Oil_MWh"].fillna(0)
    d["Other_MWh"] = d["Other_MWh"].fillna(0)
    d = d[[f"{c}_MWh" for c in CONV + RE]]
    d["Total_MWh"] = d.sum(axis=1)
    d = d[d["Total_MWh"].between(*DAY_RANGE)].round(0)
    d.index.name = "date"
    return d, len(fill.intersection(d.index))


def check(npp, iced):
    both = npp[["Coal_MWh"]].join(iced[["Coal_MWh"]], rsuffix="_iced", how="inner").dropna()
    if len(both):
        dev = (both["Coal_MWh_iced"] / both["Coal_MWh"] - 1).abs()
        out(f"  check: ICED coal vs NPP dgr2 coal on {len(both)} days: median |diff| {100 * dev.median():.2f}%, "
            f"{(dev > 0.02).sum()} days off by >2%")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", default=os.path.join(OUT_DIR, "india_power_generation_daily.xlsx"))
    ap.add_argument("--npp", default=os.path.join(OUT_DIR, "india_npp_generation_daily.xlsx"))
    args = ap.parse_args()

    old_iced, old_dem = read_sheet(args.out, "ICED daily"), read_sheet(args.out, "Demand")
    out(f"Saved: ICED daily {len(old_iced)} days, Demand {len(old_dem)} days")
    try:
        new_iced, new_dem = fetch_iced(passphrase())
    except Exception as e:  # noqa: BLE001  (keep the saved history; the workbook is still rebuilt)
        out(f"ICED fetch failed: {type(e).__name__}: {e}")
        new_iced = new_dem = pd.DataFrame()
    iced, dem = merge(old_iced, new_iced), merge(old_dem, new_dem)
    if iced.empty:
        sys.exit("no ICED data, nothing saved yet")
    npp = read_sheet(args.npp, "Daily")
    out(f"NPP conventional: {len(npp)} days {npp.index.min():%Y-%m-%d}..{npp.index.max():%Y-%m-%d}" if len(npp)
        else "NPP conventional workbook missing: conventional series from ICED's copy of the CEA report")
    if len(npp):
        check(npp, iced)
    daily, filled = build_daily(npp, iced)
    dem = dem[[c for c in ("Demand_peak_MW", "Energy_met_MWh", "Demand_avg_MW") if c in dem]] if len(dem) else dem

    notes = ["UNITS",
             "MWh per day, all-India (sources publish MU = GWh; x1,000). Daily: Coal = coal + lignite steam plants, "
             "Gas = gas turbines / CCGTs, Oil = diesel (CEA daily generation report, gross generation of stations "
             "above 25 MW; Hydro = large hydro, excluding imports from Bhutan); Wind, Solar and Other = 'Other Res' "
             "(biomass, small hydro and other renewables as reported daily). Total = sum. Demand: Demand_peak_MW = "
             "peak demand met (MW), Energy_met_MWh = energy met (MWh/day), Demand_avg_MW = energy met / 24. ICED "
             "daily: the seven series exactly as ICED publishes them (MU x 1,000).",
             "", "COVERAGE",
             f"Daily from {daily.index.min():%Y-%m-%d} to {daily.index.max():%Y-%m-%d} ({len(daily)} days; a day "
             "appears once both the CEA conventional report and the renewables figures are out). "
             f"{filled} day(s) of conventional generation missing from the NPP pull were taken from ICED's copy of "
             "the same CEA figures. " + (f"Demand from {dem.index.min():%Y-%m-%d} to {dem.index.max():%Y-%m-%d}."
                                         if len(dem) else ""),
             "'Other' (about 10-12 TWh a year) covers only part of what CEA's monthly RE report counts as small "
             "hydro + biomass + bagasse + other (about 20-25 TWh a year), so Total understates utility generation "
             "by roughly 1%. Rooftop / behind-the-meter solar and captive plants are not included. Small hydro is "
             "in Other where reported, not in Hydro.",
             "", "SOURCE",
             "Conventional: CEA (Central Electricity Authority) daily generation report via the National Power "
             "Portal, https://npp.gov.in/public-reports/cea/daily/dgr/DD-MM-YYYY/dgr2-YYYY-MM-DD.xls "
             "(india_npp_generation_daily.xlsx, asia/INDIA_NPP.py).",
             "Wind, solar, other RE and demand: NITI Aayog, India Climate & Energy Dashboard (ICED), daily "
             "generation and daily demand series (compiled from CEA / Grid-India daily reports): "
             "https://iced.niti.gov.in/energy/electricity/generation#daily ; API "
             f"{GEN_URL} and {DEMAND_URL} (responses AES-encoded with the key in the site's public JS bundle). "
             "Grid-India's own daily PSP report is not reachable from the pull's servers, and CEA's daily RE "
             "report stopped on 2025-11-18.",
             "Whole series re-read on every run (one request each) and merged over the saved sheets."]
    xlsx_notes.write_workbook(args.out, {"Daily": daily, "Demand": dem, "ICED daily": iced.round(0)}, notes,
                              {"UNITS", "COVERAGE", "SOURCE"})
    out(f"Saved {args.out}: Daily {len(daily)} days, Demand {len(dem)}, ICED daily {len(iced)}")
    if len(daily):
        y = (daily.groupby(daily.index.year).sum() / 1e6).round(1)
        y["days"] = daily.groupby(daily.index.year).size()
        out("TWh per year:\n" + y.rename(columns=lambda c: c.replace("_MWh", "")).to_string())
        out("Last days (MWh):\n" + daily.tail(4).to_string())
    if len(dem):
        out("Demand, last days:\n" + dem.tail(3).to_string())


if __name__ == "__main__":
    main()
