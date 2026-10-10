"""
Macro demand drivers for the long-term fundamentals view in the master workbooks.

    python3 MACRO_DRIVERS.py [--out "output/Data and Chart Outputs/macro_drivers.xlsx"]

Writes macro_drivers.xlsx (wide sheets: index `year` or `date`, one column per ISO3 code):
  GDP_real_USD_bn, GDP_PPP_USD_bn, Population_m, Industry_share_pct, Urban_pct, Electricity_access_pct
      World Bank WDI API, all countries + WLD (aggregates dropped). Re-downloaded every run.
  GDP_growth_IMF_pct, Population_IMF_m
      IMF DataMapper API (WEO), all countries + WLD, including forecasts to the end of the WEO horizon.
  CDD_18, HDD_18 (annual, complete years only), CDD_18_monthly, HDD_18_monthly (complete months only)
      Degree days, base 18 C, from daily mean 2 m temperature (NASA POWER, MERRA-2), population-weighted over each
      country's 3 largest cities (1 city for small countries). Only the master-workbook countries.
  City_T2M_daily
      The raw daily city temperatures (C) - the incremental history store: each run fetches only the days after the
      last saved one, minus a 10-day revision window (full history for a city not saved yet).
  Countries
      iso3, name, region (World Bank), Forecast_from (first WEO projection year).
"""
import argparse
import calendar
import datetime as dt
import os
import sys
import time
from concurrent.futures import ThreadPoolExecutor

import pandas as pd
import requests

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import xlsx_notes  # noqa: E402

OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "output", "Data and Chart Outputs", "macro_drivers.xlsx")
UA = {"User-Agent": "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36",
      "Accept": "application/json"}
WB = "https://api.worldbank.org/v2"
IMF = "https://www.imf.org/external/datamapper/api/v1"
NASA = "https://power.larc.nasa.gov/api/temporal/daily/point"
OPEN_METEO = "https://archive-api.open-meteo.com/v1/archive"
WEATHER_START = dt.date(1991, 1, 1)
REVISION_DAYS = 10
BASE_C = 18.0

WB_INDICATORS = [  # sheet, indicator, divisor
    ("GDP_real_USD_bn", "NY.GDP.MKTP.KD", 1e9),
    ("GDP_PPP_USD_bn", "NY.GDP.MKTP.PP.KD", 1e9),
    ("Population_m", "SP.POP.TOTL", 1e6),
    ("Industry_share_pct", "NV.IND.TOTL.ZS", 1),
    ("Urban_pct", "SP.URB.TOTL.IN.ZS", 1),
    ("Electricity_access_pct", "EG.ELC.ACCS.ZS", 1),
]
IMF_INDICATORS = [("GDP_growth_IMF_pct", "NGDP_RPCH"), ("Population_IMF_m", "LP")]

# Degree-day cities: iso3 -> [(city, lat, lon, population in millions, approx. metro)]. The countries of the master
# workbooks (South & Southeast Asia, South & Central America + Caribbean, North America, Europe incl. GB, Turkey, Cyprus,
# Australia & New Zealand).
CITIES = {
    # South & Southeast Asia
    "IND": [("Delhi", 28.61, 77.21, 33.0), ("Mumbai", 19.08, 72.88, 21.0), ("Kolkata", 22.57, 88.36, 15.5)],
    "PAK": [("Karachi", 24.86, 67.01, 17.0), ("Lahore", 31.55, 74.34, 13.5), ("Faisalabad", 31.42, 73.08, 3.7)],
    "BGD": [("Dhaka", 23.81, 90.41, 23.0), ("Chittagong", 22.36, 91.78, 5.4), ("Khulna", 22.85, 89.54, 1.0)],
    "LKA": [("Colombo", 6.93, 79.86, 2.3), ("Kandy", 7.29, 80.63, 0.15), ("Jaffna", 9.66, 80.01, 0.09)],
    "NPL": [("Kathmandu", 27.72, 85.32, 3.0), ("Pokhara", 28.21, 83.99, 0.5), ("Biratnagar", 26.45, 87.27, 0.25)],
    "BTN": [("Thimphu", 27.47, 89.64, 0.12)],
    "THA": [("Bangkok", 13.76, 100.50, 11.0), ("Chiang Mai", 18.79, 98.98, 1.2),
            ("Nakhon Ratchasima", 14.97, 102.10, 0.5)],
    "VNM": [("Ho Chi Minh City", 10.82, 106.63, 9.3), ("Hanoi", 21.03, 105.85, 8.4), ("Hai Phong", 20.84, 106.69, 2.1)],
    "PHL": [("Manila", 14.60, 120.98, 14.0), ("Cebu", 10.32, 123.89, 3.0), ("Davao", 7.19, 125.46, 1.8)],
    "IDN": [("Jakarta", -6.21, 106.85, 34.0), ("Surabaya", -7.25, 112.75, 10.0), ("Bandung", -6.92, 107.61, 8.5)],
    "MYS": [("Kuala Lumpur", 3.14, 101.69, 8.4), ("George Town", 5.41, 100.33, 2.8), ("Johor Bahru", 1.49, 103.74, 1.8)],
    "SGP": [("Singapore", 1.35, 103.82, 5.9)],
    "MMR": [("Yangon", 16.87, 96.20, 5.6), ("Mandalay", 21.96, 96.08, 1.5), ("Naypyidaw", 19.76, 96.08, 1.2)],
    "KHM": [("Phnom Penh", 11.56, 104.92, 2.3), ("Siem Reap", 13.36, 103.86, 0.25), ("Battambang", 13.10, 103.20, 0.2)],
    "LAO": [("Vientiane", 17.98, 102.63, 0.95), ("Savannakhet", 16.56, 104.75, 0.12), ("Pakse", 15.12, 105.80, 0.09)],
    "BRN": [("Bandar Seri Begawan", 4.90, 114.94, 0.24)],
    "TLS": [("Dili", -8.56, 125.57, 0.28)],
    # North America
    "USA": [("New York", 40.71, -74.01, 19.5), ("Los Angeles", 34.05, -118.24, 12.8), ("Chicago", 41.88, -87.63, 9.3)],
    "CAN": [("Toronto", 43.65, -79.38, 6.7), ("Montreal", 45.50, -73.57, 4.3), ("Vancouver", 49.28, -123.12, 2.6)],
    "MEX": [("Mexico City", 19.43, -99.13, 22.0), ("Guadalajara", 20.67, -103.35, 5.3), ("Monterrey", 25.69, -100.32, 5.3)],
    # South America
    "ARG": [("Buenos Aires", -34.60, -58.38, 15.5), ("Cordoba", -31.42, -64.18, 1.6), ("Rosario", -32.95, -60.65, 1.4)],
    "BOL": [("La Paz-El Alto", -16.50, -68.15, 1.9), ("Santa Cruz", -17.78, -63.18, 1.8),
            ("Cochabamba", -17.39, -66.16, 1.3)],
    "BRA": [("Sao Paulo", -23.55, -46.63, 22.4), ("Rio de Janeiro", -22.91, -43.17, 13.6),
            ("Belo Horizonte", -19.92, -43.94, 6.0)],
    "CHL": [("Santiago", -33.45, -70.67, 7.0), ("Valparaiso", -33.05, -71.62, 1.0), ("Concepcion", -36.83, -73.05, 1.0)],
    "COL": [("Bogota", 4.71, -74.07, 11.3), ("Medellin", 6.24, -75.58, 4.1), ("Cali", 3.45, -76.53, 2.8)],
    "ECU": [("Guayaquil", -2.19, -79.89, 3.1), ("Quito", -0.18, -78.47, 2.0), ("Cuenca", -2.90, -79.00, 0.6)],
    "PRY": [("Asuncion", -25.26, -57.58, 3.5), ("Ciudad del Este", -25.51, -54.61, 0.4),
            ("Encarnacion", -27.33, -55.87, 0.15)],
    "PER": [("Lima", -12.05, -77.04, 11.0), ("Arequipa", -16.41, -71.54, 1.1), ("Trujillo", -8.11, -79.03, 1.0)],
    "URY": [("Montevideo", -34.90, -56.16, 1.8), ("Salto", -31.38, -57.96, 0.1), ("Paysandu", -32.32, -58.08, 0.08)],
    "VEN": [("Caracas", 10.48, -66.90, 2.9), ("Maracaibo", 10.65, -71.64, 2.3), ("Valencia", 10.16, -68.00, 1.8)],
    "GUY": [("Georgetown", 6.80, -58.16, 0.2)],
    "SUR": [("Paramaribo", 5.85, -55.20, 0.24)],
    # Central America
    "BLZ": [("Belize City", 17.50, -88.20, 0.07)],
    "CRI": [("San Jose", 9.93, -84.08, 1.4), ("Limon", 9.99, -83.03, 0.1), ("Liberia", 10.63, -85.44, 0.07)],
    "SLV": [("San Salvador", 13.69, -89.22, 1.1), ("Santa Ana", 13.99, -89.56, 0.25), ("San Miguel", 13.48, -88.18, 0.25)],
    "GTM": [("Guatemala City", 14.63, -90.51, 3.0), ("Quetzaltenango", 14.83, -91.52, 0.2),
            ("Escuintla", 14.30, -90.78, 0.15)],
    "HND": [("Tegucigalpa", 14.07, -87.19, 1.3), ("San Pedro Sula", 15.50, -88.03, 1.0), ("La Ceiba", 15.76, -86.78, 0.2)],
    "NIC": [("Managua", 12.11, -86.24, 1.1), ("Leon", 12.43, -86.88, 0.2), ("Masaya", 11.97, -86.09, 0.17)],
    "PAN": [("Panama City", 8.98, -79.52, 1.9), ("Colon", 9.36, -79.90, 0.25), ("David", 8.43, -82.43, 0.15)],
    # Caribbean
    "TTO": [("Port of Spain", 10.66, -61.51, 0.55)],
    "PRI": [("San Juan", 18.47, -66.11, 2.0), ("Ponce", 18.01, -66.61, 0.13), ("Mayaguez", 18.20, -67.14, 0.07)],
    "JAM": [("Kingston", 17.97, -76.79, 1.2), ("Montego Bay", 18.47, -77.92, 0.11), ("Mandeville", 18.04, -77.50, 0.05)],
    "DOM": [("Santo Domingo", 18.49, -69.93, 3.5), ("Santiago de los Caballeros", 19.45, -70.69, 1.0),
            ("La Romana", 18.43, -68.97, 0.15)],
    # Europe
    "DEU": [("Berlin", 52.52, 13.40, 3.7), ("Hamburg", 53.55, 9.99, 1.9), ("Munich", 48.14, 11.58, 1.5)],
    "FRA": [("Paris", 48.86, 2.35, 11.0), ("Lyon", 45.76, 4.84, 2.3), ("Marseille", 43.30, 5.37, 1.9)],
    "ESP": [("Madrid", 40.42, -3.70, 6.8), ("Barcelona", 41.39, 2.17, 5.6), ("Valencia", 39.47, -0.38, 1.6)],
    "ITA": [("Rome", 41.90, 12.50, 4.3), ("Milan", 45.46, 9.19, 4.3), ("Naples", 40.85, 14.27, 3.1)],
    "NLD": [("Amsterdam", 52.37, 4.90, 2.5), ("Rotterdam", 51.92, 4.48, 1.0), ("The Hague", 52.08, 4.30, 0.8)],
    "BEL": [("Brussels", 50.85, 4.35, 2.1), ("Antwerp", 51.22, 4.40, 1.1), ("Liege", 50.63, 5.57, 0.75)],
    "LUX": [("Luxembourg", 49.61, 6.13, 0.13)],
    "POL": [("Warsaw", 52.23, 21.01, 3.1), ("Krakow", 50.06, 19.94, 1.5), ("Lodz", 51.76, 19.46, 1.0)],
    "AUT": [("Vienna", 48.21, 16.37, 2.9), ("Linz", 48.31, 14.29, 0.8), ("Graz", 47.07, 15.44, 0.6)],
    "CHE": [("Zurich", 47.38, 8.54, 1.4), ("Geneva", 46.20, 6.14, 0.6), ("Basel", 47.56, 7.59, 0.55)],
    "CZE": [("Prague", 50.08, 14.44, 2.7), ("Brno", 49.20, 16.61, 0.7), ("Ostrava", 49.82, 18.26, 0.5)],
    "SVK": [("Bratislava", 48.15, 17.11, 0.66), ("Kosice", 48.72, 21.26, 0.24), ("Presov", 49.00, 21.24, 0.09)],
    "HUN": [("Budapest", 47.50, 19.04, 3.0), ("Debrecen", 47.53, 21.63, 0.2), ("Szeged", 46.25, 20.15, 0.16)],
    "ROU": [("Bucharest", 44.43, 26.10, 2.3), ("Cluj-Napoca", 46.77, 23.59, 0.4), ("Iasi", 47.16, 27.59, 0.37)],
    "BGR": [("Sofia", 42.70, 23.32, 1.5), ("Plovdiv", 42.14, 24.75, 0.35), ("Varna", 43.21, 27.91, 0.34)],
    "GRC": [("Athens", 37.98, 23.73, 3.7), ("Thessaloniki", 40.64, 22.94, 1.1), ("Patras", 38.25, 21.73, 0.2)],
    "PRT": [("Lisbon", 38.72, -9.14, 2.9), ("Porto", 41.15, -8.61, 1.8), ("Braga", 41.55, -8.42, 0.2)],
    "HRV": [("Zagreb", 45.82, 15.98, 0.8), ("Split", 43.51, 16.44, 0.18), ("Rijeka", 45.33, 14.44, 0.12)],
    "SVN": [("Ljubljana", 46.06, 14.51, 0.29), ("Maribor", 46.55, 15.65, 0.1), ("Celje", 46.24, 15.27, 0.05)],
    "DNK": [("Copenhagen", 55.68, 12.57, 2.1), ("Aarhus", 56.16, 10.20, 0.35), ("Odense", 55.40, 10.39, 0.2)],
    "SWE": [("Stockholm", 59.33, 18.07, 2.4), ("Gothenburg", 57.71, 11.97, 1.1), ("Malmo", 55.60, 13.00, 0.75)],
    "NOR": [("Oslo", 59.91, 10.75, 1.1), ("Bergen", 60.39, 5.32, 0.29), ("Trondheim", 63.43, 10.40, 0.2)],
    "FIN": [("Helsinki", 60.17, 24.94, 1.5), ("Tampere", 61.50, 23.76, 0.4), ("Turku", 60.45, 22.27, 0.33)],
    "IRL": [("Dublin", 53.35, -6.26, 1.5), ("Cork", 51.90, -8.47, 0.3), ("Limerick", 52.66, -8.63, 0.1)],
    "GBR": [("London", 51.51, -0.13, 9.6), ("Birmingham", 52.49, -1.89, 2.9), ("Manchester", 53.48, -2.24, 2.8)],
    "ISL": [("Reykjavik", 64.15, -21.94, 0.24)],
    "EST": [("Tallinn", 59.44, 24.75, 0.45), ("Tartu", 58.38, 26.72, 0.1), ("Narva", 59.38, 28.19, 0.05)],
    "LVA": [("Riga", 56.95, 24.11, 0.6), ("Daugavpils", 55.87, 26.54, 0.08), ("Liepaja", 56.51, 21.01, 0.07)],
    "LTU": [("Vilnius", 54.69, 25.28, 0.6), ("Kaunas", 54.90, 23.90, 0.3), ("Klaipeda", 55.71, 21.13, 0.16)],
    "SRB": [("Belgrade", 44.79, 20.45, 1.7), ("Novi Sad", 45.27, 19.83, 0.35), ("Nis", 43.32, 21.90, 0.26)],
    "BIH": [("Sarajevo", 43.86, 18.41, 0.42), ("Banja Luka", 44.77, 17.19, 0.19), ("Tuzla", 44.54, 18.67, 0.11)],
    "MNE": [("Podgorica", 42.44, 19.26, 0.19)],
    "MKD": [("Skopje", 41.99, 21.43, 0.53), ("Kumanovo", 42.13, 21.71, 0.1), ("Bitola", 41.03, 21.33, 0.08)],
    "ALB": [("Tirana", 41.33, 19.82, 0.8), ("Durres", 41.32, 19.45, 0.2), ("Vlore", 40.47, 19.49, 0.13)],
    "XKX": [("Pristina", 42.66, 21.17, 0.2)],
    "MLT": [("Valletta", 35.90, 14.51, 0.5)],
    "CYP": [("Nicosia", 35.17, 33.36, 0.34)],
    "TUR": [("Istanbul", 41.01, 28.98, 15.7), ("Ankara", 39.93, 32.86, 5.8), ("Izmir", 38.42, 27.14, 4.5)],
    # Australia & New Zealand
    "AUS": [("Sydney", -33.87, 151.21, 5.5), ("Melbourne", -37.81, 144.96, 5.3), ("Brisbane", -27.47, 153.03, 2.7)],
    "NZL": [("Auckland", -36.85, 174.76, 1.7), ("Wellington", -41.29, 174.78, 0.43)],
    # China (3 largest cities: a coarse proxy for a country of this climatic range)
    "CHN": [("Shanghai", 31.23, 121.47, 24.9), ("Beijing", 39.90, 116.41, 21.5), ("Guangzhou", 23.13, 113.26, 18.7)],
    # South Korea (four largest metropolitan cities, populations approximate city-proper figures, millions)
    "KOR": [("Seoul", 37.57, 126.98, 9.4), ("Busan", 35.18, 129.08, 3.3), ("Incheon", 37.46, 126.71, 3.0), ("Daegu", 35.87, 128.60, 2.4)],
    # Taiwan (Greater Taipei, Kaohsiung, Taichung; approximate metropolitan populations, million)
    "TWN": [("Taipei", 25.03, 121.57, 7.0), ("Kaohsiung", 22.63, 120.30, 2.7), ("Taichung", 24.15, 120.67, 2.8)],
}


def col_name(iso3, city):
    return f"{iso3} {city}"


def get_json(url, params=None, tries=4, timeout=120):
    last = None
    for i in range(tries):
        try:
            r = requests.get(url, params=params, headers=UA, timeout=timeout)
            if r.status_code == 429:
                time.sleep(30 * (i + 1))
                continue
            r.raise_for_status()
            return r.json()
        except Exception as exc:  # noqa: BLE001
            last = exc
            time.sleep(5 * (i + 1))
    raise RuntimeError(f"{url}: {last}")


# ---------------------------------------------------------------- World Bank
def wb_countries():
    """iso3 -> (name, region) for real countries (aggregates dropped), plus WLD."""
    js = get_json(f"{WB}/country", {"format": "json", "per_page": 1000})
    out = {}
    for c in js[1]:
        region = (c.get("region") or {}).get("value", "").strip()
        if region == "Aggregates" and c["id"] != "WLD":
            continue
        out[c["id"]] = (c["name"], "World" if c["id"] == "WLD" else region)
    return out


def wb_indicator(code, keep):
    rows, page, pages = [], 1, 1
    while page <= pages:
        js = get_json(f"{WB}/country/all/indicator/{code}",
                      {"format": "json", "per_page": 20000, "date": "1990:2030", "page": page})
        pages = int(js[0].get("pages", 1) or 1)
        rows += js[1] or []
        page += 1
    df = pd.DataFrame([{"iso3": r.get("countryiso3code"), "year": int(r["date"]), "value": r["value"]}
                       for r in rows if r.get("value") is not None and r.get("countryiso3code") in keep])
    if df.empty:
        return pd.DataFrame()
    w = df.pivot_table(index="year", columns="iso3", values="value", aggfunc="first").sort_index()
    w.index = w.index.astype(int)
    w.index.name = "year"
    return w[[c for c in ["WLD"] + sorted(c for c in w.columns if c != "WLD") if c in w.columns]]


# ---------------------------------------------------------------- IMF
# imf.org (Akamai) answers 403 to plain Python clients from GitHub, so the DataMapper API is read with curl_cffi's
# browser impersonation; if that fails too, the WEO is read from the IMF SDMX 3.0 API (api.imf.org, CSV).
IMF_SDMX = "https://api.imf.org/external/sdmx/3.0/data/dataflow/IMF.RES/WEO/+/*.{code}.A"
WORLD_CODES = {"WEOWORLD", "WORLD", "WLD", "G001", "W00", "001", "W001"}


def imf_get(url, accept="application/json"):
    last = None
    try:
        from curl_cffi import requests as creq
        for i in range(3):
            try:
                r = creq.get(url, impersonate="chrome", timeout=120, headers={"Accept": accept})
                r.raise_for_status()
                return r
            except Exception as exc:  # noqa: BLE001
                last = exc
                time.sleep(5 * (i + 1))
    except ImportError:
        pass
    for i in range(2):
        try:
            r = requests.get(url, headers={**UA, "Accept": accept}, timeout=120)
            r.raise_for_status()
            return r
        except Exception as exc:  # noqa: BLE001
            last = exc
            time.sleep(5)
    raise RuntimeError(f"{url}: {last}")


IMF_LABELS = {}


def imf_meta():
    """Country codes (ISO3) IMF lists as countries, and the WEO vintage from the indicator metadata."""
    js = imf_get(f"{IMF}/countries").json().get("countries", {})
    IMF_LABELS.update({k: (v or {}).get("label", "") for k, v in js.items()})
    vintage = ""
    try:
        ind = imf_get(f"{IMF}/indicators").json().get("indicators", {})
        vintage = (ind.get("NGDP_RPCH") or {}).get("source", "") or ""
    except Exception as exc:  # noqa: BLE001
        print(f"IMF indicators metadata failed: {exc}")
    return set(js), vintage


def _wide(rows):
    df = pd.DataFrame(rows)
    if df.empty:
        return pd.DataFrame()
    w = df.pivot_table(index="year", columns="iso3", values="value", aggfunc="first").sort_index()
    w.index = w.index.astype(int)
    w.index.name = "year"
    w = w[w.index >= 1990]
    return w[[c for c in ["WLD"] + sorted(c for c in w.columns if c != "WLD") if c in w.columns]]


def imf_datamapper(code, imf_countries, keep):
    vals = imf_get(f"{IMF}/{code}").json().get("values", {}).get(code, {})
    rows = []
    for c, series in vals.items():
        iso = "WLD" if c in WORLD_CODES else c
        if iso != "WLD" and not (c in imf_countries or c in keep):
            continue
        for y, v in (series or {}).items():
            if v is not None:
                rows.append({"iso3": iso, "year": int(y), "value": float(v)})
    return _wide(rows), {}


def imf_sdmx(code, keep):
    """WEO from the SDMX 3.0 API as CSV. Returns (wide frame, {iso3: first forecast year}) when the latest-actual
    attribute is present."""
    import io
    r = imf_get(IMF_SDMX.format(code=code), accept="application/vnd.sdmx.data+csv;version=2.0.0")
    df = pd.read_csv(io.StringIO(r.text), dtype=str)
    print(f"  SDMX {code}: columns {list(df.columns)[:20]}", flush=True)
    area = next(c for c in ("COUNTRY", "REF_AREA", "JURISDICTION") if c in df.columns)
    df["value"] = pd.to_numeric(df["OBS_VALUE"], errors="coerce")
    df["year"] = pd.to_numeric(df["TIME_PERIOD"].str[:4], errors="coerce")
    df["iso3"] = df[area].where(~df[area].isin(WORLD_CODES), "WLD")
    df = df[df["iso3"].isin(keep | {"WLD"}) | (df["iso3"].str.len() == 3)].dropna(subset=["value", "year"])
    df = df[df["iso3"].str.fullmatch(r"[A-Z]{3}")]
    if "SCALE" in df.columns and code == "LP":   # population: millions
        sc = pd.to_numeric(df["SCALE"], errors="coerce")
        df.loc[sc.notna(), "value"] = df.loc[sc.notna(), "value"] * (10.0 ** sc[sc.notna()]) / 1e6
    ff = {}
    lat = next((c for c in df.columns if "LATEST_ACTUAL" in c.upper()), None)
    if lat:
        for iso, v in df.groupby("iso3")[lat].first().items():
            y = pd.to_numeric(str(v)[:4], errors="coerce")
            if pd.notna(y):
                ff[iso] = int(y) + 1
    return _wide(df[["iso3", "year", "value"]].to_dict("records")), ff


def imf_indicator(code, imf_countries, keep):
    try:
        w, ff = imf_datamapper(code, imf_countries, keep)
        if not w.empty:
            return w, ff, "IMF DataMapper API"
    except Exception as exc:  # noqa: BLE001
        print(f"  DataMapper {code} failed: {exc}; trying SDMX", flush=True)
    w, ff = imf_sdmx(code, keep)
    return w, ff, "IMF SDMX 3.0 API (WEO dataflow)"


# ---------------------------------------------------------------- weather
def fetch_nasa(lat, lon, start, end):
    js = get_json(NASA, {"parameters": "T2M", "community": "RE", "longitude": lon, "latitude": lat,
                         "start": start.strftime("%Y%m%d"), "end": end.strftime("%Y%m%d"), "format": "JSON"},
                  tries=3, timeout=300)
    t = js["properties"]["parameter"]["T2M"]
    s = pd.Series({pd.Timestamp(k): (None if v is None or v <= -998 else float(v)) for k, v in t.items()}, dtype=float)
    return s.dropna()


def fetch_open_meteo(lat, lon, start, end):
    js = get_json(OPEN_METEO, {"latitude": lat, "longitude": lon, "start_date": start.isoformat(),
                               "end_date": end.isoformat(), "daily": "temperature_2m_mean", "timezone": "UTC"},
                  tries=4, timeout=300)
    d = js["daily"]
    return pd.Series(d["temperature_2m_mean"], index=pd.to_datetime(d["time"]), dtype=float).dropna()


SOURCE_USED = {}


def fetch_city(job):
    iso, city, lat, lon, start, end = job
    try:
        s = fetch_nasa(lat, lon, start, end)
        if s.empty:
            raise RuntimeError("no data")
        SOURCE_USED[col_name(iso, city)] = "NASA POWER"
        return col_name(iso, city), s
    except Exception as exc:  # noqa: BLE001
        print(f"  NASA POWER failed for {iso} {city}: {exc}; trying Open-Meteo", flush=True)
    try:
        time.sleep(8)   # Open-Meteo's free tier is rate limited (weighted by days requested)
        s = fetch_open_meteo(lat, lon, start, end)
        SOURCE_USED[col_name(iso, city)] = "Open-Meteo ERA5"
        return col_name(iso, city), s
    except Exception as exc:  # noqa: BLE001
        print(f"  Open-Meteo failed for {iso} {city}: {exc}", flush=True)
        return col_name(iso, city), pd.Series(dtype=float)


def load_saved_temps(path):
    try:
        df = pd.read_excel(path, sheet_name="City_T2M_daily", index_col=0)
    except (FileNotFoundError, ValueError) as exc:
        print(f"No saved city temperatures ({exc}); full history from {WEATHER_START}")
        return pd.DataFrame()
    df.index = pd.to_datetime(df.index)
    df.index.name = "date"
    return df.apply(pd.to_numeric, errors="coerce")


def update_temps(saved, workers):
    today = dt.date.today()
    end = today - dt.timedelta(days=1)
    jobs = []
    for iso, cities in CITIES.items():
        for city, lat, lon, _pop in cities:
            c = col_name(iso, city)
            if c in saved and saved[c].notna().any():
                last = saved[c].dropna().index.max().date()
                start = max(WEATHER_START, last - dt.timedelta(days=REVISION_DAYS))
            else:
                start = WEATHER_START
            if start <= end:
                jobs.append((iso, city, lat, lon, start, end))
    print(f"Weather: {len(jobs)} city requests ({sum(j[4] == WEATHER_START for j in jobs)} full-history)", flush=True)
    new = {}
    with ThreadPoolExecutor(max_workers=workers) as ex:
        for i, (c, s) in enumerate(ex.map(fetch_city, jobs), 1):
            if not s.empty:
                new[c] = s.round(2)
            if i % 20 == 0:
                print(f"  {i}/{len(jobs)} cities done", flush=True)
    out = pd.DataFrame(new).combine_first(saved) if new else saved.copy()   # new values win (revision window)
    out = out.sort_index()
    out.index.name = "date"
    cols = [col_name(iso, city) for iso, cities in CITIES.items() for city, *_ in cities]
    return out[[c for c in cols if c in out] + [c for c in out.columns if c not in cols]]


def degree_days(temps):
    """Population-weighted daily CDD/HDD per country (a day counts only when every city of the country has data)."""
    cdd, hdd = {}, {}
    for iso, cities in CITIES.items():
        names = [col_name(iso, c[0]) for c in cities]
        if not all(n in temps for n in names):
            continue
        w = pd.Series({col_name(iso, c[0]): c[3] for c in cities})
        w = w / w.sum()
        t = temps[names]
        ok = t.notna().all(axis=1)
        cdd[iso] = ((t - BASE_C).clip(lower=0) * w).sum(axis=1).where(ok)
        hdd[iso] = ((BASE_C - t).clip(lower=0) * w).sum(axis=1).where(ok)
    return pd.DataFrame(cdd), pd.DataFrame(hdd)


def aggregate(daily):
    """Monthly sums for complete months, annual sums for complete years."""
    if daily.empty:
        return pd.DataFrame(), pd.DataFrame()
    n_m = daily.notna().resample("MS").sum()
    s_m = daily.resample("MS").sum(min_count=1)
    days_in_m = pd.Series([calendar.monthrange(d.year, d.month)[1] for d in s_m.index], index=s_m.index)
    monthly = s_m.where(n_m.eq(days_in_m, axis=0)).round(1)
    monthly.index.name = "date"
    n_y = daily.notna().groupby(daily.index.year).sum()
    s_y = daily.groupby(daily.index.year).sum(min_count=1)
    days_in_y = pd.Series([366 if calendar.isleap(y) else 365 for y in s_y.index], index=s_y.index)
    annual = s_y.where(n_y.eq(days_in_y, axis=0)).dropna(how="all").round(0)
    annual.index = annual.index.astype(int)
    annual.index.name = "year"
    return monthly.dropna(how="all"), annual


# ---------------------------------------------------------------- main
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", default=OUT)
    ap.add_argument("--workers", type=int, default=4)
    args = ap.parse_args()
    os.makedirs(os.path.dirname(args.out), exist_ok=True)

    sheets, failed = {}, []
    countries = wb_countries()
    keep = set(countries)
    print(f"World Bank: {len(keep)} countries (incl. WLD)", flush=True)
    for sheet, code, div in WB_INDICATORS:
        try:
            w = wb_indicator(code, keep) / div
            sheets[sheet] = w.round(4)
            print(f"  {sheet}: {w.shape[1]} countries, {w.index.min()}-{w.index.max()}", flush=True)
        except Exception as exc:  # noqa: BLE001
            failed.append(f"World Bank {code}: {exc}")
            print(f"  {sheet} FAILED: {exc}", flush=True)

    vintage, imf_countries, imf_ff, imf_src = "", set(), {}, set()
    try:
        imf_countries, vintage = imf_meta()
    except Exception as exc:  # noqa: BLE001
        print(f"IMF DataMapper countries list failed: {exc}")
    for sheet, code in IMF_INDICATORS:
        try:
            w, ff, src = imf_indicator(code, imf_countries, keep)
            if code == "NGDP_RPCH":
                imf_ff = ff
            imf_src.add(src)
            sheets[sheet] = w.round(4)
            print(f"  {sheet}: {w.shape[1]} countries, {w.index.min()}-{w.index.max()}", flush=True)
        except Exception as exc:  # noqa: BLE001
            failed.append(f"IMF {code}: {exc}")
            print(f"  {sheet} FAILED: {exc}", flush=True)

    # first WEO projection year: the DataMapper API does not give it per country, so the WEO-wide one is used
    # (April / October WEO of year Y -> projections from Y; the API's vintage string names the WEO)
    vy = [int(t) for t in vintage.replace("(", " ").replace(")", " ").split() if t.isdigit() and len(t) == 4]
    forecast_from = vy[0] if vy else dt.date.today().year

    saved = load_saved_temps(args.out)
    temps = update_temps(saved, args.workers)
    d_cdd, d_hdd = degree_days(temps)
    cdd_m, cdd_y = aggregate(d_cdd)
    hdd_m, hdd_y = aggregate(d_hdd)
    sheets.update({"CDD_18": cdd_y, "HDD_18": hdd_y, "CDD_18_monthly": cdd_m, "HDD_18_monthly": hdd_m})

    rows = []
    iso_all = set(keep)
    for s in sheets.values():
        iso_all |= set(map(str, s.columns))
    for iso in sorted(iso_all, key=lambda c: (c != "WLD", c)):
        name, region = countries.get(iso, (IMF_LABELS.get(iso, ""), ""))
        rows.append({"iso3": iso, "name": name, "region": region, "Forecast_from": imf_ff.get(iso, forecast_from)})
    sheets["Countries"] = pd.DataFrame(rows).set_index("iso3")
    sheets["City_T2M_daily"] = temps

    srcs = sorted(set(SOURCE_USED.values())) or ["(none fetched this run - saved history)"]
    last_day = temps.dropna(how="all").index.max().date() if not temps.empty else None
    missing = [col_name(i, c[0]) for i, cs in CITIES.items() for c in cs if col_name(i, c[0]) not in temps]
    notes = [
        "UNITS",
        "GDP_real_USD_bn: GDP, constant 2015 US$, billion (World Bank NY.GDP.MKTP.KD / 1e9)",
        "GDP_PPP_USD_bn: GDP, PPP, constant 2021 international $, billion (World Bank NY.GDP.MKTP.PP.KD / 1e9)",
        "Population_m: population, million (World Bank SP.POP.TOTL / 1e6)",
        "Industry_share_pct: industry incl. construction, value added, % of GDP (World Bank NV.IND.TOTL.ZS)",
        "Urban_pct: urban population, % of total (World Bank SP.URB.TOTL.IN.ZS)",
        "Electricity_access_pct: access to electricity, % of population (World Bank EG.ELC.ACCS.ZS)",
        "GDP_growth_IMF_pct: real GDP growth, % y/y (IMF WEO NGDP_RPCH), includes forecasts",
        "Population_IMF_m: population, million (IMF WEO LP), includes forecasts",
        "CDD_18 / HDD_18: cooling / heating degree days, base 18 C, degree-C days per year",
        "CDD_18_monthly / HDD_18_monthly: the same per month (date = first of the month)",
        "City_T2M_daily: daily mean 2 m air temperature, deg C, per city (column = ISO3 + city)",
        "",
        "LAYOUT",
        "Wide sheets: index column `year` (int; `date` for the monthly and daily sheets), one column per ISO3 code. "
        "WLD = World. World Bank aggregates (regions, income groups) are dropped.",
        "",
        "COVERAGE",
        "World Bank: 1990 to the latest WDI year, all countries + WLD. IMF: 1990 to the end of the WEO horizon.",
        f"IMF source this run: {', '.join(sorted(imf_src)) or 'FAILED'}. WEO vintage: {vintage or 'not given by the API'}. "
        + (f"Forecast_from (Countries sheet): first year after each country's latest actual (WEO attribute), "
           f"{forecast_from} where none is given." if imf_ff else
           f"Forecast_from (Countries sheet) = {forecast_from}: the WEO-wide first projection year - the API does "
           "not give the last actual year per country, and some countries' latest actuals are older."),
        f"Degree days: {len(CITIES)} master-workbook countries, from 1991. Annual sheets hold complete calendar years "
        "only (the current, partial year is left out); monthly sheets hold complete months only.",
        f"Latest daily temperature: {last_day}. Weather source this run: {', '.join(srcs)}.",
        "Countries with no degree days: those not in the masters (S&SE Asia, S&C America + Caribbean, North "
        "America, Europe incl. GB, Turkey, Cyprus, Australia & NZ).",
        "",
        "METHOD",
        "Degree days: per city, CDD = max(0, Tmean - 18), HDD = max(0, 18 - Tmean) from the daily mean; the country "
        "value is the population-weighted mean of its cities (3 largest; 2 for NZ; 1 for small countries). A day counts only "
        "when every city has data. Three cities can misstate large, climatically varied countries (USA, BRA, IND, "
        "IDN, CAN) - an index of year-on-year change, not a national average.",
        "Cities (ISO3: city [population, million]): " + "; ".join(
            f"{i}: " + ", ".join(f"{c[0]} [{c[3]}]" for c in cs) for i, cs in CITIES.items()),
        "Incremental: each run reads City_T2M_daily and fetches only days after the last saved one, minus a "
        f"{REVISION_DAYS}-day revision window. World Bank and IMF are re-downloaded every run.",
        "",
        "SOURCE",
        "World Bank World Development Indicators API: https://api.worldbank.org/v2/",
        "IMF DataMapper API (World Economic Outlook): https://www.imf.org/external/datamapper/api/v1/",
        "NASA POWER daily point API, T2M, community RE (MERRA-2 / GEOS): https://power.larc.nasa.gov/ ; fallback "
        "Open-Meteo historical archive (ERA5): https://open-meteo.com/",
    ]
    if failed or missing:
        notes += ["", "FAILED THIS RUN"] + failed + ([f"No temperature data: {', '.join(missing)}"] if missing else [])
    xlsx_notes.write_workbook(args.out, sheets, notes,
                              {"UNITS", "LAYOUT", "COVERAGE", "METHOD", "SOURCE", "FAILED THIS RUN"})
    print(f"Wrote {args.out}")
    for f in failed:
        print(f"FAILED: {f}")


if __name__ == "__main__":
    main()
