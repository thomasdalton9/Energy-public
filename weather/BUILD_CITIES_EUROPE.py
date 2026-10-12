"""
Builds weather/cities_europe.csv and weather/countries_europe.csv (run ONCE, or when a source is revised, in GitHub Actions
via .github/workflows/weather_cities_europe.yml; the csv files are committed and read by weather/degree_days.py).

The 4 most populous cities per country, by an OFFICIAL population source:
  * Eurostat Urban Audit, table urb_cpop1, indicator DE1001V (population on 1 January, total) - the national statistical
    offices' figures for the Urban Audit "cities" (cores or "greater cities"); latest year of each city, so the vintage differs
    by country (stated in population_year). Coordinates: GISCO URAU label points.
  * Eurostat/GISCO LAU municipality population (POP_<year>, from the national statistical offices) where the Urban Audit lists
    fewer than 4 cities (Estonia, Slovenia, Luxembourg, Malta, Cyprus) or none (Serbia, North Macedonia, Albania). Coordinates:
    centroid of the LAU polygon. Serbia's Belgrade and Nis and North Macedonia's Skopje are sums of their city municipalities.
    Malta, Luxembourg and Cyprus have fewer than 4 meaningful cities: the 4 largest settlements/municipalities are used.
  * Countries Eurostat does not cover (see EXTRA below): national statistics office tables, or - where none could be read -
    populations left blank (equal weights, flagged ASSUMPTION by degree_days.py).
Country populations (weights of the aggregate): Eurostat demo_pjan (1 January 2025); United Kingdom: ONS mid-year estimate.
"""
import csv
import io
import json
import os
import re
import sys

import pandas as pd
import requests

HERE = os.path.dirname(os.path.abspath(__file__))
UA = {"User-Agent": "Mozilla/5.0"}
ES_API = "https://ec.europa.eu/eurostat/api/dissemination/statistics/1.0/data/"
GISCO = "https://gisco-services.ec.europa.eu/distribution/v2/"
UA_URL = "https://ec.europa.eu/eurostat/databrowser/view/urb_cpop1/default/table"
LAU_URL = "https://ec.europa.eu/eurostat/web/gisco/geodata/statistical-units/local-administrative-units"
PJAN_URL = "https://ec.europa.eu/eurostat/databrowser/view/demo_pjan/default/table"

# country code (ISO2, GB for the UK) -> (name, Eurostat code, method)
COUNTRIES = [
    ("AT", "Austria", "AT", "UA"), ("BE", "Belgium", "BE", "UA"), ("BG", "Bulgaria", "BG", "UA"), ("HR", "Croatia", "HR", "UA"),
    ("CY", "Cyprus", "CY", "LAU"), ("CZ", "Czechia", "CZ", "UA"), ("DK", "Denmark", "DK", "UA"), ("EE", "Estonia", "EE", "LAU"),
    ("FI", "Finland", "FI", "UA"), ("FR", "France", "FR", "UA"), ("DE", "Germany", "DE", "UA"), ("GR", "Greece", "EL", "UA"),
    ("HU", "Hungary", "HU", "UA"), ("IE", "Ireland", "IE", "UA"), ("IT", "Italy", "IT", "UA"), ("LV", "Latvia", "LV", "UA"),
    ("LT", "Lithuania", "LT", "UA"), ("LU", "Luxembourg", "LU", "LAU"), ("MT", "Malta", "MT", "LAU"), ("NL", "Netherlands", "NL", "UA"),
    ("PL", "Poland", "PL", "UA"), ("PT", "Portugal", "PT", "UA"), ("RO", "Romania", "RO", "UA"), ("SK", "Slovakia", "SK", "UA"),
    ("SI", "Slovenia", "SI", "LAU"), ("ES", "Spain", "ES", "UA"), ("SE", "Sweden", "SE", "UA"),
    ("GB", "United Kingdom", "UK", "UA"), ("NO", "Norway", "NO", "UA"), ("CH", "Switzerland", "CH", "UA"),
    ("TR", "Turkey", "TR", "UA"),
    ("RS", "Serbia", "RS", "LAU"), ("MK", "North Macedonia", "MK", "LAU"), ("AL", "Albania", "AL", "LAU"),
    ("ME", "Montenegro", "ME", "EXTRA"), ("BA", "Bosnia and Herzegovina", "BA", "EXTRA"), ("XK", "Kosovo", "XK", "EXTRA"),
]
EU27_UK = {"AT", "BE", "BG", "HR", "CY", "CZ", "DK", "EE", "FI", "FR", "DE", "GR", "HU", "IE", "IT", "LV", "LT", "LU", "MT", "NL",
           "PL", "PT", "RO", "SK", "SI", "ES", "SE", "GB"}
CY_NAMES = {"Λεμεσός": "Lemesos", "Στρόβολος": "Strovolos", "Λευκωσία": "Lefkosia", "Λάρνακα": "Larnaka", "Πάφος": "Pafos",
            "Λακατάμεια": "Lakatamia"}
SKOPJE = ["Аеродром", "Гази Баба", "Карпош", "Чаир", "Кисела Вода", "Ѓорче Петров", "Центар", "Сарај", "Бутел", "Шуто Оризари"]
MK_NAMES = {"Куманово": "Kumanovo", "Тетово": "Tetovo", "Битола": "Bitola", "Гостивар": "Gostivar", "Прилеп": "Prilep"}
# Urban Audit cities without a GISCO label point (Turkey is not in URAU; Jelgava was dropped in the 2021 vintage): approximate
# city-centre coordinates (lon, lat) - geographic facts, not populations.
FALLBACK_COORDS = {"TR012C": (28.9784, 41.0082), "TR001C": (32.8597, 39.9334), "TR013C": (27.1428, 38.4237), "TR005C": (29.0610, 40.1885),
                   "TR002C": (35.3213, 37.0000), "LV003C": (23.7213, 56.6500)}
OUT_COLS = ["country_code", "country", "city", "lat", "lon", "population", "population_year", "population_source", "url", "note"]


def get(url, **kw):
    r = requests.get(url, headers=UA, timeout=600, **kw)
    r.raise_for_status()
    return r


def ua_cities():
    js = get(ES_API + "urb_cpop1", params={"indic_ur": "DE1001V", "format": "JSON", "lang": "en"}).json()
    codes = list(js["dimension"]["cities"]["category"]["index"])
    lab = js["dimension"]["cities"]["category"]["label"]
    years = list(js["dimension"]["time"]["category"]["index"])
    pop = {}
    for k, v in js["value"].items():
        k = int(k)
        c, y = codes[k // len(years)], years[k % len(years)]
        if c.endswith("C") and (c not in pop or y > pop[c][1]):
            pop[c] = (v, y)
    coords = {}
    for yr in ("2024", "2021", "2020"):
        try:
            g = get(f"{GISCO}urau/geojson/URAU_LB_{yr}_4326.geojson").json()
        except Exception as e:  # noqa: BLE001
            print("URAU", yr, e)
            continue
        for f in g["features"]:
            code = f["properties"]["URAU_CODE"]
            if re.match(r"^[A-Z]{2}\d{3}C\d$", code):   # the 2020 vintage appends a version digit (UK001C1)
                code = code[:6]
            coords.setdefault(code, f["geometry"]["coordinates"])
    for code, (lon, lat) in FALLBACK_COORDS.items():
        coords.setdefault(code, [lon, lat])
    return pop, lab, coords


def clean(name):
    name = re.sub(r"\s*\((greater city)\)", "", name)
    return name.replace("Greater ", "").strip()


def lau_data(wanted):
    """{iso2: DataFrame(name, pop, year, lon, lat)} from the GISCO LAU geojson (polygons -> centroid)."""
    from shapely.geometry import shape
    out = {}
    for yr in ("2023", "2021"):
        js = get(f"{GISCO}lau/geojson/LAU_RG_01M_{yr}_4326.geojson").json()
        for f in js["features"]:
            p = f["properties"]
            cc = p.get("CNTR_CODE")
            if cc not in wanted:
                continue
            pop = p.get(f"POP_{yr}")
            if not pop:
                continue
            c = shape(f["geometry"]).representative_point() if False else shape(f["geometry"]).centroid
            out.setdefault((cc, yr), []).append({"name": p["LAU_NAME"], "pop": float(pop), "lon": c.x, "lat": c.y})
        print("LAU", yr, "read", {k: len(v) for k, v in out.items() if k[1] == yr}, flush=True)
    return {k: pd.DataFrame(v) for k, v in out.items()}


def lau_rows(cc, name, df, yr, merge=None, top=4):
    """Top-4 LAU units (merged groups first) by population."""
    df = df.copy()
    df["city"] = df["name"]
    if merge:
        for city, members in merge.items():
            m = df[df["name"].isin(members)] if isinstance(members, list) else df[df["name"].str.startswith(members)]
            if len(m):
                w = m["pop"]
                row = {"name": city, "city": city, "pop": w.sum(), "lon": (m["lon"] * w).sum() / w.sum(), "lat": (m["lat"] * w).sum() / w.sum()}
                df = pd.concat([df.drop(m.index), pd.DataFrame([row])], ignore_index=True)
    df = df.sort_values("pop", ascending=False).head(top)
    return [dict(country_code=cc, country=name, city=r.city, lat=round(r.lat, 4), lon=round(r.lon, 4), population=int(r.pop),
                 population_year=int(yr), population_source=f"Eurostat/GISCO LAU {yr} municipality population (national statistics office)",
                 url=LAU_URL, note="") for r in df.itertuples()]


def country_populations():
    rows = []
    pj = get(ES_API + "demo_pjan", params={"time": "2025", "sex": "T", "age": "TOTAL", "format": "JSON", "lang": "en"}).json()
    geo = pj["dimension"]["geo"]["category"]["index"]
    geo = {k: i for i, k in enumerate(geo)} if isinstance(geo, list) else geo
    inv = {v: k for k, v in geo.items()}
    es = {inv[int(k)]: v for k, v in pj["value"].items()}
    for cc, name, ecode, _m in COUNTRIES:
        v = es.get(ecode)
        rows.append(dict(country_code=cc, country=name, population=v, population_year=2025 if v else None,
                         source="Eurostat demo_pjan, population on 1 January 2025" if v else "", url=PJAN_URL,
                         aggregates="EUROPE" if cc in EU27_UK else ""))
    df = pd.DataFrame(rows)
    # United Kingdom: ONS mid-year population estimate (series UKPOP); not in Eurostat since 2020
    try:
        t = get("https://www.ons.gov.uk/generator", params={"format": "csv", "uri": "/peoplepopulationandcommunity/populationandmigration/populationestimates/timeseries/ukpop/pop"}).text
        yrs = [(int(a), int(b)) for a, b in re.findall(r'^"(\d{4})","(\d+)"', t, flags=re.M)]
        y, v = max(yrs)
        df.loc[df.country_code == "GB", ["population", "population_year", "source", "url"]] = [
            v, y, f"ONS mid-{y} population estimates, United Kingdom (series UKPOP)",
            "https://www.ons.gov.uk/peoplepopulationandcommunity/populationandmigration/populationestimates/timeseries/ukpop/pop"]
    except Exception as e:  # noqa: BLE001
        print("ONS UK population failed:", e)
        try:
            u = get(ES_API + "demo_pjan", params={"geo": "UK", "sex": "T", "age": "TOTAL", "format": "JSON", "lang": "en"}).json()
            ty = list(u["dimension"]["time"]["category"]["index"])
            vals = {ty[int(k)]: v for k, v in u["value"].items()}
            y = max(vals)
            df.loc[df.country_code == "GB", ["population", "population_year", "source"]] = [vals[y], int(y), f"Eurostat demo_pjan, United Kingdom 1 January {y} (last year Eurostat holds; ONS fetch failed)"]
        except Exception as e2:  # noqa: BLE001
            print("Eurostat UK population failed:", e2)
    return df


def main():
    sys.path.insert(0, HERE)
    pop, lab, coords = ua_cities()
    rows, log = [], []
    need_lau = {"SI", "LU", "MT", "EE", "CY", "RS", "MK", "AL"}
    lau = lau_data(need_lau)
    for cc, name, ecode, method in COUNTRIES:
        if method == "UA":
            codes = sorted([c for c in pop if c.startswith(ecode) and c in coords], key=lambda c: -pop[c][0])[:4]
            if len(codes) < 4:
                log.append(f"{cc}: only {len(codes)} Urban Audit cities")
            for c in codes:
                v, y = pop[c]
                rows.append(dict(country_code=cc, country=name, city=clean(lab[c]), lat=round(coords[c][1], 4), lon=round(coords[c][0], 4),
                                 population=int(v), population_year=int(y),
                                 population_source=f"Eurostat Urban Audit urb_cpop1 DE1001V ({'greater city' if 'greater' in lab[c] else 'city'}), national statistics office figure",
                                 url=UA_URL, note="Turkey: the Urban Audit has no later Turkish figures than 2004" if cc == "TR" else ""))
        elif method == "LAU":
            yr = max((y for y in ("2023", "2021") if (ecode, y) in lau and len(lau[(ecode, y)]) >= 4),
                     key=lambda y: lau[(ecode, y)]["pop"].sum())   # the vintage with the more complete population column
            df = lau[(ecode, yr)]
            merge = None
            if cc == "RS":
                merge = {"Belgrade": "Belgrade - ", "Niš": "Niš - "}
            if cc == "MK":
                merge = {"Skopje": SKOPJE}
            r = lau_rows(cc, name, df, yr, merge)
            for x in r:
                c = CY_NAMES.get(x["city"], MK_NAMES.get(x["city"], x["city"])).replace(" linn", "")
                x["city"] = c.title() if cc == "AL" else c
                if cc in ("LU", "MT", "SI", "CY"):
                    x["note"] = "fewer than 4 meaningful cities: the 4 largest municipalities/localities"
                if cc in ("RS", "MK"):
                    x["note"] = "city = sum of its city municipalities" if x["city"] in ("Belgrade", "Niš", "Skopje") else ""
            rows += r
    try:
        import EXTRA_CITIES_EUROPE as extra   # optional module with ME / BA / XK / TR national-statistics rows
        rows += extra.rows()
    except ImportError:
        pass
    have = {r["country_code"] for r in rows}
    for cc, name, ecode, method in COUNTRIES:
        if cc not in have:
            log.append(f"{cc}: NO cities")
    cities = pd.DataFrame(rows)[OUT_COLS]
    cities.to_csv(os.path.join(HERE, "cities_europe.csv"), index=False)
    cp = country_populations()
    cp.to_csv(os.path.join(HERE, "countries_europe.csv"), index=False)
    print(cities.to_string())
    print(cp.to_string())
    print("\n".join(log))


if __name__ == "__main__":
    main()
