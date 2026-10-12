"""Rows for countries outside Eurostat's Urban Audit / GISCO LAU (used by BUILD_CITIES_EUROPE.py).

Kosovo: Kosovo Agency of Statistics (ASK) PxWeb API, table 'Estimated data by municipality, 2012 - 2025' (population estimate,
latest year, the 4 largest municipalities). City-centre coordinates below are approximate geographic coordinates, not populations.
Montenegro, Bosnia and Herzegovina: MONSTAT (Census 2023 results) and the Agency for Statistics of BiH publish their tables as
PDF/Excel files that could not be read reliably, so NO population is entered: the four cities are the largest by the last
censuses as generally known (NOT an extracted official figure), populations are blank and degree_days.py gives them equal weight,
flagged ASSUMPTION. Replace the blanks with sourced figures to switch the weights to POPULATION.
"""
import requests

UA = {"User-Agent": "Mozilla/5.0"}
BASE = ("https://askdata.rks-gov.net/api/v1/en/ASKdata/Population/Estimate%2C%20projection%20and%20structure%20of%20population/"
        "Population%20estimate/tab001.px")
URL = "https://askdata.rks-gov.net/pxweb/en/ASKdata/ASKdata__Population__Estimate,%20projection%20and%20structure%20of%20population__Population%20estimate/"
XK_COORDS = {"prish": ("Pristina", 42.6629, 21.1655), "prizren": ("Prizren", 42.2139, 20.7397), "ferizaj": ("Ferizaj", 42.3702, 21.1483),
             "gjilan": ("Gjilan", 42.4635, 21.4694), "pej": ("Peja", 42.6593, 20.2883), "gjakov": ("Gjakova", 42.3803, 20.4308),
             "mitrovic": ("Mitrovica", 42.8914, 20.8660), "podujev": ("Podujeva", 42.9107, 21.1930)}
ASSUMED = {
    "ME": ("Montenegro", [("Podgorica", 42.4304, 19.2594), ("Niksic", 42.7731, 18.9445), ("Bijelo Polje", 43.0383, 19.7478), ("Bar", 42.0937, 19.1003)]),
    "BA": ("Bosnia and Herzegovina", [("Sarajevo", 43.8563, 18.4131), ("Banja Luka", 44.7722, 17.1910), ("Tuzla", 44.5384, 18.6671), ("Zenica", 44.2034, 17.9077)]),
}


def kosovo():
    meta = requests.get(BASE, headers=UA, timeout=120).json()
    var = {v["code"]: dict(zip(v["values"], v["valueTexts"])) for v in meta["variables"]}
    d = requests.post(BASE, json={"query": [], "response": {"format": "json"}}, headers=UA, timeout=120).json()
    years = var["Year"]
    last = max(years, key=lambda k: int(str(years[k])[:4]))
    rows = []
    for r in d["data"]:
        mk, yk = r["key"]
        if yk == last and r["values"][0] not in ("", "-", ".."):
            rows.append((var["Municipality"][mk], float(r["values"][0])))
    rows = [(n, p) for n, p in rows if n.strip().lower() not in ("kosovo", "total")]
    rows.sort(key=lambda x: -x[1])
    out = []
    for n, p in rows[:4]:
        key = next((k for k in XK_COORDS if n.lower().startswith(k) or k in n.lower()), None)
        if key is None:
            raise RuntimeError(f"no coordinates for Kosovo municipality {n}")
        nm, lat, lon = XK_COORDS[key]
        out.append(dict(country_code="XK", country="Kosovo", city=nm, lat=lat, lon=lon, population=int(p), population_year=int(str(years[last])[:4]),
                        population_source="Kosovo Agency of Statistics (ASK), population estimate by municipality", url=URL,
                        note="municipality population (city + rural area); coordinates approximate city centre"))
    return out


def rows():
    out = []
    try:
        out += kosovo()
    except Exception as e:  # noqa: BLE001
        print("Kosovo failed:", e)
        out += [dict(country_code="XK", country="Kosovo", city=c, lat=la, lon=lo, population=None, population_year=None, population_source="",
                     url="", note="ASSUMPTION: population not sourced") for c, la, lo in
                [("Pristina", 42.6629, 21.1655), ("Prizren", 42.2139, 20.7397), ("Ferizaj", 42.3702, 21.1483), ("Gjilan", 42.4635, 21.4694)]]
    for cc, (name, cities) in ASSUMED.items():
        for c, la, lo in cities:
            out.append(dict(country_code=cc, country=name, city=c, lat=la, lon=lo, population=None, population_year=None, population_source="",
                            url="", note="ASSUMPTION: official city populations not extracted; city choice from general knowledge of the last censuses; equal weights"))
    return out
