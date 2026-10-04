"""
Probe for the Europe power supply/load gaps (Italy, Poland, Bulgaria, Romania, Great Britain). Prints only.
 1. Eurostat nrg_cb_em (monthly electricity balance): which nrg_bal lines exist for IT/BG/RO/PL/GB and their annual totals 2022-25
 2. Eurostat nrg_cb_pem Italy by siec, annual
 3. PSE open-data API (api.raporty.pse.pl): which endpoints answer for dates in 2022-2024
 4. Terna download / public API reachability
 5. NESO data portal: demand datasets (ND, TSD, embedded) for 2024
"""
import json
import sys

import pandas as pd
import requests

H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0 Safari/537.36", "Accept": "*/*"}
ES = "https://ec.europa.eu/eurostat/api/dissemination/statistics/1.0/data/"


def js(j):
    ids = j["id"]
    cats = [list(j["dimension"][d]["category"]["index"].keys()) if isinstance(j["dimension"][d]["category"]["index"], dict)
            else list(j["dimension"][d]["category"]["index"]) for d in ids]
    s = pd.Series(float("nan"), index=pd.MultiIndex.from_product(cats, names=ids), dtype=float)
    v = j["value"]
    for k, x in (v.items() if isinstance(v, dict) else enumerate(v)):
        if x is not None:
            s.iloc[int(k)] = x
    lab = {d: j["dimension"][d]["category"].get("label", {}) for d in ids}
    return s.dropna(), lab


def eurostat(ds, **p):
    try:
        r = requests.get(ES + ds, params={"format": "JSON", "lang": "EN", **p}, headers=H, timeout=120)
        print(ds, p, r.status_code, flush=True)
        if r.status_code != 200:
            print(r.text[:300])
            return None, None
        return js(r.json())
    except Exception as e:  # noqa: BLE001
        print(ds, "ERR", e)
        return None, None


def part1():
    for geo in ("IT", "BG", "RO", "PL", "UK"):
        s, lab = eurostat("nrg_cb_em", geo=geo, unit="GWH", sinceTimePeriod="2022-01")
        if s is None:
            continue
        df = s.reset_index(name="v")
        df["year"] = df["time"].str[:4]
        df["n"] = df.groupby(["nrg_bal", "siec", "year"])["v"].transform("count")
        t = df[df["siec"].isin(["E7000", "TOTAL"])].pivot_table(index="nrg_bal", columns="year", values="v", aggfunc="sum")
        names = lab.get("nrg_bal", {})
        t.insert(0, "label", [names.get(i, "")[:60] for i in t.index])
        print(geo, "nrg_cb_em siecs:", sorted(df["siec"].unique())[:20])
        print(t.round(0).to_string(), flush=True)
    s, lab = eurostat("nrg_cb_pem", geo="IT", unit="GWH", sinceTimePeriod="2022-01")
    if s is not None:
        df = s.reset_index(name="v")
        df["year"] = df["time"].str[:4]
        t = df.pivot_table(index="siec", columns="year", values="v", aggfunc="sum")
        t.insert(0, "label", [lab["siec"].get(i, "")[:40] for i in t.index])
        print(t.round(0).to_string(), flush=True)


def part3():
    base = "https://api.raporty.pse.pl/api/"
    for ep in ("his-wlk-cal", "kse-load", "zap-kse", "pk5l-wp", "his-gen-pal", "gen-jw", "poze-redoze", "dem-kse"):
        for day in ("2022-06-01", "2023-06-01", "2024-01-15", "2024-06-02", "2024-07-15", "2025-06-01"):
            try:
                r = requests.get(base + ep, params={"$filter": f"doba eq '{day}'", "$first": 3}, headers=H, timeout=40)
                v = r.json().get("value", []) if r.headers.get("content-type", "").startswith("application/json") else []
                print("PSE", ep, day, r.status_code, len(v), (json.dumps(v[0])[:200] if v else r.text[:100]), flush=True)
            except Exception as e:  # noqa: BLE001
                print("PSE", ep, day, "ERR", str(e)[:80], flush=True)
                break


def part4():
    for u in ("https://www.terna.it/it/sistema-elettrico/transparency-report/total-load",
              "https://download.terna.it/terna/0000/0000/00/00.xlsx",
              "https://api.terna.it/transparency/v1.0/getactualtotalload",
              "https://www.terna.it/it/sistema-elettrico/statistiche/pubblicazioni-statistiche",
              "https://dati.terna.it/en/download-center", "https://www.gme.it/en"):
        try:
            r = requests.get(u, headers=H, timeout=40)
            print("TERNA", u, r.status_code, len(r.content), flush=True)
        except Exception as e:  # noqa: BLE001
            print("TERNA", u, "ERR", str(e)[:80], flush=True)


def part5():
    try:
        r = requests.get("https://api.neso.energy/api/3/action/package_search", params={"q": "historic demand", "rows": 10}, headers=H, timeout=60)
        print("NESO", r.status_code, flush=True)
        for p in r.json()["result"]["results"][:10]:
            print(" ", p["name"], [(x["name"], x["format"], x["url"][-60:]) for x in p["resources"][:12]], flush=True)
    except Exception as e:  # noqa: BLE001
        print("NESO ERR", str(e)[:120])


for f in (part1, part3, part4, part5):
    try:
        f()
    except Exception as e:  # noqa: BLE001
        print(f.__name__, "failed", e, flush=True)
sys.exit(0)
