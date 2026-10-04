"""Probe national power-generation sources (DE/IT/PL/RO/BG) reachable from GitHub Actions (round 2): Eurostat monthly net generation, Energy-Charts files, PSE API."""
import requests, json, sys
import pandas as pd
H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0 Safari/537.36"}


def jsonstat(j):
    ids, size = j["id"], j["size"]
    cats = [list(j["dimension"][d]["category"]["index"].keys()) if isinstance(j["dimension"][d]["category"]["index"], dict) else j["dimension"][d]["category"]["index"] for d in ids]
    idx = pd.MultiIndex.from_product(cats, names=ids)
    s = pd.Series(float("nan"), index=idx)
    vals = j["value"]
    if isinstance(vals, dict):
        for k, v in vals.items():
            s.iloc[int(k)] = v
    else:
        s[:] = vals
    return s


def eurostat():
    for geo in ["DE", "IT", "PL", "RO", "BG", "EL"]:
        u = f"https://ec.europa.eu/eurostat/api/dissemination/statistics/1.0/data/nrg_cb_pem?format=JSON&lang=EN&geo={geo}&unit=GWH&sinceTimePeriod=2022-01"
        r = requests.get(u, headers=H, timeout=60)
        print(geo, r.status_code, len(r.content), flush=True)
        if r.status_code != 200:
            print(r.text[:300]); continue
        j = r.json()
        s = jsonstat(j).dropna().reset_index(name="v")
        s["year"] = s["time"].str[:4]
        print("  last month", s["time"].max(), "sieC codes", sorted(s["siec"].unique()))
        t = s.groupby(["siec", "year"])["v"].sum().unstack().round(0)
        print(t.to_string(), flush=True)
        n = s.groupby(["siec", "year"])["time"].nunique().unstack()
        print("  months per year (TOTAL):", n.loc["TOTAL"].to_dict() if "TOTAL" in n.index else n.head(2).to_dict(), flush=True)


def ec():
    for u in ["https://energy-charts.info/charts/power/data/de/month_2024.json", "https://energy-charts.info/charts/power/data/de/year_2024.json",
              "https://energy-charts.info/charts/energy/data/de/year_2024.json", "https://energy-charts.info/charts/energy/data/de/month_2024.json",
              "https://energy-charts.info/charts/energy/data/de/month_sum_2024.json",
              "https://energy-charts.info/charts/power/data/it/week_2024_23.json", "https://energy-charts.info/charts/power/data/pl/week_2024_23.json",
              "https://api.energy-charts.info/public_power?country=de&start=2024-06-01&end=2024-06-02",
              "https://api.energy-charts.info/total_power?country=de&start=2024-06-01&end=2024-06-02",
              "https://api.energy-charts.info/ren_share_daily_avg?country=de&year=2024"]:
        try:
            r = requests.get(u, headers=H, timeout=40)
            print(r.status_code, len(r.content), u, flush=True)
            if r.status_code == 200:
                try:
                    j = r.json()
                    if isinstance(j, list):
                        for it in j[:40]:
                            nm = it.get("name", {}); nm = nm.get("en", nm) if isinstance(nm, dict) else nm
                            d = it.get("data", [])
                            print("    ", nm, len(d), str(d[:3])[:80], flush=True)
                    else:
                        print("    keys", list(j)[:20], str(j)[:300], flush=True)
                except Exception as e:
                    print("    notjson", r.text[:200], flush=True)
        except Exception as e:
            print("ERR", u, type(e).__name__, flush=True)


def pse():
    for ep in ["his-wlk-cal", "his-gen-pal", "kse-load", "gen-jw", "poze-redoze", "pdgsz"]:
        for flt in ["business_date eq '2024-06-02'", "doba eq '2024-06-02'"]:
            u = f"https://api.raporty.pse.pl/api/{ep}?$filter={flt}&$first=3"
            try:
                r = requests.get(u, headers=H, timeout=40)
                print(r.status_code, len(r.content), ep, flt, flush=True)
                print("    ", r.text[:600].replace("\n", " "), flush=True)
                if r.status_code == 200:
                    break
            except Exception as e:
                print("ERR", ep, type(e).__name__, flush=True)


eurostat(); ec(); pse()
