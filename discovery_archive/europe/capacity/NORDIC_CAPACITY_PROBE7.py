"""Probe 7 (12 Oct 2026): Norway NVE hydro/wind register fields (decommission dates) and a 1 January rebuild against ENTSO-E A68."""
import collections, json, os
import pandas as pd, requests
S = requests.Session(); S.headers["User-Agent"] = "Mozilla/5.0 (energy-data probe)"
h = S.get("https://api.nve.no/web/Powerplant/GetHydroPowerPlants", timeout=120).json()
print("hydro", len(h))
for k in ["Kraftverkstatus", "ErIDrift", "UnderBygging", "VannKVType", "VannKVTypeID"]:
    print(k, collections.Counter(str(x.get(k)) for x in h).most_common(8))
print("UteAvDrift examples", collections.Counter(str(x.get("UteAvDrift"))[:25] for x in h).most_common(8))
ng = [x for x in h if not x.get("ErIDrift")]
for x in ng[:6]:
    print({k: x.get(k) for k in ["Navn", "VannKVType", "MaksYtelse", "IDriftDato", "UteAvDrift", "UnderBygging", "Kraftverkstatus", "ForsteUtnyttelseAvFalletDato", "DatoForEldsteKraftproduserendeDel"]})
print("not in operation with IDriftDato:", sum(1 for x in ng if x.get("IDriftDato")), "of", len(ng), "MW", sum(x.get("MaksYtelse") or 0 for x in ng))
w = S.get("https://api.nve.no/web/WindPowerplant/GetWindPowerplants", timeout=120).json()
tu = [dict(t, plant=p["Navn"]) for p in w for t in p["Turbiner"]]
print("wind plants", len(w), "turbine rows", len(tu), "with DatoUtavdrift", sum(1 for t in tu if t.get("DatoUtavdrift")))
print([ (t["plant"], t["AntallTurbiner"], t["TurbinStorrelse_kW"], t["DatoIdriftsatt"], t["DatoUtavdrift"]) for t in tu if t.get("DatoUtavdrift")][:8])
print("no DatoIdriftsatt:", sum(1 for t in tu if not t.get("DatoIdriftsatt")))
print("plants with offshore-like names", [p["Navn"] for p in w if "Hywind" in p["Navn"] or "Tampen" in p["Navn"]])
def hydro_at(d):
    d = pd.Timestamp(d); hy = ps = 0
    for x in h:
        t = pd.to_datetime(x.get("IDriftDato"), errors="coerce"); mw = x.get("MaksYtelse") or 0
        if pd.isna(t) or t >= d or mw <= 0 or x.get("UnderBygging"): continue
        if not x.get("ErIDrift"): continue
        if x["VannKVType"] == "Pumpekraftverk": ps += mw
        elif x["VannKVType"] == "Kraftverk": hy += mw
    return hy, ps
def wind_at(d):
    d = pd.Timestamp(d); s = 0
    for t in tu:
        a = pd.to_datetime(t.get("DatoIdriftsatt"), errors="coerce"); b = pd.to_datetime(t.get("DatoUtavdrift"), errors="coerce")
        if pd.isna(a) or a >= d or (not pd.isna(b) and b < d): continue
        s += t["AntallTurbiner"] * t["TurbinStorrelse_kW"] / 1000
    return s
e = pd.read_excel("output/Data and Chart Outputs/norway_power_capacity.xlsx", sheet_name="Monthly"); e["date"] = pd.to_datetime(e["date"])
for _, r in e.iterrows():
    d = r["date"]; hy, ps = hydro_at(d)
    print(d.date(), "hydro rebuild/ENTSOE", round(hy), r["Hydro_MW"], "| pumped", round(ps), r["PumpedStorage_MW"], "| wind", round(wind_at(d)), r["Wind_MW"])
print("recent wind", [(str(x), round(wind_at(x))) for x in pd.date_range("2024-01-01", "2026-10-01", freq="6MS")])
