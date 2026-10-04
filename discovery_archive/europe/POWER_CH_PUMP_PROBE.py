"""Find a Swiss pumped-storage consumption (Verbrauch Speicherpumpen) series: opendata.swiss / BFE ogd catalogue and candidate CSVs."""
import requests, io
H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0 Safari/537.36", "Referer": "https://www.bfe.admin.ch/"}
for q in ["Speicherpumpen", "Elektrizitätsbilanz", "Swissgrid", "Elektrizitätsstatistik"]:
    try:
        r = requests.get("https://opendata.swiss/api/3/action/package_search", params={"q": q, "rows": 15}, headers=H, timeout=60)
        for p in r.json()["result"]["results"]:
            print(q, "|", p["name"], "|", (p.get("title") or {}).get("de"))
            for res in p.get("resources", []):
                u = res.get("download_url") or res.get("url")
                if u and u.lower().endswith(".csv"):
                    print("    ", u)
    except Exception as e:
        print(q, "ERR", e)
for n in range(95, 110):
    pass
for u in ["https://www.uvek-gis.admin.ch/BFE/ogd/18/ogd18_energiebilanz_elektrizitaet_monatlich.csv",
          "https://www.uvek-gis.admin.ch/BFE/ogd/17/ogd17_fuf_ee_gesamt.csv",
          "https://www.uvek-gis.admin.ch/BFE/ogd/102/ogd102_stromproduktion_swissgrid_stundlich.csv"]:
    try:
        r = requests.get(u, headers=H, timeout=60)
        print(u, r.status_code, r.text[:600].replace("\n", " // "))
    except Exception as e:
        print(u, "ERR", e)
