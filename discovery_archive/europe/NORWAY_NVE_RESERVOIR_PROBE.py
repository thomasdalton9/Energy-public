"""
One-off probe: Norway hydro reservoir filling (NVE magasinstatistikk, public, keyless).

  https://biapi.nve.no/magasinstatistikk/api/Magasinstatistikk/HentOffentligData
      weekly filling by area (omrType NO = whole country, EL = price areas NO1-NO5)
  https://biapi.nve.no/magasinstatistikk/api/Magasinstatistikk/HentOffentligDataMinMaxMedian
      20-year min / max / median filling per ISO week

Prints the national weekly series (from 2019) and the min/max/median table as CSV to stdout.
"""
import pandas as pd
import requests

BASE = "https://biapi.nve.no/magasinstatistikk/api/Magasinstatistikk/"
HEADERS = {"User-Agent": "Mozilla/5.0", "Accept": "application/json"}

for name in ("HentOffentligData", "HentOffentligDataMinMaxMedian"):
    r = requests.get(BASE + name, headers=HEADERS, timeout=(10, 120))
    print(name, r.status_code, len(r.content))
    r.raise_for_status()
    d = pd.DataFrame(r.json())
    print("columns:", list(d.columns))
    print(d.head(3).to_string())
    nat = d[d["omrType"] == "NO"] if "omrType" in d.columns else d
    if name == "HentOffentligData":
        nat = nat[nat["iso_aar"] >= 2019]
        keep = [c for c in ("iso_aar", "iso_uke", "dato_Id", "fyllingsgrad", "kapasitet_TWh", "fylling_TWh") if c in nat.columns]
    else:
        keep = [c for c in nat.columns if c in ("iso_uke", "minFyllingsgrad", "maxFyllingsgrad", "medianFyllingsgrad",
                                                 "minFyllingTWH", "maxFyllingTWH", "medianFylling_TWH")]
    print(f"=== {name} CSV START ===")
    print(nat[keep].sort_values(keep[:2]).to_csv(index=False))
    print(f"=== {name} CSV END ===")
