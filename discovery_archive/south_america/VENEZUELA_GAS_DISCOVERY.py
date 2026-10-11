"""
One-off probe: is there any Venezuela natural gas demand (by sector or total) series obtainable?

Sandbox cannot reach jodidata.org, opec.org, energyinst.org, pdvsa.com, mppee.gob.ve (proxy 403). Run in GitHub
Actions (discovery_archive/workflows/venezuela_gas_discovery.yml); logs status of each candidate, saves pages and
files under ve_raw/, and prints the Venezuela rows of the JODI-Gas CSV (flows, units, months covered) so
south_america/VENEZUELA_GAS.py can be confirmed against the real layout.

Candidates: JODI-Gas downloads page + CSV zip; OPEC Annual Statistical Bulletin (annual, no sectors); PDVSA
'Informe de Gestion Anual' / 'PDVSA en cifras' (last gas-by-use tables ~2016, PDF); MPPEE (ministry) site;
Energy Institute Statistical Review (annual totals); Ember/EIA international as labelled fallbacks.
"""
import io
import os
import re
import zipfile

import pandas as pd
import requests

OUT = "ve_raw"
H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0 Safari/537.36"}
PAGES = [
    "https://www.jodidata.org/gas/database/data-downloads.aspx",
    "https://www.opec.org/opec_web/en/publications/202.htm",
    "https://www.pdvsa.com/index.php?option=com_content&view=article&id=6648&Itemid=1220&lang=es",
    "https://www.pdvsa.com/index.php/es/informe-de-gestion-anual",
    "https://www.mppee.gob.ve/",
    "https://www.energyinst.org/statistical-review",
    "https://ourworldindata.org/grapher/gas-consumption-by-country.csv",
]
ZIPS = ["https://www.jodidata.org/_resources/files/downloads/gas-data/jodi_gas_csv_beta.zip",
        "https://www.jodidata.org/_resources/files/downloads/gas-data/jodi_gas_csv.zip"]


def save(name, content):
    os.makedirs(OUT, exist_ok=True)
    with open(os.path.join(OUT, re.sub(r"[^A-Za-z0-9._-]+", "_", name)[-120:]), "wb") as f:
        f.write(content)


def main():
    for u in PAGES:
        try:
            r = requests.get(u, headers=H, timeout=(20, 90))
            print(f"{r.status_code} {len(r.content):>8} {u}")
            if r.status_code == 200:
                save(u, r.content)
                for l in sorted(set(re.findall(r'href="([^"]+\.(?:xlsx?|csv|pdf|zip))"', r.text, re.I)))[:80]:
                    print("     link:", l)
        except requests.RequestException as e:
            print(f"ERR {type(e).__name__} {u}")
    for u in ZIPS:
        try:
            r = requests.get(u, headers=H, timeout=(20, 300))
            print(f"{r.status_code} {len(r.content):>10} {r.headers.get('Last-Modified')} {u}")
            if r.status_code != 200:
                continue
            z = zipfile.ZipFile(io.BytesIO(r.content))
            print("zip members:", z.namelist())
            df = pd.read_csv(z.open(next(n for n in z.namelist() if n.lower().endswith(".csv"))), dtype=str)
            print("columns:", list(df.columns))
            ve = df[df.iloc[:, 0].str.upper().isin(["VE", "VENEZUELA"])]
            print("Venezuela rows:", len(ve))
            for c in df.columns:
                if c.lower() in ("flow_breakdown", "unit_measure", "energy_product", "assessment_code"):
                    print(f"  {c}: {sorted(ve[c].dropna().unique())}")
            tp = [c for c in df.columns if "time" in c.lower()]
            if tp and len(ve):
                print("  months:", ve[tp[0]].min(), "..", ve[tp[0]].max())
            os.makedirs(OUT, exist_ok=True)
            ve.to_csv(os.path.join(OUT, "jodi_venezuela.csv"), index=False)
            break
        except (requests.RequestException, zipfile.BadZipFile, StopIteration) as e:
            print(f"ERR {type(e).__name__} {u}")


if __name__ == "__main__":
    main()
