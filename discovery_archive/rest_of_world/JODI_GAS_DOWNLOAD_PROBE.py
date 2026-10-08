"""One-off probe: find the current JODI Gas world database download (the *_csv_beta.zip file stops at 2018-08).
Prints every link on the JODI download pages and tries candidate zip URLs; also checks FRED reachability."""
import io, re, zipfile
import pandas as pd, requests

H = {"User-Agent": "Mozilla/5.0 (X11; Linux x86_64) energy-data/1.0"}
for page in ["https://www.jodidata.org/gas/database/data-downloads.aspx", "https://www.jodidata.org/gas/database/",
             "https://www.jodidata.org/oil/database/data-downloads.aspx", "https://www.jodidata.org/gas/",
             "https://www.jodidata.org/gas/database/world-database.aspx", "https://www.jodidb.org/"]:
    try:
        r = requests.get(page, headers=H, timeout=60)
        links = sorted(set(re.findall(r'(?:href|src)="([^"]+)"', r.text)))
        print(f"\n== {page} {r.status_code} {len(r.text)}B")
        for l in links:
            if any(k in l.lower() for k in ["download", ".zip", ".csv", ".xls", "_resources", "jodidb", "database", "api"]):
                print("  ", l)
        for m in re.findall(r"[^<>\"']{0,80}(?:csv|zip|xlsx?)[^<>\"']{0,80}", r.text, flags=re.I)[:40]:
            print("   txt:", m.strip()[:160])
    except Exception as e:  # noqa: BLE001
        print(page, "failed:", type(e).__name__, str(e)[:160])

base = "https://www.jodidata.org/_resources/files/downloads/gas-data/"
cands = ["jodi_gas_csv_beta.zip", "jodi_gas_csv.zip", "jodi_gas_world_csv.zip", "world_gas_csv.zip", "jodi_gas_beta_csv.zip",
         "jodi_gas_world_database.zip", "jodi_gas.zip", "jodi_gas_csv_full.zip", "jodi_gas_world.zip", "gas_world_csv.zip",
         "jodi-gas-csv.zip", "JODI_Gas_CSV.zip", "jodi_gas_csv_beta_2.zip", "jodi_gas_monthly.zip"]
for c in cands:
    try:
        r = requests.head(base + c, headers=H, timeout=30, allow_redirects=True)
        print("HEAD", c, r.status_code, r.headers.get("content-length"), r.headers.get("last-modified"))
    except Exception as e:  # noqa: BLE001
        print("HEAD", c, "failed", type(e).__name__)
for u in ["https://www.jodidata.org/_resources/files/downloads/oil-data/world_primary_csv.zip",
          "https://www.jodidata.org/_resources/files/downloads/oil-data/world_secondary_csv.zip"]:
    try:
        r = requests.head(u, headers=H, timeout=30, allow_redirects=True)
        print("HEAD oil", u.split("/")[-1], r.status_code, r.headers.get("content-length"), r.headers.get("last-modified"))
    except Exception as e:  # noqa: BLE001
        print("HEAD oil failed", type(e).__name__)
# JODI beta zip: check last-modified and the latest period per country
try:
    r = requests.get(base + "jodi_gas_csv_beta.zip", headers=H, timeout=300)
    print("beta zip last-modified:", r.headers.get("last-modified"))
    z = zipfile.ZipFile(io.BytesIO(r.content)); d = pd.read_csv(z.open(z.namelist()[0]), dtype=str)
    print(d.groupby("REF_AREA").TIME_PERIOD.max().value_counts().head(10))
except Exception as e:  # noqa: BLE001
    print("beta check failed", type(e).__name__, str(e)[:160])
# FRED
for u in ["https://fred.stlouisfed.org/graph/fredgraph.csv?id=PNGASJPUSDM", "https://api.stlouisfed.org/", "https://fred.stlouisfed.org/"]:
    try:
        r = requests.get(u, headers=H, timeout=120); print("FRED", u, r.status_code, len(r.content), r.text[:80].replace("\n", " | "))
    except Exception as e:  # noqa: BLE001
        print("FRED", u, "failed", type(e).__name__, str(e)[:100])
