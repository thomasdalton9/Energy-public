"""
One-off pull for the Asia-Pacific industrial coal-to-LNG switching study (private repo analysis). Writes raw CSVs to
discovery_archive/rest_of_world/data/ (public, keyless sources):
  ap_energy_balance_un.csv   UNSD energy balances (SDMX DF_UNData_EnergyBalance), TJ, 2010-2023: primary coal & peat
                             (B00_CL), coal products (B01_CP), natural gas (B04_NG); every transaction (supply,
                             transformation, own use, final consumption by industry subsector ...)
  ap_energy_stats_un.csv     UNSD energy statistics (SDMX DF_UNDATA_ENERGY): coking / other bituminous / sub-bituminous
                             coal / lignite / natural gas by industry subsector and autoproducer plants (native units)
  ap_prices_monthly.csv      FRED (IMF): Japan LNG import price, Australian coal, Henry Hub; Yahoo JKM=F if listed
Found by AP_INDUSTRIAL_COAL_PROBE.py.
"""
import io, os, sys, time
import pandas as pd, requests

OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data")
os.makedirs(OUT, exist_ok=True)
H = {"User-Agent": "Mozilla/5.0 (X11; Linux x86_64) energy-data/1.0"}
B = "https://data.un.org/legacy/ws/rest/data/"
AREAS = {"156": "China", "392": "Japan", "410": "South Korea", "356": "India", "586": "Pakistan", "050": "Bangladesh",
         "764": "Thailand", "458": "Malaysia", "608": "Philippines", "704": "Viet Nam", "360": "Indonesia",
         "702": "Singapore", "344": "Hong Kong", "792": "Turkey", "144": "Sri Lanka", "104": "Myanmar", "490": "Other Asia (Taiwan)"}


def sdmx(path, tries=3):
    for t in range(tries):
        try:
            r = requests.get(B + path, headers={**H, "Accept": "application/vnd.sdmx.data+csv"}, timeout=300)
            if r.status_code == 200 and r.text.startswith("DATAFLOW"):
                return pd.read_csv(io.StringIO(r.text), dtype={"REF_AREA": str})
            print(f"  {r.status_code} {len(r.content)} B: {r.text[:200]!r}", flush=True)
        except Exception as e:  # noqa: BLE001
            print(f"  try {t + 1}: {type(e).__name__} {str(e)[:150]}", flush=True)
        time.sleep(5)
    return pd.DataFrame()


frames = []
for code, name in AREAS.items():
    d = sdmx(f"DF_UNData_EnergyBalance/{code}.B00_CL+B01_CP+B04_NG../ALL/?startPeriod=2010&endPeriod=2023")
    print(f"balance {name}: {len(d)} rows {sorted(d['TIME_PERIOD'].unique())[-3:] if len(d) else ''}", flush=True)
    if len(d):
        frames.append(d.assign(country=name))
bal = pd.concat(frames, ignore_index=True)
bal.to_csv(os.path.join(OUT, "ap_energy_balance_un.csv"), index=False)
print("saved balance", bal.shape)

TRANS = "121+1211+1213+1214A+1214B+1214C+1214D+1214E+1214F+1214G+1214H+1214I+1214J+1214O+08812+08822+08811+08821+081+03+GA+01"
frames = []
for code, name in AREAS.items():
    d = sdmx(f"DF_UNDATA_ENERGY/A.{code}.0121+0129+0210+0220+3000.{TRANS}/ALL/?startPeriod=2015&endPeriod=2023")
    print(f"stats {name}: {len(d)} rows", flush=True)
    if len(d):
        frames.append(d.assign(country=name))
if frames:
    st = pd.concat(frames, ignore_index=True)
    st.to_csv(os.path.join(OUT, "ap_energy_stats_un.csv"), index=False)
    print("saved stats", st.shape, "columns", list(st.columns))

prices = {}
for s in ["PNGASJPUSDM", "PCOALAUUSDM", "PNGASUSUSDM", "PNGASEUUSDM"]:
    try:
        r = requests.get(f"https://fred.stlouisfed.org/graph/fredgraph.csv?id={s}", headers=H, timeout=60); r.raise_for_status()
        x = pd.read_csv(io.StringIO(r.text)); x.columns = ["date", s]
        prices[s] = pd.to_numeric(x.set_index(pd.to_datetime(x["date"]))[s], errors="coerce")
        print(f"FRED {s}: {prices[s].dropna().index.min().date()}..{prices[s].dropna().index.max().date()}", flush=True)
    except Exception as e:  # noqa: BLE001
        print(f"FRED {s} failed: {e}", flush=True)
try:
    r = requests.get("https://query1.finance.yahoo.com/v8/finance/chart/JKM=F?range=max&interval=1mo", headers=H, timeout=60)
    j = r.json()["chart"]["result"][0]
    prices["YAHOO_JKM"] = pd.Series(j["indicators"]["quote"][0]["close"], index=pd.to_datetime(j["timestamp"], unit="s").normalize())
    print("Yahoo JKM=F:", prices["YAHOO_JKM"].dropna().index.min(), prices["YAHOO_JKM"].dropna().index.max(), flush=True)
except Exception as e:  # noqa: BLE001
    print("Yahoo JKM failed:", type(e).__name__, str(e)[:150], flush=True)
p = pd.DataFrame(prices); p.index.name = "date"
p.to_csv(os.path.join(OUT, "ap_prices_monthly.csv"))
print("saved prices", p.shape)
