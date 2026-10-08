"""
One-off pull for the JKM -> LNG-demand lag study (private repo analysis). Writes raw CSVs to
discovery_archive/rest_of_world/data/ (public, keyless sources):
  jodi_gas_world_NewFormat.csv.gz  JODI Gas World Database (monthly, by country, 2009-01 to the latest month): all flows
                             (IMPLNG = LNG imports, IMPPIP, TOTIMPSB, INDPROD, TOTDEMO, STOCKCH, CLOSTLV ...), units M3/TJ/KTONS.
                             Direct file: https://www.jodidata.org/_resources/files/downloads/gas-data/GAS_world_NewFormat.zip
                             (the downloads page lists it through a JavaScript file-list, served from /jodi-publisher/gas/<id>/;
                             the older jodi_gas_csv_beta.zip at the same path is frozen at 2018-08 - do not use it).
                             JODI_LNG_LAG_PULL4.py is the fallback that scrapes the same table from the jodidb.org viewer.
  jkm_daily_yahoo.csv        JKM front-month futures, daily closes (Yahoo Finance JKM=F) -> monthly means in the analysis
  jkm_monthly_yahoo.csv      same, Yahoo 1mo bars (fallback / cross-check)
  lag_prices_fred_monthly.csv  FRED (IMF): Japan LNG import price (oil-indexed contract proxy), Europe gas, Henry Hub,
                             Brent, Australian coal; OECD industrial production indices for Japan, Korea, India, China
  lag_prices_wb_monthly.csv  World Bank Pink Sheet monthly (Japan LNG, Europe gas, coal) - cross-check
Runs in GitHub Actions (the Claude sandbox cannot reach these hosts).
"""
import gzip
import io
import os
import re
import zipfile

import pandas as pd
import requests

OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data")
os.makedirs(OUT, exist_ok=True)
H = {"User-Agent": "Mozilla/5.0 (X11; Linux x86_64) energy-data/1.0"}
FLOWS = {"IMPLNG", "IMPPIP", "TOTIMPSB", "INDPROD", "TOTDEMO", "TOTDEMC", "EXPLNG", "TOTEXPSB", "STOCKCH", "CLOSTLV",
         "MAINTOT", "OSOURCES"}

# ---------------------------------------------------------------- 1. JODI Gas
def jodi_candidates():
    urls = ["https://www.jodidata.org/_resources/files/downloads/gas-data/GAS_world_NewFormat.zip"]
    for page in ["https://www.jodidata.org/gas/database/data-downloads.aspx", "https://www.jodidata.org/gas/"]:
        try:
            t = requests.get(page, headers=H, timeout=60).text
            for m in re.findall(r'href="([^"]+\.zip)"', t, flags=re.I):
                u = m if m.startswith("http") else "https://www.jodidata.org" + (m if m.startswith("/") else "/" + m)
                if u not in urls:
                    urls.append(u)
            print("JODI page", page, "zip links:", [u for u in urls[1:]], flush=True)
        except Exception as e:  # noqa: BLE001
            print("JODI page failed:", page, type(e).__name__, str(e)[:160], flush=True)
    return urls


got = False
for url in jodi_candidates():
    if "gas" not in url.lower() or "csv" not in url.lower():
        continue
    try:
        r = requests.get(url, headers=H, timeout=600)
        print("JODI", url, r.status_code, len(r.content), r.headers.get("content-type"), flush=True)
        if not r.ok or len(r.content) < 10000:
            continue
        z = zipfile.ZipFile(io.BytesIO(r.content))
        names = z.namelist(); print("  zip members:", names, flush=True)
        frames = []
        for n in names:
            if not n.lower().endswith(".csv"):
                continue
            d = pd.read_csv(z.open(n), dtype=str)
            print(f"  {n}: {d.shape}; columns {list(d.columns)}", flush=True)
            frames.append(d)
        d = pd.concat(frames, ignore_index=True)
        for c in d.columns:
            if c.upper() not in ("OBS_VALUE", "TIME_PERIOD"):
                vals = d[c].unique()
                print(f"  {c}: {len(vals)} values: {list(vals)[:60]}", flush=True)
        tp = [c for c in d.columns if c.upper() == "TIME_PERIOD"][0]
        print("  TIME_PERIOD range:", d[tp].min(), "..", d[tp].max(), flush=True)
        fb = [c for c in d.columns if c.upper() == "FLOW_BREAKDOWN"]
        if fb:
            keep = d[d[fb[0]].str.upper().isin(FLOWS)]
            print(f"  kept flows {sorted(keep[fb[0]].unique())}: {keep.shape}", flush=True)
        else:
            keep = d
        with gzip.open(os.path.join(OUT, "jodi_gas_world_NewFormat.csv.gz"), "wt") as f:
            keep.to_csv(f, index=False)
        print("  wrote jodi_gas_world_NewFormat.csv.gz", os.path.getsize(os.path.join(OUT, "jodi_gas_world_NewFormat.csv.gz")), "B", flush=True)
        got = True
        break
    except Exception as e:  # noqa: BLE001
        print("JODI failed:", url, type(e).__name__, str(e)[:300], flush=True)
if not got:
    print("JODI: NO DATA SAVED", flush=True)

# ---------------------------------------------------------------- 2. JKM (Yahoo)
for interval, fname in [("1d", "jkm_daily_yahoo.csv"), ("1mo", "jkm_monthly_yahoo.csv")]:
    try:
        r = requests.get(f"https://query1.finance.yahoo.com/v8/finance/chart/JKM=F?range=max&interval={interval}",
                         headers=H, timeout=60)
        j = r.json()["chart"]["result"][0]
        q = j["indicators"]["quote"][0]
        d = pd.DataFrame({"date": pd.to_datetime(j["timestamp"], unit="s").normalize(), "open": q.get("open"),
                          "high": q.get("high"), "low": q.get("low"), "close": q.get("close"), "volume": q.get("volume")})
        d = d.dropna(subset=["close"])
        d.to_csv(os.path.join(OUT, fname), index=False)
        print(f"Yahoo JKM=F {interval}: {d.date.min().date()} .. {d.date.max().date()} n={len(d)}", flush=True)
        print(d.tail(3).to_string(), flush=True)
    except Exception as e:  # noqa: BLE001
        print(f"Yahoo JKM=F {interval} failed:", type(e).__name__, str(e)[:300], flush=True)

# ---------------------------------------------------------------- 3. FRED
FRED = {"PNGASJPUSDM": "Japan LNG import price (IMF), $/MMBtu", "PNGASEUUSDM": "Europe gas TTF-based (IMF), $/MMBtu",
        "PNGASUSUSDM": "Henry Hub (IMF), $/MMBtu", "POILBREUSDM": "Brent (IMF), $/bbl",
        "PCOALAUUSDM": "Australian coal (IMF), $/t",
        "JPNPROINDMISMEI": "Japan industrial production index (OECD MEI)",
        "KORPROINDMISMEI": "Korea industrial production index (OECD MEI)",
        "INDPROINDMISMEI": "India industrial production index (OECD MEI)",
        "CHNPRINTO01IXPYM": "China industrial production, % y/y (OECD MEI)",
        "CHNPROINDMISMEI": "China industrial production index (OECD MEI)"}
cols = {}
for s, label in FRED.items():
    try:
        r = requests.get(f"https://fred.stlouisfed.org/graph/fredgraph.csv?id={s}", headers=H, timeout=60)
        r.raise_for_status()
        d = pd.read_csv(io.StringIO(r.text)); d.columns = ["date", s]
        d["date"] = pd.to_datetime(d.date); x = pd.to_numeric(d[s], errors="coerce")
        x.index = d.date; cols[s] = x.resample("MS").mean()
        print(f"FRED {s} ({label}): {x.dropna().index.min().date()} .. {x.dropna().index.max().date()} n={x.notna().sum()}", flush=True)
    except Exception as e:  # noqa: BLE001
        print(f"FRED {s} failed:", type(e).__name__, str(e)[:200], flush=True)
if cols:
    F = pd.DataFrame(cols); F = F[F.index >= "2000-01-01"]; F.index.name = "date"
    F.round(4).to_csv(os.path.join(OUT, "lag_prices_fred_monthly.csv"))
    print("FRED saved", F.shape, flush=True)

# ---------------------------------------------------------------- 4. World Bank pink sheet
try:
    page = requests.get("https://www.worldbank.org/en/research/commodity-markets", headers=H, timeout=60).text
    m = re.search(r'https://thedocs\.worldbank\.org/[^"\']*CMO-Historical-Data-Monthly\.xlsx', page)
    url = m.group(0) if m else "https://thedocs.worldbank.org/en/doc/18675f1d1639c7a34d463f59263ba0a2-0050012025/related/CMO-Historical-Data-Monthly.xlsx"
    r = requests.get(url, headers=H, timeout=120); r.raise_for_status()
    wb = pd.read_excel(io.BytesIO(r.content), sheet_name="Monthly Prices", header=4)
    wb = wb.rename(columns={wb.columns[0]: "month"}); wb = wb[wb.month.astype(str).str.match(r"\d{4}M\d{2}")]
    keep = [c for c in wb.columns if any(k in str(c).lower() for k in ["natural gas", "liquefied", "coal, austral", "brent"])]
    wb[["month"] + keep].to_csv(os.path.join(OUT, "lag_prices_wb_monthly.csv"), index=False)
    print("World Bank:", url, keep, wb.month.iloc[-1], flush=True)
except Exception as e:  # noqa: BLE001
    print("World Bank failed:", type(e).__name__, str(e)[:200], flush=True)
print("DONE", sorted(os.listdir(OUT)), flush=True)
