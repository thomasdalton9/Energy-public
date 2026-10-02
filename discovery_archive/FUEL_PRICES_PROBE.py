"""One-off probe (runs in GitHub Actions; the Claude sandbox cannot reach these hosts):
monthly world fuel prices for the Brazil water-value study, printed to the job log as CSV.

Sources, all keyless:
  FRED (fred.stlouisfed.org/graph/fredgraph.csv?id=...):
    PNGASEUUSDM  Natural gas, Europe (TTF-based since 2020; Russian border before), USD/MMBtu, monthly (IMF)
    PNGASJPUSDM  LNG, Japan import price, USD/MMBtu, monthly (IMF)
    PNGASUSUSDM  Henry Hub, USD/MMBtu, monthly (IMF)
    POILBREUSDM  Brent, USD/bbl, monthly (IMF)
    DDFUELNYH    NY Harbor ULSD diesel spot, USD/gal, daily (EIA)  -> monthly mean
    DEXBZUS      BRL per USD, daily (Fed H.10) -> monthly mean
  World Bank pink sheet (CMO-Historical-Data-Monthly.xlsx): Brent, TTF, Japan LNG, coal - cross-check
  Yahoo Finance chart API: TTF=F (ICE Dutch TTF front month) and JKM=F if listed - cross-check / JKM
"""
import io, json, sys, time, re
import pandas as pd, requests

UA = {"User-Agent": "Mozilla/5.0 (X11; Linux x86_64) energy-data-probe/1.0"}


def fred(series):
    url = f"https://fred.stlouisfed.org/graph/fredgraph.csv?id={series}"
    r = requests.get(url, headers=UA, timeout=60); r.raise_for_status()
    df = pd.read_csv(io.StringIO(r.text)); df.columns = ["date", series]
    df["date"] = pd.to_datetime(df.date); df[series] = pd.to_numeric(df[series], errors="coerce")
    return df.set_index("date")[series]


out = {}
for s in ["PNGASEUUSDM", "PNGASJPUSDM", "PNGASUSUSDM", "POILBREUSDM", "DDFUELNYH", "DEXBZUS"]:
    try:
        x = fred(s); out[s] = x.resample("MS").mean(); print(f"FRED {s}: {x.dropna().index.min().date()} .. {x.dropna().index.max().date()} n={x.notna().sum()}", flush=True)
    except Exception as e:
        print(f"FRED {s} FAILED: {e}", flush=True)

# Yahoo: TTF=F and JKM=F monthly
for sym, name in [("TTF=F", "YAHOO_TTF_EUR_MWH"), ("JKM=F", "YAHOO_JKM_USD_MMBTU"), ("NG=F", "YAHOO_HH")]:
    try:
        r = requests.get(f"https://query1.finance.yahoo.com/v8/finance/chart/{sym}?range=max&interval=1mo", headers=UA, timeout=60)
        j = r.json()["chart"]["result"][0]; ts = pd.to_datetime(j["timestamp"], unit="s"); cl = j["indicators"]["quote"][0]["close"]
        x = pd.Series(cl, index=ts).resample("MS").mean(); out[name] = x; print(f"Yahoo {sym}: {x.dropna().index.min().date()} .. {x.dropna().index.max().date()} n={x.notna().sum()}", flush=True)
    except Exception as e:
        print(f"Yahoo {sym} FAILED: {e}", flush=True)

# World Bank pink sheet
try:
    page = requests.get("https://www.worldbank.org/en/research/commodity-markets", headers=UA, timeout=60).text
    m = re.search(r'https://thedocs\.worldbank\.org/[^"\']*CMO-Historical-Data-Monthly\.xlsx', page)
    url = m.group(0) if m else "https://thedocs.worldbank.org/en/doc/18675f1d1639c7a34d463f59263ba0a2-0050012025/related/CMO-Historical-Data-Monthly.xlsx"
    r = requests.get(url, headers=UA, timeout=120); r.raise_for_status()
    wb = pd.read_excel(io.BytesIO(r.content), sheet_name="Monthly Prices", header=4)
    wb = wb.rename(columns={wb.columns[0]: "date"}); wb = wb[wb.date.astype(str).str.match(r"\d{4}M\d{2}")]
    wb["date"] = pd.to_datetime(wb.date.astype(str).str.replace("M", "-") + "-01"); wb = wb.set_index("date")
    cols = {c: c for c in wb.columns if any(k in str(c).lower() for k in ["brent", "natural gas, europe", "liquefied natural gas", "coal, australian", "natural gas, us"])}
    for c in cols:
        out["WB_" + str(c).replace(" ", "_").replace(",", "")] = pd.to_numeric(wb[c], errors="coerce")
    print("World Bank pink sheet columns used:", list(cols), flush=True)
except Exception as e:
    print(f"World Bank FAILED: {e}", flush=True)

M = pd.DataFrame(out); M = M[M.index >= "2004-01-01"]
print("\n=====BEGIN_CSV=====")
print(M.round(4).to_csv())
print("=====END_CSV=====", flush=True)
