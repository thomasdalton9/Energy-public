"""
One-off pull, round 2, for the Asia-Pacific industrial coal-to-LNG study:
  ap_prices_wb_monthly.csv  World Bank Pink Sheet monthly (coal Australia / South Africa, LNG Japan, natural gas
                            Europe / US) - FRED's IMF series did not download in round 1
  ap_taiwan_egeda.csv       APEC EGEDA energy balance for Chinese Taipei (Taiwan is not in UNSD data): coal types and
                            natural gas by industry subsector, autoproducers, imports, 2015-2023, PJ
"""
import io, os, re
import pandas as pd, requests

OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data")
H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0"}
s = requests.Session(); s.headers.update(H)

try:
    page = s.get("https://www.worldbank.org/en/research/commodity-markets", timeout=60).text
    m = re.search(r'https://thedocs\.worldbank\.org/[^"\']*CMO-Historical-Data-Monthly\.xlsx', page)
    url = m.group(0) if m else "https://thedocs.worldbank.org/en/doc/18675f1d1639c7a34d463f59263ba0a2-0050012025/related/CMO-Historical-Data-Monthly.xlsx"
    r = s.get(url, timeout=120); r.raise_for_status()
    x = pd.read_excel(io.BytesIO(r.content), sheet_name="Monthly Prices", header=None)
    hdr = next(i for i in range(10) if any("Coal" in str(v) for v in x.iloc[i]))
    cols = [str(v).strip() for v in x.iloc[hdr]]
    d = x.iloc[hdr + 2:].copy(); d.columns = ["month"] + cols[1:]
    keep = ["month"] + [c for c in d.columns if re.search(r"Coal|Liquefied|Natural gas", c, re.I)]
    d = d[keep]; d = d[d["month"].astype(str).str.match(r"\d{4}M\d{2}")]
    d.to_csv(os.path.join(OUT, "ap_prices_wb_monthly.csv"), index=False)
    print("World Bank:", url, d.shape, keep, d.tail(2).to_string(), flush=True)
except Exception as e:  # noqa: BLE001
    print("World Bank failed:", type(e).__name__, str(e)[:200], flush=True)

EG = "https://www.egeda.ewg.apec.org/egeda/database/"
FE = {"001001002": "Coking coal", "001001001": "Other bituminous coal", "001001003": "Sub-bituminous coal",
      "001002000": "Anthracite", "001003000": "Lignite", "001000000": "Coal", "005022000": "Natural gas"}
FS = ["000005000", "000008000", "000024000", "000024161", "000024162", "000009169", "000020000", "000020003", "000020118",
      "000020017", "000020120", "000020176", "000020177", "000020178", "000020179", "000020121", "000020180", "000020181",
      "000020182", "000020122", "000022189"]
try:
    s.post(EG + "rev_newbalance_select_cond2.php", data={"OTYPE": "9"}, timeout=60)
    data = [("Y1", "2015"), ("Y2", "2023"), ("U", "002"), ("HEAD", "Y"), ("ITEM", "C"), ("OIL", "1"), ("fC[]", "016")]
    data += [("fE[]", k) for k in FE] + [("fS[]", k) for k in FS]
    r = s.post(EG + "php/rev_newbalance2/balance.php", data=data, timeout=180,
               headers={"Referer": EG + "rev_newbalance_select_cond2.php"})
    print("EGEDA:", r.status_code, len(r.content), r.headers.get("content-type"), flush=True)
    if r.ok and len(r.content) > 500:
        try:
            tables = pd.read_html(io.StringIO(r.text))
            big = max(tables, key=len)
            big.to_csv(os.path.join(OUT, "ap_taiwan_egeda.csv"), index=False)
            print("EGEDA table", big.shape); print(big.head(15).to_string()[:3000])
        except Exception as e:  # noqa: BLE001
            print("EGEDA parse failed:", e); print(re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " | ", r.text))[:3000])
    else:
        print(r.text[:500])
except Exception as e:  # noqa: BLE001
    print("EGEDA failed:", type(e).__name__, str(e)[:200])
