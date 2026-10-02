"""
One-off discovery, round 4 (prints only): NZ hydro storage (Transpower market operations hydro information page,
its linked files), and an STTM price history file on aemo.com.au.
"""
import io
import re

import pandas as pd
import requests

H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0"}


def get(url):
    try:
        r = requests.get(url, headers=H, timeout=90)
        print(f"{r.status_code} {len(r.content):>10,} B  {r.headers.get('content-type', '')[:35]:35}  {r.url[:200]}", flush=True)
        return r
    except Exception as e:  # noqa: BLE001
        print(f"FAIL {type(e).__name__}: {str(e)[:150]}  {url}", flush=True)


def links(r, pat, n=40):
    if r is None or not r.ok:
        return []
    found = sorted(set(re.findall(r'href="([^"]+)"', r.text, re.I)))
    hit = [u for u in found if re.search(pat, u, re.I)]
    for u in hit[:n]:
        print("   link:", u)
    return hit


def peek(r, label):
    if r is None or not r.ok:
        return
    c = r.content
    try:
        if c[:2] == b"PK" or c[:4] == b"\xd0\xcf\x11\xe0":
            xl = pd.ExcelFile(io.BytesIO(c))
            print("   sheets:", xl.sheet_names)
            for sh in xl.sheet_names[:3]:
                print(pd.read_excel(xl, sh, header=None).iloc[:12, :8].to_string()[:2000])
        else:
            d = pd.read_csv(io.BytesIO(c), nrows=50)
            print("   columns:", list(d.columns)[:20])
            print(d.head(5).to_string()[:1500])
    except Exception as e:  # noqa: BLE001
        print("   peek failed", type(e).__name__, str(e)[:100], c[:200])


print("=== Transpower hydro information")
for u in ("https://www.transpower.co.nz/system-operator/notices-and-reporting/market-operations-weekly-report/hydro-information",
          "https://www.transpower.co.nz/system-operator/security-supply/hydro-storage",
          "https://www.transpower.co.nz/system-operator/security-supply/electricity-risk-curves"):
    r = get(u)
    hit = links(r, r"\.(csv|xlsx?)|storage|hydro|download|media|static", 40)
    for h in [x for x in hit if re.search(r"\.(csv|xlsx?)(\?|$)", x, re.I)][:4]:
        h = h if h.startswith("http") else "https://www.transpower.co.nz" + h
        peek(get(h), h)

print("=== EMI hydro reports")
for u in ("https://www.emi.ea.govt.nz/Wholesale/Reports/W_HS_C", "https://www.emi.ea.govt.nz/Wholesale/Reports/W_HSC_C",
          "https://www.emi.ea.govt.nz/Wholesale/Reports?category=Hydrology",
          "https://www.emi.ea.govt.nz/Wholesale/Reports/W_HL_C"):
    links(get(u), r"hydro|storage|W_H", 20)

print("=== AEMO STTM data pages")
for u in ("https://aemo.com.au/energy-systems/gas/short-term-trading-market-sttm/data-sttm",
          "https://www.aemo.com.au/energy-systems/gas/short-term-trading-market-sttm/data-sttm/price-and-withdrawals",
          "https://aemo.com.au/energy-systems/gas/declared-wholesale-gas-market-dwgm/data-dwgm"):
    hit = links(get(u), r"\.(csv|xlsx?|zip)(\?|$)", 30)
    for h in [x for x in hit if re.search(r"price", x, re.I)][:2]:
        h = h if h.startswith("http") else "https://aemo.com.au" + h
        peek(get(h), h)
