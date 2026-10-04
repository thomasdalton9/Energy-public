"""
Capacity round 4: run asia/BANGLADESH_BPDB_CAPACITY.py to a temp file; probe DOE Philippines power statistics
pages (/data-and-prices/energy-statistics/electric-power-industry/2025-power-statistics) for capacity files.
"""
import os
import re
import subprocess
import sys
import tempfile

import pandas as pd
import requests

H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0 Safari/537.36",
     "Accept": "*/*"}
pd.set_option("display.width", 250)


def out(*a):
    print(*a, flush=True)


def doe():
    base = "https://doe.gov.ph/data-and-prices/energy-statistics/electric-power-industry"
    seen = set()
    for u in (base + "/2025-power-statistics", base + "/2024-power-statistics", base + "/list-of-existing-power-plants"):
        try:
            r = requests.get(u, headers=H, timeout=(20, 90))
        except Exception as e:
            out(f"{u}: {e}")
            continue
        out(f"\n{u}: {r.status_code} {len(r.content)}")
        for l in sorted(set(re.findall(r'(?:href|src)=["\']([^"\']+)["\']', r.text))):
            if not l.startswith("/_next") and re.search(r"\.(xlsx?|pdf|csv)|statistic|plant|capacity|upload|media|files", l, re.I):
                out("   ", l)
        for l in sorted(set(re.findall(r'(?:https?:)?//[^"\'\s\\<>]+?\.(?:xlsx?|pdf|csv)', r.text)))[:80]:
            out("   file:", l)
            seen.add(l)
        t = re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", r.text))
        i = t.find("Home / Energy Statistics")
        out(t[i:i + 3000])
    for l in [x for x in seen if re.search(r"xlsx?$", x)][:3]:
        r = requests.get(l if l.startswith("http") else "https:" + l, headers=H, timeout=(20, 120))
        out(f"GET {l}: {r.status_code} {len(r.content)}")
        try:
            xl = pd.ExcelFile(__import__("io").BytesIO(r.content))
            out(xl.sheet_names)
            for s in xl.sheet_names[:4]:
                out(f"--- {s}"); out(xl.parse(s, header=None).head(40).to_string(max_cols=14, max_colwidth=20))
        except Exception as e:
            out(f"  {e}")


if __name__ == "__main__":
    p = os.path.join(tempfile.mkdtemp(), "bangladesh_power_capacity.xlsx")
    subprocess.run([sys.executable, "asia/BANGLADESH_BPDB_CAPACITY.py", "--out", p], timeout=600)
    if os.path.exists(p):
        for s in ("Monthly", "Derated", "Latest"):
            out(pd.read_excel(p, sheet_name=s).to_string())
    try:
        doe()
    except Exception as e:
        out(f"!! doe: {e!r}")
