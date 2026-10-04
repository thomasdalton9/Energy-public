"""
Capacity/price round 3: run the new pulls to a temp folder (asia/SINGAPORE_USEP.py, asia/INDIA_CEA_CAPACITY.py,
asia/SRI_LANKA_CAPACITY.py) and print summaries; probe BPDB's power-generation-unit-search ajax with the page's
own option values (installed_capacity / derated_capacity / all) and DOE's electric-power-industry page.
"""
import os
import re
import subprocess
import sys
import tempfile

import pandas as pd
import requests
import urllib3

urllib3.disable_warnings()
H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0 Safari/537.36",
     "Accept": "*/*"}
pd.set_option("display.width", 250)
TMP = tempfile.mkdtemp()


def out(*a):
    print(*a, flush=True)


def run(script, name, sheets):
    p = os.path.join(TMP, name)
    out(f"\n######## {script}")
    subprocess.run([sys.executable, script, "--out", p], timeout=1200)
    if os.path.exists(p):
        for s in sheets:
            try:
                d = pd.read_excel(p, sheet_name=s, index_col=0)
                out(f"--- {s} {d.shape}")
                out(d.tail(6).to_string(max_cols=20))
                if len(d) > 12:
                    out(d.iloc[:: max(1, len(d) // 8)].to_string(max_cols=20))
            except Exception as e:
                out(f"  {s}: {e}")


def bangladesh():
    out("\n######## BPDB ajax")
    for ct in ("installed_capacity", "derated_capacity", "all"):
        try:
            r = requests.get("https://misc.bpdb.gov.bd/power-generation-unit-search", params={"capacity_type": ct},
                             headers=dict(H, **{"X-Requested-With": "XMLHttpRequest",
                                                "Referer": "https://misc.bpdb.gov.bd/power-generation-unit"}),
                             timeout=(20, 90), verify=False)
            out(f"{ct}: {r.status_code} {r.headers.get('content-type')} {len(r.content)}")
            out(r.text[:4000])
        except Exception as e:
            out(f"{ct}: {e}")
    r = requests.get("https://misc.bpdb.gov.bd/power-generation-unit", headers=H, timeout=(20, 90), verify=False)
    i = r.text.find("success:function(data)")
    out(r.text[i:i + 6000])


def philippines():
    out("\n######## DOE")
    for u in ("https://doe.gov.ph/data-and-prices/energy-statistics/electric-power-industry",):
        r = requests.get(u, headers=H, timeout=(20, 90))
        out(f"{u}: {r.status_code} {len(r.content)}")
        for l in sorted(set(re.findall(r'(?:href|src)=["\']([^"\']+)["\']', r.text))):
            if re.search(r"\.(xlsx?|pdf|csv)|statistic|plant|capacity|api|strapi|cms|upload", l, re.I):
                out("   ", l)
        for l in sorted(set(re.findall(r'https?://[^"\'\s\\]+\.(?:xlsx?|pdf|csv)', r.text)))[:80]:
            out("   file:", l)
        t = re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", r.text))
        i = t.find("Home / Energy Statistics")
        out(t[i:i + 4000])


if __name__ == "__main__":
    run("asia/SRI_LANKA_CAPACITY.py", "sri_lanka_power_capacity.xlsx", ["Monthly"])
    run("asia/INDIA_CEA_CAPACITY.py", "india_power_capacity.xlsx", ["Monthly", "Files"])
    run("asia/SINGAPORE_USEP.py", "singapore_power_prices.xlsx", ["Daily", "Monthly"])
    for f in (bangladesh, philippines):
        try:
            f()
        except Exception as e:
            out(f"!! {f.__name__}: {e!r}")
