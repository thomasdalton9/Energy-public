"""
JKT discovery round 2: Taiwan Bureau of Energy E-STAT API (no key). Fetches the monthly endpoints from a runner and
prints the JSON structure: top-level keys, every list of records (parent key, length, first and last record) so the
column meaning and period format can be read from the log.
Run by .github/workflows/jkt_discovery.yml (set the script path in that workflow) .
"""
import json
import sys

import requests

LANDING = "https://ea01.moeaea.gov.tw/a0303/02/en/database/api/"
ROOT = "https://ea01.moeaea.gov.tw/a0303/02/api"
H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/140.0 Safari/537.36",
     "Accept": "application/json, */*", "Accept-Language": "en-GB,en;q=0.9,zh-TW;q=0.8", "Referer": LANDING}
ENDPOINTS = ["/v1/zone/monthly/6/1", "/v1/zone/monthly/3/5", "/v1/zone/monthly/8/1"]


def walk(v, path=""):
    if isinstance(v, dict):
        for k, c in v.items():
            if isinstance(c, list) and c and isinstance(c[0], dict):
                print(f"  LIST at {path}/{k}: {len(c)} records")
                print("    first:", json.dumps(c[0], ensure_ascii=False)[:600])
                print("    last: ", json.dumps(c[-1], ensure_ascii=False)[:600])
            else:
                walk(c, f"{path}/{k}")
    elif isinstance(v, list):
        for i, c in enumerate(v[:5]):
            walk(c, f"{path}[{i}]")


for url in [LANDING] + [ROOT + e for e in ENDPOINTS]:
    try:
        r = requests.get(url, headers=H, timeout=(15, 120))
        print(f"[{r.status_code}] {url} {r.headers.get('content-type')} {len(r.content)} bytes", flush=True)
        if "json" in (r.headers.get("content-type") or ""):
            p = r.json()
            print("  top-level:", list(p.keys()) if isinstance(p, dict) else type(p).__name__)
            walk(p)
        else:
            print("  ", r.text[:300].replace("\n", " "))
    except Exception as e:  # noqa: BLE001
        print(f"[ERR] {url}: {type(e).__name__}: {str(e)[:200]}", flush=True)
sys.exit(0)
