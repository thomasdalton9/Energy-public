"""
India daily renewable generation discovery, round 4: the ICED daily generation endpoint.

Round 3: ICED's bundle calls dailyGeneration(r) -> GET `${B}/energy/electricity/generation/daily?${r}` (r = "source=all"
and date-range filters) and apiV1.dailyPeakDemand("energyMet"). This round: show how B / the http options are
defined, every dailyGeneration / dailyPeakDemand call site, then call the endpoint on both bases.
"""
import json
import re

import requests
import urllib3

urllib3.disable_warnings()
H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 "
                   "Safari/537.36", "Accept": "application/json, text/plain, */*",
     "Origin": "https://iced.niti.gov.in", "Referer": "https://iced.niti.gov.in/"}
T = (15, 120)


def out(*a):
    print(*a, flush=True)


def ctx(src, kw, n=4, w=400):
    for m in list(re.finditer(re.escape(kw), src))[:n]:
        out(f"  ctx {kw}: {src[max(0, m.start() - w):m.end() + w]!r}")


def call(u):
    try:
        r = requests.get(u, headers=H, timeout=T, verify=False)
        out(f"\n  GET {u} -> {r.status_code} {len(r.content)} {r.headers.get('content-type', '')[:30]}")
        if r.ok:
            t = r.text
            out("    " + t[:1500])
            try:
                j = r.json()
                d = j.get("data", j) if isinstance(j, dict) else j
                if isinstance(d, list):
                    out(f"    list of {len(d)}; first {json.dumps(d[:2])[:800]}; last {json.dumps(d[-2:])[:800]}")
                elif isinstance(d, dict):
                    out(f"    keys {list(d)[:30]}")
                    for k, v in list(d.items())[:12]:
                        out(f"      {k}: {json.dumps(v)[:500]}")
            except Exception:  # noqa: BLE001
                pass
        return r
    except Exception as e:  # noqa: BLE001
        out(f"  GET {u}: {type(e).__name__}: {e}")


def main():
    r = requests.get("https://iced.niti.gov.in/", headers=H, timeout=T, verify=False)
    js = re.findall(r'src=["\']([^"\']*main[^"\']*\.js)["\']', r.text)[0]
    src = requests.get("https://iced.niti.gov.in/" + js.lstrip("/"), headers=H, timeout=T, verify=False).text
    i = src.find("dailyGeneration(r){")
    out("service head: " + repr(src[max(0, i - 6000):i - 3000]))
    out("service near: " + repr(src[max(0, i - 3000):i + 300]))
    for kw in ("dailyGeneration(", "dailyPeakDemand(", "genDateRange", "daily-gen", "generation/daily",
               "peak-demand/daily", "dailyPeakDemand(r)", "KEY"):
        ctx(src, kw, n=5, w=350)
    for base in ("https://icedapi.niti.gov.in", "https://icedapi.niti.gov.in/v1"):
        for q in ("source=all", "source=all&rangeType=Daily", "source=all&startDate=2025-11-01&endDate=2026-09-30",
                  "source=wind", "source=solar"):
            call(f"{base}/energy/electricity/generation/daily?{q}")
    for p in ("/analytics/peak-demand/daily?type=energyMet", "/energy/electricity/demand/daily?type=energyMet",
              "/websiteLastUpdated"):
        for base in ("https://icedapi.niti.gov.in", "https://icedapi.niti.gov.in/v1"):
            call(base + p)


if __name__ == "__main__":
    main()
