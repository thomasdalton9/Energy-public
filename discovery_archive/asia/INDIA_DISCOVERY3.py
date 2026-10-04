"""
India daily renewable generation discovery, round 3: NITI Aayog ICED (India Climate & Energy Dashboard) API.

Round 2: CEA daily RE xlsx 2020-12-17..2025-11-18 (state rows; Wind / Solar / Others RES, MU); none after;
CEA monthly RE PDF page 3 = all-India by source; ICED Angular bundle main.*.js uses https://icedapi.niti.gov.in/v1.
This round: pull every API path out of ICED's bundle, show those about generation / daily / power, call them.
"""
import re

import requests
import urllib3

urllib3.disable_warnings()
H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 "
                   "Safari/537.36", "Accept": "application/json, text/plain, */*",
     "Origin": "https://iced.niti.gov.in", "Referer": "https://iced.niti.gov.in/"}
T = (15, 90)


def out(*a):
    print(*a, flush=True)


def main():
    r = requests.get("https://iced.niti.gov.in/", headers=H, timeout=T, verify=False)
    js = [j for j in re.findall(r'src=["\']([^"\']*main[^"\']*\.js)["\']', r.text)]
    out(f"main bundle: {js}")
    src = requests.get("https://iced.niti.gov.in/" + js[0].lstrip("/"), headers=H, timeout=T, verify=False).text
    out(f"bundle {len(src)} chars")
    # base-url variables and every quoted path-like string
    for m in re.finditer(r"icedapi\.niti\.gov\.in[^\"'`]*", src):
        out(f"  base: {src[max(0, m.start() - 120):m.end() + 60]!r}")
    paths = sorted(set(re.findall(r'["\'`](/?[a-zA-Z][\w\-]*(?:/[\w\-${}.]+){1,6})["\'`]', src)))
    keep = [p for p in paths if re.search(r"gener|daily|power|electric|psp|demand|renew|wind|solar|energy|mix|source",
                                          p, re.I) and not re.search(r"\.(png|svg|jpg|css)$|^assets/", p)]
    out(f"{len(paths)} path strings; {len(keep)} relevant:")
    for p in keep[:400]:
        out(f"  {p}")
    # context of 'daily' generation code
    for kw in ("dailyGeneration", "daily-generation", "dailyGen", "DailyGeneration", "generation/daily", "psp", "allIndiaGeneration"):
        for m in list(re.finditer(re.escape(kw), src))[:3]:
            out(f"  ctx {kw}: {src[max(0, m.start() - 250):m.end() + 250]!r}")
    # try likely calls
    base = "https://icedapi.niti.gov.in/v1"
    tried = set()
    for p in keep:
        if "${" in p or len(tried) > 60:
            continue
        u = base + ("" if p.startswith("/") else "/") + p
        if u in tried:
            continue
        tried.add(u)
        try:
            rr = requests.get(u, headers=H, timeout=(10, 40), verify=False)
            if rr.ok:
                out(f"  GET {u} -> {rr.status_code} {len(rr.content)}: {rr.text[:400]!r}")
            else:
                out(f"  GET {u} -> {rr.status_code}")
        except Exception as e:  # noqa: BLE001
            out(f"  GET {u}: {type(e).__name__}")


if __name__ == "__main__":
    main()
