"""
Guyana gas: the Ministry of Natural Resources' Petroleum Management Programme site (petroleum.gov.gy, page
data-visualization / 'Data Centre') shows Stabroek production - oil, gas produced, gas reinjected, used as fuel,
flared, water - by day / month / year, one month behind. This probe finds what feeds that page (iframe / embedded
dashboard / JSON API / WordPress REST) and prints the first records of any data endpoint it finds.

Usage: python3 GUYANA_GAS_DISCOVERY.py [extra URL ...]
"""
import json
import re
import sys
from urllib.parse import urljoin

import requests

S = requests.Session()
S.headers.update({"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) "
                                "Chrome/124.0 Safari/537.36", "Accept": "*/*"})
BASE = "https://petroleum.gov.gy/"
PAGES = ["data-visualization", "data-centre", "data-center", "production-data", "gas-to-energy", ""] + sys.argv[1:]
API_HINT = re.compile(r"""["'`]((?:https?://[^"'`\s]+)?/(?:api|wp-json|data|graphql|odata|_next/data)[^"'`\s]*)["'`]""", re.I)


def get(url, **kw):
    try:
        r = S.get(url, timeout=60, **kw)
        print(f"  GET {url} -> {r.status_code} {r.headers.get('content-type', '')[:40]} {len(r.content):,} B", flush=True)
        return r
    except requests.RequestException as e:
        print(f"  GET {url} -> {type(e).__name__}: {str(e)[:120]}", flush=True)
        return None


seen_js, endpoints = set(), set()
for p in PAGES:
    url = p if p.startswith("http") else BASE + p
    r = get(url)
    if r is None or r.status_code != 200:
        continue
    t = r.text
    for m in re.findall(r"<iframe[^>]+src=[\"']([^\"']+)", t, re.I):
        print(f"    iframe: {m}", flush=True)
        endpoints.add(urljoin(url, m))
    for m in re.findall(r"(powerbi\.com[^\"'\s]*|tableau[^\"'\s]*|datawrapper[^\"'\s]*|flourish[^\"'\s]*|"
                        r"lookerstudio[^\"'\s]*|arcgis[^\"'\s]*)", t, re.I)[:10]:
        print(f"    embed: {m}", flush=True)
    for m in API_HINT.findall(t):
        endpoints.add(urljoin(url, m))
    for js in re.findall(r"<script[^>]+src=[\"']([^\"']+\.js[^\"']*)", t, re.I):
        js = urljoin(url, js)
        if js in seen_js or "jquery" in js.lower() or "wp-includes" in js:
            continue
        seen_js.add(js)
        rj = get(js)
        if rj is not None and rj.status_code == 200:
            hits = set(API_HINT.findall(rj.text))
            for h in hits:
                endpoints.add(urljoin(js, h))
            for kw in ("produc", "reinject", "flare", "gas", "MMSCF", "mmscf"):
                i = rj.text.find(kw)
                if i >= 0 and ("fetch(" in rj.text or "axios" in rj.text or "ajax" in rj.text):
                    print(f"    {js.split('/')[-1]}: '{kw}' ...{rj.text[max(0, i - 120):i + 160]!r}", flush=True)
                    break
    txt = re.sub(r"\s+", " ", re.sub(r"<script.*?</script>|<style.*?</style>|<[^>]+>", " ", t, flags=re.S))
    for kw in ("reinject", "flared", "MMSCF", "Gas Produced", "gas produced"):
        i = txt.find(kw)
        if i >= 0:
            print(f"    text '{kw}': ...{txt[max(0, i - 200):i + 300]}", flush=True)

print("\n== WordPress REST", flush=True)
r = get(BASE + "wp-json/")
if r is not None and r.status_code == 200:
    try:
        routes = list(r.json().get("routes", {}))
        print(f"  {len(routes)} routes; non-core: {[x for x in routes if not x.startswith(('/wp/', '/oembed', '/wp-site'))][:60]}",
              flush=True)
    except ValueError:
        pass

print("\n== candidate endpoints", flush=True)
for e in sorted(endpoints)[:40]:
    r = get(e)
    if r is None or r.status_code != 200:
        continue
    body = r.text.strip()
    if body[:1] in "[{":
        try:
            j = json.loads(body)
            print(f"    JSON: {json.dumps(j)[:800]}", flush=True)
        except ValueError:
            print(f"    {body[:300]!r}", flush=True)
    else:
        print(f"    {body[:200]!r}", flush=True)
