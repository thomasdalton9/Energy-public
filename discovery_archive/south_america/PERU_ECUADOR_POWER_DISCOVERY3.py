"""
Round 3 (Ecuador history only; Peru is settled - see PERU_COES_GENERATION.py).

Round 2: COES' ExportarGeneracion gives per-plant MWh by technology but no
fuel split, so the per-day JSON stays the source for Peru. For CENACE the
Wayback CDX call got a 503, the R1/R3 xlsx files are programmed
re-dispatch (not actuals), and the biblioteca / transparencia pages
(2 MB each) weren't searched properly.

This round: Wayback CDX with retries (and the availability API as a
fallback), counting snapshots of InformacionOperativa.htm per month and
fetching a few; data-file links on CENACE's biblioteca / transparencia /
indicadores pages.
"""

print("STARTING", flush=True)

import collections
import os
import re
import time

import requests
import urllib3

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
OUT = "ecpe_raw"
os.makedirs(OUT, exist_ok=True)
S = requests.Session()
S.headers["User-Agent"] = "Mozilla/5.0 (energy-data research; contact via github thomasdalton9/Energy)"


def get(url, name=None, tries=4, **kw):
    for i in range(tries):
        try:
            r = S.get(url, timeout=kw.get("timeout", 120), verify=kw.get("verify", True), params=kw.get("params"))
            print(f"== GET {r.url}: {r.status_code} {r.headers.get('Content-Type', '')} {len(r.content):,} bytes",
                  flush=True)
            if r.status_code in (429, 502, 503, 504) and i < tries - 1:
                time.sleep(20 * (i + 1))
                continue
            if name and r.ok:
                open(os.path.join(OUT, name), "wb").write(r.content)
            return r
        except requests.RequestException as e:
            print(f"== GET {url}: FAILED {type(e).__name__}: {str(e)[:200]}", flush=True)
            time.sleep(20 * (i + 1))
    return None


print("\n############ Wayback", flush=True)
snaps = []
for url in ["cenace.gob.ec/info-operativa/InformacionOperativa.htm", "www.cenace.gob.ec/info-operativa/*"]:
    r = get("https://web.archive.org/cdx/search/cdx", "ec_wayback_cdx_" + str(len(snaps)) + ".txt",
            params={"url": url, "from": "2019", "fl": "timestamp,original,statuscode,length"}, timeout=240)
    if r is None or not r.ok:
        continue
    lines = [l.split() for l in r.text.splitlines() if l.strip()]
    print(f"  {url}: {len(lines)} rows", flush=True)
    origs = collections.Counter(l[1] for l in lines)
    for o, n in origs.most_common(30):
        print(f"    {n:5d} {o}", flush=True)
    snaps += [l for l in lines if "InformacionOperativa.htm" in l[1] and l[2] == "200"]
per_month = collections.Counter(s[0][:6] for s in snaps)
print("  InformacionOperativa 200-snapshots per month:", dict(sorted(per_month.items())), flush=True)
days = sorted({s[0][:8] for s in snaps})
print(f"  distinct days {len(days)}: first {days[:5]} last {days[-5:]}", flush=True)
for s in [snaps[0], snaps[len(snaps) // 2]] if snaps else []:
    rr = get(f"https://web.archive.org/web/{s[0]}id_/{s[1]}", f"ec_wayback_{s[0]}.html")
    if rr is not None and rr.ok:
        t = re.sub(r"\s+", " ", re.sub(r"<script.*?</script>|<style.*?</style>|<[^>]+>", " ", rr.text, flags=re.S))
        for m in re.finditer(r"INFORMACI[ÓO]N OPERATIVA \w+", t):
            print("   ", t[m.start():m.start() + 300], flush=True)
if not snaps:
    get("https://archive.org/wayback/available", "ec_wayback_available.json",
        params={"url": "cenace.gob.ec/info-operativa/InformacionOperativa.htm", "timestamp": "20240101"})

print("\n############ CENACE library pages", flush=True)
for u in ["https://www.cenace.gob.ec/biblioteca/", "https://www.cenace.gob.ec/transparencia/",
          "https://www.cenace.gob.ec/indicadores/"]:
    r = get(u, verify=False)
    if r is None or not r.ok:
        continue
    links = set(re.findall(r'(?:href|data-href|data-url|src)=["\']([^"\']+\.(?:xlsx?|csv|pdf|zip)[^"\']*)["\']',
                           r.text, re.I))
    print(f"  {len(links)} file links", flush=True)
    for l in sorted(links):
        if re.search(r"estad|operat|operac|produc|generac|anual|mensual|diari|balance|energ|informe", l, re.I):
            print("    FILE", l, flush=True)
    for m in list(re.finditer(r"estad[íi]stic|informaci[óo]n operativa|producci[óo]n", r.text, re.I))[:10]:
        print("    CTX", re.sub(r"\s+", " ", r.text[max(0, m.start() - 200):m.start() + 200]), flush=True)

print("\nDONE", flush=True)
