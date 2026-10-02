"""
Round 2 (round 1: the EBY WordPress REST API needs a login; CAMMESA's missing days read fine now). Lists every post
in the sitemaps of www.eby.gov.py and www.eby.org.ar, keeps those whose URL looks like a generation / energy report,
and for the months PARAGUAY_POWER.py estimates (May-22, Mar-23..Nov-24 bar Aug/Nov-23) prints any ANDE / SINP /
Paraguay figure the post gives.

Usage: python3 PARAGUAY_YACYRETA_SINP_DISCOVERY2.py
"""
import html
import os
import re
import sys

import pandas as pd
import requests

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "..", "south_america"))
sys.argv = [sys.argv[0]]
import PARAGUAY_POWER as P  # noqa: E402

S = requests.Session()
S.headers.update(P.UA)
WANT = {"2022-05", "2023-03", "2023-04", "2023-05", "2023-06", "2023-07", "2023-09", "2023-10", "2023-12", "2024-01",
        "2024-02", "2024-03", "2024-04", "2024-05", "2024-06", "2024-07", "2024-08", "2024-09", "2024-10", "2024-11"}
KEY = re.compile(r"generaci|energ|produc|datos-oficiales|mwh|gwh|record|entreg", re.I)


def get(url, verify=True):
    try:
        r = S.get(url, timeout=60, verify=verify)
        return r if r.status_code == 200 else None
    except requests.RequestException as e:
        print(f"  {url}: {type(e).__name__}", flush=True)
        return None


def sitemap_urls(root, verify=True):
    seen, urls, todo = set(), [], [root + p for p in ("/sitemap_index.xml", "/wp-sitemap.xml", "/sitemap.xml",
                                                       "/post-sitemap.xml")]
    while todo:
        u = todo.pop(0)
        if u in seen:
            continue
        seen.add(u)
        r = get(u, verify)
        if r is None:
            continue
        locs = re.findall(r"<loc>\s*([^<\s]+)\s*</loc>", r.text)
        lastmods = re.findall(r"<lastmod>\s*([^<\s]+)\s*</lastmod>", r.text)
        print(f"  {u}: {len(locs)} entries", flush=True)
        for loc in locs:
            (todo if loc.endswith(".xml") else urls).append(loc)
    return sorted(set(urls))


bundle = None
try:
    bundle = P.chain_bundle()   # eby.org.ar omits its intermediate certificate (verification stays on)
except Exception as e:  # noqa: BLE001
    print(f"chain bundle failed: {e}", flush=True)

for root, verify in [("https://www.eby.gov.py", True), ("https://www.eby.org.ar", bundle or True)]:
    print(f"\n===== {root}", flush=True)
    urls = sitemap_urls(root, verify)
    cand = [u for u in urls if KEY.search(u)]
    print(f"  {len(urls)} URLs, {len(cand)} look like energy/generation posts", flush=True)
    for u in cand:
        r = get(u, verify)
        if r is None:
            continue
        body = html.unescape(r.text)
        title = re.search(r"<title>([^<]+)", body)
        title = title.group(1).strip() if title else u
        when = re.search(r'datetime="(\d{4}-\d{2}-\d{2})', body)
        posted = pd.Timestamp(when.group(1)) if when else pd.Timestamp("2024-01-01")
        txt = re.sub(r"\s+", " ", re.sub(r"<script.*?</script>|<style.*?</style>|<[^>]+>", " ", body, flags=re.S))
        y, m = P.eby_period(title, u, posted)
        key = f"{y}-{m:02d}"
        sadi, sinp = P.parse_eby(txt)
        mark = "<== WANTED" if key in WANT else ""
        print(f"  [{key}] {posted:%Y-%m-%d} SADI={sadi} SINP={sinp} {mark} | {title[:80]} | {u}", flush=True)
        if key in WANT and sinp is None:
            for mm in re.finditer(r"(ANDE|SINP|Paraguay|margen derecha)", txt):
                i = mm.start()
                print(f"      ...{txt[max(0, i - 150):i + 250]}...", flush=True)
                break
