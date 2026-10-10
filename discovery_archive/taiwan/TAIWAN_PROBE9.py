"""Probe 9: ESIST database-query API (long monthly history) - find the endpoints in the page scripts."""
import re
import requests
UA = {"User-Agent": "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 Chrome/124 Safari/537.36"}
S = requests.Session()
S.headers.update(UA)
D = "discovery_archive/results/taiwan"
LOG = []
B = "https://ea01.moeaea.gov.tw"


def get(u, **k):
    try:
        return S.get(u, timeout=90, **k)
    except Exception as e:  # noqa: BLE001
        LOG.append(f"ERR {u} {type(e).__name__} {str(e)[:100]}")


for page in ("/a0303/02/database/search/electric-generation/", "/a0303/02/database/search/", "/a0303/02/newest/monthly/"):
    h = get(B + page)
    if h is None:
        continue
    scripts = sorted(set(re.findall(r'(?:src|href)="(/a0303/02/_astro/[^"]+\.js)"', h.text)))
    LOG.append(f"PAGE {page} {len(h.text)} scripts {len(scripts)}")
    for s in scripts:
        t = get(B + s)
        if t is None:
            continue
        for m in set(re.findall(r'["\'`]((?:/a0303/02)?/api/[^"\'`]{2,120})', t.text)):
            LOG.append(f"  {s[-24:]} API {m[:150]}")
    for m in set(re.findall(r'["\'`]((?:/a0303/02)?/api/[^"\'`]{2,120})', h.text)):
        LOG.append(f"  html API {m[:150]}")
    for m in re.findall(r'href="([^"]+\.(?:xlsx?|ods|zip|csv)[^"]*)"', h.text)[:10]:
        LOG.append(f"  file {m}")
for u in ("/a0303/02/api/pages/database/search/electric-generation", "/a0303/02/api/pages/database/search",
          "/a0303/02/api/pages/database", "/a0303/02/api/v1/database/search", "/a0303/02/api/files/月報.zip"):
    r = get(B + u)
    if r is not None:
        LOG.append(f"TRY {r.status_code} {len(r.content)} {u} {r.text[:200] if len(r.content) < 5000 else ''}")
open(D + "/probe9_log.txt", "w", encoding="utf-8").write("\n".join(LOG))
print("\n".join(l[:200] for l in LOG[:8]))
