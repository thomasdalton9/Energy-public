"""Probe 4 (compact output): ESIST API spec, WRA open API reservoir data and dataset list."""
import re
import requests
UA = {"User-Agent": "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 Chrome/124 Safari/537.36"}
S = requests.Session()
S.headers.update(UA)
OUT = []


def p(*a):
    s = " ".join(str(x) for x in a).replace("\n", " ")
    OUT.append(s[:420])


def get(u, **k):
    try:
        return S.get(u, timeout=60, **k)
    except Exception as e:  # noqa: BLE001
        p("ERR", u, type(e).__name__, str(e)[:80])


E = "https://ea01.moeaea.gov.tw/a0303/02"
r = get(E + "/api/pages/database/api")
if r is not None:
    p("SPEC", r.status_code, len(r.content), r.headers.get("content-type"))
    for i in range(0, min(len(r.text), 4200), 420):
        p("  spec>", r.text[i:i + 420])
r = get(E + "/_astro/ApiSpecView.BCR7bTsp.js")
if r is not None:
    p("JS", r.status_code, len(r.content), sorted(set(re.findall(r'["\'`](/[A-Za-z0-9_/\-{}.]*(?:api|Api)[A-Za-z0-9_/\-{}.?=&]*)', r.text)))[:20])
for u in (E + "/api/", E + "/api/pages/database/search/electric-generation", E + "/api/pages/newest/monthly"):
    r = get(u)
    if r is not None:
        p("TRY", r.status_code, len(r.content), u, r.text[:300])

W = "https://opendata.wra.gov.tw"
for q in ("?format=JSON&limit=2", "?sort=_importdate%20asc&format=JSON&limit=2", "?format=JSON&limit=2&offset=0", "?format=CSV&limit=3"):
    r = get(f"{W}/api/v2/2be9044c-6e44-4856-aad5-dd108c2e6679{q}")
    if r is not None:
        p("WRA", r.status_code, len(r.content), q, r.headers.get("content-type"), r.text[:500])
r = get(f"{W}/datasets?topic_name=%E6%B0%B4%E5%BA%AB%E8%88%87%E5%A0%B0%E5%A3%A9&page=1")
if r is not None:
    for h, t in re.findall(r'<a[^>]+href="(/datasets/[^"]+)"[^>]*>(.*?)</a>', r.text, re.S)[:40]:
        p("  DS", h, re.sub(r"<[^>]+>|\s+", " ", t).strip()[:60])
    p("  ds-links-total", len(re.findall(r'href="/datasets/', r.text)))
r = get(W + "/openapi/swagger/index.html")
if r is not None:
    p("SW", re.findall(r'url\s*[:=]\s*["\']([^"\']+)', r.text)[:5], re.findall(r'src="([^"]+\.js)"', r.text)[:5])
print("\n".join(OUT))
