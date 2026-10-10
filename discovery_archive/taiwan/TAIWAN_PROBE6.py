"""Probe 6 (compact): ESIST monthly table rows; WRA reservoir names / dataset catalogue."""
import json
import re
import requests
UA = {"User-Agent": "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 Chrome/124 Safari/537.36"}
S = requests.Session()
S.headers.update(UA)
OUT = []


def p(*a):
    OUT.append(" ".join(str(x) for x in a).replace("\n", " ")[:330])


def chunks(tag, s, n=300, maxn=4):
    for i in range(0, min(len(s), n * maxn), n):
        p(tag, s[i:i + n])


def get(u, **k):
    try:
        return S.get(u, timeout=90, **k)
    except Exception as e:  # noqa: BLE001
        p("ERR", u, type(e).__name__, str(e)[:80])


E = "https://ea01.moeaea.gov.tw/a0303/02/api"
r = get(E + "/pages/database/api")
paths = json.loads(r.text)["paths"]
p("PATHS-rest", " | ".join(f"{k.replace('/monthly/', '')}={v['get']['summary'][5:].strip()[:12]}" for k, v in list(paths.items())[22:]))
for t, sec in (("3/2", "全國"), ("6/1", None), ("3/3", "全國")):
    r = get(f"{E}/v1/zone/monthly/{t}")
    j = r.json()
    p("TABLE", t, "sections", list(j.keys()))
    k = sec or list(j.keys())[0]
    rows = j[k]
    p("  section", k, "rows", len(rows))
    for idx in (0, 1, 2, 3, len(rows) // 2, len(rows) - 2, len(rows) - 1):
        row = rows[idx]
        chunks(f"  r{idx}", " ; ".join(f"{a.replace('Column', 'c')[:14]}={str(b).replace(chr(10), '/')[:22]}" for a, b in row.items()), 300, 2)

r = get("https://opendata.wra.gov.tw/datasets?topic_name=%E6%B0%B4%E5%BA%AB%E8%88%87%E5%A0%B0%E5%A3%A9&page=1")
if r is not None:
    h = r.text
    p("WRA-page", len(h), "uuids", len(set(re.findall(r"[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}", h))),
      "nuxt", "__NUXT" in h, "水庫每日" in h)
    for m in re.finditer(r"水庫每日營運[^\"<]{0,40}", h):
        p("  hit", m.group(0)[:80]); break
    for js in re.findall(r'(?:src|href)="(/_nuxt/[^"]+\.js)"', h)[:3]:
        p("  js", js)
    i = h.find("window.__NUXT__")
    p("  nuxt-snippet", h[i:i + 300] if i >= 0 else "none")
hits = []
for i in range(45380, 45620):
    r = get(f"https://data.gov.tw/api/v2/rest/dataset/{i}")
    if r is not None:
        try:
            j = r.json()["result"]
            if re.search("水庫|蓄水", j["title"]):
                hits.append((i, j["title"], [d.get("resourceDownloadUrl", "")[:100] for d in j["distribution"]][:1]))
        except Exception:  # noqa: BLE001
            pass
for h in hits:
    p("DG", *h)
print("\n".join(OUT))
