"""Probe 11: WRA reservoir ID -> name: legacy dataset GUIDs on the new platform, fhy site scripts."""
import re
import requests
UA = {"User-Agent": "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 Chrome/124 Safari/537.36"}
S = requests.Session()
S.headers.update(UA)
D = "discovery_archive/results/taiwan"
LOG = []


def get(u, **k):
    try:
        return S.get(u, timeout=90, **k)
    except Exception as e:  # noqa: BLE001
        LOG.append(f"ERR {u} {type(e).__name__} {str(e)[:100]}")


for g in ("1602ca19-b224-4cc3-aa31-11b1b124530f", "50c8256d-30c5-4b8d-9b84-2e14d5c6df71"):
    r = get(f"https://opendata.wra.gov.tw/api/v2/{g}?format=JSON&limit=3")
    LOG.append(f"GUID {g} {r.status_code if r is not None else None} {r.text[:600].replace(chr(10), ' ') if r is not None else ''}")
for u in ("http://data.wra.gov.tw/Service/OpenData.aspx?format=json&id=1602CA19-B224-4CC3-AA31-11B1B124530F",
          "https://data.wra.gov.tw/", "https://fhy.wra.gov.tw/fhyv2/monitor/reservoir"):
    r = get(u)
    LOG.append(f"URL {u} {r.status_code if r is not None else None} {len(r.content) if r is not None else ''} {r.text[:200].replace(chr(10), ' ') if r is not None else ''}")
h = get("https://fhy.wra.gov.tw/ReservoirPage_2011/Statistics.aspx")
js = re.findall(r'src="(/fhyv2/js/[^"]+\.js)"', h.text) if h is not None else []
LOG.append(f"fhy scripts {js}")
for s in js:
    t = get("https://fhy.wra.gov.tw" + s)
    if t is None:
        continue
    txt = t.text
    LOG.append(f"JS {s} {len(txt)} has10501={'10501' in txt} 石門={'石門' in txt}")
    for m in list(re.finditer(r'.{60}10501.{160}', txt))[:3]:
        LOG.append("  10501> " + m.group(0).replace("\n", " "))
    for m in sorted(set(re.findall(r'["`](/?(?:Api|api|WraApi)/[A-Za-z0-9_/{}\-.?=$&]+)', txt)))[:25]:
        LOG.append("  api> " + m)
open(D + "/probe11_log.txt", "w", encoding="utf-8").write("\n".join(LOG))
print(len(LOG))
