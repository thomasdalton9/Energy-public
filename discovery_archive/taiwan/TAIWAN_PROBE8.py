"""Probe 8: WRA open-data site API (catalogue) to find the reservoir name/capacity dataset; ESIST nothing more."""
import json
import os
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


h = get("https://opendata.wra.gov.tw/datasets?topic_name=%E6%B0%B4%E5%BA%AB%E8%88%87%E5%A0%B0%E5%A3%A9&page=1").text
i = h.find("window.__NUXT__.config")
LOG.append("CONFIG " + h[i:i + 1500])
bundle = ""
for js in re.findall(r'(?:src|href)="(/_nuxt/[^"]+\.js)"', h):
    t = get("https://opendata.wra.gov.tw" + js)
    if t is not None:
        bundle += "\n" + t.text
open(D + "/wra_bundle.js", "w", encoding="utf-8").write(bundle)
for pat in (r'.{80}api/v\d.{120}', r'.{60}topic_name.{120}', r'.{60}\$fetch.{120}', r'.{60}apiUrl.{120}'):
    for m in list(re.finditer(pat, bundle))[:4]:
        LOG.append("B> " + m.group(0).replace("\n", " "))
open(D + "/probe8_log.txt", "w", encoding="utf-8").write("\n".join(LOG))
print("\n".join(l[:300] for l in LOG[:5]))
