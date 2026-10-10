"""Probe: what Taiwan generation-by-fuel history is reachable keyless (Zenodo 7537890, Taipower open-data file ids, data.gov.tw metadata)."""
import json, os, re, sys, requests
OUT = "discovery_archive/results/taiwan_gen"; os.makedirs(OUT, exist_ok=True)
UA = {"User-Agent": "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 Chrome/124 Safari/537.36"}
log = open(f"{OUT}/probe1_log.txt", "w", encoding="utf-8")
def P(*a):
    s = " ".join(str(x) for x in a); print(s); log.write(s + "\n"); log.flush()
def tryget(url, n=600, rng=None, **kw):
    h = dict(UA); 
    if rng: h["Range"] = rng
    try:
        r = requests.get(url, headers=h, timeout=(15, 60), stream=True, **kw)
        b = next(r.iter_content(n), b"")
        P("URL", url, r.status_code, r.headers.get("content-type"), r.headers.get("content-length"), r.headers.get("content-range"), r.headers.get("last-modified"))
        P("   ", b.decode("utf-8-sig", "replace")[:n].replace("\n", " | "))
        r.close(); return r
    except Exception as e:
        P("ERR", url, type(e).__name__, str(e)[:120])
# Zenodo
for u in ["https://zenodo.org/api/records/7537890", "https://zenodo.org/records/7537890"]:
    r = tryget(u, 6000)
try:
    j = requests.get("https://zenodo.org/api/records/7537890", headers=UA, timeout=60).json()
    json.dump(j, open(f"{OUT}/zenodo_7537890.json", "w"), ensure_ascii=False, indent=1)
    P("LICENSE", j.get("metadata", {}).get("license"), "TITLE", j.get("metadata", {}).get("title"))
    for f in j.get("files", []):
        P("FILE", f.get("key"), f.get("size"), f.get("links", {}).get("self"))
        url = f.get("links", {}).get("self")
        if url and f.get("size", 0) < 400e6:
            fn = f"{OUT}/zenodo_{f['key']}"
            r = requests.get(url, headers=UA, timeout=(15, 300), stream=True)
            with open(fn, "wb") as fh:
                for c in r.iter_content(1 << 20): fh.write(c)
            P("saved", fn, os.path.getsize(fn))
except Exception as e:
    P("ZENODO ERR", type(e).__name__, e)
# Taipower file ids
B = "https://service.taipower.com.tw/data/opendata/apply/file/"
for i in range(1, 31):
    for ext in ("json", "csv"):
        u = f"{B}d0060{i:02d}/001.{ext}"
        try:
            r = requests.get(u, headers=UA, timeout=(10, 40), stream=True)
            b = next(r.iter_content(300), b"")
            if r.status_code == 200 and b:
                P("OK", u, r.headers.get("content-length"), r.headers.get("last-modified"), b.decode("utf-8-sig", "replace")[:250].replace("\n", " | "))
            r.close()
        except Exception as e:
            pass
# data.gov.tw metadata
for ds in (37331, 8931, 19995, 37330, 33462, 25465, 29935):
    tryget(f"https://data.gov.tw/api/v2/rest/dataset/{ds}", 2500)
for u in ["https://www.taipower.com.tw/2764/2826/2828/", "https://www.taipower.com.tw/d006/loadGraph/loadGraph/data/genary.json",
          "https://www.taipower.com.tw/d006/loadGraph/loadGraph/data/genloadareaperc.csv", "https://data.gov.tw/en/datasets/37331",
          "https://data.gov.tw/dataset/37331", "https://data.gov.tw/en/datasets/search?query=taipower+generation"]:
    tryget(u, 1500)
