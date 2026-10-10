"""Probe 3: Internet Archive (Wayback) captures of Taipower's rolling d006010 (and d006009 / d006001) files - do they extend history?"""
import json, os, re, time, requests
OUT = "discovery_archive/results/taiwan_gen"; os.makedirs(OUT, exist_ok=True)
UA = {"User-Agent": "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 Chrome/124 Safari/537.36"}
log = open(f"{OUT}/probe3_log.txt", "w", encoding="utf-8")
def P(*a):
    s = " ".join(str(x) for x in a); print(s); log.write(s + "\n"); log.flush()
def get(u, **kw):
    for i in range(4):
        try:
            r = requests.get(u, headers=UA, timeout=(20, 120), **kw)
            if r.status_code in (503, 429): time.sleep(8 * (i + 1)); continue
            return r
        except Exception as e:
            P("  retry", type(e).__name__); time.sleep(5)
caps = {}
for name in ("d006010/001.json", "d006010/001.csv", "d006009/001.json", "d006001/001.json"):
    u = "http://web.archive.org/cdx/search/cdx?url=service.taipower.com.tw/data/opendata/apply/file/" + name + "&output=json&filter=statuscode:200"
    r = get(u)
    rows = json.loads(r.text) if r is not None and r.status_code == 200 and r.text.strip() else []
    P("CDX", name, r.status_code if r is not None else None, "captures", max(len(rows) - 1, 0))
    caps[name] = rows[1:]
    for row in rows[1:60]: P("   ", row[1], row[3], "len", row[6])
    if len(rows) > 61: P("    ... last", rows[-1][1], "len", rows[-1][6])
# open the head of each d006010 capture: the window it holds
seen = set()
for row in caps["d006010/001.json"]:
    ts, dig = row[1], row[5]
    if dig in seen: continue
    seen.add(dig)
    u = f"http://web.archive.org/web/{ts}id_/https://service.taipower.com.tw/data/opendata/apply/file/d006010/001.json"
    try:
        r = requests.get(u, headers=UA, timeout=(20, 120), stream=True)
        head = next(r.iter_content(700), b"").decode("utf-8-sig", "replace"); 
        s, e = re.search(r'"START_DATE":"([^"]+)"', head), re.search(r'"END_DATE":"([^"]+)"', head)
        P("CAPTURE", ts, r.status_code, r.headers.get("content-length"), s and s.group(1), e and e.group(1)); r.close()
    except Exception as ex: P("CAPTURE ERR", ts, type(ex).__name__)
