"""Probe 2: data.gov.tw metadata of Taipower generation datasets, label inventory of d006010 and d006001, Wayback copies of older d006010 windows."""
import collections, json, os, re, requests
OUT = "discovery_archive/results/taiwan_gen"; os.makedirs(OUT, exist_ok=True)
UA = {"User-Agent": "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 Chrome/124 Safari/537.36"}
log = open(f"{OUT}/probe2_log.txt", "w", encoding="utf-8")
def P(*a):
    s = " ".join(str(x) for x in a); print(s); log.write(s + "\n"); log.flush()
for ds in (37331, 8931):
    try:
        r = requests.get(f"https://data.gov.tw/api/v2/rest/dataset/{ds}", headers=UA, timeout=60)
        json.dump(r.json(), open(f"{OUT}/datagovtw_{ds}.json", "w"), ensure_ascii=False, indent=1); P("meta", ds, r.status_code)
    except Exception as e: P("meta ERR", ds, e)
for u in ["https://data.gov.tw/api/v2/rest/dataset?keyword=%E5%8F%B0%E7%81%A3%E9%9B%BB%E5%8A%9B%E5%85%AC%E5%8F%B8%20%E7%99%BC%E8%B3%BC%E9%9B%BB",
          "https://data.gov.tw/api/front/dataset/list?q=%E7%99%BC%E8%B3%BC%E9%9B%BB",
          "https://data.gov.tw/api/front/dataset/search?q=%E7%99%BC%E8%B3%BC%E9%9B%BB",
          "https://data.gov.tw/datasets/search?query=%E7%99%BC%E8%B3%BC%E9%9B%BB",
          "http://archive.org/wayback/available?url=service.taipower.com.tw/data/opendata/apply/file/d006010/001.json",
          "http://web.archive.org/cdx/search/cdx?url=service.taipower.com.tw/data/opendata/apply/file/d006010/*&output=json&limit=50",
          "http://web.archive.org/cdx/search/cdx?url=www.taipower.com.tw/d006/loadGraph/loadGraph/data/genary.json&output=json&limit=20"]:
    try:
        r = requests.get(u, headers=UA, timeout=60)
        P("URL", u[:110], r.status_code, r.headers.get("content-type"), len(r.content)); P("   ", r.text[:700].replace("\n", " "))
    except Exception as e: P("ERR", u[:100], type(e).__name__)
# d006001 live
try:
    r = requests.get("https://service.taipower.com.tw/data/opendata/apply/file/d006001/001.json", headers=UA, timeout=60)
    open(f"{OUT}/d006001_live.json", "wb").write(r.content); P("d006001", r.status_code, len(r.content))
except Exception as e: P("d006001 ERR", e)
# d006010 label inventory
REC = re.compile(r'\{"FUEL_TYPE":"([^"]*)","UNIT_NAME":"([^"]*)","DATETIME":"([^"]*)","NET_P":"([^"]*)"\}')
U = "https://service.taipower.com.tw/data/opendata/apply/file/d006010/001.json"
r = requests.get(U, headers=UA, timeout=(15, 600), stream=True)
sums = collections.defaultdict(lambda: [0, 0.0, 0]); stamps = collections.Counter(); buf = ""; first = last = None
for ch in r.iter_content(8 << 20):
    buf += ch.decode("utf-8", "ignore"); k = buf.rfind("}"); part, buf = buf[:k + 1], buf[k + 1:]
    for lab, unit, dt, v in REC.findall(part):
        s = sums[(lab, unit)]
        try: s[1] += float(v); s[0] += 1
        except ValueError: s[2] += 1
        stamps[dt[:7]] += 1
        first = first or dt; last = dt
P("window", first, last, dict(stamps))
bylab = collections.defaultdict(lambda: [0, 0.0])
with open(f"{OUT}/d006010_units.csv", "w", encoding="utf-8") as f:
    f.write("fuel_type,unit,n,mean_MW,n_blank\n")
    for (lab, unit), (n, tot, bl) in sorted(sums.items()):
        m = tot / n if n else float("nan"); f.write(f"{lab},{unit},{n},{m:.1f},{bl}\n")
        bylab[lab][0] += 1; bylab[lab][1] += m if n else 0
for lab, (nu, m) in bylab.items(): P("LABEL", lab, "units", nu, "sum of unit means MW", round(m, 0))
