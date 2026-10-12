"""Probe 4: which resources do the nuclear outage documents (A80) name in each country? (zone-level query, all fuels, B14 kept locally)"""
import io, os, sys, zipfile, re, collections, time, threading
import xml.etree.ElementTree as ET
from concurrent.futures import ThreadPoolExecutor
import requests
HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(os.path.dirname(os.path.dirname(HERE)))
sys.path.insert(0, os.path.join(ROOT, "europe")); sys.path.insert(0, ROOT)
import entsoe_common as C

_lock = threading.Lock(); _last = [0.0]
def get(params):
    with _lock:
        w = _last[0] + 0.35 - time.time()
        if w > 0: time.sleep(w)
        _last[0] = time.time()
    for k in range(4):
        try:
            return requests.get(C.URL, params=dict(params, securityToken=C.api_key()), headers=C.UA, timeout=(10, 240))
        except requests.RequestException:
            time.sleep(5)
    return None

def strip(t): return t.split("}")[-1]
def parse_doc(x):
    root = ET.fromstring(x)
    d = {strip(c.tag): c for c in root}
    ts = [c for c in root if strip(c.tag) == "TimeSeries"]
    out = []
    for t in ts:
        o = {}
        for e in t.iter():
            n = strip(e.tag)
            if e.text and e.text.strip() and n not in ("Point", "quantity", "position"):
                o.setdefault(n, e.text.strip())
        out.append(o)
    return root.find("{*}mRID").text, root.find("{*}revisionNumber").text, out

def pages(params):
    docs, off = [], 0
    while True:
        r = get(dict(params, offset=off))
        if r is None or r.content[:2] != b"PK":
            break
        z = zipfile.ZipFile(io.BytesIO(r.content))
        docs += [z.read(n).decode("utf8", "replace") for n in z.namelist()]
        if len(z.namelist()) < 200: break
        off += 200
        if off >= 4800: break
    return docs

zones = {"FR": "10YFR-RTE------C", "BE": "10YBE----------2", "ES": "10YES-REE------0", "SE1": "10Y1001A1001A44P", "SE2": "10Y1001A1001A45N", "SE3": "10Y1001A1001A46L", "SE4": "10Y1001A1001A47J",
         "FI": "10YFI-1--------U", "CH": "10YCH-SWISSGRIDZ", "CZ": "10YCZ-CEPS-----N", "SK": "10YSK-SEPS-----K", "HU": "10YHU-MAVIR----U", "SI": "10YSI-ELES-----O",
         "RO": "10YRO-TEL------P", "BG": "10YCA-BULGARIA-R", "NL": "10YNL----------L"}
wins = [("202412312300", "202503312200"), ("202503312200", "202506302200"), ("202506302200", "202509302200"), ("202509302200", "202512312300")]
jobs = [(cc, z, w) for cc, z in zones.items() for w in wins]
def run(j):
    cc, z, w = j
    out = []
    for bt in ("A53", "A54"):
        out += pages({"documentType": "A80", "BiddingZone_Domain": z, "businessType": bt, "periodStart": w[0], "periodEnd": w[1]})
    return cc, out
res = collections.defaultdict(list)
with ThreadPoolExecutor(8) as ex:
    for cc, out in ex.map(run, jobs):
        res[cc] += out
for cc, docs in res.items():
    nuc = collections.Counter(); ver = collections.Counter(); status = collections.Counter(); n = 0
    for x in docs:
        if "psrType>B14<" not in x: continue
        n += 1
        mrid, rev, tss = parse_doc(x)
        status[re.findall(r"<docStatus>\s*<value>(\w+)<", x)[0] if "<docStatus>" in x else "none"] += 1
        for o in tss:
            nuc[(o.get("production_RegisteredResource.mRID"), o.get("production_RegisteredResource.name"),
                 o.get("production_RegisteredResource.pSRType.powerSystemResources.mRID"), o.get("production_RegisteredResource.pSRType.powerSystemResources.name"),
                 o.get("production_RegisteredResource.pSRType.powerSystemResources.nominalP"))] += 1
    print(f"## {cc}: {len(docs)} docs total (all fuels), {n} nuclear docs, status {dict(status)}, distinct nuclear resources {len(nuc)}")
    for k, v in sorted(nuc.items(), key=lambda kv: str(kv[0])):
        print("   ", v, k)
fr = [x for x in res["FR"] if "psrType>B14<" in x]
print(fr[0] if fr else "no FR nuclear")
print("multi-TimeSeries docs:", sum(1 for x in sum(res.values(), []) if x.count("<TimeSeries>") > 1))
print("multi-Available_Period docs:", sum(1 for x in sum(res.values(), []) if x.count("<Available_Period>") > 1))
ex2 = [x for x in sum(res.values(), []) if x.count("<Available_Period>") > 1 and "psrType>B14<" in x]
print(ex2[0][:6000] if ex2 else "no multi-period nuclear doc")
