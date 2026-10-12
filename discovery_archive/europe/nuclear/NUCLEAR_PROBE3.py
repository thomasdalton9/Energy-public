"""Probe 3: ENTSO-E RegisteredResource filter + paging; Elexon REMIT by-event semantics and message detail; ODRE eco2mix daily; EDF guidance release."""
import io, os, sys, zipfile, re, json, collections
import requests
HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(os.path.dirname(os.path.dirname(HERE)))
sys.path.insert(0, os.path.join(ROOT, "europe")); sys.path.insert(0, ROOT)
import entsoe_common as C

def get(params):
    C._throttle()
    return requests.get(C.URL, params=dict(params, securityToken=C.api_key()), headers=C.UA, timeout=(10, 240))

def docs(r):
    if r.content[:2] != b"PK":
        return None
    z = zipfile.ZipFile(io.BytesIO(r.content))
    return [z.read(n).decode("utf8", "replace") for n in z.namelist()]

def tr(tag, params):
    r = get(params); d = docs(r)
    if d is None:
        print(f"== {tag}: HTTP {r.status_code} {C.redact(' '.join(re.findall(r'<text>(.*?)</text>', r.text)))[:180]}"); return []
    res = collections.Counter(re.findall(r"production_RegisteredResource.mRID[^>]*>([^<]+)<", " ".join(d)))
    print(f"== {tag}: {len(d)} docs; resources {dict(res.most_common(6))}")
    return d

FR = "10YFR-RTE------C"
CAT3 = "17W100B100P0189S"
for dt in ("A80", "A77"):
    d = tr(f"FR {dt} Cattenom3 whole 2025", {"documentType": dt, "BiddingZone_Domain": FR, "RegisteredResource": CAT3, "periodStart": "202412312300", "periodEnd": "202512312300"})
d = tr("FR A80 Cattenom3 2027", {"documentType": "A80", "BiddingZone_Domain": FR, "RegisteredResource": CAT3, "periodStart": "202612312300", "periodEnd": "202712312300"})
if d:
    print(d[0][:3200])
d = tr("FR A80 Cattenom3 no zone 2025", {"documentType": "A80", "RegisteredResource": CAT3, "periodStart": "202412312300", "periodEnd": "202512312300"})
# paging on a month with offset
for off in (0, 200, 400, 600):
    d = tr(f"FR A80 Sep planned offset {off}", {"documentType": "A80", "BiddingZone_Domain": FR, "businessType": "A53", "periodStart": "202609010000", "periodEnd": "202610010000", "offset": off})
    nuc = sum(1 for x in d if "<production_RegisteredResource.pSRType.psrType>B14<" in x or "psrType>B14<" in x)
    print("   nuclear docs in page:", nuc)
# Elexon
H = "https://data.elexon.co.uk/bmrs/api/v1/"
def ex(path):
    r = requests.get(H + path, timeout=90); return r
r = ex("remit/list/by-event?from=2026-10-01T00:00Z&to=2026-10-07T00:00Z&fuelType=Nuclear&format=json")
lst = r.json()["data"]; print("Elexon nuclear list 1-7 Oct:", len(lst), [(x["id"], x["mrid"][-14:], x["revisionNumber"]) for x in lst])
r = ex("remit/list/by-event?from=2026-12-01T00:00Z&to=2026-12-07T00:00Z&fuelType=Nuclear&format=json")
print("Elexon nuclear list 1-7 Dec 2026 (future):", len(r.json()["data"]))
r = ex("remit/list/by-event?from=2022-03-01T00:00Z&to=2022-03-07T00:00Z&fuelType=Nuclear&format=json")
print("Elexon nuclear list Mar 2022:", len(r.json()["data"]))
if lst:
    r = ex(f"remit/{lst[0]['id']}?format=json"); print("detail:", r.status_code, r.text[:2500])
    r = ex(f"remit/search?mrid={lst[0]['mrid']}&format=json"); j = r.json(); print("search by mrid:", r.status_code, len(j.get('data', j)) if isinstance(j, dict) else len(j), r.text[:600])
    r = ex(f"remit/revisions?mrid={lst[0]['mrid']}&format=json"); print("revisions:", r.status_code, r.text[:400])
# ODRE daily nuclear
for ds, extra in (("eco2mix-national-cons-def", ""), ("eco2mix-national-tr", "")):
    u = (f"https://opendata.reseaux-energies.fr/api/explore/v2.1/catalog/datasets/{ds}/records?select=date,avg(nucleaire) as nuc,count(nucleaire) as n"
         f"&group_by=date&where=date_heure>=date'2026-09-25'&order_by=date&limit=20")
    r = requests.get(u, timeout=60); print("ODRE", ds, r.status_code, r.text[:700])
# EDF guidance release
for u in ["https://www.globenewswire.com/news-release/2026/07/31/", "https://www.globenewswire.com/news-release/2025/12/18/3208040/0/en/edf-estimated-nuclear-generation-in-france.html",
          "https://www.edf.fr/en/the-edf-group/dedicated-sections/investors/financial-results"]:
    try:
        r = requests.get(u, headers=C.UA, timeout=40)
        t = re.sub(r"<[^>]+>", " ", r.text); t = re.sub(r"\s+", " ", t)
        print("EDF", r.status_code, u[:100], len(t))
        for m in re.finditer(r"[^.]{0,200}\bTWh\b[^.]{0,200}", t):
            print("   ", m.group(0)[:400]); 
            break
    except Exception as e:
        print("EDF err", u[:60], type(e).__name__)
