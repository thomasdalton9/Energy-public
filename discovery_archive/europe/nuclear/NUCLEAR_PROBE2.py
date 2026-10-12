"""Probe 2: ENTSO-E filters (PsrType, RegisteredResource, update window), fleet via A71, Elexon REMIT nuclear, ODRE fields, EDF press page."""
import io, os, sys, zipfile, re, json, collections
import xml.etree.ElementTree as ET
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

def summarise(tag, params):
    r = get(params)
    d = docs(r)
    if d is None:
        print(f"== {tag}: HTTP {r.status_code} {C.redact(' '.join(re.findall(r'<text>(.*?)</text>', r.text)))[:200]}")
        return []
    psr = collections.Counter(re.findall(r"pSRType.psrType>(\w+)<", " ".join(d)))
    types = collections.Counter(re.findall(r"<type>(A\d+)</type>", " ".join(d)))
    print(f"== {tag}: {len(d)} docs, psr {dict(psr)}, type {dict(types)}")
    return d

FR = "10YFR-RTE------C"
base = {"BiddingZone_Domain": FR, "periodStart": "202609010000", "periodEnd": "202610010000"}
for dt in ("A77", "A80"):
    summarise(f"FR {dt} Sep PsrType=B14", dict(base, documentType=dt, PsrType="B14"))
    summarise(f"FR {dt} Sep PsrType=B14 planned", dict(base, documentType=dt, PsrType="B14", businessType="A53"))
    summarise(f"FR {dt} Sep PsrType=B14 forced", dict(base, documentType=dt, PsrType="B14", businessType="A54"))
d = summarise("FR A80 Sep B14 + update window", dict(base, documentType="A80", PsrType="B14", periodStartUpdate="202610050000", periodEndUpdate="202610130000"))
d = summarise("FR A77 Sep B14 + update window", dict(base, documentType="A77", PsrType="B14", periodStartUpdate="202610050000", periodEndUpdate="202610130000"))
# one nuclear doc, full text
d = summarise("FR A80 Sep B14 (again)", dict(base, documentType="A80", PsrType="B14"))
nuc = [x for x in d if "B14" in x]
if nuc:
    print(nuc[0][:4500])
# docStatus handling: any with docStatus
print("docStatus values:", collections.Counter(re.findall(r"<docStatus>\s*<value>(\w+)</value>", " ".join(d))))
# wide year window
summarise("FR A80 2025 B14 whole-year", {"documentType": "A80", "BiddingZone_Domain": FR, "PsrType": "B14", "periodStart": "202412312300", "periodEnd": "202512312300"})
summarise("FR A80 2 years", {"documentType": "A80", "BiddingZone_Domain": FR, "PsrType": "B14", "periodStart": "202612312300", "periodEnd": "202812312300"})
# fleet A71 for each nuclear country
zones = {"FR": FR, "BE": "10YBE----------2", "ES": "10YES-REE------0", "SE3": "10Y1001A1001A46L", "SE": "10YSE-1--------K", "FI": "10YFI-1--------U", "CH": "10YCH-SWISSGRIDZ",
         "CZ": "10YCZ-CEPS-----N", "SK": "10YSK-SEPS-----K", "HU": "10YHU-MAVIR----U", "SI": "10YSI-ELES-----O", "RO": "10YRO-TEL------P", "BG": "10YCA-BULGARIA-R", "NL": "10YNL----------L", "GB": "10YGB----------A", "DE": "10Y1001A1001A82H"}
for cc, z in zones.items():
    r = get({"documentType": "A71", "processType": "A33", "in_Domain": z, "PsrType": "B14", "periodStart": "202512312300", "periodEnd": "202612312300"})
    if r.status_code != 200:
        print("A71", cc, r.status_code, C.redact(" ".join(re.findall(r"<text>(.*?)</text>", r.text)))[:120]); continue
    root = ET.fromstring(r.text)
    out = []
    for ts in root:
        if ts.tag.endswith("TimeSeries"):
            k = {c.tag.split("}")[-1]: c for c in ts}
            q = [p.text for p in ts.iter() if p.tag.endswith("quantity")]
            out.append((k["registeredResource.name"].text, k["registeredResource.mRID"].text, q[:1]))
    print("A71", cc, len(out), out[:80])
# Elexon REMIT
H = "https://data.elexon.co.uk/bmrs/api/v1/"
for url in [H + "remit/list/by-event?from=2026-10-01T00:00Z&to=2026-10-07T00:00Z&fuelType=Nuclear&format=json",
            H + "remit/list/by-event?from=2026-10-01T00:00Z&to=2026-10-07T00:00Z&format=json&profile=Full",
            H + "remit/search?assetId=T_HEYM11&format=json",
            H + "datasets/REMIT/stream?publishDateTimeFrom=2026-10-10T00:00Z&publishDateTimeTo=2026-10-11T00:00Z&fuelType=Nuclear&format=json"]:
    try:
        r = requests.get(url, timeout=90)
        print("ELEXON", r.status_code, url[len(H):][:110], len(r.text)); print(r.text[:1800])
    except Exception as e:
        print("ELEXON err", e)
# ODRE
for ds in ("eco2mix-national-cons-def", "eco2mix-national-tr"):
    u = f"https://opendata.reseaux-energies.fr/api/explore/v2.1/catalog/datasets/{ds}/records?limit=2&order_by=date_heure%20desc"
    r = requests.get(u, timeout=60); print("ODRE", ds, r.status_code, r.text[:900])
# EDF press page
r = requests.get("https://www.edf.fr/en/the-edf-group/dedicated-sections/journalists/all-press-releases", headers=C.UA, timeout=40)
for m in re.finditer(r'href="([^"]+)"[^>]*>\s*([^<]{10,200}(?:nuclear|output|production)[^<]{0,120})<', r.text, re.I):
    print("EDF", m.group(1), re.sub(r"\s+", " ", m.group(2))[:160])
