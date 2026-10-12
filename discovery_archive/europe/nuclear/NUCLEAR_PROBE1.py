"""Probe 1: ENTSO-E unavailability documents (A77 production units, A80 generation units) for nuclear (B14);
Elexon REMIT; ODRE eco2mix nuclear."""
import io, os, sys, zipfile, re, json, time
import requests
HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(os.path.dirname(os.path.dirname(HERE)))
sys.path.insert(0, os.path.join(ROOT, "europe")); sys.path.insert(0, ROOT)
import entsoe_common as C

def get(params):
    C._throttle()
    r = requests.get(C.URL, params=dict(params, securityToken=C.api_key()), headers=C.UA, timeout=(10, 180))
    return r

def show(tag, params):
    r = get(params)
    ct = r.headers.get("content-type", "")
    print(f"== {tag}: HTTP {r.status_code} {ct} {len(r.content)} bytes")
    c = r.content
    if c[:2] == b"PK":
        z = zipfile.ZipFile(io.BytesIO(c))
        names = z.namelist()
        print("  zip files:", len(names), names[:3])
        x = z.read(names[0]).decode("utf8", "replace")
        print(x[:3500])
        return names, z
    print(C.redact(r.text[:1500]))
    return [], None

zones = {"FR": "10YFR-RTE------C", "BE": "10YBE----------2", "GB": "10YGB----------A"}
for dt in ("A80", "A77"):
    p = {"documentType": dt, "BiddingZone_Domain": zones["FR"], "PsrType": "B14",
         "periodStart": "202610010000", "periodEnd": "202611010000"}
    show(f"FR {dt} Oct 2026 (no business type)", p)
# future window, planned
p = {"documentType": "A80", "BiddingZone_Domain": zones["FR"], "PsrType": "B14", "businessType": "A53",
     "periodStart": "202701010000", "periodEnd": "202801010000"}
names, z = show("FR A80 planned 2027", p)
# count docs by offset
for off in (0, 200):
    r = get(dict(p, offset=off))
    if r.content[:2] == b"PK":
        print("offset", off, "docs", len(zipfile.ZipFile(io.BytesIO(r.content)).namelist()))
    else:
        print("offset", off, C.redact(r.text[:300]))
# updated-since
p = {"documentType": "A80", "BiddingZone_Domain": zones["FR"], "PsrType": "B14",
     "periodStartUpdate": "202610010000", "periodEndUpdate": "202610120000"}
show("FR A80 updated 1-12 Oct", p)
for cc in ("BE", "GB"):
    show(f"{cc} A80 2026-11", {"documentType": "A80", "BiddingZone_Domain": zones[cc], "PsrType": "B14",
                              "periodStart": "202611010000", "periodEnd": "202612010000"})
# installed capacity per production unit A71
show("FR A71 2026", {"documentType": "A71", "processType": "A33", "in_Domain": zones["FR"], "PsrType": "B14",
                      "periodStart": "202512312300", "periodEnd": "202612312300"})
# Elexon REMIT
for url in ["https://data.elexon.co.uk/bmrs/api/v1/remit/list/by-event?from=2026-10-01T00:00Z&to=2026-12-31T00:00Z&fuelType=Nuclear&format=json",
            "https://data.elexon.co.uk/bmrs/api/v1/datasets/REMIT?publishDateTimeFrom=2026-10-01T00:00Z&publishDateTimeTo=2026-10-12T00:00Z&format=json"]:
    try:
        r = requests.get(url, timeout=60)
        print("ELEXON", r.status_code, url[:90]); print(r.text[:2500])
    except Exception as e:
        print("ELEXON err", e)
# ODRE eco2mix
for url in ["https://opendata.reseaux-energies.fr/api/explore/v2.1/catalog/datasets/eco2mix-national-cons-def/records?select=date,avg(nucleaire)&group_by=date&where=date%3E%272026-09-01%27&order_by=date&limit=5",
            "https://opendata.reseaux-energies.fr/api/explore/v2.1/catalog/datasets/eco2mix-national-tr/records?select=date,avg(nucleaire)&group_by=date&where=date%3E%272026-10-01%27&order_by=date&limit=5"]:
    try:
        r = requests.get(url, timeout=60)
        print("ODRE", r.status_code); print(r.text[:600])
    except Exception as e:
        print("ODRE err", e)
# EDF / RTE REMIT-ish pages
for url in ["https://www.edf.fr/en/the-edf-group/who-we-are/activities/optimisation-and-trading/remit",
            "https://www.edf.fr/en/the-edf-group/dedicated-sections/journalists/all-press-releases",
            "https://www.services-rte.com/en/view-data-published-by-rte/unavailability-generation.html",
            "https://www.edf.fr/groupe-edf/produire-une-energie-respectueuse-du-climat/lenergie-nucleaire/disponibilite-du-parc-nucleaire",
            "https://data.rte-france.com/catalog/-/api/doc/user-guide/Generation+Unavailability/1.1"]:
    try:
        r = requests.get(url, headers=C.UA, timeout=40)
        print("WEB", r.status_code, url[:100], len(r.text), re.sub(r"\s+", " ", r.text[:200]))
    except Exception as e:
        print("WEB err", url[:80], type(e).__name__)
