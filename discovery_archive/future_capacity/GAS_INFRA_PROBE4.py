"""Probe 4: TYNDP 2024 site crawl + A0 map, PCI documents, Sodir field status layer, NSTA own datasets."""
import os, re, json, subprocess
import requests
OUT = "discovery_archive/results/future_capacity"
UA = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36"}
def get(u, hdr=None, **kw):
    try:
        r = requests.get(u, headers=dict(UA, **(hdr or {})), timeout=(10, 180), **kw)
        print(r.status_code, len(r.content), r.headers.get("content-type", "")[:30], u[:150]); return r
    except Exception as e:
        print("FAIL", u[:120], type(e).__name__); return None
# TYNDP site
r = get("https://tyndp2024.entsog.eu/")
if r is not None:
    ls = sorted(set(re.findall(r'href="(https://tyndp2024\.entsog\.eu/[^"#]+)"', r.text)))
    for l in ls:
        if not re.search(r"wp-content|wp-json|xmlrpc|privacy|site-notice|feed|oembed", l): print("  T", l)
r = get("https://tyndp2024.entsog.eu/wp-json/wp/v2/pages?per_page=100&_fields=link,title")
if r is not None and r.status_code == 200:
    for p in r.json():
        if re.search(r"annex|project|capacit|lng|storage|transmission|pipeline|infrastructure", p["link"], re.I): print("  P", p["link"])
r = get("https://www.entsog.eu/sites/default/files/2024-12/ENTSOG_TYNDP_2024_A0_NG_FULL_5005_TRA_UGS_LNG.pdf")
if r is not None and r.status_code == 200:
    open("/tmp/tyndp_map.pdf", "wb").write(r.content)
    subprocess.run(["pdftotext", "-layout", "/tmp/tyndp_map.pdf", OUT + "/tyndp2024_a0_map.txt"])
    print("  map text bytes", os.path.getsize(OUT + "/tyndp2024_a0_map.txt"))
for u in ["https://tyndp2024.entsog.eu/downloads/", "https://tyndp2024.entsog.eu/infrastructure-report-introduction/annex/"]:
    r = get(u)
    if r is not None and r.status_code == 200:
        for l in sorted(set(re.findall(r'href="([^"]+\.(?:xlsx?|csv|pdf|zip))"', r.text, re.I))): print("  X", l[:200])
for name, u in [("btg_tyndp_2026.txt", "https://bulgartransgaz.bg/public/files/useruploads/%D0%BD%D0%BE%D0%B2%D0%B8%D0%BD%D0%B8_%D1%81%D0%B0%D0%B9%D1%82/27042026/TYNDP%20BTG%202026-2035%20%D0%B3.%20-%20public%20consultation%20EN.pdf"),
                ("transgaz_dp_2026.txt", "https://www.transgaz.ro/sites/default/files/DEVELOPMENT%20PLAN%20FOR%20THE%20NATIONAL%20GAS%20TRANSMISSION%20SYSTEM%202026-2035.pdf")]:
    r = get(u)
    if r is not None and r.status_code == 200 and r.content[:4] == b"%PDF":
        open("/tmp/x.pdf", "wb").write(r.content)
        subprocess.run(["pdftotext", "-layout", "/tmp/x.pdf", OUT + "/" + name]); print("  saved", name, os.path.getsize(OUT + "/" + name))
for u in ["https://www.norskpetroleum.no/en/facts/field/", "https://www.norskpetroleum.no/en/developments-and-operations/recent-activity/", "https://www.norskpetroleum.no/en/production-and-exports/exports-of-oil-and-gas/"]:
    r = get(u)
    if r is not None and r.status_code == 200: open(OUT + "/" + re.sub(r"\W+", "_", u[8:])[:60] + ".html", "w").write(r.text)
# PCI
r = get("https://energy.ec.europa.eu/topics/infrastructure/projects-common-interest-and-projects-mutual-interest/pci-and-pmi-selection-process_en")
if r is not None:
    for l in sorted(set(re.findall(r'href="([^"]*document/download[^"]+)"', r.text))): print("  D", l[:230])
r = get("https://energy.ec.europa.eu/document/download/944b96b9-4efd-44a3-bbfe-45b752b0b55f_en?filename=Technical_document_1st%20PCI_PMI%20list_15.05.2024_FINAL.pdf")
if r is not None and r.status_code == 200:
    open("/tmp/pci.pdf", "wb").write(r.content)
    subprocess.run(["pdftotext", "-layout", "/tmp/pci.pdf", OUT + "/pci_technical_document_2024.txt"])
# Sodir
base = "https://factmaps.sodir.no/api/rest/services/Factmaps/FactMapsWGS84/MapServer"
r = get(base + "/502?f=json")
if r is not None and r.status_code == 200:
    print("  502 fields:", [f["name"] for f in r.json().get("fields", [])])
r = get(base + "/502/query?where=1%3D1&outFields=fldName,fldCurrentActivitySatus,cmpLongName,fldDiscoveryYear&returnGeometry=false&resultRecordCount=2000&f=json")
if r is not None and r.status_code == 200:
    rows = [x["attributes"] for x in r.json().get("features", [])]
    json.dump(rows, open(OUT + "/sodir_fields_status.json", "w"))
    from collections import Counter
    print("  status counts", Counter(x.get("fldCurrentActivitySatus") for x in rows))
    for x in rows:
        if x.get("fldCurrentActivitySatus") not in ("Producing", "Shut down"): print("   F", x)
r = get(base + "/311?f=json")
if r is not None and r.status_code == 200:
    print("  311 fields:", [f["name"] for f in r.json().get("fields", [])])
r = get(base + "/311/query?where=1%3D1&outFields=*&returnGeometry=false&resultRecordCount=3&f=json")
if r is not None and r.status_code == 200: print("  ", r.text[:1500])
# NSTA
r = get("https://opendata-nstauthority.hub.arcgis.com/api/v3/datasets?q=NSTA&page%5Bsize%5D=100")
if r is not None and r.status_code == 200:
    for d in r.json().get("data", []):
        a = d.get("attributes", {})
        if re.search(r"nsta|ukcs|north sea", (a.get("owner", "") + a.get("orgName", "") + (a.get("url") or "")), re.I):
            print("   N", a.get("name"), "|", (a.get("url") or "")[:120])
