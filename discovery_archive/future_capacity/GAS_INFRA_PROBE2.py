"""Probe 2: GIE LNG/storage map PDFs (text), ENTSOG TYNDP links, PCI list, Sodir factmaps, NSTA data pages."""
import os, re, io, json
import requests
OUT = "discovery_archive/results/future_capacity"
UA = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36"}
def get(u, **kw):
    try:
        r = requests.get(u, headers=UA, timeout=(10, 90), **kw)
        print(r.status_code, len(r.content), r.headers.get("content-type", "")[:30], u[:150])
        return r
    except Exception as e:
        print("FAIL", u[:120], type(e).__name__); return None
def links(r, pat):
    for l in sorted(set(re.findall(r'href="([^"]+)"', r.text))):
        if re.search(pat, l, re.I): print("   L", l[:200])
# GIE maps
for name, u in [("gie_lng_map.pdf", "https://www.gie.eu/wp-content/uploads/filr/10797/GIE_LNG_2026_A0_1189x841_FULL_v110_TOPO.pdf"),
                ("gie_stor_map.pdf", "https://www.gie.eu/wp-content/uploads/filr/3997/GIE_STOR_2026_A0_1189x841_FULL_v110_TOPO.pdf")]:
    r = get(u)
    if r is not None and r.status_code == 200:
        open(os.path.join(OUT, name), "wb").write(r.content)
        try:
            import pdfplumber
            with pdfplumber.open(io.BytesIO(r.content)) as pdf:
                txt = "\n".join((p.extract_text() or "") for p in pdf.pages)
            open(os.path.join(OUT, name.replace(".pdf", ".txt")), "w").write(txt)
            print("   pages", len(pdf.pages), "chars", len(txt))
        except Exception as e: print("   pdf err", e)
r = get("https://www.gie.eu/publications/maps/gie-lng-map/")
if r is not None: links(r, r"filr|xlsx|dataset|download|\.csv")
r = get("https://www.gie.eu/publications/maps/")
if r is not None: links(r, r"filr|xlsx|dataset|download|\.csv")
# TYNDP
r = get("https://www.entsog.eu/tyndp")
if r is not None: links(r, r"2024|2026|2025|project|annex")
for u in ["https://tyndp.entsog.eu", "https://tyndp2024.entsog.eu", "https://www.entsog.eu/tyndp-2024", "https://www.entsog.eu/tyndp-2026",
          "https://www.entsog.eu/tyndp-2024-annex-a-project-tables", "https://tyndp.entsog.eu/api/v1/projects", "https://transparency.entsog.eu/api/v1/tyndpprojects"]:
    r = get(u)
    if r is not None and r.status_code == 200 and "html" in r.headers.get("content-type", ""): links(r, r"2024|2026|annex|project|xlsx")
# PCI
r = get("https://energy.ec.europa.eu/topics/infrastructure/projects-common-interest-and-projects-mutual-interest_en")
if r is not None: links(r, r"pci|annex|list|xlsx|2024|2023|delegated")
for u in ["https://data.europa.eu/eli/reg_del/2024/1041/oj", "https://eur-lex.europa.eu/legal-content/EN/TXT/HTML/?uri=OJ:L_202401041",
          "https://eur-lex.europa.eu/eli/reg_del/2024/1041/oj"]:
    r = get(u, allow_redirects=True)
    if r is not None and r.status_code == 200: print("   ", re.sub(r"\s+", " ", r.text[:300]))
# Sodir
for u in ["https://factmaps.sodir.no/api/rest/services/Factmaps/FactMapsWGS84/MapServer?f=json",
          "https://factmaps.sodir.no/api/rest/services/Factmaps/FactMapsWGS84/MapServer/201?f=json",
          "https://factpages.sodir.no/en/field/TableView/Producing",
          "https://factpages.sodir.no/en/field/PageView/All/43506",
          "https://data.norge.no/datasets?q=sodir",
          "https://sodir.no/en/", "https://www.sodir.no/en/", "https://www.sodir.no/en/facts/projects/"]:
    r = get(u)
    if r is not None and r.status_code == 200 and "json" in r.headers.get("content-type", ""):
        print("   ", r.text[:600])
# NSTA
r = get("https://www.nstauthority.co.uk/data-and-insights/data/themes/")
if r is not None: links(r, r"field|project|csv|xlsx|production|arcgis|hub")
for u in ["https://opendata-nstauthority.hub.arcgis.com/", "https://opendata-nstauthority.hub.arcgis.com/api/v3/datasets?q=fields"]:
    r = get(u)
    if r is not None and r.status_code == 200: print("   ", re.sub(r"\s+", " ", r.text[:300]))
