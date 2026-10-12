"""Probe 5 (12 Oct 2026): sitemaps of Elering / AST / Litgrid for capacity pages; Swedish Vindbrukskollen ArcGIS services."""
import re, json
import requests
S = requests.Session(); S.headers["User-Agent"] = "Mozilla/5.0 (energy-data probe)"
def get(url, **kw):
    try: return S.get(url, timeout=60, **kw)
    except Exception as e: print("  ERR", url[:100], type(e).__name__); return None
PAT = r"capacit|galia|jauda|tootmisv|installed|generation|duomen|statist|production|gener|elektrostac|power-?plant|open-?data|register"
for name, urls in {"ELERING": ["https://elering.ee/sitemap.xml", "https://www.elering.ee/sitemap.xml", "https://www.elering.ee/en/sitemap.xml"],
                   "AST": ["https://www.ast.lv/sitemap.xml", "https://www.ast.lv/en/sitemap.xml"],
                   "LITGRID": ["https://www.litgrid.eu/sitemap.xml", "https://www.litgrid.eu/wp-sitemap.xml", "https://www.litgrid.eu/sitemap_index.xml", "https://www.litgrid.eu/robots.txt"]}.items():
    print("===", name)
    try:
        for u in urls:
            r = get(u); print(u, getattr(r, "status_code", None), len(r.text) if r is not None else "")
            if r is not None and r.ok:
                locs = re.findall(r"<loc>([^<]+)</loc>", r.text); print(len(locs), "locs")
                print([l for l in locs if re.search(PAT, l, re.I)][:50])
                subs = [l for l in locs if l.endswith(".xml")][:8]
                for s in subs:
                    rr = get(s)
                    if rr is not None and rr.ok:
                        ll = re.findall(r"<loc>([^<]+)</loc>", rr.text); print("  sub", s, len(ll), [l for l in ll if re.search(PAT, l, re.I)][:40])
    except Exception as e: print("ERROR", e)
print("=== SE Vindbrukskollen")
try:
    for u in ["https://ext-geodata-nationella.lansstyrelsen.se/arcgis/rest/services/LST?f=json", "https://ext-geodata-nationella.lansstyrelsen.se/arcgis/rest/services?f=json",
              "https://ext-geodata-nationella.lansstyrelsen.se/arcgis/rest/services/LST/lst_vbk_verk/MapServer?f=json",
              "https://ext-geodata.lansstyrelsen.se/arcgis/rest/services?f=json", "https://vbk.lansstyrelsen.se/api", "https://vbk.lansstyrelsen.se/config.json"]:
        r = get(u); print(u, getattr(r, "status_code", None), (r.text[:700] if r is not None else ""))
except Exception as e: print("ERROR", e)
print("=== SE other")
try:
    for u in ["https://www.svk.se/om-kraftsystemet/kraftsystemdata/", "https://www.svenskvindenergi.org/statistik", "https://svenskvindenergi.org/statistik/", "https://www.svk.se/en/national-grid/the-control-room/"]:
        r = get(u); print(u, getattr(r, "status_code", None))
        if r is not None and r.ok: print(sorted(set(re.findall(r'href="([^"]*(?:xls|csv|statistik|kapacitet)[^"]*)"', r.text)))[:20])
except Exception as e: print("ERROR", e)
