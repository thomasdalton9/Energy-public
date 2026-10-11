"""
Third pass: print the contents of the HTML data tables found by pass 2 (ANARE-CI wpDataTables, ministry
oil and gas tables, SENELEC key figures, CRSE tables), the data.gouv.ci energy dataset list, and the ARE Benin
document library API behind the SBPE daily production reports. Prints only - writes nothing.
"""
import html as ihtml
import json
import re
from urllib.parse import urljoin

import requests

requests.packages.urllib3.disable_warnings()
H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0"}


def get(url, **kw):
    try:
        r = requests.get(url, headers=H, timeout=60, verify=False, **kw)
        print(f"\n#### {r.status_code} {len(r.content):,} B {r.headers.get('content-type','')[:30]} {url}", flush=True)
        return r
    except Exception as e:  # noqa: BLE001
        print(f"\nFAIL {type(e).__name__} {url}", flush=True)
        return None


def clean(s):
    return re.sub(r"\s+", " ", ihtml.unescape(re.sub(r"<[^>]+>", " ", s))).strip()


def tables(url, maxrows=14, maxtab=8):
    r = get(url)
    if r is None or not r.ok:
        return None
    for i, t in enumerate(re.findall(r"<table.*?</table>", r.text, re.S | re.I)[:maxtab]):
        rows = re.findall(r"<tr.*?</tr>", t, re.S | re.I)
        print(f"  -- table {i}: {len(rows)} rows; id={re.search(r'id=.([\w-]+)', t).group(1) if re.search(r'id=.([\w-]+)', t) else ''}")
        for row in rows[:maxrows]:
            cells = [clean(c) for c in re.findall(r"<t[hd][^>]*>(.*?)</t[hd]>", row, re.S | re.I)]
            print("    ", " | ".join(cells)[:260])
    for m in re.findall(r"(?:data-wpdatatable_id|wdt_ID|table_id)[=\"':\s]+(\d+)", r.text)[:5]:
        print("  wpdatatable id:", m)
    for m in re.findall(r"[\"'](https?:[^\"']*(?:get_wdtable|wdt_[a-z_]+)[^\"']*)", r.text)[:3]:
        print("  wdt endpoint:", m)
    return r


for u in ["https://anare.ci/big-data-2/production-brute-delectricite-2/", "https://anare.ci/big-data-2/pointe-mw-2/",
          "https://anare.ci/big-data-2/consommation-nat-2/", "https://anare.ci/big-data-2/puissance-installee-totale-2/",
          "https://anare.ci/production-de-la-centrale-daggreko/", "https://anare.ci/test-affichage-tableau/",
          "https://anare.ci/espaces-operateurs-prives/les-fournisseurs-du-combustible-gaz-naturel/",
          "https://www.energie.gouv.ci/petrole/petrole-et-gaz", "https://www.senelec.sn/qui-sommes-nous/chiffres-cles/",
          "https://www.senelec.sn/nos-metiers/production/", "https://www.crse.sn/secteurs/secteur-electricite/",
          "https://edg.com.gn/production/"]:
    tables(u)

# wpDataTables ajax on ANARE
r = get("https://anare.ci/big-data-2/production-brute-delectricite-2/")
if r is not None and r.ok:
    ids = set(re.findall(r'data-wpdatatable_id=["\'](\d+)', r.text)) | set(re.findall(r'id=["\']wdt-?table[_-]?(\d+)', r.text, re.I)) | set(re.findall(r'table_(\d+)', r.text))
    print("candidate table ids:", sorted(ids)[:10])
    nonce = re.search(r'wdtNonceFrontendServerSide["\']?\s*[:=]\s*["\']([0-9a-f]+)', r.text) or re.search(r'wdtNonce["\']?\s*[:=]\s*["\']([0-9a-f]+)', r.text)
    print("nonce:", nonce.group(1) if nonce else None)
    for tid in sorted(ids)[:2]:
        try:
            rr = requests.post("https://anare.ci/wp-admin/admin-ajax.php?action=get_wdtable&table_id=" + tid, headers=H, verify=False, timeout=60,
                               data={"draw": 1, "start": 0, "length": 5, "wdtNonce": nonce.group(1) if nonce else ""})
            print("ajax", tid, rr.status_code, rr.text[:600])
        except Exception as e:  # noqa: BLE001
            print("ajax fail", e)
    print("WP REST pages/1156:")
    rp = get("https://anare.ci/wp-json/wp/v2/pages/1156")
    if rp is not None and rp.ok:
        print(clean(json.loads(rp.text)["content"]["rendered"])[:800])

# data.gouv.ci
r = get("https://data.gouv.ci/datasets?q=electricite")
if r is not None and r.ok:
    for m in sorted(set(re.findall(r'/datasets/([\w-]+)', r.text)))[:25]:
        print("  dataset:", m)
    txt = clean(re.sub(r"<script.*?</script>", " ", r.text, flags=re.S))
    print(txt[:1500])
for u in ["https://data.gouv.ci/api/v1/portals/yCWsyaGpA/datasets?limit=30", "https://data.gouv.ci/api/v1/datasets?limit=30"]:
    rr = get(u)
    if rr is not None and rr.ok:
        print(rr.text[:1500])

# ARE Benin library
r = get("https://www.are.bj/documents/categorie/rapports-de-production-journalier-de-la-sbpe")
if r is not None and r.ok:
    for m in sorted(set(re.findall(r'(?:https?:)?//[^"\'\s]*backoffice\.are\.bj[^"\'\s]*', r.text)))[:15]:
        print("  backoffice:", m)
    for m in sorted(set(re.findall(r'["\'](/?(?:api|_next|documents)[^"\']{0,120})', r.text)))[:25]:
        print("  path:", m)
for u in ["https://backoffice.are.bj/api/documents?category=rapports-de-production-journalier-de-la-sbpe",
          "https://backoffice.are.bj/api/documents", "https://backoffice.are.bj/"]:
    rr = get(u)
    if rr is not None:
        print(rr.text[:800])
