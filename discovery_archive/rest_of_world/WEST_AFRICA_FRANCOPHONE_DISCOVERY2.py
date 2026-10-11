"""
Second pass: open the data-looking pages found by WEST_AFRICA_FRANCOPHONE_DISCOVERY.py and print text snippets,
file links (xls/xlsx/csv/pdf/json), iframes, embedded chart/JSON endpoints. Prints only - writes nothing.
"""
import re
from urllib.parse import urljoin

import requests

requests.packages.urllib3.disable_warnings()
H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0"}
URLS = [
    "https://anare.ci/big-data-2/", "https://anare.ci/big-data-2/production-brute-delectricite-2/",
    "https://anare.ci/big-data-2/pointe-mw-2/", "https://anare.ci/pages/open-data/",
    "https://anare.ci/production-de-la-centrale-de-ciprel/", "https://anare.ci/les-infrastructures/production/",
    "https://anare.ci/espaces-operateurs-prives/les-fournisseurs-du-combustible-gaz-naturel/",
    "https://anare.ci/documents/bulletins-dentreprise/",
    "https://www.energie.gouv.ci/energie/statistique-activites", "https://www.energie.gouv.ci/petrole/statistique-activites",
    "https://www.energie.gouv.ci/petrole/petrole-et-gaz", "https://www.petroci.ci/production/",
    "https://www.cie.ci/nos-activites/production", "https://data.gouv.ci/datasets?q=electricite",
    "https://www.are.bj/documents/categorie/rapports-de-production-journalier-de-la-sbpe",
    "https://nigelec.ne/web/production", "https://edg.com.gn/production/", "https://edg.com.gn/rapports/",
    "https://www.arse.tg/espace-operateurs/electricite/production/", "https://www.arse.bf/documentation/bulletin-officiel/",
    "https://www.senelec.sn/qui-sommes-nous/chiffres-cles/", "https://www.senelec.sn/nos-metiers/production/",
    "https://www.senelec.sn/rapports/", "https://www.crse.sn/secteurs/secteur-electricite/",
    "https://www.crse.sn/secteurs/aval-intermediaire-gazier/", "https://www.crse.sn/documents/?cats=rapports",
    "https://www.somelec.mr", "https://www.edmsa.ml", "https://energie.gouv.ml/",
    "https://www.ceet.tg/tg/", "https://sbee.bj/",
]


def show(url):
    try:
        r = requests.get(url, headers=H, timeout=45, verify=False)
    except Exception as e:  # noqa: BLE001
        print(f"FAIL {type(e).__name__} {url}", flush=True)
        return
    print(f"\n#### {r.status_code} {len(r.content):,} B {url}", flush=True)
    if not r.ok:
        return
    html = r.text
    body = re.sub(r"<(script|style)[^>]*>.*?</\1>", " ", html, flags=re.S | re.I)
    txt = re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", body))
    print("TEXT:", txt[:700])
    files = sorted({urljoin(r.url, h) for h in re.findall(r'(?:href|src|data-src)=["\']([^"\']+\.(?:xlsx?|csv|json|pdf|zip)[^"\']*)', html, re.I)})
    for f in files[:25]:
        print("  FILE:", f)
    if len(files) > 25:
        print(f"  ... {len(files)} files total")
    for f in sorted(set(re.findall(r'<iframe[^>]+src=["\']([^"\']+)', html, re.I)))[:8]:
        print("  IFRAME:", f)
    for f in sorted(set(re.findall(r'["\'](https?://[^"\']*(?:api|json|chart|powerbi|datawrapper|tableau|flourish|admin-ajax)[^"\']*)', html, re.I)))[:10]:
        print("  ENDPOINT:", f[:200])
    for t in re.findall(r"<table", html, re.I)[:1]:
        print("  HAS HTML TABLE(s):", len(re.findall(r"<table", html, re.I)))
    for m in re.findall(r"(?:var|const)\s+\w*(?:data|chart)\w*\s*=\s*[\[{].{0,200}", html, re.I)[:3]:
        print("  JSDATA:", m[:220])


for u in URLS:
    show(u)
