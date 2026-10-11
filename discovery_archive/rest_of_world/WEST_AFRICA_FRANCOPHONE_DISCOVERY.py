"""
One-off discovery: do francophone West African grid operators / utilities / regulators / ministries publish
generation, load or dispatch data in structured form? Prints only - writes nothing.
For each site: GET the home page (and a few likely sub-paths), print status, then list links that look like
data (xls/xlsx/csv/json/api/dashboard/bulletin/statistique/rapport/production/dispatch/conduite).
"""
import re
import sys
from urllib.parse import urljoin

import requests

requests.packages.urllib3.disable_warnings()
H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0"}
PAT = re.compile(r"\.(xlsx?|csv|json|zip)(\?|$)|api|dashboard|tableau|bulletin|statisti|rapport|report|production|"
                 r"dispatch|conduite|dispatching|donn|data|open|energie-electrique|gaz|gas|GTA|annuaire|chiffres", re.I)

SITES = {
    "Senegal": ["https://www.senelec.sn", "https://www.senelec.sn/statistiques", "https://crse.sn", "https://www.crse.sn/publications",
                "https://aner.sn", "https://www.aner.sn", "https://www.energie.gouv.sn", "https://www.petrosen.sn",
                "https://data.gouv.sn", "https://www.ansd.sn", "https://www.ansd.sn/rubriques/energie"],
    "Cote d'Ivoire": ["https://www.ci-energies.ci", "https://ci-energies.ci", "https://www.cie.ci", "https://www.anare.ci",
                      "https://www.energie.gouv.ci", "https://www.petroci.ci", "https://data.gouv.ci", "https://www.ins.ci"],
    "Mali": ["https://www.edmsa.ml", "https://edm-sa.com", "https://www.cree.ml", "https://cree.ml", "https://www.energie.gouv.ml"],
    "Burkina Faso": ["https://www.sonabel.bf", "https://sonabel.bf", "https://www.arse.bf", "https://arse.bf"],
    "Benin": ["https://www.sbee.bj", "https://sbee.bj", "https://www.ceb-benin.org", "https://www.cebenin.org", "https://www.are.bj", "https://are.bj"],
    "Togo": ["https://www.ceet.tg", "https://ceet.tg", "https://www.arse.tg", "https://arse.tg"],
    "Niger": ["https://www.nigelec.ne", "https://nigelec.ne", "https://www.energie.gouv.ne"],
    "Guinea": ["https://www.edg.com.gn", "https://edg.com.gn", "https://www.edg-guinee.com"],
    "Mauritania": ["https://www.somelec.mr", "https://somelec.mr", "https://www.are.mr", "https://are.mr", "https://www.petrole.gov.mr", "https://www.mpem.gov.mr"],
    "Regional / gas": ["https://www.bp.com/en/global/corporate/what-we-do/oil-and-gas/gta.html", "https://www.energycharter.org",
                       "https://www.ecowapp.org", "https://www.erera.arrec.org", "https://www.arrec.org", "https://www.wapp-ecowas.org"],
}


def get(url):
    try:
        r = requests.get(url, headers=H, timeout=40, verify=False, allow_redirects=True)
        print(f"{r.status_code} {len(r.content):>9,} B {r.headers.get('content-type', '')[:30]:30} {url}"
              + (f" -> {r.url}" if r.url.rstrip('/') != url.rstrip('/') else ""), flush=True)
        return r
    except Exception as e:  # noqa: BLE001
        print(f"FAIL {type(e).__name__}: {str(e)[:100]} {url}", flush=True)
        return None


for country, urls in SITES.items():
    print("\n" + "=" * 20, country, "=" * 20, flush=True)
    for u in urls:
        r = get(u)
        if r is None or not r.ok or "html" not in r.headers.get("content-type", "").lower():
            continue
        t = re.search(r"<title[^>]*>(.*?)</title>", r.text, re.I | re.S)
        print("   title:", re.sub(r"\s+", " ", t.group(1))[:100] if t else "-")
        seen = []
        for h in re.findall(r'href=["\']([^"\']+)["\']', r.text, re.I):
            full = urljoin(r.url, h)
            if PAT.search(full) and full not in seen and not re.search(r"facebook|twitter|linkedin|youtube|instagram", full):
                seen.append(full)
        for full in seen[:30]:
            print("   link:", full)
sys.stdout.flush()
