"""Round 2: follow the promising pages found by WEST_AFRICA_OTHER_DISCOVERY.py
(ARSEL production dashboard, PURA statistics, NAWEC, ENEO production/annual report,
SNH, EG LNG, EDSA, LERC, Cape Verde ARME electricity categories, Electra, WAPP)."""
import re, sys
from urllib.parse import urljoin
import requests
H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36"}
URLS = [
 "https://arsel-cm.org/index.php/tableau-de-bord_production/",
 "https://arsel-cm.org/wp-json/wp/v2/pages?search=production&per_page=20&_fields=link,title",
 "https://arsel-cm.org/wp-json/wp/v2/media?per_page=30&_fields=source_url,mime_type&search=production",
 "https://arsel-cm.org/wp-json/wp/v2/pages?search=statistique&per_page=20&_fields=link,title",
 "https://pura.gm/statistics/", "https://pura.gm/economic-regulations/statistics/",
 "https://pura.gm/economic-regulations/reports/",
 "https://pura.gm/wp-json/wp/v2/media?per_page=40&_fields=source_url,mime_type&search=electricity",
 "https://nawec.gm/", "https://nawec.gm/load-shedding-schedule/",
 "https://nawec.gm/wp-json/wp/v2/pages?per_page=30&_fields=link,title",
 "https://www.eneocameroon.cm/index.php/en/production-en/la-production-a-eneo-en",
 "https://www.eneocameroon.cm/index.php/en/2024-annual-report",
 "https://www.snh.cm/wp-json/wp/v2/pages?search=production&per_page=20&_fields=link,title",
 "https://www.snh.cm/wp-json/wp/v2/media?per_page=30&_fields=source_url,mime_type&search=production",
 "https://eglng.com/en", "https://www.edsa.sl", "https://www.lerc.gov.lr/",
 "https://www.lerc.gov.lr/sitemap.xml", "https://www.mme.gov.lr/",
 "https://www.arme.cv/index.php?option=com_jdownloads&view=category&catid=146&Itemid=741",
 "https://www.arme.cv/index.php?option=com_jdownloads&view=category&catid=21&Itemid=741",
 "https://www.arme.cv/index.php/energia", "https://www.arme.cv/index.php/electricidade",
 "https://www.electra.cv/", "https://electra.cv/", "https://www.ecowapp.org/sitemap.xml",
 "https://eagb.gw/", "https://www.eagb.gw/",
]
KEY = re.compile(r"(\.csv|\.xlsx?|\.json|\.pdf|\.zip|api|dashboard|statist|bulletin|bolet|report|relat|generation|produ|load|dispatch|energ|electr|data|dados|demand)", re.I)
out = []
def p(*a):
    s = " ".join(str(x) for x in a); print(s, flush=True); out.append(s)
for u in URLS:
    p(f"\n== {u}")
    try:
        r = requests.get(u, headers=H, timeout=(10, 25))
    except requests.RequestException as e:
        p(f"  ERR {type(e).__name__}"); continue
    ct = r.headers.get("content-type", "")
    p(f"  status={r.status_code} final={r.url} bytes={len(r.content)} ct={ct}")
    if r.status_code != 200: continue
    if "json" in ct or u.endswith("sitemap.xml") or "xml" in ct:
        p("  BODY:", r.text[:1500].replace("\n", " ")); continue
    txt = re.sub(r"<script.*?</script>|<style.*?</style>", " ", r.text, flags=re.S | re.I)
    plain = re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", txt))
    p("  TEXT:", plain[:700])
    iframes = re.findall(r'<(?:iframe|embed|object)[^>]+src=["\']([^"\']+)', r.text, re.I)
    for i in iframes[:10]: p("  IFRAME:", urljoin(r.url, i))
    seen = set()
    for m in re.finditer(r'(?:href|src|data-[a-z-]*url)=["\']([^"\'#]+)', r.text, re.I):
        l = urljoin(r.url, m.group(1))
        if l in seen or not KEY.search(l) or re.search(r"(fonts\.g|\.(png|jpe?g|webp|gif|svg|css)\b|facebook|twitter|linkedin|wp-includes|/feed)", l, re.I): continue
        seen.add(l)
    for l in sorted(seen)[:50]: p("  link:", l)
open("west_africa_other_output2.txt", "w").write("\n".join(out))
