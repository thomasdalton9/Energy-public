"""Sarawak Energy annual and sustainability reports (round 2: 2020-2022): can GitHub Actions reach them (directly or via the Internet Archive),
and do they give generation (GWh) by fuel / installed capacity by plant for Sarawak? Prints status per URL, then for each
PDF obtained the pages with generation / capacity / fuel-mix keywords. PDFs themselves are not saved; keyword pages go to
discovery_archive/results/sarawak/<name>_pages.txt."""
import io, os, re
from curl_cffi import requests as cr
import pypdf

OUT = "discovery_archive/results/sarawak"
os.makedirs(OUT, exist_ok=True)
S = cr.Session()
URLS = {
    "ASR2020": "https://www.sarawakenergy.com/assets/pdf/INTERACTIVE-Sarawak-Energy-ASR20.pdf",
    "ASR2021": "https://www.sarawakenergy.com/assets/pdf/ASR21-Interactive.pdf",
    "ASR2022": "https://www.sarawakenergy.com/assets/pdf/SarawakEnergy-ARSR2022.pdf",
}
LISTING = "https://sarawakenergy.com/investors/annual-and-sustainability-reports"
KEY = re.compile(r"GWh|generation mix|energy mix|installed capacity|Bakun|Murum|Batang Ai|Balingian|Mukah|Sejingkat|Tanjung Kidurong|Sri Aman|diesel|hydro", re.I)


def fetch(name, url):
    tries = [("direct", url, {}),
             ("direct+referer", url, {"Referer": "https://www.sarawakenergy.com/"}),
             ("wayback", "https://web.archive.org/web/2025/" + url, {})]
    for label, u, h in tries:
        try:
            r = S.get(u, impersonate="chrome", timeout=90, headers=h, allow_redirects=True)
            print(f"[{name}] {label}: {r.status_code} {len(r.content)}B {r.headers.get('content-type','')[:30]}")
            if r.status_code == 200 and r.content[:4] == b"%PDF":
                return r.content, label
        except Exception as e:  # noqa: BLE001
            print(f"[{name}] {label}: ERR {type(e).__name__} {str(e)[:100]}")
    return None, None


# listing page (skipped in round 2)
for label, u in ():
    try:
        r = S.get(u, impersonate="chrome", timeout=60)
        print(f"[listing] {label}: {r.status_code} {len(r.content)}B")
        if r.status_code == 200:
            pdfs = sorted(set(re.findall(r'href="([^"]+\.pdf)"', r.text, flags=re.I)))
            print("  pdf links:", len(pdfs))
            for p in pdfs[:25]:
                print("   ", p)
            break
    except Exception as e:  # noqa: BLE001
        print(f"[listing] {label}: ERR {e}")

for name, url in URLS.items():
    b, how = fetch(name, url)
    if not b:
        continue
    rd = pypdf.PdfReader(io.BytesIO(b))
    print(f"  {name}: {len(rd.pages)} pages via {how}")
    keep = []
    for i, pg in enumerate(rd.pages):
        try:
            tx = pg.extract_text() or ""
        except Exception:  # noqa: BLE001
            continue
        hits = len(KEY.findall(tx))
        if hits >= 4 and re.search(r"GWh|MW", tx):
            keep.append((i + 1, tx))
    print(f"  {name}: {len(keep)} pages with generation/capacity keywords: {[k[0] for k in keep][:20]}")
    with open(f"{OUT}/{name}_pages.txt", "w") as f:
        for pn, tx in keep[:25]:
            f.write(f"\n===== {name} page {pn} =====\n{tx}\n")
