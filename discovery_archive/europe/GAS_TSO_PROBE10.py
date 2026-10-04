"""
Probe 10: Enagas statistical bulletin PDF list + text of a Dec-2022 and a latest bulletin (demand tables);
Transgaz fragment raw HTML; Gasgrid transparency pages; Plinacro menu links. Prints only.
"""
import re
import subprocess
import sys

import requests

UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0 Safari/537.36"
H = {"User-Agent": UA}
EN = "https://www.enagas.es"
FRAG = EN + "/content/enagas/en/gestion-tecnica-sistema/energy-data/publicaciones/boletin-estadistico-gas/_jcr_content/responsiveGrid/container/filedownloadpaginati.nocache.html/enagas/components/content/filedownloadpagination"


def enagas():
    print("=== ENAGAS bulletins")
    pdfs = []
    for pg in range(0, 26):
        try:
            r = requests.get(FRAG, params={"page_": pg}, headers=H, timeout=60)
            ls = re.findall(r'href="(/content/dam[^"]+\.pdf)"', r.text)
            print(" page", pg, r.status_code, len(ls), ls[:1])
            pdfs += [x for x in ls if x not in pdfs]
            if not ls and pg > 1:
                break
        except Exception as e:  # noqa: BLE001
            print(" page", pg, type(e).__name__)
    print(len(pdfs), "pdfs")
    for p in pdfs:
        print("  ", p.split("/")[-1])
    pick = [p for p in pdfs if re.search(r"dic22|dec22|dic_22", p, re.I)] or [p for p in pdfs if "22" in p.split("/")[-1]][:1]
    pick = pick[:1] + pdfs[:1]
    for p in pick:
        try:
            r = requests.get(EN + p, headers=H, timeout=120)
            print("PDF", p.split("/")[-1], r.status_code, len(r.content))
            open("/tmp/b.pdf", "wb").write(r.content)
            out = subprocess.run(["pdftotext", "-layout", "/tmp/b.pdf", "-"], capture_output=True, text=True).stdout
            print(" text chars", len(out))
            for m in re.finditer(r"(?i)demand|consumption", out):
                pass
            idx = [m.start() for m in re.finditer(r"(?i)conventional|total demand|demanda", out)]
            print(" hits", len(idx))
            seen = 0
            for i in idx[:6]:
                print("-----", i)
                print(out[max(0, i - 300): i + 1800])
                seen += 1
        except Exception as e:  # noqa: BLE001
            print("PDF fail", type(e).__name__, str(e)[:200])


def transgaz():
    print("=== TRANSGAZ raw")
    r = requests.get("https://www.transgaz.ro/new-tabel-transparenta-masuratori_en.php?poz=197", headers={**H, "Referer": "https://www.transgaz.ro/en/clients/operational-data/physical-flows"}, timeout=60)
    t = r.text
    print(len(t), r.encoding)
    i = t.find("<body")
    print(re.sub(r"\s+", " ", t[i:i + 6000]))
    print("...TAIL...")
    print(re.sub(r"\s+", " ", t[-3000:]))


def gasgrid():
    print("=== GASGRID pages")
    for u in ("https://gasgrid.fi/en/gas-business/transparency-and-market-information/", "https://gasgrid.fi/en/topic/gas-consumption/",
              "https://gasgrid.fi/en/bulletins/gas-market-review-the-role-of-gas-infrastructure-in-balancing-power-highlighted-strong-growth-in-biomethane-gas-use-in-finland/"):
        try:
            r = requests.get(u, headers=H, timeout=60)
            print(u, r.status_code, len(r.text))
            for m in re.finditer(r'href="([^"]+)"[^>]*>([^<]{0,90})', r.text):
                h = m.group(1)
                if re.search(r"\.(xlsx?|csv|pdf|zip)|wp-content/uploads|transparen|entsog|data|statistic|review|consumption", h, re.I) and "cookie" not in h:
                    print("   ", h[:170], "|", m.group(2).strip()[:70])
            body = re.sub(r"<script.*?</script>|<style.*?</style>", "", r.text, flags=re.S)
            txt = re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", body))
            j = txt.find("onsumption")
            print("  TEXT:", txt[max(0, j - 300): j + 1200])
        except Exception as e:  # noqa: BLE001
            print(u, type(e).__name__)


def plinacro():
    print("=== PLINACRO menu")
    r = requests.get("https://www.plinacro.hr/default.aspx?id=6", headers=H, timeout=60)
    seen = set()
    for m in re.finditer(r'href="([^"]*default\.aspx\?id=\d+[^"]*)"[^>]*>([^<]{2,90})', r.text):
        if m.group(1) not in seen:
            seen.add(m.group(1))
            print("  ", m.group(1)[:60], "|", m.group(2).strip())


for fn in (enagas, transgaz, gasgrid, plinacro):
    try:
        fn()
    except Exception as e:  # noqa: BLE001
        print(fn.__name__, "FAILED", type(e).__name__, str(e)[:300])
sys.exit(0)
