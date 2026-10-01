"""
Round 3 (Oct-2026): gas use for the Dominican Republic and Jamaica, Jamaica
generation, and a check of LUMA's daily report for Puerto Rico.
  do3   MEM monthly "Boletin Informativo Generacion y Gestion de Energia" (text of
        the PDF: does it carry fuel / gas consumption?), OC document tree
        (Informe Mensual / Diario de Operacion), CNE data portal (datacne.gob.do)
        scripts / API, Banco Central sector externo links
  jm3   MSET statistics-data / petroleum / energy-balances document lists, OUR
        quarterly performance report text, MSET energy balance 2024 text
  pr3   LUMA daily availability report text (does it hold generation by unit?)
Pages and PDFs are saved under probe_files/.
Usage: python3 CARIBBEAN_POWER_GAS_PROBE3.py do3|jm3|pr3
"""

print("STARTING", flush=True)

import io
import os
import re
import sys
from urllib.parse import urljoin

import requests

H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) "
                   "Chrome/124.0 Safari/537.36", "Accept-Language": "es-ES,es;q=0.9,en;q=0.8"}
S = requests.Session()
S.headers.update(H)
OUT = "probe_files"
os.makedirs(OUT, exist_ok=True)


def get(url, save=None, **kw):
    try:
        r = S.get(url, timeout=90, **kw)
        print(f"[{r.status_code}] {r.url[:200]} {r.headers.get('content-type')} {len(r.content):,} B", flush=True)
        if save:
            with open(os.path.join(OUT, save), "wb") as f:
                f.write(r.content)
        return r
    except requests.RequestException as e:
        print(f"ERR {url}: {type(e).__name__}: {str(e)[:200]}", flush=True)
        return None


def pdf_text(content, name, pages=None, pattern=None, maxchars=3500):
    import pdfplumber
    try:
        with pdfplumber.open(io.BytesIO(content)) as pdf:
            print(f"  {name}: {len(pdf.pages)} pages", flush=True)
            with open(os.path.join(OUT, name + ".txt"), "w") as f:
                for i, p in enumerate(pdf.pages):
                    t = p.extract_text() or ""
                    f.write(f"\n===== page {i + 1}\n{t}")
                    if (pages is None or i < pages) and (pattern is None or re.search(pattern, t, re.I)):
                        print(f"  --- page {i + 1}\n{t[:maxchars]}", flush=True)
    except Exception as e:
        print("  pdf error", e)


def links(r, pat):
    seen = set()
    for m in re.finditer(r"<a\b[^>]*href=[\"']([^\"']+)[\"'][^>]*>(.*?)</a>", r.text, re.I | re.S):
        link = urljoin(r.url, m.group(1))
        txt = re.sub(r"<[^>]+>|\s+", " ", m.group(2)).strip()
        if link not in seen and re.search(pat, link + " " + txt, re.I):
            seen.add(link)
            print("   A", link[:220], "|", txt[:90])
    return seen


def do3():
    r = get("https://mem.gob.do/wp-content/uploads/2026/07/05.-Boletin-Informativo-Generacion-y-Gestion-Energia-Mayo-2026.pdf")
    if r is not None and r.ok:
        pdf_text(r.content, "mem_boletin_2026_05", pages=30)
    for node, tab, mod in [("ent115735", 260, 1174), ("ent115727", 259, 1168), ("ent115728", 259, 1168)]:
        r = get(f"https://www.oc.org.do/DesktopModules/Bring2mind/DMX/GetTreeviewContents.aspx?NodeId={node}&TabId={tab}"
                f"&ModuleId={mod}", save=f"oc_tree_{node}.txt")
        if r is not None:
            print("   ", r.text[:2000])
    for u in ["https://www.oc.org.do/tabid/260/Default.aspx?Command=Core_ViewFolder&EntryId=115735",
              "https://www.oc.org.do/DesktopModules/Bring2mind/DMX/API/Entries/List?folderId=115735",
              "https://www.oc.org.do/DesktopModules/Bring2mind/DMX/API/Folders/115735/Entries"]:
        r = get(u, save="oc_try_" + re.sub(r"\W", "_", u[-30:]) + ".html")
        if r is not None:
            links(r, r"EntryId=\d+|Download|\.pdf|\.xls")
    r = get("https://datacne.gob.do/repositorio-estadistico", save="datacne_repo.html")
    if r is not None:
        for js in re.findall(r"(?:src|href)=[\"']([^\"']+\.js[^\"']*)", r.text):
            u = urljoin(r.url, js)
            j = get(u)
            if j is not None and j.ok:
                for m in set(re.findall(r"[\"'`](https?://[^\"'`\s]{6,150}|/api/[^\"'`\s]{2,120})[\"'`]", j.text)):
                    print("     JSURL", m)
        print(r.text[:3000])
    for u in ["https://datacne.gob.do/api/repositorio", "https://datacne.gob.do/api/repositorio-estadistico",
              "https://datacne.gob.do/tablero-dinamico", "https://datacne.gob.do/"]:
        r = get(u, save="datacne_" + re.sub(r"\W", "_", u[22:]) + ".html")
        if r is not None and "html" in (r.headers.get("content-type") or ""):
            links(r, r"tablero|repositorio|gas|combust|import|generac|estad|\.xlsx?|\.csv|\.pdf")
        elif r is not None:
            print("   ", r.text[:800])
    r = get("https://www.bancentral.gov.do/a/CustomView/2532-sector-externo", save="bcrd_sector_externo.html")
    if r is not None:
        links(r, r"petrol|import|combust|\.xlsx?|gas")


def jm3():
    for u in ["https://www.mset.gov.jm/document-category/statistics-data/",
              "https://www.mset.gov.jm/document-category/petroleum/",
              "https://www.mset.gov.jm/document-category/energy-balances/",
              "https://www.mset.gov.jm/documents/quarterly-performance-reports/"]:
        r = get(u, save="mset_" + u.rstrip("/").split("/")[-1] + ".html")
        if r is not None:
            links(r, r"\.(pdf|xlsx?|csv)|/documents/|statistic|petroleum|generation|electricity|data")
    r = get("https://our.org.jm/wp-content/uploads/2026/09/Quarterly-Performance_April-to-June-2026-FINAL_September.pdf")
    if r is not None and r.ok:
        pdf_text(r.content, "our_qpr_2026q2", pattern=r"generat|fuel|gas|LNG|MWh|GWh|IPP|renewable", maxchars=2500)
    r = get("https://www.mset.gov.jm/wp-content/uploads/2021/07/MSETT-National-Energy-Balance-2024_Update.pdf")
    if r is not None and r.ok:
        pdf_text(r.content, "mset_balance_2024", pattern=r"natural gas|LNG|generation", maxchars=2500)


def pr3():
    r = get("https://lumapr.com/wp-content/uploads/2026/09/20260929-Daily-Availability-Report.pdf")
    if r is not None and r.ok:
        pdf_text(r.content, "luma_daily_20260929", pages=6)
    for y, m, d in [(2021, 6, 15), (2022, 6, 15), (2024, 6, 15)]:
        get(f"https://lumapr.com/wp-content/uploads/{y}/{m:02d}/{y}{m:02d}{d:02d}-Daily-Availability-Report.pdf")


if __name__ == "__main__":
    {"do3": do3, "jm3": jm3, "pr3": pr3}[sys.argv[1]]()
    print("DONE", flush=True)
