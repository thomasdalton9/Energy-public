"""
Morocco gas probe, pass 4: per-plant figures.

Targets:
  hcp     HCP 'Annuaire statistique du Maroc' (Excel), chapter 'Energie et Eau': dump every table that mentions gas, thermal plants,
          Tahaddart or Ain Beni Mathar, with headers, for every yearbook the download page lists (each covers three years)
  sites   Ministry of Energy Transition static site (chiffres cles, rapports), Observatoire Marocain de l'Energie, ONEE (no certificate
          check: the certificate fails verification from Actions)

Usage: python3 MOROCCO_GAS_PROBE4.py <target>
"""
import io
import os
import re
import sys
import time
import zipfile

import pandas as pd
import requests
import urllib3

urllib3.disable_warnings()
UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36"
H = {"User-Agent": UA}
OUT = "probe_files"
os.makedirs(OUT, exist_ok=True)


def get(url, tries=2, timeout=(20, 240), **kw):
    kw.setdefault("headers", H)
    last = None
    for _ in range(tries):
        try:
            return requests.get(url, timeout=timeout, **kw)
        except requests.RequestException as e:
            last = e
            time.sleep(3)
    print(f"  ERR {url[:140]} {type(last).__name__}: {str(last)[:160]}")
    return None


def page_links(url, pat=r".", n=80, **kw):
    r = get(url, **kw)
    print(f"\n=== {url} -> {r.status_code if r is not None else None} {len(r.content) if r is not None else ''} {r.url if r is not None else ''}")
    if r is None or r.status_code != 200:
        return None, []
    out, seen = [], set()
    for h, t in re.findall(r'href=["\']([^"\'#]+)["\'][^>]*>(.{0,300}?)</a', r.text, re.S | re.I):
        t = re.sub(r"<[^>]+>|\s+", " ", t).strip()
        if re.search(pat, h + " " + t, re.I) and h not in seen:
            seen.add(h)
            out.append((requests.compat.urljoin(r.url, h), t))
    for h, t in out[:n]:
        print("   ", h[:170], "|", t[:90])
    return r, out


SHEET_PAT = r"gaz|thermique|Tahaddart|Mathar|cycle combin"


def hcp():
    r = get("https://www.hcp.ma/downloads/Annuaires-statistiques-du-Maroc-Format-Excel_t22392.html")
    items = []
    if r is not None and r.ok:
        for m in re.finditer(r'href="(/file/(\d+)/)"', r.text):
            ctx = re.sub(r"<[^>]+>|\s+", " ", r.text[m.end():m.end() + 600])
            t = re.search(r"Annuaire[^|]{0,80}?ann[ée]e\s*(\d{4})", ctx)
            items.append((m.group(2), t.group(1) if t else ""))
    seen, uniq = set(), []
    for fid, y in items:
        if fid not in seen:
            seen.add(fid)
            uniq.append((fid, y))
    print("files:", uniq)
    for fid, y in uniq[:14]:
        rr = get(f"https://www.hcp.ma/file/{fid}/")
        if rr is None or not rr.ok:
            continue
        cd = requests.utils.unquote(rr.headers.get("content-disposition", ""))
        print(f"\n##### file {fid} {len(rr.content)} {cd[:140]}")
        books = []
        if rr.content[:2] == b"PK":
            z = zipfile.ZipFile(io.BytesIO(rr.content))
            names = z.namelist()
            if any(n.startswith("xl/") for n in names):
                books = [(cd, rr.content)]
            else:
                books = [(n, z.read(n)) for n in names if re.search(r"nergie", n, re.I) and n.lower().endswith((".xlsx", ".xls"))]
                if not books:
                    print("   no energy workbook:", names[:30])
        elif rr.content[:4] == b"\xd0\xcf\x11\xe0":
            books = [(cd, rr.content)]
        for name, content in books:
            try:
                x = pd.read_excel(io.BytesIO(content), sheet_name=None, header=None)
            except Exception as e:  # noqa: BLE001
                print("   read error", name, type(e).__name__)
                continue
            tag = re.sub(r"[^0-9A-Za-z]+", "_", name)[-40:]
            for s, d in x.items():
                txt = d.astype(str)
                if not txt.apply(lambda col: col.str.contains(SHEET_PAT, case=False, regex=True)).any().any():
                    continue
                if re.search(r"SOMMAIRE|^pg$", str(s), re.I):
                    continue
                d = d.dropna(how="all").dropna(axis=1, how="all")
                fn = f"{OUT}/hcp_{fid}_{tag}_{re.sub(r'[^0-9A-Za-z]+', '_', str(s))}.csv"
                d.to_csv(fn, index=False, header=False)
                print(f"\n--- {name} sheet {s!r} -> {fn}")
                for _, row in d.iterrows():
                    vals = [str(v).strip() for v in row.tolist() if str(v) != "nan"]
                    # keep the Arabic column out of the log
                    vals = [v for v in vals if not re.search(r"[؀-ۿ]", v)]
                    if vals:
                        print("    ", " | ".join(v[:60] for v in vals)[:300])


def sites():
    for u in ["https://www.mem.gov.ma/Pages/art7b30.html?v=4", "https://www.mem.gov.ma/Pages/rapports.html",
              "https://www.mem.gov.ma/Pages/secteur0a89.html?e=1"]:
        r, out = page_links(u, r"pdf|xls|chiffre|rapport|statist|bilan|gaz|electr|energ", 80)
        if r is not None and "art7b30" in u:
            t = re.sub(r"<script.*?</script>|<style.*?</style>", " ", r.text, flags=re.S)
            t = re.sub(r"<[^>]+>", " ", t)
            t = re.sub(r"\s+", " ", t)
            i = t.lower().find("gaz")
            print("   TEXT:", t[max(0, i - 1500):i + 2500])
        for h, tt in out:
            if h.lower().endswith(".pdf") and re.search(r"chiffre|bilan|statist|tableau|energ", h + tt, re.I):
                rr = get(h)
                if rr is not None and rr.ok and rr.content[:4] == b"%PDF":
                    import pdfplumber
                    with pdfplumber.open(io.BytesIO(rr.content)) as pdf:
                        print(f"   PDF {h[:150]} pages={len(pdf.pages)}")
                        for i, pg in enumerate(pdf.pages[:40]):
                            for ln in (pg.extract_text() or "").split("\n"):
                                if re.search(r"gaz naturel|Tahaddart|Mathar|cycle combin|GME|gazoduc", ln, re.I):
                                    print(f"     p{i + 1}: {ln.strip()[:200]}")
    r, out = page_links("https://www.observatoirenergie.ma/", r".", 80)
    for u in ["https://www.observatoirenergie.ma/donnees/", "https://www.observatoirenergie.ma/publications/",
              "https://www.observatoirenergie.ma/bilan-energetique/", "https://www.observatoirenergie.ma/wp-json/wp/v2/pages?per_page=50"]:
        page_links(u, r"pdf|xls|csv|bilan|donn|statist|gaz|electr", 60)
    r, out = page_links("https://www.one.org.ma/fr/pages/index.asp", r"rapport|annuel|activit|chiffre|statist|pdf|production|publication", 80,
                        verify=False, timeout=(25, 60))
    for u in ["https://www.one.org.ma/FR/pages/interne.asp?esp=1&id1=3&id2=114&t2=1"]:
        rr = get(u, verify=False, timeout=(25, 60))
        if rr is not None and rr.ok:
            t = re.sub(r"<script.*?</script>|<style.*?</style>", " ", rr.text, flags=re.S)
            print("   ONEE page text:", re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", t))[:3000])
            for h, tt in re.findall(r'href=["\']([^"\'#]+)["\'][^>]*>(.{0,200}?)</a', rr.text, re.S | re.I):
                if re.search(r"pdf|rapport|chiffre", h + tt, re.I):
                    print("    ", requests.compat.urljoin(rr.url, h)[:170], "|", re.sub(r"<[^>]+>|\s+", " ", tt)[:80])


if __name__ == "__main__":
    {"hcp": hcp, "sites": sites}[sys.argv[1]]()
