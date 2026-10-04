"""
Indonesia raw power data, discovery round 1 (run from a GitHub Actions runner; the sandbox is blocked).
Candidates: ESDM homepage data cards (a GitHub scraper reads a daily gas card there - is there an
electricity one?), Ditjen Gatrik download index (Statistik Ketenagalistrikan, xlsx?), PLN statistics page,
BPS statistics tables / web API, Satu Data catalogue (CKAN), guessed PLN UIP2B / ESDM dashboard hosts.
Prints status, size, text snippets around power keywords and data-like links. Nothing is written.
"""
import re
import sys
from urllib.parse import urljoin

import requests
import urllib3

urllib3.disable_warnings()
H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) "
                   "Chrome/124.0 Safari/537.36", "Accept-Language": "id,en;q=0.8"}
T = 40
KW = re.compile(r"(listrik|beban|MW\b|GWh|TWh|electric|power|pembangkit|daya mampu|produksi)", re.I)
LINK = re.compile(r"""(?:href|src)=["']([^"']+)["']""", re.I)
DATA = re.compile(r"\.(xlsx?|csv|json|pdf|zip)(\?|$)|api|statistik|data|download", re.I)


def out(s):
    print(s, flush=True)


def probe(name, url, show=800, links=60, kw=True, **kwa):
    out(f"\n{'=' * 90}\n{name}: {url}")
    try:
        r = requests.get(url, headers=H, timeout=T, verify=False, **kwa)
    except Exception as e:
        out(f"  ERROR {e!r}")
        return None
    out(f"  -> {r.status_code} {r.headers.get('content-type')} {len(r.content)} bytes final={r.url}")
    ct = r.headers.get("content-type", "")
    if "html" in ct or "json" in ct or "text" in ct:
        txt = r.text
        if show:
            out("  HEAD: " + re.sub(r"\s+", " ", txt[:show]))
        if kw:
            plain = re.sub(r"<script.*?</script>|<style.*?</style>", " ", txt, flags=re.S | re.I)
            plain = re.sub(r"<[^>]+>", " ", plain)
            plain = re.sub(r"\s+", " ", plain)
            hits = [m.start() for m in KW.finditer(plain)]
            seen = -1000
            n = 0
            for h in hits:
                if h - seen < 200:
                    continue
                seen = h
                out("  KW: ..." + plain[max(0, h - 120):h + 180])
                n += 1
                if n > 25:
                    break
        ls = []
        for l in LINK.findall(r.text):
            u = urljoin(r.url, l)
            if DATA.search(u) and u not in ls:
                ls.append(u)
        for u in ls[:links]:
            out("  LINK " + u)
    return r


def main():
    # ESDM homepage cards
    for u in ("https://www.esdm.go.id/en", "https://www.esdm.go.id/id", "https://www.esdm.go.id/"):
        probe("ESDM home", u, show=300, links=80)
    # Gatrik
    probe("Gatrik home", "https://gatrik.esdm.go.id/", links=100)
    for cat in ("statistik", "publikasi", "buku", "data"):
        probe(f"Gatrik download_index {cat}",
              f"https://gatrik.esdm.go.id/frontend/download_index?kode_category={cat}", show=300, links=120, kw=False)
    probe("Gatrik download_index all", "https://gatrik.esdm.go.id/frontend/download_index", show=300, links=150, kw=False)
    # PLN
    for u in ("https://web.pln.co.id/stakeholder/laporan-statistik", "https://web.pln.co.id/statics/",
              "https://web.pln.co.id/", "https://www.pln.co.id/stakeholder/laporan-statistik"):
        probe("PLN", u, show=300, links=80)
    # guessed dispatch / dashboards
    for u in ("https://uip2b.pln.co.id/", "https://p2b.pln.co.id/", "https://uip2bjb.pln.co.id/",
              "https://dashboard.esdm.go.id/", "https://satudata.esdm.go.id/", "https://data.esdm.go.id/",
              "https://ebtke.esdm.go.id/", "https://simebtke.esdm.go.id/", "https://geoportal.esdm.go.id/",
              "https://momi.esdm.go.id/", "https://www.esdm.go.id/id/data-statistik"):
        probe("guess", u, show=300, links=40)
    # BPS
    for u in ("https://www.bps.go.id/id/statistics-table/2/ODU5IzI=/listrik-yang-didistribusikan-menurut-provinsi--gwh-.html",
              "https://www.bps.go.id/id/statistics-table/2/MzIxIzI=/kapasitas-terpasangpln-menurut-jenis-pembangkit-listrik.html",
              "https://webapi.bps.go.id/v1/api/list/model/var/lang/ind/domain/0000/subject/7/"):
        probe("BPS", u, show=600, links=40)
    # Satu Data CKAN
    for u in ("https://katalog.data.go.id/api/3/action/package_search?q=produksi%20listrik&rows=20",
              "https://data.go.id/api/3/action/package_search?q=produksi%20listrik&rows=20",
              "https://katalog.satudata.go.id/api/3/action/package_search?q=produksi%20listrik%20pln&rows=20"):
        probe("CKAN", u, show=3000, links=0, kw=False)


if __name__ == "__main__":
    main()
