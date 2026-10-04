"""
Probe 9: (1) Gaz-System KspRealization (actual gas transmitted) per zone: zones list, one day, history depth;
(2) Transgaz masuratori fragment source (forms/scripts); (3) Enagas bulletin file list; (4) Gasgrid sitemap;
(5) Plinacro transparency table rows. Prints only.
"""
import json
import re
import sys

import requests

UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0 Safari/537.36"
H = {"User-Agent": UA}
GS = "https://swi.gaz-system.pl/mir/api/v1/en/KspRealization/"


def gazsystem():
    print("=== GAZ-SYSTEM")
    try:
        r = requests.get(GS + "init", headers=H, timeout=60)
        print("init", r.status_code, len(r.text))
        j = r.json()
        for f in j.get("filters", []):
            print(" filter", f.get("filterName"), [(d["value"], d.get("subtext")) for d in f.get("data", [])][:80])
    except Exception as e:  # noqa: BLE001
        print("init fail", type(e).__name__, str(e)[:200])

    def q(day, length=500, to=None):
        flt = [{"filterName": "gasDay$from", "values": day}]
        if to:
            flt.append({"filterName": "gasDay$to", "values": to})
        body = {"start": 0, "length": length, "globalFilter": {"value": ""}, "lang": "en",
                "sort": [{"column": "gasDay", "order": "asc"}, {"column": "zoneCode", "order": "asc"}],
                "customFilters": {"filters": flt}}
        r = requests.post(GS + "data", json=body, headers=H, timeout=90)
        return r

    for day in ("2025-01-15", "2021-01-15", "2018-01-15"):
        try:
            r = q(day, 200, day)
            print(f"data {day}: {r.status_code} len={len(r.text)}")
            j = r.json()
            print(" keys", list(j.keys()), "total", j.get("recordsTotal") or j.get("total"))
            rows = j.get("data", [])
            tot = 0
            for x in rows:
                print("  ", x.get("gasDay"), x.get("zoneCode"), x.get("zoneName"), x.get("direction"), x.get("gasType"), x.get("realization"), x.get("status"))
        except Exception as e:  # noqa: BLE001
            print(f"data {day} fail", type(e).__name__, str(e)[:300])


def transgaz():
    print("=== TRANSGAZ fragment")
    r = requests.get("https://www.transgaz.ro/new-tabel-transparenta-masuratori_en.php?poz=197", headers={**H, "Referer": "https://www.transgaz.ro/en/clients/operational-data/physical-flows"}, timeout=60)
    t = r.text
    for m in re.finditer(r"<script[^>]*>(.*?)</script>", t, re.S):
        s = m.group(1).strip()
        if s and "_gaq" not in s[:80]:
            print("SCRIPT:", re.sub(r"\s+", " ", s)[:2500])
    for m in re.finditer(r"<(form|select|input|option)[^>]*>", t):
        print("  ", m.group(0)[:200])
    for m in re.finditer(r'href="([^"]+)"', t):
        if not m.group(1).endswith(".css"):
            print("   href", m.group(1)[:200])
    body = re.sub(r"<script.*?</script>|<style.*?</style>", "", t, flags=re.S)
    print("TEXT:", re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", body))[:1500])


def enagas():
    print("=== ENAGAS bulletin fragment")
    u = "https://www.enagas.es/content/enagas/en/gestion-tecnica-sistema/energy-data/publicaciones/boletin-estadistico-gas/_jcr_content/responsiveGrid/container/filedownloadpaginati.nocache.html/enagas/components/content/filedownloadpagination"
    r = requests.get("https://www.enagas.es/en/technical-management-system/energy-data/publications/gas-statistical-bulletin/", headers=H, timeout=60)
    print("page", r.status_code)
    ms = re.findall(r'["\'(]([^"\'()\s]*filedownloadpaginati[^"\'()\s]*)', r.text)
    print("refs", ms[:5])
    for m in re.finditer(r'<a[^>]+href="([^"]+)"[^>]*>(.{0,120}?)</a>', r.text, re.S):
        h = m.group(1)
        if re.search(r"/dam/|\.pdf|\.xls|download|boletin|bulletin", h, re.I):
            print("   page-link", h[:200], "|", re.sub(r"\s+", " ", re.sub(r"<[^>]+>", "", m.group(2)))[:80])
    frag = "https://www.enagas.es/content/enagas/en/gestion-tecnica-sistema/energy-data/publicaciones/boletin-estadistico-gas/_jcr_content/responsiveGrid/container/filedownloadpaginati.nocache.html/enagas/components/content/filedownloadpagination"
    for u2 in ms[:2] + [frag]:
        uu = u2 if u2.startswith("http") else "https://www.enagas.es" + u2
        try:
            r2 = requests.get(uu, headers=H, timeout=60)
            print("frag", r2.status_code, len(r2.text), uu[:200])
            for m in re.finditer(r'<a[^>]+href="([^"]+)"[^>]*>(.{0,200}?)</a>', r2.text, re.S):
                print("   frag-link", m.group(1)[:200], "|", re.sub(r"\s+", " ", re.sub(r"<[^>]+>", "", m.group(2)))[:80])
            print(" total pages hints:", re.findall(r'data-[a-z-]+="[^"]{0,80}"', r2.text)[:15])
        except Exception as e:  # noqa: BLE001
            print("frag fail", type(e).__name__)


def gasgrid():
    print("=== GASGRID sitemap")
    for u in ("https://gasgrid.fi/wp-sitemap.xml", "https://gasgrid.fi/sitemap_index.xml", "https://gasgrid.fi/sitemap.xml"):
        try:
            r = requests.get(u, headers=H, timeout=40)
            print(u, r.status_code, len(r.text))
            if r.status_code == 200:
                subs = re.findall(r"<loc>([^<]+)</loc>", r.text)
                print(subs[:30])
                for s in subs[:12]:
                    rr = requests.get(s, headers=H, timeout=40)
                    for l in re.findall(r"<loc>([^<]+)</loc>", rr.text):
                        if re.search(r"data|consum|market|statistic|flow|transparen|volume|publication|report", l, re.I):
                            print("   ", l)
                break
        except Exception as e:  # noqa: BLE001
            print(u, type(e).__name__)


def plinacro():
    print("=== PLINACRO table")
    r = requests.get("https://www.plinacro.hr/default.aspx?id=109", headers=H, timeout=60)
    for m in re.finditer(r"<tr[^>]*>(.*?)</tr>", r.text, re.S):
        row = m.group(1)
        txt = re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", row)).strip()
        hs = re.findall(r'href="([^"]+)"', row)
        print("  ", txt[:110], "|", hs[:2])


for fn in (gazsystem, transgaz, enagas, gasgrid, plinacro):
    try:
        fn()
    except Exception as e:  # noqa: BLE001
        print(fn.__name__, "FAILED", type(e).__name__, str(e)[:300])
sys.exit(0)
