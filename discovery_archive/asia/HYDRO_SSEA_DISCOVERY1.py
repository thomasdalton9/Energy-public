"""
South & Southeast Asia hydro reservoir discovery, round 1.
  Pakistan  WAPDA river-flow-data page (scripts / ajax / wp-json), PMD FFD monthly dam sheets
            (pmd.gov.pk/FFD/index_files/daily/dams<mon><yy>_files/sheet001.htm), ffd.pmd.gov.pk river-state, IRSA
  Vietnam   EVN reservoir page (raw html, scripts), hochuathuydien.evn.com.vn, DWRM quanly.dwrm.gov.vn/hochua, vndms
  Sri Lanka PUCSL GenData JS bundle -> /api/ paths
"""
import re
from urllib.parse import urljoin

import requests
import urllib3

urllib3.disable_warnings()
H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 "
                   "Safari/537.36", "Accept": "text/html,application/json,*/*;q=0.8", "Accept-Language": "en-US,en;q=0.9"}
T = (15, 60)
s = requests.Session()
s.headers.update(H)


def out(*a):
    print(*a, flush=True)


def get(u, **k):
    try:
        r = s.get(u, timeout=k.pop("timeout", T), verify=False, **k)
        out(f"GET {u} -> {r.status_code} {r.headers.get('content-type')} {len(r.content)}")
        return r
    except Exception as e:  # noqa: BLE001
        out(f"GET {u} ERR {type(e).__name__}: {str(e)[:200]}")
        return None


def text(h, n=3000):
    h = re.sub(r"(?is)<(script|style)[^>]*>.*?</\1>", " ", h)
    return re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", h))[:n]


def scripts(base, h):
    return [urljoin(base, x) for x in re.findall(r'<script[^>]+src=["\']([^"\']+)', h, re.I)]


def apis(js):
    return sorted(set(re.findall(r'["\'`](/?(?:api|wp-json|ajax|Api|API)[/\w\-\.{}$?=&]*)["\'`]', js)))[:200]


def wapda():
    out("\n=========== WAPDA")
    r = get("https://wapda.gov.pk/river-flow-data/")
    if r is not None and r.ok:
        h = r.text
        sc = scripts(r.url, h)
        out("scripts:", [x for x in sc if "wapda" in x][:60])
        for k in ("admin-ajax", "wp-json", "tablepress", "wpdatatable", "ninja", "iframe", "ajaxurl", "fetch(", "$.ajax",
                  "river", "Inflow", "inflow", "TARBELA", "Tarbela", "cusecs", "Cusecs", "data-url", "data-src", "json"):
            for m in list(re.finditer(re.escape(k), h))[:4]:
                out(f"  [{k}] ..." + h[max(0, m.start() - 300):m.start() + 400].replace("\n", " ") + "...")
        i = h.find("elementor-widget-container", h.find("River Flow"))
        main = re.search(r'(?is)<main.*?</main>', h)
        if main:
            out("MAIN html (raw, 6000):", main.group(0)[:6000])
        for u in sc:
            if "wapda.gov.pk" in u and not re.search(r"jquery|elementor/assets|wp-includes", u):
                j = get(u)
                if j is not None and j.ok:
                    for m in re.finditer(r"(river|flow|tarbela|inflow|ajax|fetch|api)", j.text, re.I):
                        out("   js: " + j.text[max(0, m.start() - 150):m.start() + 250].replace("\n", " "))
                        break
    for u in ("https://wapda.gov.pk/wp-json/wp/v2/pages?slug=river-flow-data",
              "https://wapda.gov.pk/wp-json/", "https://wapda.gov.pk/wp-json/wp/v2/types"):
        r = get(u)
        if r is not None and r.ok:
            out(r.text[:4000])


def ffd():
    out("\n=========== PMD FFD")
    for u in ("https://ffd.pmd.gov.pk/river-state", "https://ffd.pmd.gov.pk/", "https://ffd.pmd.gov.pk/dams",
              "https://ffd.pmd.gov.pk/reservoir", "http://ffd.pmd.gov.pk/", "https://www.pmd.gov.pk/FFD/index.html",
              "https://pmd.gov.pk/FFD/index_files/daily/damsnov15_files/sheet001.htm",
              "https://pmd.gov.pk/FFD/index_files/daily/damssep26_files/sheet001.htm",
              "https://pmd.gov.pk/FFD/index_files/daily/damsoct26_files/sheet001.htm",
              "https://pmd.gov.pk/FFD/index_files/daily/damsaug26_files/sheet001.htm",
              "https://pmd.gov.pk/FFD/index_files/daily/damsjan24_files/sheet001.htm",
              "https://pmd.gov.pk/FFD/index_files/daily/dams.htm",
              "https://pmd.gov.pk/FFD/index_files/daily/"):
        r = get(u)
        if r is None:
            continue
        h = r.text
        out("  text:", text(h, 2500))
        out("  scripts:", scripts(r.url, h)[:30])
        out("  links:", sorted(set(re.findall(r'href=["\']([^"\']+)', h)))[:80])
        for m in re.finditer(r"(fetch\(|axios|\$\.(?:ajax|get|getJSON)|/api/[\w/\-]+)", h):
            out("   js: " + h[max(0, m.start() - 100):m.start() + 250].replace("\n", " "))
        for js in scripts(r.url, h):
            if "pmd.gov.pk" in js and not re.search(r"jquery|bootstrap|leaflet", js, re.I):
                j = get(js)
                if j is not None and j.ok:
                    out("   api strings:", apis(j.text))
                    for m in list(re.finditer(r"(fetch\(|axios\.|/api/)", j.text))[:15]:
                        out("    js: " + j.text[max(0, m.start() - 120):m.start() + 200].replace("\n", " "))


def irsa():
    out("\n=========== IRSA")
    for u in ("http://pakirsa.gov.pk/", "https://pakirsa.gov.pk/", "http://pakirsa.gov.pk/WaterSituation.aspx",
              "http://pakirsa.gov.pk/DailyData.aspx"):
        r = get(u)
        if r is not None:
            out("  text:", text(r.text, 1500))
            out("  links:", sorted(set(re.findall(r'href=["\']([^"\']+)', r.text)))[:120])


def evn():
    out("\n=========== EVN / Vietnam")
    r = get("https://www.evn.com.vn/vi-VN/thong-tin-ho-thuy-dien/Muc-nuoc-cac-ho-thuy-dien-60-123")
    if r is not None and r.ok:
        h = r.text
        out("  all scripts:", scripts(r.url, h))
        out("  inline scripts:")
        for sc in re.findall(r"(?is)<script(?![^>]*src)[^>]*>(.*?)</script>", h):
            sc = sc.strip()
            if sc:
                out("   ---", sc[:1500].replace("\n", " "))
        out("  iframes:", re.findall(r'(?is)<iframe[^>]*>', h))
        i = h.find("Mực nước các hồ thủy điện", h.find("Trang chủ"))
        out("  raw html around title:", h[i:i + 5000])
    for u in ("https://hochuathuydien.evn.com.vn/", "http://hochuathuydien.evn.com.vn/",
              "https://hochuathuydien.evn.com.vn/PageHoChuaThuyDienEmbedEVN.aspx",
              "http://hochuathuydien.evn.com.vn/PageHoChuaThuyDienEmbedEVN.aspx",
              "https://quanly.dwrm.gov.vn/hochua", "https://quanly.dwrm.gov.vn/", "http://quanly.dwrm.gov.vn/hochua",
              "https://vndms.dmc.gov.vn/", "https://vndms.dmc.gov.vn/home/hochua", "http://thuyloivietnam.vn/",
              "https://e15.moit.gov.vn/", "https://hochua.moit.gov.vn/", "http://hochua.moit.gov.vn/"):
        r = get(u, timeout=(20, 60))
        if r is not None:
            out("  text:", text(r.text, 1200))
            out("  scripts:", scripts(r.url, r.text)[:20])
            out("  iframes:", re.findall(r'(?is)<iframe[^>]*>', r.text)[:5])
            for m in list(re.finditer(r"(\$\.(?:ajax|get|post|getJSON)|fetch\(|url\s*:|/api/[\w/\-]+)", r.text))[:15]:
                out("   js: " + r.text[max(0, m.start() - 80):m.start() + 250].replace("\n", " "))


def pucsl():
    out("\n=========== PUCSL GenData")
    r = get("https://gendata.pucsl.gov.lk/")
    if r is None:
        return
    sc = scripts(r.url, r.text)
    out("  scripts:", sc)
    for js in sc:
        j = get(js)
        if j is not None and j.ok:
            a = apis(j.text)
            out("   api strings:", a)
            for m in list(re.finditer(r"reservoir|Reservoir|storage|waterLevel|water-level", j.text))[:20]:
                out("    js: " + j.text[max(0, m.start() - 150):m.start() + 200].replace("\n", " "))
            # chunk files
            for c in sorted(set(re.findall(r'["\']((?:static/js|assets|_next/static)[^"\']+\.js)["\']', j.text)))[:40]:
                k = get(urljoin(r.url, c))
                if k is not None and k.ok and re.search(r"reservoir", k.text, re.I):
                    out("   chunk with reservoir:", c, apis(k.text))
                    for m in list(re.finditer(r"reservoir", k.text, re.I))[:10]:
                        out("    js: " + k.text[max(0, m.start() - 200):m.start() + 200].replace("\n", " "))


for f in (wapda, ffd, irsa, evn, pucsl):
    try:
        f()
    except Exception as e:  # noqa: BLE001
        out(f"!! {f.__name__}: {e}")
