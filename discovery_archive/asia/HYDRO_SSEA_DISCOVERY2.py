"""
South & Southeast Asia hydro reservoir discovery, round 2 (after HYDRO_SSEA_DISCOVERY1.py):
  Sri Lanka PUCSL GenData /api/reservoir/storage-rainfall (dateAggregation values, history depth), /api/reservoir/inflow,
            /api/daily-inflow
  Pakistan  IRSA daily water situation PDFs (pakirsa.gov.pk/Doc/DataDD-MM-YYYY.pdf): text, archive depth;
            WAPDA /water-situation/ and /river-flow/ pages
  Vietnam   DWRM hochua app scripts (API), vndms.gov.vn, thuyloivietnam tlvn.js, MOIT daily reservoir news posts,
            hochuathuydien.evn.com.vn with a long timeout
"""
import io
import json
import re
from datetime import date, timedelta
from urllib.parse import urljoin

import pdfplumber
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


def get(u, quiet=False, **k):
    try:
        r = s.get(u, timeout=k.pop("timeout", T), verify=False, **k)
        if not quiet:
            out(f"GET {r.url} -> {r.status_code} {r.headers.get('content-type')} {len(r.content)}")
        return r
    except Exception as e:  # noqa: BLE001
        out(f"GET {u} ERR {type(e).__name__}: {str(e)[:200]}")
        return None


def text(h, n=3000):
    h = re.sub(r"(?is)<(script|style)[^>]*>.*?</\1>", " ", h)
    return re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", h))[:n]


def pucsl():
    out("\n=========== PUCSL")
    A = "https://gendata.pucsl.gov.lk/api/"
    hd = {"Accept": "application/json", "Referer": "https://gendata.pucsl.gov.lk/"}
    f, t = "2026-09-01T00:00:00.000Z", "2026-10-03T00:00:00.000Z"
    for agg in ("day", "daily", "Day", "1day", "15min", "hour", "month", "week"):
        r = get(A + "reservoir/storage-rainfall", params={"dateAggregation": agg, "from": f, "to": t}, headers=hd)
        if r is not None:
            out("   ", r.text[:800])
    for p in ("reservoir/inflow", "daily-inflow"):
        r = get(A + p, params={"from": f, "to": t}, headers=hd)
        if r is not None:
            out("   ", r.text[:1500])
    for y0 in ("2015-01-01", "2018-01-01", "2020-01-01", "2022-01-01", "2023-01-01", "2024-01-01"):
        d = date.fromisoformat(y0)
        r = get(A + "reservoir/storage-rainfall", params={"dateAggregation": "day", "from": f"{d}T00:00:00.000Z",
                                                          "to": f"{d + timedelta(days=10)}T00:00:00.000Z"}, headers=hd)
        if r is not None:
            out("   ", r.text[:500])
    # one long range: how many rows, which reservoirs, earliest date
    r = get(A + "reservoir/storage-rainfall", params={"dateAggregation": "day", "from": "2010-01-01T00:00:00.000Z",
                                                      "to": "2026-10-04T00:00:00.000Z"}, headers=hd, timeout=(15, 180))
    if r is not None and r.ok:
        try:
            j = r.json()
            out(f"    long range rows {len(j)}; first {j[:2]}; last {j[-2:]}")
            out("    reservoirs:", sorted(set(x.get("reservoirName") for x in j)))
            out("    dates:", min(x.get("reportDate") for x in j), max(x.get("reportDate") for x in j))
        except Exception as e:  # noqa: BLE001
            out("    json err", e, r.text[:300])


def irsa():
    out("\n=========== IRSA")
    today = date(2026, 10, 3)
    first = None
    for d in [today, date(2026, 9, 1), date(2026, 6, 15), date(2026, 1, 10), date(2025, 10, 1), date(2025, 4, 1),
              date(2024, 10, 1), date(2023, 7, 1), date(2022, 1, 5), date(2020, 1, 5), date(2018, 6, 1)]:
        for fmt in ("Data{:%d-%m-%Y}.pdf", "Data{:%d-%m-%y}.pdf", "Data{d.day}-{d.month}-{d.year}.pdf"):
            name = fmt.format(d, d=d) if "{d." in fmt else fmt.format(d)
            r = get("http://pakirsa.gov.pk/Doc/" + name)
            if r is not None and r.ok and r.content[:4] == b"%PDF":
                if first is None:
                    first = r.content
                with pdfplumber.open(io.BytesIO(r.content)) as pdf:
                    tx = "\n".join((p.extract_text() or "") for p in pdf.pages)
                out(f"   {name}: {len(pdf.pages)} pages; text:\n{tx[:2500 if d == today else 900]}")
                if d == today:
                    with pdfplumber.open(io.BytesIO(r.content)) as pdf:
                        for tb in pdf.pages[0].extract_tables():
                            out("   TABLE:")
                            for row in tb[:40]:
                                out("    ", row)
                break


def wapda():
    out("\n=========== WAPDA pages")
    for u in ("https://wapda.gov.pk/water-situation/", "https://wapda.gov.pk/river-flow/"):
        r = get(u)
        if r is None or not r.ok:
            continue
        h = r.text
        m = re.search(r'(?is)<main.*?</main>', h)
        body = m.group(0) if m else h
        out("  main text:", text(body, 5000))
        out("  main raw:", body[:4000])
        out("  imgs/pdfs/iframes:", re.findall(r'(?i)(?:src|href|data)=["\']([^"\']+\.(?:png|jpe?g|pdf|xlsx?|htm|aspx)[^"\']*)', body)[:30])
        out("  iframes:", re.findall(r'(?is)<iframe[^>]*>', body))
        for k in ("tablepress", "wpdatatable", "ajax", "fetch(", "data-src", "Tarbela", "TARBELA"):
            for mm in list(re.finditer(re.escape(k), body))[:3]:
                out(f"  [{k}] " + body[max(0, mm.start() - 300):mm.start() + 600].replace("\n", " "))
    for u in ("https://wapda.gov.pk/wp-json/wp/v2/posts?slug=water-situation", "https://wapda.gov.pk/wp-json/wp/v2/posts?slug=river-flow",
              "https://wapda.gov.pk/wp-json/wp/v2/pages?slug=water-situation", "https://wapda.gov.pk/wp-json/wp/v2/pages?slug=river-flow"):
        r = get(u)
        if r is not None and r.ok:
            try:
                for it in r.json():
                    c = it.get("content", {}).get("rendered", "")
                    out(f"   id {it.get('id')} modified {it.get('modified')}: {text(c, 2000)}")
                    out("   raw:", c[:3000])
            except Exception as e:  # noqa: BLE001
                out("  ", e)


def vietnam():
    out("\n=========== Vietnam")
    base = "https://quanly.dwrm.gov.vn/hochua/"
    for js in ("joiner/init.version.js", "joiner/init.head.js", "joiner/init.body.js"):
        r = get(base + js)
        if r is not None and r.ok:
            out("   ", r.text[:1500].replace("\n", " "))
            for u in re.findall(r'["\']([^"\']+\.js[^"\']*)["\']', r.text)[:40]:
                k = get(urljoin(base + js, u))
                if k is not None and k.ok:
                    a = sorted(set(re.findall(r'["\'`]((?:https?://[^"\'`]+)?/?(?:api|Api|API|service|services|rest)/[\w/\-\.{}]+)', k.text)))[:80]
                    if a:
                        out(f"    {u}: {a}")
    for u in ("https://vndms.gov.vn/", "https://vndms.gov.vn/home/hochua", "https://vndms.gov.vn/hochua"):
        r = get(u, timeout=(20, 60))
        if r is not None:
            out("  text:", text(r.text, 800))
            sc = [urljoin(r.url, x) for x in re.findall(r'<script[^>]+src=["\']([^"\']+)', r.text, re.I)]
            out("  scripts:", sc[:30])
            for m in list(re.finditer(r"(\$\.(?:ajax|get|post|getJSON)|fetch\(|url\s*:\s*['\"])", r.text))[:15]:
                out("   js: " + r.text[max(0, m.start() - 80):m.start() + 250].replace("\n", " "))
            for js in sc:
                if "vndms" in js and not re.search(r"jquery|bootstrap|leaflet|lib", js, re.I):
                    k = get(js)
                    if k is not None and k.ok:
                        for m in list(re.finditer(r"(hochua|HoChua|reservoir|ajax|fetch\()", k.text))[:12]:
                            out("    js: " + k.text[max(0, m.start() - 120):m.start() + 220].replace("\n", " "))
    r = get("http://thuyloivietnam.vn/Scripts/Customers/Home/tlvn.js?v=202001140159")
    if r is not None and r.ok:
        out("   tlvn.js:", r.text[:2500])
    r = get("https://moit.gov.vn/tin-tuc/phat-trien-nang-luong")
    if r is not None and r.ok:
        ls = sorted(set(re.findall(r'href=["\']([^"\']*ho-thu[^"\']*)["\']', r.text, re.I)))
        out("  MOIT reservoir posts:", ls[:20])
        if ls:
            p = get(urljoin(r.url, ls[0]))
            if p is not None:
                out("  post text:", text(p.text, 4000))
    for u in ("https://hochuathuydien.evn.com.vn/PageHoChuaThuyDienEmbedEVN.aspx",):
        get(u, timeout=(60, 90))


for f in (pucsl, irsa, wapda, vietnam):
    try:
        f()
    except Exception as e:  # noqa: BLE001
        out(f"!! {f.__name__}: {e}")
