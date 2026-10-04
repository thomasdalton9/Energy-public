"""
South & Southeast Asia gas discovery, round 3 (GitHub Actions):
  PK  PBS newest Monthly Bulletin (media listing ordered by date, publication / energy-mining pages); table 3.5;
      SSGC monthly capacity-information PDF text
  ID  ESDM highlight page, Ditjen Migas lifting page / Buku Statistik Migas
  MY  DOSM 'download release' PDF for petroleum & natural gas (monthly table?)
  VN  NSO monthly statistical tables workbook (Bieu thang): sheets and natural gas rows; media history depth
  PH  DOE energy statistics page structure (API?)
"""
import io
import re
from urllib.parse import urljoin

import pandas as pd
import requests
import urllib3

urllib3.disable_warnings()
H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) "
                   "Chrome/124.0 Safari/537.36", "Accept-Language": "en,id;q=0.8,vi;q=0.6"}
T = (15, 120)
LINK = re.compile(r"""(?:href|src|data-src|data-url)=["']([^"']+)["']""", re.I)
S = requests.Session()
S.headers.update(H)


def out(s=""):
    print(s, flush=True)


def get(u, **kw):
    try:
        r = S.get(u, timeout=T, verify=False, **kw)
        out(f"\n### {u}\n  -> {r.status_code} {r.headers.get('content-type')} {len(r.content)}b final={r.url} "
            f"LM={r.headers.get('last-modified')}")
        return r
    except Exception as e:  # noqa: BLE001
        out(f"\n### {u}\n  ERROR {type(e).__name__}: {str(e)[:200]}")
        return None


def plain(html):
    t = re.sub(r"<script.*?</script>|<style.*?</style>", " ", html, flags=re.S | re.I)
    return re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", t))


def links(r, pat, n=80):
    ls = []
    for l in LINK.findall(r.text):
        if re.search(pat, l, re.I):
            a = urljoin(r.url, l)
            if a not in ls:
                ls.append(a)
    out(f"  links /{pat}/ ({len(ls)}):")
    for l in ls[:n]:
        out(f"    {l}")
    return ls


def ctx(text, pat, n=6, w=300):
    for i, m in enumerate(re.finditer(pat, text, re.I)):
        if i >= n:
            break
        out(f"  CTX: ...{text[max(0, m.start() - w):m.end() + w]}...")


def pdf_text(content, pat, maxp=6, chars=3500, pages=None):
    import pdfplumber
    with pdfplumber.open(io.BytesIO(content)) as pdf:
        out(f"  pdf pages: {len(pdf.pages)}")
        hits = 0
        for i, p in enumerate(pdf.pages):
            if pages and i not in pages:
                continue
            t = p.extract_text() or ""
            if pat is None or re.search(pat, t, re.I):
                out(f"  ---- page {i + 1} ----\n{t[:chars]}")
                hits += 1
                if hits >= maxp:
                    break


def wp_media(base, q, n=100, page=1):
    r = get(f"{base}/wp-json/wp/v2/media?search={q}&per_page={n}&page={page}&orderby=date&order=desc&_fields=date,source_url")
    res = []
    if r is not None and r.status_code == 200:
        try:
            res = [(m["date"], m["source_url"]) for m in r.json()]
        except ValueError:
            pass
        out(f"  total={r.headers.get('X-WP-Total')} pages={r.headers.get('X-WP-TotalPages')}")
        for d, u in res:
            out(f"    {d} {u}")
    return res


def pakistan():
    for q in ("MBS_", "MBS-", "Bulletin_of_Statistics", "Monthly-Bulletin", "Bulletin 2025", "MBS 2025", "MBS 2026"):
        wp_media("https://www.pbs.gov.pk", q, 30)
    for u in ("https://www.pbs.gov.pk/publication-2/", "https://www.pbs.gov.pk/energy-mining-2/"):
        r = get(u)
        if r is not None:
            ctx(plain(r.text), r"bulletin|natural gas", n=8, w=200)
            links(r, r"\.pdf|\.xls|bulletin|mbs|energy|gas", n=60)
    r = get("https://www.pbs.gov.pk/wp-content/uploads/2020/07/MBS_Sep_2024.pdf")
    if r is not None and r.status_code == 200:
        pdf_text(r.content, r"3\.5 Production of Crude Oil, Natural Gas|3\.3 Company Wise", maxp=3)
    r = get("https://www.ssgc.com.pk/web/uploads/capacityInfo/2026_07_cap_info_dist_trans.pdf")
    if r is not None and r.status_code == 200 and r.content[:4] == b"%PDF":
        pdf_text(r.content, None, maxp=4, chars=4000)
    r = get("https://www.ssgc.com.pk/web/capacity-information-v2-gas-transmission/")
    if r is not None:
        i = r.text.find("capacityInfo")
        out("  RAW: " + r.text[max(0, i - 3000):i + 1500])


def indonesia():
    for u in ("https://www.esdm.go.id/id/highlight", "https://migas.esdm.go.id/post/buku-statistik-migas"):
        r = get(u)
        if r is not None and r.status_code == 200:
            t = plain(r.text)
            i = t.find("Gas")
            out("  TEXT: " + t[:200] + " ... " + t[max(0, i - 500): i + 3000])
            links(r, r"\.pdf|\.xls|\.csv|api|lifting|produksi|statist|json", n=60)
            out("  ajax-ish: " + str(sorted(set(re.findall(r"""["']([^"'\s]*(?:api|ajax|json|chart|data)[^"'\s]*)["']""",
                                                           r.text, re.I)))[:60]))
    r = get("https://migas.esdm.go.id/page/harga-minyak-mentah-dan-lifting-migas")
    if r is not None:
        t = plain(r.text)
        i = t.find("Lifting", 1500)
        out("  TEXT: " + t[1400:6000])
        ctx(t, r"MMSCFD|gas bumi", n=6)


def malaysia():
    r = get("https://www.dosm.gov.my/site/downloadrelease?id=mining-of-petroleum-and-natural-gas-statistics-q22026&lang=English&admin_view=")
    if r is not None and r.status_code == 200:
        if r.content[:4] == b"%PDF":
            pdf_text(r.content, r"natural gas|month|Jan", maxp=8, chars=3000)
        else:
            out("  " + r.text[:1500])
            links(r, r"\.pdf|\.xls|download", n=30)


def vietnam():
    hist = []
    for p in (1, 2, 3):
        hist += wp_media("https://www.nso.gov.vn", "Bieu-thang", 100, p)
    for q in ("Bieu thang", "Bieu", "So lieu thang"):
        wp_media("https://www.nso.gov.vn", q, 20)
    r = get("https://www.nso.gov.vn/wp-content/uploads/2026/10/02.-Bieu-thang-9.2026.xlsx")
    if r is not None and r.status_code == 200:
        x = pd.read_excel(io.BytesIO(r.content), sheet_name=None, header=None)
        out(f"  sheets: {list(x)}")
        for name, df in x.items():
            m = df.apply(lambda c: c.astype(str).str.contains("Khí|khí đốt|Dầu thô", regex=True)).any(axis=1)
            if m.any():
                out(f"  --- sheet {name} shape {df.shape}; header rows:\n{df.head(8).to_string()[:2500]}")
                out(df[m].to_string()[:2500])


def philippines():
    r = get("https://doe.gov.ph/data-and-prices/energy-statistics")
    if r is not None:
        t = r.text
        i = t.find("Official Energy Statistics")
        out("  scripts: " + str(re.findall(r'<script[^>]+src="([^"]+)"', t)[:40]))
        out("  ajax-ish: " + str(sorted(set(re.findall(r"""["']([^"'\s]*(?:api|ajax|json|views|sites/default/files)[^"'\s]*)["']""",
                                                       t, re.I)))[:80]))
        out("  TEXT: " + plain(t)[600:5000])
        links(r, r"data-and-prices|statist|natural|gas|upstream|downstream", n=80)
    for u in ("https://doe.gov.ph/data-and-prices/energy-statistics/natural-gas",
              "https://doe.gov.ph/data-and-prices/energy-statistics/upstream-natural-gas",
              "https://doe.gov.ph/data-and-prices/natural-gas", "https://doe.gov.ph/energy-resources/natural-gas",
              "https://doe.gov.ph/data-and-prices/energy-statistics/electric-power-industry"):
        r = get(u)
        if r is not None and r.status_code == 200:
            out("  TEXT: " + plain(r.text)[600:2500])
            links(r, r"\.pdf|\.xls|natural|gas|statist", n=60)


if __name__ == "__main__":
    for f in (pakistan, indonesia, malaysia, vietnam, philippines):
        out(f"\n\n{'#' * 30} {f.__name__} {'#' * 30}")
        try:
            f()
        except Exception as e:  # noqa: BLE001
            out(f"!! {f.__name__}: {type(e).__name__}: {e}")
