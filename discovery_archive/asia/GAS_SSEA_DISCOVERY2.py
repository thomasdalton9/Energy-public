"""
South & Southeast Asia gas discovery, round 2 (GitHub Actions). Follows round 1:
  PK  PBS WordPress media search for Monthly Bulletin of Statistics PDFs + natural gas table text from one;
      SSGC / SNGPL capacity-information pages (Pakistan Gas Network Code); Petroleum Division downloads
  ID  ESDM homepage 'Highlight' card source (API?), datamigas.esdm.go.id, Ditjen Migas lifting page
  MY  DOSM petroleum & natural gas release download links; MEIH natural gas production/consumption pages
  VN  NSO monthly report page attachments and the natural gas line
  PH  DOE energy statistics page, natural gas links
"""
import io
import re
from urllib.parse import urljoin

import requests
import urllib3

urllib3.disable_warnings()
H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) "
                   "Chrome/124.0 Safari/537.36", "Accept-Language": "en,id;q=0.8,vi;q=0.6"}
T = (15, 90)
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


def pdf_pages(content, pat, maxp=6):
    import pdfplumber
    with pdfplumber.open(io.BytesIO(content)) as pdf:
        out(f"  pdf pages: {len(pdf.pages)}")
        hits = 0
        for i, p in enumerate(pdf.pages):
            t = p.extract_text() or ""
            if re.search(pat, t, re.I):
                out(f"  ---- page {i + 1} ----\n{t[:3500]}")
                hits += 1
                if hits >= maxp:
                    break


def pakistan():
    for q in ("MBS", "Monthly Bulletin", "bulletin"):
        r = get(f"https://www.pbs.gov.pk/wp-json/wp/v2/media?search={q}&per_page=100&_fields=date,source_url")
        if r is not None and r.status_code == 200:
            try:
                for m in r.json()[:100]:
                    out(f"    {m.get('date')} {m.get('source_url')}")
            except ValueError:
                out(r.text[:500])
    r = get("https://www.pbs.gov.pk/")
    if r is not None:
        links(r, r"bulletin|mbs|energy|publication")
    for q in ("monthly-bulletin-of-statistics", "monthly bulletin"):
        r = get(f"https://www.pbs.gov.pk/wp-json/wp/v2/pages?search={q}&_fields=link,title")
        if r is not None:
            out("  " + r.text[:1500])
    r = get("https://www.pbs.gov.pk/wp-content/uploads/2020/07/MBS_Sep_2024.pdf")
    if r is not None and r.status_code == 200:
        pdf_pages(r.content, r"natural gas|LNG", maxp=6)
    for u in ("https://www.ssgc.com.pk/web/capacity-information-v2-gas-transmission/",
              "https://www.ssgc.com.pk/web/gas-infrastructure/"):
        r = get(u)
        if r is not None:
            ctx(plain(r.text), r"MMCFD|MMSCFD|capacity|flow", n=8)
            links(r, r"\.pdf|\.xls|capacity|daily|uploads")
    r = get("https://www.sngpl.com.pk/")
    if r is not None:
        links(r, r"capacity|network|code|transmission|daily|supply|statist|annual", n=60)
    for u in ("https://www.sngpl.com.pk/web/capacity-information", "https://www.sngpl.com.pk/capacity-information.jsp",
              "https://www.sngpl.com.pk/web/networkcode"):
        get(u)
    r = get("https://petroleum.gov.pk/Download")
    if r is not None:
        ctx(plain(r.text), r"year ?book|gas|statist", n=10, w=200)
        links(r, r"\.pdf|\.xls|download|year", n=80)


def indonesia():
    r = get("https://www.esdm.go.id/id")
    if r is not None:
        i = r.text.find("Produksi Gas")
        out("  RAW around 'Produksi Gas':\n" + r.text[max(0, i - 2500):i + 1500])
        out("  scripts: " + str(re.findall(r'<script[^>]+src="([^"]+)"', r.text)[:40]))
        out("  ajax-ish: " + str(sorted(set(re.findall(r"""["'](/?[\w/.-]*(?:api|ajax|highlight|json)[\w/.?=&-]*)["']""",
                                                       r.text, re.I)))[:40]))
    for u in ("https://datamigas.esdm.go.id", "https://datamigas.esdm.go.id/", "https://migas.esdm.go.id/page/harga-minyak-mentah-dan-lifting-migas",
              "https://www.esdm.go.id/id/download", "https://migas.esdm.go.id/post/index"):
        r = get(u)
        if r is not None and r.status_code == 200:
            out("  TEXT: " + plain(r.text)[:1500])
            links(r, r"\.pdf|\.xls|\.csv|api|lifting|produksi|gas|statist|data", n=60)


def malaysia():
    r = get("https://www.dosm.gov.my/portal-main/release-content/mining-of-petroleum-and-natural-gas-statistics-q22026")
    if r is not None:
        links(r, r"download|\.pdf|\.xls|uploads|release-content/file", n=60)
        ctx(r.text, r"Download release", n=3, w=600)
        ctx(plain(r.text), r"month|monthly|bulan", n=6)
    r = get("https://meih.st.gov.my/statistics")
    if r is not None:
        links(r, r"natural|gas|production|statistic", n=80)
    for u in ("https://open.dosm.gov.my/data-catalogue/ipi", "https://api.data.gov.my/data-catalogue?id=ipi&limit=5"):
        r = get(u)
        if r is not None:
            out("  " + r.text[:800])


def vietnam():
    r = get("https://www.nso.gov.vn/tin-tuc-thong-ke/2026/09/mot-so-net-chinh-tinh-hinh-kinh-te-xa-hoi-thang-tam-va-8-thang-nam-2026/")
    if r is not None:
        links(r, r"wp-content/uploads|\.xlsx?|\.pdf|\.zip|\.docx?", n=40)
        ctx(plain(r.text), r"khí đốt|khí thiên nhiên|triệu m3|m3", n=6)
    r = get("https://www.nso.gov.vn/bao-cao-tinh-hinh-kinh-te-xa-hoi-hang-thang/")
    if r is not None:
        links(r, r"tin-tuc-thong-ke/20|bao-cao|\?p=|du-lieu|so-lieu", n=40)
    r = get("https://www.nso.gov.vn/so-lieu-thong-ke/")
    if r is not None:
        ctx(plain(r.text), r"khí đốt|Sản phẩm công nghiệp|sản phẩm chủ yếu", n=6, w=200)
        links(r, r"cong-nghiep|san-pham|industr", n=40)
    r = get("https://www.nso.gov.vn/wp-json/wp/v2/media?search=xlsx&per_page=50&_fields=date,source_url")
    if r is not None:
        out("  " + r.text[:3000])


def philippines():
    r = get("https://doe.gov.ph/data-and-prices/energy-statistics")
    if r is not None:
        out("  TEXT: " + plain(r.text)[:3000])
        links(r, r"gas|statist|\.pdf|\.xls|energy-statistics", n=100)
    for u in ("https://doe.gov.ph/data-and-prices/energy-statistics?q=natural+gas", "https://doe.gov.ph/search?q=natural+gas"):
        r = get(u)
        if r is not None:
            links(r, r"gas", n=40)


if __name__ == "__main__":
    for f in (pakistan, indonesia, malaysia, vietnam, philippines):
        out(f"\n\n{'#' * 30} {f.__name__} {'#' * 30}")
        try:
            f()
        except Exception as e:  # noqa: BLE001
            out(f"!! {f.__name__}: {type(e).__name__}: {e}")
