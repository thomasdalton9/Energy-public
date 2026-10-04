"""
Grid-operator data discovery, round 2 (after SCADA_DISCOVERY1.py): open the leads.
  Vietnam   EVN 'Thong tin tom tat van hanh HTD Quoc gia' (daily operation summary posts): list + one post's text;
            EVN 'Muc nuoc cac ho thuy dien' (hydro reservoir levels): table / data endpoint
  Pakistan  WAPDA river-flow-data (Tarbela / Mangla levels, inflows): table text; CPPA xwdiscos energy purchase data,
            service-operation; NEPRA FCA 'Plant wise Units.xlsx' (monthly generation by plant)
  Sri Lanka CEB plant-operation page (telemetered generation); PUCSL daily reports
  Philippines DOE 2025 power statistics page (Next.js: data in the page payload?)
  Nepal     NEA site (JS app): script bundles -> api endpoints
  India     MERIT / Grid-India once more with a browser-like session (were reset in round 1)
"""
import io
import re

import pandas as pd
import requests
import urllib3

urllib3.disable_warnings()
H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 "
                   "Safari/537.36", "Accept": "text/html,application/xhtml+xml,application/json,*/*;q=0.8",
     "Accept-Language": "en-US,en;q=0.9"}
T = (15, 60)


def out(*a):
    print(*a, flush=True)


def get(u, **kw):
    try:
        r = requests.get(u, headers=H, timeout=T, verify=False, **kw)
        out(f"GET {u} -> {r.status_code} {r.headers.get('content-type', '')[:40]} {len(r.content)}")
        return r
    except Exception as e:  # noqa: BLE001
        out(f"GET {u} ERROR {type(e).__name__}: {str(e)[:120]}")
        return None


def text_of(html, n=3000):
    t = re.sub(r"<script.*?</script>|<style.*?</style>", " ", html, flags=re.S | re.I)
    return re.sub(r"<[^>]+>|\s+", " ", t)[:n]


def tables(html, k=4, rows=15):
    try:
        ts = pd.read_html(io.StringIO(html))
    except Exception as e:  # noqa: BLE001
        out(f"    read_html: {e!r}")
        return
    out(f"    {len(ts)} tables")
    for t in ts[:k]:
        out(f"    table {t.shape}: " + t.head(rows).to_string(max_cols=14, max_colwidth=18)[:2500])


def links(html, pat):
    return sorted(set(l for l in re.findall(r'(?:href|src)\s*=\s*["\']([^"\'#]+)', html) if re.search(pat, l, re.I)))


def vietnam():
    r = get("https://www.evn.com.vn/vi-VN/news-l/Thong-tin-tom-tat-van-hanh-HTD-Quoc-gia-60-2015")
    if r is not None and r.ok:
        posts = links(r.text, r"/d/vi-VN/news/")
        out(f"  {len(posts)} posts: {posts[:15]}")
        out("  page text: " + text_of(r.text, 1500))
        for p in posts[:2]:
            x = get(requests.compat.urljoin(r.url, p))
            if x is not None and x.ok:
                body = text_of(x.text, 20000)
                i = max(0, body.find("sản lượng") - 300)
                out("    post: " + body[i:i + 3000])
                tables(x.text, 3)
                imgs = links(x.text, r"userfile.*\.(png|jpe?g)")
                out(f"    images: {imgs[:5]}")
    r = get("https://www.evn.com.vn/vi-VN/thong-tin-ho-thuy-dien/Muc-nuoc-cac-ho-thuy-dien-60-123")
    if r is not None and r.ok:
        out("  text: " + text_of(r.text, 2500))
        tables(r.text, 3, 30)
        out(f"  scripts/ajax: {links(r.text, r'ajax|api|ashx|asmx|json|Handler')[:20]}")
        for m in re.finditer(r"(\$\.(?:ajax|get|post)|fetch\()\s*\(?\s*[{'\"][^;]{0,300}", r.text):
            out("    js: " + m.group(0)[:300])


def pakistan():
    r = get("https://wapda.gov.pk/river-flow-data/")
    if r is not None and r.ok:
        i = r.text.find("Tarbela")
        out("  text near Tarbela: " + text_of(r.text[max(0, i - 3000):i + 6000], 3500))
        tables(r.text, 4, 20)
        frames = re.findall(r"<iframe[^>]+src=[\"']([^\"']+)", r.text)
        out(f"  iframes: {frames[:5]}")
    for u in ("https://cppa.gov.pk/xwdiscos-energy-purchase-data", "https://cppa.gov.pk/service-operation",
              "https://cppa.gov.pk/policies-reports"):
        x = get(u)
        if x is not None and x.ok:
            files = links(x.text, r"\.(pdf|xlsx?|csv)")
            out(f"    files: {files[:30]}")
            out("    text: " + text_of(x.text, 1200))
    x = get("https://nepra.org.pk/Admission%20Notices/2026/09%20Sep/01-%20Plant%20wise%20Units.xlsx")
    if x is not None and x.ok:
        xl = pd.ExcelFile(io.BytesIO(x.content))
        out(f"    sheets {xl.sheet_names}")
        for sh in xl.sheet_names[:2]:
            out(xl.parse(sh, header=None, nrows=45).to_string(max_cols=12, max_colwidth=22)[:4500])


def sri_lanka():
    r = get("https://www.ceb.lk/plant-operation/en")
    if r is not None and r.ok:
        i = r.text.find("Telemetered")
        out("  text: " + text_of(r.text[max(0, i - 6000):i + 2000], 3500))
        tables(r.text, 4, 30)
        dl = links(r.text, r"ajax|api|json|\.php|\.xlsx?|\.csv")
        out(f"  data links: {dl[:20]}")
        for m in re.finditer(r"(\$\.(?:ajax|get|post|getJSON)|fetch\(|url\s*:)\s*[^;]{0,250}", r.text):
            out("    js: " + m.group(0)[:250])
    r = get("https://www.pucsl.gov.lk/daily-reports/")
    if r is not None and r.ok:
        files = links(r.text, r"\.(pdf|xlsx?|csv)")
        out(f"  files: {files[:20]}")
        out("  text: " + text_of(r.text, 1500))


def philippines():
    r = get("https://doe.gov.ph/data-and-prices/energy-statistics/electric-power-industry/2025-power-statistics")
    if r is not None and r.ok:
        docs = sorted(set(re.findall(r'https?://prod-cms\.doe\.gov\.ph/[^"\'\s\\]+', r.text)))
        out(f"  prod-cms docs: {docs[:40]}")
        out("  text: " + text_of(r.text, 2000))
        for d in docs[:2]:
            x = get(d)
            if x is not None and x.ok and x.content[:4] == b"%PDF":
                import pdfplumber
                with pdfplumber.open(io.BytesIO(x.content)) as pdf:
                    out("    PDF: " + (pdf.pages[0].extract_text() or "")[:2000].replace("\n", " | "))


def nepal():
    r = get("https://www.nea.org.np/")
    if r is None:
        return
    js = links(r.text, r"\.js($|\?)")
    out(f"  scripts: {js}")
    for j in js[:6]:
        x = get(requests.compat.urljoin(r.url, j))
        if x is None or not x.ok:
            continue
        for m in sorted(set(re.findall(r'["\'`](https?://[^"\'`\s]{6,120}|/api/[^"\'`\s]{1,100})["\'`]', x.text))):
            if re.search(r"api|cms|ldc|report|upload", m, re.I):
                out(f"    {m}")


def india():
    s = requests.Session()
    s.headers.update(H)
    for u in ("https://meritindia.in/", "https://www.meritindia.in/", "https://grid-india.in/en/",
              "https://webapi.grid-india.in/api/v1/file", "https://posoco.in/"):
        try:
            r = s.get(u, timeout=T, verify=False)
            out(f"  {u} -> {r.status_code} {len(r.content)}: " + text_of(r.text, 600))
        except Exception as e:  # noqa: BLE001
            out(f"  {u}: {type(e).__name__}: {str(e)[:120]}")


if __name__ == "__main__":
    import sys
    for f in sys.argv[1:] or ["vietnam", "pakistan", "sri_lanka", "philippines", "nepal", "india"]:
        out(f"\n==================== {f}")
        try:
            globals()[f]()
        except Exception as e:  # noqa: BLE001
            out(f"!! {f}: {e!r}")
