"""Probe 2: ldc.nea.org.np (NEA Load Dispatch Center, Laravel site): links, script.js, report pages."""
import io
import re
from urllib.parse import urljoin
import requests
import urllib3
urllib3.disable_warnings()
H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0 Safari/537.36"}
S = requests.Session()
S.headers.update(H)
BASE = "https://ldc.nea.org.np/"


def g(url, **kw):
    try:
        r = S.get(url, timeout=(15, 60), verify=False, **kw)
        print(f"{r.status_code} {len(r.content):>8} {r.headers.get('content-type','')[:30]:30} {url}", flush=True)
        return r
    except Exception as e:  # noqa
        print(f"ERR {type(e).__name__}: {str(e)[:120]} {url}", flush=True)
        return None


def links(r):
    return sorted(set(urljoin(r.url, h) for h in re.findall(r'href=["\']([^"\'#]+)["\']', r.text)))


r = g(BASE)
home = links(r)
for u in home:
    print("  L", u)
print(re.sub(r"\s+", " ", re.sub(r"<script.*?</script>|<style.*?</style>|<[^>]+>", " ", r.text, flags=re.S))[:3000])
js = g(BASE + "frontend/script.js")
print(js.text[:6200])
pdf = None
seen = set()
for u in home:
    if "ldc.nea.org.np" not in u or u in seen or re.search(r"\.(css|png|jpg|ico|js)$", u):
        continue
    seen.add(u)
    p = g(u)
    if p is None or not p.ok:
        continue
    if p.content[:4] == b"%PDF":
        pdf = pdf or p
        continue
    sub = links(p)
    interesting = [x for x in sub if re.search(r"pdf|report|ndor|daily|download|storage|api|page=", x, re.I)]
    for x in interesting[:60]:
        print("    ->", x)
    t = re.sub(r"\s+", " ", re.sub(r"<script.*?</script>|<style.*?</style>|<[^>]+>", " ", p.text, flags=re.S))
    print("    TEXT:", t[:1200])
    if pdf is None:
        for x in interesting:
            if x.lower().endswith(".pdf") or "storage" in x:
                q = g(x)
                if q is not None and q.ok and q.content[:4] == b"%PDF":
                    pdf = q
                    break
if pdf is not None:
    import pdfplumber
    print("PDF", pdf.url)
    with pdfplumber.open(io.BytesIO(pdf.content)) as d:
        print("pages", len(d.pages))
        for pg in d.pages[:2]:
            print("-----"); print(pg.extract_text())
