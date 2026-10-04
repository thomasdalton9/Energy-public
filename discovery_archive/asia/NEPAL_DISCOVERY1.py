"""Probe 1: NEA Load Dispatch Centre daily operation reports (NDOR PDFs). Tries URL forms, the site's JS bundles
for API routes, and the ldc subdomain."""
import io
import re
import requests
import urllib3
urllib3.disable_warnings()
H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0 Safari/537.36"}


def g(url, n=300, **kw):
    try:
        r = requests.get(url, headers=H, timeout=(15, 60), verify=False, **kw)
        print(f"{r.status_code} {len(r.content):>8} {r.headers.get('content-type','')[:40]:40} {url}", flush=True)
        if n and "pdf" not in r.headers.get("content-type", ""):
            print("   ", re.sub(r"\s+", " ", r.text[:n]))
        return r
    except Exception as e:  # noqa
        print(f"ERR {type(e).__name__}: {str(e)[:120]} {url}", flush=True)
        return None


def pdftext(r):
    import pdfplumber
    with pdfplumber.open(io.BytesIO(r.content)) as pdf:
        print("   pages", len(pdf.pages))
        for p in pdf.pages[:3]:
            print("----- page"); print(p.extract_text())
            for t in p.extract_tables():
                print("TABLE", t[:40])


got = None
for host in ("https://nea.org.np", "https://www.nea.org.np", "https://cms.nea.org.np"):
    for name in ("NDOR 2081_05_30.pdf", "NDOR 2080_03_15.pdf"):
        for sep in ("\\", "%5C", "/"):
            path = sep.join(["admin", "assets", "uploads", "ldc", name.replace(" ", "%20")])
            r = g(f"{host}/{path}", 0)
            if r is not None and r.ok and r.content[:4] == b"%PDF" and got is None:
                got = r
if got:
    pdftext(got)

# site and bundles
for u in ("https://www.nea.org.np/", "https://nea.org.np/", "https://ldc.nea.org.np/", "https://cms.nea.org.np/",
          "https://www.nea.org.np/ldc", "https://ldc.nea.org.np/api", "https://nea.org.np/api"):
    r = g(u, 600)
    if r is not None and r.ok:
        for s in sorted(set(re.findall(r'(?:src|href)="([^"]+\.(?:js|css))"', r.text)))[:20]:
            su = s if s.startswith("http") else u.split("/", 3)[0] + "//" + u.split("/")[2] + "/" + s.lstrip("/")
            jr = g(su, 0)
            if jr is not None and jr.ok and su.endswith(".js"):
                hits = set(re.findall(r'["\'`]((?:https?://[^"\'`\s]{3,120})|(?:/?api/[^"\'`\s]{1,100}))["\'`]', jr.text))
                for h in sorted(hits)[:150]:
                    print("    JS:", h)
                for m in re.finditer(r"ldc|NDOR|uploads|daily.?operation", jr.text, re.I):
                    print("    CTX:", jr.text[max(0, m.start()-150):m.end()+150].replace("\n", " "))
                    break
