"""Discovery 3: NNPC monthly report PDFs, NUPRC Next.js pages/API, NBS elibrary titles."""
import re, sys, io, json
import requests
UA = {"User-Agent": "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 Chrome/124 Safari/537.36"}
def get(u, **k):
    try:
        return requests.get(u, headers=UA, timeout=60, **k)
    except Exception as e:
        print("   ERR", u, type(e).__name__, str(e)[:80]); return None
def strip(t):
    t = re.sub(r"<(script|style).*?</\1>", " ", t, flags=re.S|re.I)
    return re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", t))

print("=== NNPC insights list")
r = get("https://www.nnpcgroup.com/insights")
slugs = sorted(set(re.findall(r'/insights/([a-z0-9-]*report[a-z0-9-]*)', r.text))) if r else []
print(slugs)
# nuxt payload may hold more posts
for u in ["https://www.nnpcgroup.com/insights?page=2", "https://www.nnpcgroup.com/insights/page/2"]:
    rr = get(u); print(u, rr.status_code if rr else None)
pdfs = []
for s in slugs[:40]:
    rr = get("https://www.nnpcgroup.com/insights/" + s)
    if rr and rr.status_code == 200:
        for m in re.findall(r'https://[^"\' ]+\.pdf', rr.text): pdfs.append((s, m))
for s, m in sorted(set(pdfs)): print("PDF", s, m)
# Try a monthly-report listing in nuxt JSON / strapi
for u in ["https://www.nnpcgroup.com/insights/monthly-reports", "https://www.nnpcgroup.com/investors", "https://www.nnpcgroup.com/investors/reports", "https://www.nnpcgroup.com/investors/monthly-reports", "https://www.nnpcgroup.com/investors/financial-reports"]:
    rr = get(u); print(u, rr.status_code if rr else None, len(rr.content) if rr else 0)
    if rr and rr.status_code == 200:
        for m in sorted(set(re.findall(r'(?:href|src)=["\']([^"\']*(?:pdf|report)[^"\']*)', rr.text, re.I)))[:40]: print("   ", m)
        print("   text:", strip(rr.text)[:600])

print("\n=== NNPC Aug 2026 PDF")
pdf_url = "https://fde-nnpc-web-cms-prod-dwfrd0hraahrbhhg.a02.azurefd.net/uploads/NNPC_Monthly_Report_for_August_2026_7f78b232ea.pdf"
rr = get(pdf_url)
if rr is not None:
    print(rr.status_code, rr.headers.get("content-type"), len(rr.content))
    if rr.status_code == 200:
        open("/tmp/nnpc.pdf", "wb").write(rr.content)
        import subprocess
        subprocess.run(["pip", "install", "-q", "pypdf"], check=False)
        from pypdf import PdfReader
        rd = PdfReader("/tmp/nnpc.pdf")
        print("pages", len(rd.pages))
        for i, p in enumerate(rd.pages[:12]):
            print(f"--- page {i+1}\n", (p.extract_text() or "")[:2500])

print("\n=== NUPRC")
for u in ["https://www.nuprc.gov.ng/development-production", "https://www.nuprc.gov.ng/reports", "https://www.nuprc.gov.ng/publications"]:
    rr = get(u)
    if rr is None: continue
    print("\n--", u, rr.status_code, len(rr.content))
    print(strip(rr.text)[:1500])
    for m in sorted(set(re.findall(r'["\'](/[^"\' ]*(?:pdf|xlsx?|csv|api)[^"\' ]*)', rr.text, re.I)))[:40]: print("   ref", m)
r = get("https://www.nuprc.gov.ng/")
js = sorted(set(re.findall(r'/_next/static/[^"\' ]+\.js', r.text)))
print("js chunks", len(js))
found = set()
for j in js[:30]:
    rj = get("https://www.nuprc.gov.ng" + j)
    if rj is None: continue
    for m in re.findall(r'https?://[a-z0-9./_-]*(?:api|nuprc|cms|strapi|supabase|firebase)[a-z0-9./_?=&-]*', rj.text, re.I): found.add(m)
for m in sorted(found)[:60]: print("   urlref", m)

print("\n=== NBS elibrary titles")
rr = get("https://www.nigerianstat.gov.ng/elibrary")
if rr is not None:
    rows = re.findall(r'<tr[^>]*>(.*?)</tr>', rr.text, re.S)
    print("rows", len(rows))
    for row in rows:
        t = strip(row)
        if re.search(r'oil|gas|petrol|crude|energy|fuel|mining', t, re.I) and not re.search(r'Price Watch', t):
            href = re.findall(r'href=["\']([^"\']+)', row)
            print("  ", t[:150], href[:2])
