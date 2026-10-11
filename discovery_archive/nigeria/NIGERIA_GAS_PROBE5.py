"""Discovery 5: NNPC monthly report slugs/PDFs by month, NLNG site data, NUPRC chunk endpoints."""
import re, json
import requests
UA = {"User-Agent": "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 Chrome/124 Safari/537.36"}
def get(u, **k):
    try: return requests.get(u, headers=UA, timeout=40, **k)
    except Exception as e:
        print("   ERR", u, type(e).__name__, str(e)[:80]); return None
MONTHS = ["january","february","march","april","may","june","july","august","september","october","november","december"]
print("=== NNPC monthly report pages")
r = get("https://www.nnpcgroup.com/insights")
print("insights hrefs:", sorted(set(re.findall(r'/insights/[a-z0-9-]+', r.text)))[:60])
ok = 0
for y in (2026, 2025, 2024):
    for m in MONTHS:
        if (y, MONTHS.index(m)) > (2026, 8): continue
        found = None
        for pat in ["nnpc-limited-monthly-report-summary-{m}-{y}", "nnpc-ltd-monthly-report-summary-{m}-{y}", "nnpc-limited-monthly-report-{m}-{y}", "nnpc-ltd-monthly-report-{m}-{y}", "nnpc-monthly-report-summary-{m}-{y}", "nnpc-limited-monthly-report-summary-for-{m}-{y}"]:
            s = pat.format(m=m, y=y)
            rr = get("https://www.nnpcgroup.com/insights/" + s)
            if rr is not None and rr.status_code == 200:
                pdfs = re.findall(r'https://[^"\' ]+\.pdf', rr.text)
                found = (s, pdfs[:2]); break
        print(y, m, found)
print("\n=== NLNG")
r = get("https://www.nlng.com/")
if r is not None:
    t = r.text
    for m in sorted(set(re.findall(r'https?://[a-z0-9./_-]*(?:strapi|cms|api|azurefd|cloudinary|amazonaws)[a-z0-9./_?=&-]*', t, re.I)))[:30]: print("  ref", m)
    for m in sorted(set(re.findall(r'href="(/[a-z0-9/_-]+)"', t)))[:80]: print("  link", m)
for u in ["https://www.nlng.com/sitemap.xml", "https://www.nlng.com/robots.txt"]:
    rr = get(u); print(u, rr.status_code if rr else None)
    if rr is not None and rr.status_code == 200: print(rr.text[:2500])
print("\n=== NUPRC")
for u in ["https://www.nuprc.gov.ng/sitemap.xml", "https://www.nuprc.gov.ng/robots.txt"]:
    rr = get(u); print(u, rr.status_code if rr else None)
    if rr is not None and rr.status_code == 200: print(rr.text[:1500])
r = get("https://www.nuprc.gov.ng/reports")
js = sorted(set(re.findall(r'/_next/static/[^"\' ]+\.js', r.text)))
for j in js:
    rj = get("https://www.nuprc.gov.ng" + j)
    if rj is None: continue
    for m in re.findall(r'.{60}(?:supabase|firebase|amazonaws|blob\.core|fetch\(|axios|/api/)[^;]{0,100}', rj.text)[:6]:
        print("  chunk", j[-20:], m[:200])
