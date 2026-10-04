"""Probe 11: how far do NDOR PDFs go? The listing (transd 'daily-operational-reports-1') runs 2080-01-01 ..
2081-09-27 only. Construct uploads/shares/Daily_op_reports/NDOR%20YYYY_MM_DD.pdf for sample days 2021..2026 on the
vendor host (plain requests work there) and on transd (browser context, behind F5); also other category slugs."""
import re
from datetime import date, timedelta
import requests
import urllib3
import nepali_datetime as nd
from playwright.sync_api import sync_playwright
urllib3.disable_warnings()
UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36"
PATH = "/uploads/shares/Daily_op_reports/NDOR%20{}.pdf"
VENDOR = "https://td.neasite.dryicesolutions.net"
TRANSD = "https://transd.nea.org.np"


def bs(g):
    d = nd.date.from_datetime_date(g)
    return f"{d.year}_{d.month:02d}_{d.day:02d}"


print("check: 2025-01-11 ->", bs(date(2025, 1, 11)), "(PDF says 2081/09/27); 2023-04-14 ->", bs(date(2023, 4, 14)))
days = []
g = date(2021, 1, 1)
while g <= date(2026, 10, 3):
    days.append(g)
    g += timedelta(days=9)
S = requests.Session()
S.headers["User-Agent"] = UA
found = []
for g in days:
    u = VENDOR + PATH.format(bs(g))
    try:
        r = S.get(u, timeout=40, verify=False, stream=True)
        head = next(r.iter_content(8), b"")
        ok = r.status_code == 200 and head[:4] == b"%PDF"
        r.close()
    except Exception as e:
        ok = f"ERR {type(e).__name__}"
    print("VENDOR", g, bs(g), ok, flush=True)
    if ok is True:
        found.append(g)
print("vendor found", len(found), "of", len(days), "first", found[:1], "last", found[-1:])

with sync_playwright() as p:
    b = p.chromium.launch(args=["--disable-blink-features=AutomationControlled"])
    ctx = b.new_context(user_agent=UA, ignore_https_errors=True)
    page = ctx.new_page()
    page.goto(TRANSD + "/en", wait_until="domcontentloaded", timeout=60000)
    page.wait_for_timeout(8000)
    tf = []
    for g in days:
        u = TRANSD + PATH.format(bs(g))
        try:
            r = ctx.request.get(u, timeout=60000)
            body = r.body()
            ok = body[:4] == b"%PDF"
            info = f"{r.status} {len(body)}"
        except Exception as e:
            ok, info = False, f"ERR {type(e).__name__}"
        print("TRANSD", g, bs(g), ok, info, flush=True)
        if ok:
            tf.append(g)
    print("transd found", len(tf), "of", len(days), "first", tf[:1], "last", tf[-1:])
    for slug in ("daily-operational-reports", "daily-operational-reports-2", "daily-operational-report",
                 "monthly-operational-reports", "yearly-operational-reports-1"):
        u = f"{TRANSD}/en/category/{slug}"
        page.goto(u, wait_until="domcontentloaded", timeout=60000)
        html = ""
        for _ in range(15):
            page.wait_for_timeout(1500)
            html = page.content()
            if "/detail/" in html or "404" in page.title():
                break
        det = list(dict.fromkeys(re.findall(r'href="[^"]+/detail/([^"]+)"', html)))
        pages = [int(x) for x in re.findall(r"[?&]page=(\d+)", html)]
        print("CAT", slug, page.title()[:50], "details", det[:6], "maxpage", max(pages) if pages else None)
    b.close()
