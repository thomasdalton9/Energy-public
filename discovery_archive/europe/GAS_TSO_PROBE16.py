"""
Probe 16: Enagas bulletin archive via ?month=&year= (all unique PDFs 2020-2026), sample parse of the demand table in
old and new formats; Plinacro SUKAP flow page buttons / search XHR. Prints only (short).
"""
import io
import json
import re
import sys

import requests

UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0 Safari/537.36"
H = {"User-Agent": UA}
EN = "https://www.enagas.es"
PAGE = EN + "/en/technical-management-system/energy-data/publications/gas-statistical-bulletin/"


def enagas():
    print("=== ENAGAS archive")
    got = {}
    for y in range(2019, 2027):
        for m in range(1, 13):
            try:
                r = requests.get(PAGE, params={"category": "", "month": m, "year": y}, headers=H, timeout=60)
                for x in re.findall(r'href="(/content/dam[^"]+\.pdf)"', r.text):
                    got.setdefault(x, (y, m))
            except Exception as e:  # noqa: BLE001
                print(y, m, type(e).__name__)
    print(len(got), "unique pdfs")
    for x, ym in sorted(got.items(), key=lambda kv: kv[1]):
        print("  ", ym, x.split("/")[-1])
    try:
        import pdfplumber
    except Exception:  # noqa: BLE001
        return
    names = list(got)
    sample = [n for n in names if re.search(r"(?i)december-2022|mayo_INGLES|diciembre20|Nov23|may-2022", n)][:4]
    for x in sample:
        r = requests.get(EN + x, headers=H, timeout=120)
        with pdfplumber.open(io.BytesIO(r.content)) as pdf:
            print("PDF", x.split("/")[-1], "pages", len(pdf.pages))
            for i, pgx in enumerate(pdf.pages[:8]):
                t = pgx.extract_text() or ""
                if re.search(r"(?i)national market|conventional", t) and "GWh" in t:
                    print(f"--- page {i + 1}\n{t[:900]}")
                    break


def plinacro():
    print("=== PLINACRO flow page")
    from playwright.sync_api import sync_playwright
    with sync_playwright() as p:
        b = p.chromium.launch()
        pg = b.new_page(user_agent=UA)
        calls = []
        pg.on("request", lambda rq: calls.append((rq.url, rq.post_data)) if rq.resource_type in ("xhr", "fetch") and rq.method == "POST" and "loadRegistry" not in rq.url and "getPointList" not in rq.url else None)
        pg.goto("https://www.sukap.plinacro.hr/pub/flow", wait_until="networkidle", timeout=60000)
        pg.wait_for_timeout(3000)
        print(" buttons:", pg.eval_on_selector_all("button", "els => els.map(e => e.innerText.trim().slice(0,30))"))
        print(" inputs:", pg.eval_on_selector_all("input", "els => els.map(e => [e.type, e.placeholder, e.value, e.name, e.id].join('|'))"))
        try:
            pg.locator("button:has-text('Pretra')").first.click(timeout=6000)
            pg.wait_for_timeout(5000)
        except Exception as e:  # noqa: BLE001
            print(" click fail", type(e).__name__)
        for u, d in calls:
            print(" call", u, (d or "")[:600])
        b.close()
    for u, d in calls[-1:]:
        try:
            body = json.loads(d)
        except Exception:  # noqa: BLE001
            continue
        txt = json.dumps(body)
        print(" body", txt[:500])
        for variant in (re.sub(r"20\d\d-\d\d-\d\d", "2025-01-15", txt), re.sub(r"(\d\d)\.(\d\d)\.20\d\d", "15.01.2025", txt)):
            try:
                r = requests.post(u, data=variant, headers={**H, "Content-Type": "application/json"}, timeout=60)
                print(" replay", r.status_code, variant[:260], "->", r.text[:700])
            except Exception as e:  # noqa: BLE001
                print(" replay fail", type(e).__name__)


for fn in (plinacro, enagas):
    try:
        fn()
    except Exception as e:  # noqa: BLE001
        print(fn.__name__, "FAILED", type(e).__name__, str(e)[:300])
sys.exit(0)
