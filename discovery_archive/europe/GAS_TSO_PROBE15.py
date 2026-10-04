"""
Probe 15: Enagas bulletin list via the page's ?month=&year= filter (+ Dec-2022 bulletin demand table);
Plinacro SUKAP flow search request (playwright click Pretrazi, then replay with other dates). Prints only (short).
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
    print("=== ENAGAS month/year filter")
    got = {}
    for y in (2021, 2022, 2023):
        for m in (1, 6, 12):
            try:
                r = requests.get(PAGE, params={"category": "", "month": m, "year": y}, headers=H, timeout=60)
                ls = re.findall(r'href="(/content/dam[^"]+\.(?:pdf|xlsx?))"', r.text)
                print(y, m, r.status_code, len(r.text), [x.split("/")[-1] for x in ls][:6])
                for x in ls:
                    got[x] = 1
            except Exception as e:  # noqa: BLE001
                print(y, m, type(e).__name__)
    pick = [x for x in got if re.search(r"dic.?22|dec.?22|dic.?21|dec.?21", x, re.I)]
    print("pick", pick)
    try:
        import pdfplumber
        for x in pick[:1]:
            r = requests.get(EN + x, headers=H, timeout=120)
            with pdfplumber.open(io.BytesIO(r.content)) as pdf:
                print(x.split("/")[-1], "pages", len(pdf.pages))
                for i, pg in enumerate(pdf.pages[:8]):
                    t = pg.extract_text() or ""
                    if re.search(r"(?i)evolution of gas demand|evoluci", t) and re.search(r"TOTAL", t):
                        print(f"--- page {i + 1}\n{t[:1400]}")
                        break
    except Exception as e:  # noqa: BLE001
        print("pdf fail", type(e).__name__, str(e)[:200])


def plinacro():
    print("=== PLINACRO flow search")
    from playwright.sync_api import sync_playwright
    with sync_playwright() as p:
        b = p.chromium.launch()
        pg = b.new_page(user_agent=UA)
        calls = []

        def on(rq):
            if rq.resource_type in ("xhr", "fetch") and rq.method == "POST" and "loadRegistry" not in rq.url and "getPointList" not in rq.url:
                calls.append((rq.url, rq.post_data))
        pg.on("request", on)
        pg.goto("https://www.sukap.plinacro.hr/pub/flow", wait_until="networkidle", timeout=60000)
        pg.wait_for_timeout(3000)
        try:
            pg.get_by_text("Pretraži").first.click(timeout=5000)
            pg.wait_for_timeout(5000)
        except Exception as e:  # noqa: BLE001
            print("click fail", type(e).__name__)
        for u, d in calls:
            print(" call", u, (d or "")[:500])
        b.close()
    for u, d in calls[-2:]:
        try:
            body = json.loads(d)
        except Exception:  # noqa: BLE001
            continue
        print(" replay keys:", list(body.keys())[:20])
        txt = json.dumps(body)
        txt2 = re.sub(r"20\d\d-\d\d-\d\d", "2025-01-15", txt)
        txt2 = re.sub(r'(\d\d)\.(\d\d)\.2026', r'15.01.2025', txt2)
        for variant in (txt, txt2):
            try:
                r = requests.post(u, data=variant, headers={**H, "Content-Type": "application/json"}, timeout=60)
                print(" replay", r.status_code, variant[:200], "->", r.text[:600])
            except Exception as e:  # noqa: BLE001
                print(" replay fail", type(e).__name__)


for fn in (plinacro, enagas):
    try:
        fn()
    except Exception as e:  # noqa: BLE001
        print(fn.__name__, "FAILED", type(e).__name__, str(e)[:300])
sys.exit(0)
