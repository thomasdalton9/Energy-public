"""
Probe 17: Plinacro SUKAP 'Ostvareni fizicki protoci' search request (JS click), replay for other dates/categories. Prints only (short).
"""
import json
import re
import sys

import requests

UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0 Safari/537.36"
H = {"User-Agent": UA}


def main():
    from playwright.sync_api import sync_playwright
    calls = []
    with sync_playwright() as p:
        b = p.chromium.launch()
        pg = b.new_page(user_agent=UA, viewport={"width": 1400, "height": 1200})
        pg.on("request", lambda rq: calls.append((rq.url, rq.post_data)) if rq.resource_type in ("xhr", "fetch") and rq.method == "POST" and "loadRegistry" not in rq.url and "getPointList" not in rq.url else None)
        pg.goto("https://www.sukap.plinacro.hr/pub/flow", wait_until="networkidle", timeout=60000)
        pg.wait_for_timeout(3000)
        ins = pg.locator("input[placeholder='DD.MM.GGGG']")
        print("date inputs", ins.count())
        try:
            ins.nth(0).fill("15.01.2025")
            ins.nth(1).fill("16.01.2025")
        except Exception as e:  # noqa: BLE001
            print("fill fail", type(e).__name__)
        pg.wait_for_timeout(500)
        print("click:", pg.evaluate("() => { const b=[...document.querySelectorAll('button')].find(x=>x.innerText.trim()==='Pretraži'); if(b){b.click(); return 'clicked'} return 'none'}"))
        pg.wait_for_timeout(6000)
        for u, d in calls:
            print(" call", u, (d or "")[:900])
        print("text:", pg.inner_text("body")[:1500].replace("\n", " | ")[-900:])
        b.close()
    for u, d in calls[-1:]:
        try:
            body = json.loads(d)
        except Exception:  # noqa: BLE001
            continue
        print("body keys", list(body.keys()))
        for ctx in ("END_BUYER", "DISTRIBUTION", "ALL_CATEGORIES"):
            bb = json.loads(d)
            for k in list(bb):
                if "ontext" in k:
                    bb[k] = ctx
            r = requests.post(u, json=bb, headers=H, timeout=60)
            print(ctx, r.status_code, len(r.text), r.text[:500])
        # depth
        bb = json.loads(d)
        txt = json.dumps(bb)
        for old in set(re.findall(r"2025-01-1\d|1\d\.01\.2025", txt)):
            pass


try:
    main()
except Exception as e:  # noqa: BLE001
    print("FAILED", type(e).__name__, str(e)[:300])
sys.exit(0)
