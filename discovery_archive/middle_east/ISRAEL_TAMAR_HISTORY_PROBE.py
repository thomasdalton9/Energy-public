"""
Probe (Oct 2026) for the one-off Tamar sales history backfill (middle_east/ISRAEL_TAMAR_HISTORY.py):
  1. Tamar Petroleum's own site (wp-json media list): every PDF before 2024, text saved, sold-quantity snippets.
  2. TASE MAYA (Playwright Chromium): Tamar Petroleum's company id, its report list 2017-2025 and the text of its
     periodic / quarterly reports (sold-quantity snippets).
Texts are saved to probe_out/ (uploaded as a workflow artifact).

Usage: python3 ISRAEL_TAMAR_HISTORY_PROBE.py <round>
"""
import io
import json
import os
import re
import sys

import logging

import requests

logging.getLogger("pypdf").setLevel(logging.ERROR)

ROUND = int(sys.argv[1]) if len(sys.argv) > 1 else 1
UA = {"User-Agent": "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36"}
OUT = "probe_out"
os.makedirs(OUT, exist_ok=True)
SNIP = re.compile(r".{0,250}(?:BCM|BCM\)|bcm|מיליארד מ\"ק|מיליארד מ״ק).{0,250}")


def pdf_text(content):
    from pypdf import PdfReader
    return "\n".join((p.extract_text() or "") for p in PdfReader(io.BytesIO(content)).pages)


def save(name, text):
    with open(os.path.join(OUT, re.sub(r"[^\w.-]+", "_", name)[:150] + ".txt"), "w", encoding="utf-8") as f:
        f.write(text)


def snippets(text, n=12):
    t = re.sub(r"\s+", " ", text)
    out = []
    for m in SNIP.finditer(t):
        s = m.group(0)
        if re.search(r"תמר|Tamar|נמכר|sold|מקומי|יצוא|export|domestic", s):
            out.append(s)
    return out[:n]


def site():
    items = []
    for page in range(1, 10):
        r = requests.get(f"https://www.tamarpetroleum.co.il/wp-json/wp/v2/media?per_page=100&page={page}", headers=UA,
                         timeout=90)
        if r.status_code != 200:
            break
        js = r.json()
        items += js
        if len(js) < 100:
            break
    print(f"##### Tamar site: {len(items)} media items")
    pdfs = sorted({(it.get("date", "")[:10], it.get("source_url", "")) for it in items
                   if str(it.get("source_url", "")).lower().endswith(".pdf")})
    for d, u in pdfs:
        print("   ", d, u)
    for d, u in pdfs:
        if d >= "2024-06" or re.search(r"reserve|NSAI|immediate|presentation|מצגת", u, re.I):
            continue
        try:
            t = pdf_text(requests.get(u, headers=UA, timeout=120).content)
        except Exception as e:  # noqa: BLE001
            print("    ERR", u, e)
            continue
        save("site_" + u.rsplit("/", 1)[-1], t)
        print(f"\n  == {d} {u.rsplit('/', 1)[-1]} ({len(t):,} chars)")
        for s in snippets(t):
            print("     |", s)


def maya():
    from playwright.sync_api import sync_playwright
    with sync_playwright() as p:
        b = p.chromium.launch()
        pg = b.new_page(locale="he-IL")
        got = {}

        def on_resp(resp):
            if "/api/v1/" in resp.url:
                try:
                    got[resp.url + " " + resp.request.method] = (resp.request.post_data, resp.json())
                except Exception:  # noqa: BLE001
                    pass
        pg.on("response", on_resp)
        pg.goto("https://maya.tase.co.il/he/companies/475/reports", timeout=90000, wait_until="networkidle")
        pg.wait_for_timeout(4000)
        cid = 1689   # found in round 2 via /api/v1/reports/1575974 (reporterId; name 'תמר פטרוליום')
        for k, (_, js) in got.items():
            print("   captured", k, json.dumps(js, ensure_ascii=False)[:300])
            if "companies" in k and "GET" in k:
                items = js if isinstance(js, list) else next((v for v in js.values() if isinstance(v, list)), []) \
                    if isinstance(js, dict) else []
                print(f"      {len(items)} items")
                for it in items:
                    s = json.dumps(it, ensure_ascii=False)
                    if re.search(r"תמר|Tamar", s):
                        print("      MATCH", s[:300])

        def reports(company, page=1, limit=20):
            body = json.dumps({"pageNumber": page, "limit": limit, "companyId": company, "offset": 0})
            return pg.evaluate("""async (body) => { const r = await fetch('/api/v1/reports/companies',
                {method: 'POST', headers: {'Content-Type': 'application/json'}, body: body});
                return [r.status, await r.text()]; }""", body)

        st, txt = reports(475, 1, 2)
        print("   sample 475:", st, txt[:1500])
        # search: try metadata endpoints seen on page, else scan ids
        for u in ["/api/v1/metadata/companies", "/api/v1/companies/search?q=%D7%AA%D7%9E%D7%A8",
                  "/api/v1/autocomplete?term=%D7%AA%D7%9E%D7%A8"]:
            res = pg.evaluate("""async (u) => { const r = await fetch(u); return [r.status, (await r.text())]; }""", u)
            print("   GET", u, res[0], len(res[1]))
            for m in re.finditer(r"\{[^{}]*(?:תמר פטרוליום|Tamar Petroleum)[^{}]*\}", res[1]):
                print("      MATCH", m.group(0)[:400])
                mm = re.search(r'"(?:id|companyId|value)"\s*:\s*"?(\d+)', m.group(0))
                if mm and cid is None:
                    cid = int(mm.group(1))
        if cid is None:  # noqa
            # a MAYA report of Tamar Petroleum's that its own site links (P1575974-00.pdf, Mar 2024): its details page
            # names the reporter
            got.clear()
            allx = []
            pg.on("response", lambda resp: allx.append((resp.status, resp.request.method, resp.url,
                                                        resp.request.post_data)))
            for u in ["/api/v1/reports/1575974", "/api/v1/reports/details/1575974", "/api/v1/report/1575974",
                      "/api/v1/reports/1575974/details"]:
                res = pg.evaluate("""async (u) => { const r = await fetch(u); return [r.status, (await r.text())]; }""", u)
                print("   GET", u, res[0], res[1][:600])
                mm = re.search(r'"(?:reporterId|companyId)"\s*:\s*(\d+)', res[1])
                if mm and cid is None and res[0] == 200:
                    cid = int(mm.group(1))
            for u in ["https://maya.tase.co.il/he/reports/1575974", "https://maya.tase.co.il/he/companies",
                      "https://maya.tase.co.il/he/reports/companies"]:
                try:
                    pg.goto(u, timeout=90000, wait_until="networkidle")
                    pg.wait_for_timeout(3000)
                    print("   page", u, pg.inner_text("body")[:600].replace("\n", " | "))
                    inp = pg.query_selector("input")
                    if inp and "companies" in u:
                        inp.fill("תמר פטרו")
                        pg.wait_for_timeout(4000)
                        print("   after typing:", pg.inner_text("body")[:1500].replace("\n", " | "))
                        hrefs = pg.eval_on_selector_all("a", "els => els.map(e => e.href)")
                        print("   links:", [h for h in hrefs if "compan" in h][:40])
                except Exception as e:  # noqa: BLE001
                    print("   goto", u, e)
            for st, me, url, pdata in allx:
                if "tase.co.il" in url and not re.search(r"\.(js|css|woff2?|png|svg|ico|jpg|gif)(\?|$)", url):
                    print("    XHR", st, me, url[:200], (pdata or "")[:200])
            for k, (_, js) in got.items():
                sj = json.dumps(js, ensure_ascii=False)
                if "תמר" in sj:
                    for m in re.finditer(r"\{[^{}]*תמר[^{}]*\}", sj):
                        print("    TAMAR JSON", k[:80], m.group(0)[:400])
                        mm = re.search(r'"(?:companyId|id|value|key)"\s*:\s*"?(\d+)', m.group(0))
                        if mm and "פטרו" in m.group(0) and cid is None:
                            cid = int(mm.group(1))
            for rid in (1575974, 1273577):
                try:
                    pg.goto(f"https://maya.tase.co.il/he/reports/details/{rid}", timeout=90000, wait_until="networkidle")
                    pg.wait_for_timeout(4000)
                except Exception as e:  # noqa: BLE001
                    print("   goto", rid, e)
                for k, (_, js) in got.items():
                    sj = json.dumps(js, ensure_ascii=False)
                    print("   details", rid, k, sj[:700])
                    for mm in re.finditer(r'"(?:reporterId|companyId|issuerId)"\s*:\s*(\d+)', sj):
                        if cid is None and int(mm.group(1)) not in (475, 394):
                            cid = int(mm.group(1))
                hrefs = pg.eval_on_selector_all("a", "els => els.map(e => e.href)")
                print("   company links:", sorted({h for h in hrefs if "/companies/" in h})[:10])
                for h in hrefs:
                    mm = re.search(r"/companies/(\d+)", h)
                    if mm and cid is None:
                        cid = int(mm.group(1))
                got.clear()
                if cid:
                    break
        print("   Tamar Petroleum company id:", cid)
        if cid is None:
            b.close()
            return
        # the list POST stops after ~50 reports; the report page's own 'siblings' call pages through the whole
        # company history: GET /api/v1/reports/<any report id>/siblings?offset=N&limit=M
        allrep, ids = [], set()
        for off in range(0, 1500, 25):
            st, txt = pg.evaluate("""async (u) => { const r = await fetch(u); return [r.status, await r.text()]; }""",
                                  f"/api/v1/reports/1575974/siblings?offset={off}&limit=25")
            if off == 0:
                print("   siblings head:", st, txt[:800])
            if st != 200:
                print("   siblings", off, st, txt[:300])
                break
            js = json.loads(txt)
            lst = js if isinstance(js, list) else next((v for v in js.values() if isinstance(v, list)), []) \
                if isinstance(js, dict) else []
            new_items = [x for x in lst if x.get("id") not in ids]
            if not new_items:
                print("   siblings: nothing new at offset", off)
                break
            for x in new_items:
                ids.add(x.get("id"))
            allrep += new_items
            if min(str(x.get("publishDate", "")) for x in lst) < "2021-12":
                break
        print(f"   {len(allrep)} reports listed")
        with open(os.path.join(OUT, "maya_reports.json"), "w", encoding="utf-8") as f:
            json.dump(allrep, f, ensure_ascii=False, indent=1)
        b.close()
    keep = re.compile(r"רבעוני|תקופתי|חצי שנתי|דוחות כספיים|דוח כספי|רבעון|Quarterly|Periodic|Annual|Financial|תרגום|"
                      r"translation", re.I)
    for r in allrep:
        title = str(r.get("title", ""))
        d = str(r.get("publishDate", ""))[:10]
        atts = [a.get("url") for a in (r.get("attachments") or []) if a.get("url")]
        print("   ", d, r.get("id"), title[:120], atts[:4])
    for r in allrep:
        title = str(r.get("title", ""))
        d = str(r.get("publishDate", ""))[:10]
        if not keep.search(title) or not ("2022-03" <= d <= "2025-07"):
            continue
        for a in (r.get("attachments") or []):
            u = a.get("url") or ""
            if not u.lower().endswith(".pdf"):
                continue
            url = u if u.startswith("http") else "https://mayafiles.tase.co.il/" + u.lstrip("/")
            try:
                rr = requests.get(url, headers=UA, timeout=180)
                t = pdf_text(rr.content) if rr.content.startswith(b"%PDF") else ""
            except Exception as e:  # noqa: BLE001
                print("    ERR", url, e)
                continue
            save(f"maya_{d}_{r.get('id')}_" + url.rsplit("/", 1)[-1], t)
            print(f"\n  == {d} {r.get('id')} {title[:80]} {url} ({len(t):,} chars)")
            for s in snippets(t):
                print("     |", s)


if __name__ == "__main__":
    if ROUND in (1, 2):
        try:
            maya()
        except Exception as e:  # noqa: BLE001
            print("MAYA failed:", type(e).__name__, e)
    if ROUND in (1, 3):
        site()
