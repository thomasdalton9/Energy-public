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
        cid = None
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
            body = json.dumps({"pageNumber": page, "limit": limit, "companyId": company, "offset": (page - 1) * limit})
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
        if cid is None:
            # a MAYA report of Tamar Petroleum's that its own site links (P1575974-00.pdf, Mar 2024): its details page
            # names the reporter
            got.clear()
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
        allrep = []
        for page in range(1, 80):
            st, txt = reports(cid, page, 50)
            if st != 200:
                print("   page", page, st, txt[:200])
                break
            js = json.loads(txt)
            lst = js if isinstance(js, list) else next((v for v in js.values() if isinstance(v, list)), [])
            if not lst:
                break
            allrep += lst
            if min(str(x.get("publishDate", "")) for x in lst) < "2017-01":
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
