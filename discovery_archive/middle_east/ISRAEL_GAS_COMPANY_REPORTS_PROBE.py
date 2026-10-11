"""
Probe (Oct 2026): can Israel's natural gas production / sales by field (Leviathan, Tamar, Karish) and exports to
Egypt / Jordan be built from the field owners' periodic company reports, fetched from GitHub Actions?

Usage: python3 ISRAEL_GAS_COMPANY_REPORTS_PROBE.py <round>

Findings are recorded below per run.
"""
import json
import re
import sys
import io

import requests

try:
    from curl_cffi import requests as creq
except Exception:  # noqa: BLE001
    creq = None

ROUND = int(sys.argv[1]) if len(sys.argv) > 1 else 1
UA = {"User-Agent": "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36",
      "Accept-Language": "en-US,en;q=0.9"}
SEC_UA = {"User-Agent": "EnergyPublicProbe research-probe@example.org", "Accept-Encoding": "gzip, deflate", "Accept": "*/*"}
LINK_RE = re.compile(r'href=["\']([^"\']+)["\']', re.I)
CHALLENGE = ("cf-chl", "Just a moment", "challenge-platform", "you have been blocked", "Attention Required")


def get(url, method="plain", headers=None, **kw):
    h = dict(UA)
    h.update(headers or {})
    try:
        if method == "plain":
            r = requests.get(url, headers=h, timeout=45, **kw)
        else:
            r = creq.get(url, impersonate="chrome", headers=headers or {}, timeout=45, **kw)
    except Exception as e:  # noqa: BLE001
        print(f"  [{method}] {url}: {type(e).__name__} {str(e)[:120]}")
        return None
    t = r.text if "pdf" not in str(r.headers.get("content-type", "")).lower() else ""
    ch = any(k in t for k in CHALLENGE)
    title = re.search(r"<title>(.*?)</title>", t, re.S)
    print(f"  [{method}] {url}: HTTP {r.status_code}, {len(r.content):,} B, {r.headers.get('content-type')}, "
          f"title {title.group(1).strip()[:70] if title else None!r}, challenge={ch}")
    r.challenge = ch
    return r


def fetch(url, **kw):
    r = get(url, "plain", **kw)
    if (r is None or r.status_code != 200 or r.challenge) and creq is not None:
        r = get(url, "cffi", **kw)
    return r


def links(r, pat=r"\.pdf|report|quarter|present|investor|financ|annual|bond|maya|sales|production", base=""):
    if r is None or r.status_code != 200:
        return []
    out = []
    for h in LINK_RE.findall(r.text):
        if re.search(pat, h, re.I):
            if h.startswith("/"):
                h = base + h
            out.append(h)
    return sorted(set(out))


def pdf_text(content, maxpages=80):
    import pdfplumber
    txt = []
    with pdfplumber.open(io.BytesIO(content)) as pdf:
        for p in pdf.pages[:maxpages]:
            txt.append(p.extract_text() or "")
    return "\n".join(txt)


VOL_RE = re.compile(r"[^\n]{0,160}(?:BCM|bcm|billion cubic|MMCF|MMcf|mmscf|kboe|BCF|Bcf)[^\n]{0,160}")


def round1():
    sites = {
        "NewMed": ["https://www.newmedenergy.com/", "https://newmedenergy.com/investor-relations/financial-reports/",
                   "https://newmedenergy.com/investor-relations/presentations/",
                   "https://newmedenergy.com/investor-relations/leviathan-bond-reports/",
                   "https://www.newmedenergy.com/investors/"],
        "Ratio": ["https://www.ratioenergies.com/", "https://www.ratioenergies.com/en/",
                  "https://www.ratioenergies.com/investors-relations/", "https://ratio.co.il/"],
        "TamarPetroleum": ["https://www.tamarpetroleum.co.il/en/", "https://www.tamarpetroleum.co.il/en/investor-relations/",
                           "https://www.tamarpetroleum.co.il/en/financial-reports/", "https://www.tamarpetroleum.co.il/"],
        "Isramco": ["https://www.isramco.co.il/", "https://www.isramco.co.il/en/", "https://www.isramconegev.co.il/"],
        "Delek": ["https://www.delek-group.com/", "https://www.delek-group.com/investor-relations/"],
        "Everest/Dor": ["https://www.dor-alon.co.il/", "https://www.dorgas.co.il/"],
        "Energean": ["https://www.energean.com/investors/results-reports-and-presentations/",
                     "https://www.investegate.co.uk/company/ENOG",
                     "https://www.investegate.co.uk/company/ENOG/announcements",
                     "https://www.londonstockexchange.com/stock/ENOG/energean-plc/analysis",
                     "https://api.londonstockexchange.com/api/gw/lse/instruments/alldata/ENOG"],
        "MAYA": ["https://maya.tase.co.il/", "https://maya.tase.co.il/en/reports/companies",
                 "https://mayaapi.tase.co.il/api/company/alllist",
                 "https://mayaapi.tase.co.il/api/report/filter",
                 "https://api.tase.co.il/api/company/securitydata?securityId=1141571&lang=1",
                 "https://market.tase.co.il/en/market_data/company/1983/major_data"],
    }
    for name, urls in sites.items():
        print(f"\n##### {name}")
        for u in urls:
            r = fetch(u)
            base = re.match(r"https?://[^/]+", u).group(0)
            ls = links(r, base=base)
            if ls:
                print(f"      links ({len(ls)}):")
                for l in ls[:60]:
                    print("        ", l)

    print("\n##### MAYA POST filter")
    for url in ["https://mayaapi.tase.co.il/api/report/filter", "https://maya.tase.co.il/api/v1/reports/companies"]:
        body = {"Page": 1, "GroupData": [], "DateFrom": "2026-01-01T00:00:00.000Z", "DateTo": "2026-10-10T00:00:00.000Z",
                "IsBreakingAnnouncement": False, "IsForTaseMember": False, "Events": [], "Subevents": [], "QOpt": 1,
                "Q": "ניו-מד"}
        for m in ("plain", "cffi"):
            try:
                if m == "plain":
                    r = requests.post(url, json=body, headers={**UA, "X-Maya-With": "allow", "Origin": "https://maya.tase.co.il",
                                                               "Referer": "https://maya.tase.co.il/"}, timeout=45)
                else:
                    r = creq.post(url, json=body, impersonate="chrome", headers={"X-Maya-With": "allow",
                                                                                "Origin": "https://maya.tase.co.il",
                                                                                "Referer": "https://maya.tase.co.il/"},
                                  timeout=45)
                print(f"  [{m}] POST {url}: {r.status_code} {r.headers.get('content-type')} {r.text[:600]!r}")
            except Exception as e:  # noqa: BLE001
                print(f"  [{m}] POST {url}: {type(e).__name__} {str(e)[:120]}")

    print("\n##### Chevron EDGAR")
    r = get("https://data.sec.gov/submissions/CIK0000093410.json", headers=SEC_UA)
    if r is not None and r.status_code == 200:
        j = r.json()
        rec = j["filings"]["recent"]
        tenk = [(rec["form"][i], rec["filingDate"][i], rec["accessionNumber"][i], rec["primaryDocument"][i])
                for i in range(len(rec["form"])) if rec["form"][i] in ("10-K", "10-Q")][:6]
        print("   ", tenk)
        for form, d, acc, doc in tenk[:3]:
            u = f"https://www.sec.gov/Archives/edgar/data/93410/{acc.replace('-', '')}/{doc}"
            rr = get(u, headers=SEC_UA)
            if rr is not None and rr.status_code == 200:
                txt = re.sub(r"<[^>]+>", " ", rr.text)
                txt = re.sub(r"&#160;|&nbsp;", " ", txt)
                txt = re.sub(r"\s+", " ", txt)
                for m in re.finditer(r"Israel", txt):
                    print("      ", form, d, "...", txt[max(0, m.start() - 200): m.start() + 300])
                    break
                n = len(re.findall("Israel", txt))
                print(f"      {form} {d}: 'Israel' x{n}")
                for m in list(re.finditer(r"Israel[^.]{0,250}(?:MMCF|cubic feet|per day|production)", txt))[:5]:
                    print("        >>", m.group(0)[:300])


def pdf_hits(url, pat=VOL_RE, maxpages=120, maxhits=40, method="plain"):
    r = fetch(url)
    if r is None or r.status_code != 200 or not r.content[:5].startswith(b"%PDF"):
        print("      not a PDF")
        return ""
    try:
        t = pdf_text(r.content, maxpages)
    except Exception as e:  # noqa: BLE001
        print("      pdf error", e)
        return ""
    hits = [m.group(0).strip() for m in pat.finditer(t)]
    print(f"      text {len(t):,} chars, {len(hits)} volume hits")
    seen = set()
    for h in hits:
        if h in seen:
            continue
        seen.add(h)
        print("        |", h[:300])
        if len(seen) >= maxhits:
            break
    return t


def sec_get(url):
    h = {"User-Agent": SEC_UA["User-Agent"], "Accept-Encoding": "gzip, deflate"}
    try:
        r = requests.get(url, headers=h, timeout=60)
        print(f"  [sec] {url}: HTTP {r.status_code}, {len(r.content):,} B")
        return r
    except Exception as e:  # noqa: BLE001
        print("  [sec]", url, e)
        return None


def round2():
    NM = "https://newmedenergy.com/wp-content/uploads/"
    docs = {
        "NewMed FS H1 2026": NM + "2026/08/FS330062026ACCESSIBLE.pdf",
        "NewMed FS Q1 2026": NM + "2026/05/FS31032026ACCESSIBLE.pdf",
        "NewMed pres 2Q26": NM + "2026/08/2Q-2026-Financial-Results-Presentation.pdf",
        "NewMed pres 4Q25": NM + "2026/03/NWMD_4Q25_VF.pdf",
        "Leviathan Bond FS H1 2026": NM + "2026/08/Leviathan-Bond-Ltd-FS-30.6.2026.pdf",
        "Tamar Petroleum Q2 2026": "https://www.tamarpetroleum.co.il/wp-content/uploads/2026/09/Tamar-Petroleum-FS-Q2_2026.pdf",
        "Tamar Petroleum Q1 2026": "https://www.tamarpetroleum.co.il/wp-content/uploads/2026/06/Tamar-Petroleum-FS-Q1_2026_26.05.2026_Final.pdf",
        "Ratio pres Aug 2026 (he)": "https://ratioenergies.com/wp-content/uploads/2026/08/מצגת-שוק-ההון-אוגוסט-2026-.pdf",
    }
    for k, u in docs.items():
        print(f"\n##### {k}")
        pdf_hits(u)

    print("\n##### NewMed immediate reports page")
    r = fetch("https://newmedenergy.com/investor-relations/immediate-reports/")
    for l in links(r, pat=r"\.pdf")[:80]:
        print("        ", l)
    print("\n##### Tamar Petroleum EN reports list + presentations (he)")
    for u in ["https://www.tamarpetroleum.co.il/en/investors-relations/financial-reports/",
              "https://www.tamarpetroleum.co.il/investors/מצגות/",
              "https://www.tamarpetroleum.co.il/investors/%d7%93%d7%95%d7%97%d7%95%d7%aa-%d7%9b%d7%a1%d7%a4%d7%99%d7%99%d7%9d/"]:
        r = fetch(u)
        for l in links(r, pat=r"\.pdf")[:80]:
            print("        ", l)
    print("\n##### Ratio pages")
    for u in ["https://ratioenergies.com/en/", "https://ratioenergies.com/investors/", "https://ratioenergies.com/en/investors/",
              "https://ratioenergies.com/%d7%93%d7%95%d7%97%d7%95%d7%aa-%d7%9b%d7%a1%d7%a4%d7%99%d7%99%d7%9d/",
              "https://ratioenergies.com/wp-json/wp/v2/media?per_page=100&mime_type=application/pdf"]:
        r = fetch(u)
        if r is not None and r.status_code == 200 and "json" in str(r.headers.get("content-type")):
            try:
                for m in r.json():
                    print("        ", m.get("date"), m.get("source_url"))
            except Exception as e:  # noqa: BLE001
                print("     json err", e, r.text[:300])
        else:
            for l in links(r, pat=r"\.pdf|invest|report")[:60]:
                print("        ", l)

    print("\n##### Energean via investegate")
    allann = []
    for page in range(1, 6):
        r = fetch(f"https://www.investegate.co.uk/company/ENOG?page={page}")
        ls = links(r, pat=r"/announcement/")
        allann += ls
    allann = sorted(set(allann))
    print(f"   {len(allann)} announcements")
    for l in allann:
        print("        ", l)
    pick = [l for l in allann if re.search(r"result|trading|operational|update|production", l, re.I)][:6]
    for l in pick:
        r = fetch(l)
        if r is not None and r.status_code == 200:
            txt = re.sub(r"<[^>]+>", " ", r.text)
            txt = re.sub(r"&nbsp;|&#160;", " ", txt)
            txt = re.sub(r"[ \t]+", " ", txt)
            hits = [m.group(0) for m in VOL_RE.finditer(txt)]
            print(f"      {l}: {len(hits)} hits")
            for h in hits[:15]:
                print("        |", h.strip()[:300])

    print("\n##### LSE news API")
    r = get("https://api.londonstockexchange.com/api/gw/lse/instruments/alldata/ENOG")
    if r is not None:
        print("     ", r.text[:800])
    body = {"path": "news", "parameters": "tidm%3DENOG%26tab%3Dnews-explorer%26tabId%3D58734a12-d97c-40cb-8047-df76e660f23f",
            "components": [{"componentId": "block_content%3A431d02ac-09b8-40c9-aba6-04a72a4f2e49",
                            "parameters": "sortby%3Ddatetime%26size%3D50%26page%3D0%26tidm%3DENOG"}]}
    try:
        r = requests.post("https://api.londonstockexchange.com/api/v1/components/refresh", json=body, headers=UA, timeout=45)
        print("   LSE refresh:", r.status_code, r.text[:1500])
    except Exception as e:  # noqa: BLE001
        print("   LSE refresh err", e)

    print("\n##### Chevron EDGAR")
    r = sec_get("https://data.sec.gov/submissions/CIK0000093410.json")
    if r is not None and r.status_code == 200:
        rec = r.json()["filings"]["recent"]
        lst = [(rec["form"][i], rec["filingDate"][i], rec["accessionNumber"][i], rec["primaryDocument"][i])
               for i in range(len(rec["form"])) if rec["form"][i] in ("10-K", "10-Q")][:5]
        print("   ", lst)
        for form, d, acc, doc in lst[:3]:
            rr = sec_get(f"https://www.sec.gov/Archives/edgar/data/93410/{acc.replace('-', '')}/{doc}")
            if rr is None or rr.status_code != 200:
                continue
            txt = re.sub(r"<[^>]+>", " ", rr.text)
            txt = re.sub(r"&#160;|&nbsp;|&#xa0;", " ", txt)
            txt = re.sub(r"\\s+", " ", txt)
            print(f"      {form} {d}: 'Israel' x{txt.count('Israel')}")
            for m in list(re.finditer(r"Israel", txt))[:8]:
                print("        >>", txt[max(0, m.start() - 150): m.start() + 250])


def maya_playwright():
    from playwright.sync_api import sync_playwright
    print("\n##### MAYA via Playwright (capture XHR)")
    with sync_playwright() as p:
        b = p.chromium.launch()
        pg = b.new_page(locale="en-US")
        seen = []
        pg.on("response", lambda resp: seen.append((resp.status, resp.url, resp.headers.get("content-type", ""))))
        for u in ["https://maya.tase.co.il/en/reports/companies", "https://maya.tase.co.il/en/companies/475"]:
            try:
                pg.goto(u, timeout=60000, wait_until="networkidle")
            except Exception as e:  # noqa: BLE001
                print("   goto", u, e)
            print("   page", u, pg.title()[:80], len(pg.content()))
            print("   text:", pg.inner_text("body")[:1500].replace("\n", " | "))
        for s, url, ct in seen:
            if "json" in ct or "api" in url:
                print("    ", s, url[:200], ct)
        b.close()


def fast_text(content, maxpages=400):
    from pypdf import PdfReader
    rd = PdfReader(io.BytesIO(content))
    return "\n".join((pg.extract_text() or "") for pg in rd.pages[:maxpages])


def ctx(t, pat, before=200, after=700, maxn=4):
    out = []
    for m in list(re.finditer(pat, t))[:maxn]:
        out.append(re.sub(r"\s+", " ", t[max(0, m.start() - before): m.end() + after]))
    return out


def round3():
    # 1. NewMed quarterly presentations: Leviathan sales by destination table
    r = fetch("https://newmedenergy.com/investor-relations/presentations/")
    pres = [l for l in links(r, pat=r"\.pdf") if re.search(r"/20(2[0-6])/", l)]
    print(f"\n##### NewMed presentations 2020+: {len(pres)}")
    for u in pres:
        rr = get(u)
        if rr is None or rr.status_code != 200 or not rr.content.startswith(b"%PDF"):
            continue
        try:
            t = pdf_text(rr.content, 60)
        except Exception as e:  # noqa: BLE001
            print("    pdf err", e)
            continue
        rows = re.findall(r"[^\n]{0,60}\(BCM\)[^\n]{0,80}", t)
        rows += re.findall(r"[^\n]{0,80}(?:sales|Sales)[^\n]{0,20}(?:MMcf/d|M Mcf/d|bcf/d)[^\n]{0,60}", t)
        print("   ", u.rsplit("/", 1)[-1], "|", " || ".join(x.strip() for x in rows[:12]))
        if u.endswith("2Q-2026-Financial-Results-Presentation.pdf"):
            for c in ctx(t, r"\(BCM\)", 900, 300, 3):
                print("      CTX:", c[:1300])

    # 2. NewMed financial statements: Leviathan quarterly sentences + Karish royalty volumes
    r = fetch("https://newmedenergy.com/investor-relations/financial-reports/")
    fs = [l for l in links(r, pat=r"\.pdf") if re.search(r"/20(2[0-6])/", l)]
    print(f"\n##### NewMed FS 2020+: {len(fs)}")
    for u in fs:
        rr = get(u)
        if rr is None or rr.status_code != 200 or not rr.content.startswith(b"%PDF"):
            continue
        try:
            t = fast_text(rr.content)
        except Exception as e:  # noqa: BLE001
            print("    pdf err", e)
            continue
        t1 = re.sub(r"\s+", " ", t)
        print("   ", u.rsplit("/", 1)[-1], f"{len(t1):,} chars")
        for pat in [r"quantit[^.]{0,80}sold from the Leviathan[^.]{0,400}",
                    r"Leviathan reservoir in Q[1-4][^.]{0,250}",
                    r"[^.]{0,200}sold by Energean in [QH][^.]{0,250}",
                    r"[^.]{0,120}Tamar reservoir[^.]{0,60}BCM[^.]{0,200}",
                    r"Total Offtake \(BCM\)[^A-Za-z]{0,60}",
                    r"[^.]{0,250}BCM, of which[^.]{0,300}"]:
            for m in list(re.finditer(pat, t1))[:3]:
                print("        |", m.group(0)[:420])

    # 3. Tamar Petroleum: sales table context
    for u in ["https://www.tamarpetroleum.co.il/wp-content/uploads/2026/09/Tamar-Petroleum-FS-Q2_2026.pdf",
              "https://www.tamarpetroleum.co.il/wp-content/uploads/2026/03/Tamar-Petroleum-Financial-Information-2025-Final.pdf",
              "https://www.tamarpetroleum.co.il/wp-content/uploads/2025/11/Tamar-Petroleum-Q32025-Report-final-clean.pdf"]:
        print(f"\n##### Tamar {u.rsplit('/', 1)[-1]}")
        rr = get(u)
        if rr is None or rr.status_code != 200:
            continue
        t = pdf_text(rr.content, 120)
        for c in ctx(t, r"\(in BCM\)|in BCM\)|BCM \(100%\)", 300, 900, 4):
            print("      CTX:", c[:1300])
    print("\n##### Tamar Petroleum: wp-json media / older report URLs")
    for u in ["https://www.tamarpetroleum.co.il/wp-json/wp/v2/media?per_page=100&mime_type=application/pdf",
              "https://www.tamarpetroleum.co.il/wp-json/wp/v2/media?per_page=100&search=Report",
              "https://newmedenergy.com/wp-json/wp/v2/media?per_page=5&mime_type=application/pdf",
              "https://ratioenergies.com/wp-json/wp/v2/media?per_page=100"]:
        rr = get(u)
        if rr is not None and rr.status_code == 200 and "json" in str(rr.headers.get("content-type")):
            try:
                js = rr.json()
                print(f"      {len(js)} items")
                for m in js[:100]:
                    print("        ", m.get("date"), m.get("source_url"))
            except Exception as e:  # noqa: BLE001
                print("      json err", e)
    # Playwright for Tamar's tabbed report archive
    try:
        from playwright.sync_api import sync_playwright
        with sync_playwright() as p:
            b = p.chromium.launch()
            pg = b.new_page()
            for u in ["https://www.tamarpetroleum.co.il/en/investors-relations/financial-reports/",
                      "https://www.tamarpetroleum.co.il/investors/%d7%93%d7%95%d7%97%d7%95%d7%aa-%d7%9b%d7%a1%d7%a4%d7%99%d7%99%d7%9d/",
                      "https://ratioenergies.com/"]:
                pg.goto(u, timeout=60000, wait_until="networkidle")
                for el in pg.query_selector_all("[role=tab], .elementor-tab-title, .e-n-tab-title, button")[:40]:
                    try:
                        el.click(timeout=1500)
                        pg.wait_for_timeout(300)
                    except Exception:  # noqa: BLE001
                        pass
                hrefs = pg.eval_on_selector_all("a", "els => els.map(e => e.href)")
                pdfs = sorted(set(h for h in hrefs if ".pdf" in h.lower()))
                print(f"   PW {u}: {len(pdfs)} pdf links")
                for h in pdfs[:120]:
                    print("        ", h)
            b.close()
    except Exception as e:  # noqa: BLE001
        print("   PW err", e)

    # 4. Energean RNS: Israel production lines
    ids = ["results-for-half-year-ended-30-june-2026/9761841", "trading-statement-operational-update/9576511",
           "2025-full-year-results/9481429", "trading-statement-operational-update/9388069",
           "9-months-2025-production-and-adjusted-ebitdax/9196364", "results-for-half-year-ended-30-june-2025/9100543",
           "trading-statement-operational-update/9257244", "2024-full-year-results/8788032"]
    print("\n##### Energean RNS Israel lines")
    for i in ids:
        r = fetch("https://www.investegate.co.uk/announcement/rns/energean--enog/" + i)
        if r is None or r.status_code != 200:
            continue
        txt = re.sub(r"<(?:td|th)[^>]*>", " | ", r.text)
        txt = re.sub(r"</tr>", "\n", txt)
        txt = re.sub(r"<[^>]+>", " ", txt)
        txt = re.sub(r"&nbsp;|&#160;", " ", txt).replace("&#183;", "*").replace("&amp;", "&")
        txt = re.sub(r"[ \t]+", " ", txt)
        print("   ", i)
        for m in re.finditer(r"[^\n]{0,200}(?:Israel|Karish)[^\n]{0,250}", txt):
            line = m.group(0).strip()
            if re.search(r"kboe|bcm|mmscf|\d{2,3}\s*\|", line, re.I):
                print("        |", line[:400])

    # 5. Natural Gas Authority / Ministry annual review quoted in NewMed reports ('BCM, of which'): done in 2


def round4():
    NM = "https://newmedenergy.com/wp-content/uploads/"
    print("\n##### A. NewMed FS: Leviathan quantities by customer location (table text)")
    for f in ["2026/08/FS330062026ACCESSIBLE.pdf", "2026/03/FS31122025ACCESSIBLE.pdf", "2025/11/FS30092025ENGACCESSIBLE.pdf",
              "2024/03/NM_FS31122023_ACCESSIBLE.pdf", "2022/03/NewMedFS31122021EN.pdf", "2021/12/FS-Q1-2021-English.pdf"]:
        rr = get(NM + f)
        if rr is None or rr.status_code != 200:
            continue
        for eng, fn in (("pypdf", fast_text), ("pdfplumber", pdf_text)):
            try:
                t = fn(rr.content, 400) if eng == "pypdf" else fn(rr.content, 400)
            except Exception as e:  # noqa: BLE001
                print("    err", e)
                continue
            m = re.search(r"geographic location", t) or re.search(r"customer location", t)
            if m:
                print(f"   [{eng}] {f}:\n" + t[m.start() - 200: m.start() + 1300])
            m2 = re.search(r"sold by Energean in [QH]", t)
            if m2:
                print(f"   [{eng}] Karish royalty note: " + re.sub(r"\s+", " ", t[m2.start() - 900: m2.start() + 400]))
            m3 = re.search(r"BCM, of which", t)
            if m3:
                print(f"   [{eng}] national review: " + re.sub(r"\s+", " ", t[m3.start() - 500: m3.start() + 300]))
            break

    print("\n##### B. Tamar Petroleum: full media list (paged)")
    tamar = []
    for page in range(1, 8):
        rr = get(f"https://www.tamarpetroleum.co.il/wp-json/wp/v2/media?per_page=100&page={page}")
        if rr is None or rr.status_code != 200:
            break
        print("      total pages", rr.headers.get("X-WP-TotalPages"), "total", rr.headers.get("X-WP-Total"))
        try:
            js = rr.json()
        except Exception:  # noqa: BLE001
            break
        tamar += [(m.get("date"), m.get("source_url")) for m in js]
        if len(js) < 100:
            break
    rep = sorted(x for x in tamar if x[1] and x[1].lower().endswith(".pdf") and
                 re.search(r"report|רבעון|תקופתי|FS|Financial|Interim|Quarterly|Annual", x[1], re.I) and
                 not re.search(r"שכר|immediate|Reserve|NSAI|עתודות", x[1], re.I))
    for d, u in rep:
        print("        ", d, u)
    print("\n##### B2. Tamar sales sentences per report")
    for d, u in rep[-30:]:
        rr = get(u)
        if rr is None or rr.status_code != 200 or not rr.content.startswith(b"%PDF"):
            continue
        try:
            t = re.sub(r"\s+", " ", fast_text(rr.content, 200))
        except Exception as e:  # noqa: BLE001
            print("    err", e)
            continue
        hits = re.findall(r"[^()]{0,160}BCM \(100%\)[^()]{0,200}", t)
        hits += re.findall(r"[^ ]{0,10}.{0,120}(?:Egypt|Jordan|export|יצוא|מצרים|ירדן).{0,40}BCM.{0,120}", t)[:3]
        hits += re.findall(r".{0,60}BCM.{0,20}100%.{0,140}", t)[:4]
        print("    ", d, u.rsplit("/", 1)[-1])
        for h in hits[:8]:
            print("        |", h[:380])

    print("\n##### C. Energean RNS: Israel gas bcm per period (all trading updates and results)")
    allann = []
    for page in range(1, 8):
        r = get(f"https://www.investegate.co.uk/company/ENOG?page={page}")
        allann += links(r, pat=r"/announcement/")
    allann = sorted(set(a for a in allann if re.search(r"trading-statement|results|9-months|production-and", a)))
    print("    ", len(allann))
    for l in allann:
        r = get(l)
        if r is None or r.status_code != 200:
            continue
        txt = re.sub(r"<[^>]+>", " ", r.text)
        txt = re.sub(r"&nbsp;|&#160;", " ", txt).replace("&#8209;", "-")
        txt = re.sub(r"\s+", " ", txt)
        date = re.search(r"\b(\d{1,2} \w+ 20\d\d)\b", txt)
        hits = re.findall(r"[^.]{0,160}\d\.\d+ bcm[^.]{0,160}(?:Q[1-4]|H[12]|20\d\d|month)[^.]{0,80}", txt)
        hits = [h for h in hits if re.search(r"Israel|includes|sold|equivalent", h, re.I)]
        print("    ", l.rsplit("energean--enog/", 1)[-1], date.group(1) if date else "")
        for h in hits[:6]:
            print("        |", h.strip()[:360])


SENT = r"(?:[^.]|\.\d)"  # sentence char: anything but a full stop, decimals allowed


def round5():
    NM = "https://newmedenergy.com/wp-content/uploads/"
    print("\n##### A. Energean RNS: Israel table rows and bcm sentences")
    ids = ["trading-statement-operational-update/9576511", "results-for-half-year-ended-30-june-2026/9761841",
           "trading-statement-operational-update/9388069", "trading-statement-operational-update/9257244",
           "results-for-half-year-ended-30-june-2025/9100543", "trading-statement-operational-update/8891249",
           "trading-statement-operational-update/8701793", "trading-statement-operational-update/8577177",
           "trading-statement-operational-update/8215919", "2023-full-year-results/8098916"]
    for i in ids:
        r = get("https://www.investegate.co.uk/announcement/rns/energean--enog/" + i)
        if r is None or r.status_code != 200:
            continue
        txt = re.sub(r"<(?:td|th)[^>]*>", " | ", r.text)
        txt = re.sub(r"</tr>", "\n", txt)
        txt = re.sub(r"<[^>]+>", " ", txt)
        txt = re.sub(r"&nbsp;|&#160;", " ", txt).replace("&#183;", "*").replace("&amp;", "&").replace("&#8209;", "-")
        txt = re.sub(r"[ \t]+", " ", txt)
        print("   ", i)
        rows = [l.strip() for l in txt.split("\n") if re.search(r"Israel|Q[1-4] 20\d\d|H[12] 20\d\d|kboed", l)
                and l.count("|") >= 3]
        for l in rows[:10]:
            print("        row|", l[:300])
        flat = re.sub(r"\s+", " ", txt)
        for m in list(re.finditer(SENT + r"{0,200}\d+\.\d+ bcm" + SENT + r"{0,200}", flat))[:8]:
            if re.search(r"Israel|includes|sold|equivalent|Karish", m.group(0)):
                print("        bcm|", m.group(0).strip()[:380])

    print("\n##### B. NewMed FS: Karish royalty volumes + national review + offtake")
    for f in ["2026/08/FS330062026ACCESSIBLE.pdf", "2026/05/FS31032026ACCESSIBLE.pdf", "2026/03/FS31122025ACCESSIBLE.pdf",
              "2025/11/FS30092025ENGACCESSIBLE.pdf", "2025/03/FS31122024ENGACCESSIBLE.pdf", "2024/03/NM_FS31122023_ACCESSIBLE.pdf"]:
        rr = get(NM + f)
        if rr is None or rr.status_code != 200:
            continue
        t = re.sub(r"\s+", " ", fast_text(rr.content))
        print("   ", f)
        for pat in [SENT + r"{0,300}sold by Energean in" + SENT + r"{0,300}",
                    SENT + r"{0,200}(?:Karish|Tanin)" + SENT + r"{0,80}royalt" + SENT + r"{0,80}\d+\.\d+ BCM" + SENT + r"{0,200}",
                    SENT + r"{0,300}\d+\.\d+ BCM, of which" + SENT + r"{0,400}",
                    r"Total Offtake \(BCM\).{0,200}",
                    r"gov\.il/BlobFolder[^ ]{0,120}"]:
            for m in list(re.finditer(pat, t))[:3]:
                print("        |", m.group(0).strip()[:600])

    print("\n##### C. Tamar Petroleum EN: export/domestic split sentences")
    for u in ["https://www.tamarpetroleum.co.il/wp-content/uploads/2026/09/Tamar-Petroleum-FS-Q2_2026.pdf",
              "https://www.tamarpetroleum.co.il/wp-content/uploads/2026/03/Tamar-Petroleum-Financial-Information-2025-Final.pdf"]:
        rr = get(u)
        if rr is None or rr.status_code != 200:
            continue
        t = re.sub(r"\s+", " ", fast_text(rr.content))
        print("   ", u.rsplit("/", 1)[-1])
        for m in list(re.finditer(SENT + r"{0,250}(?:export|Egypt|Jordan|domestic market)" + SENT + r"{0,250}", t))[:40]:
            if re.search(r"\d+\.\d+ ?BCM|BCM ?\d|%", m.group(0)) and re.search(r"sold|sales|quantit", m.group(0)):
                print("        |", m.group(0).strip()[:450])

    print("\n##### D. gov.il BlobFolder: Natural Gas Authority annual reviews")
    cands = []
    for y in range(2019, 2026):
        cands += [f"https://www.gov.il/BlobFolder/reports/ng_{y}/he/ng_{y}.pdf",
                  f"https://www.gov.il/BlobFolder/reports/ng_{y}/he/ng_{y}_summary.pdf",
                  f"https://www.gov.il/BlobFolder/reports/natural_gas_{y}/he/natural_gas_{y}.pdf",
                  f"https://www.gov.il/BlobFolder/reports/ng-{y}/he/ng-{y}.pdf"]
    for u in cands:
        rr = get(u) if creq is None else get(u, "cffi")
        if rr is not None and rr.status_code == 200 and rr.content.startswith(b"%PDF"):
            try:
                t = re.sub(r"\s+", " ", fast_text(rr.content, 80))
            except Exception as e:  # noqa: BLE001
                print("    err", e)
                continue
            for m in list(re.finditer(r".{0,120}(?:תמר|לוויתן|כריש).{0,40}BCM.{0,120}|.{0,120}BCM.{0,40}(?:תמר|לוויתן|כריש).{0,120}", t))[:10]:
                print("        |", m.group(0)[:300])

    print("\n##### E. MAYA files direct (report ids used by NewMed site)")
    for rid in ["1742160", "1764416", "1663875"]:
        lo = (int(rid) - 1) // 1000 * 1000 + 1
        for u in [f"https://mayafiles.tase.co.il/rpdf/{lo}-{lo + 999}/P{rid}-00.pdf",
                  f"https://mayafiles.tase.co.il/RPdf/{lo}-{lo + 999}/P{rid}-00.pdf",
                  f"https://maya.tase.co.il/reports/details/{rid}",
                  f"https://mayaapi.tase.co.il/api/report/details/{rid}"]:
            rr = fetch(u)
            if rr is not None and rr.status_code == 200 and not rr.content.startswith(b"%PDF"):
                print("        head:", rr.text[:200].replace("\n", " "))


def round6():
    print("\n##### A. Tamar FY2025: domestic/export table 2025 + 2024")
    rr = get("https://www.tamarpetroleum.co.il/wp-content/uploads/2026/03/Tamar-Petroleum-Financial-Information-2025-Final.pdf")
    t = re.sub(r"\s+", " ", fast_text(rr.content))
    m = re.search(r"broken down by domestic market and export", t)
    print("   ", t[m.start(): m.start() + 700] if m else "not found")

    print("\n##### B. gov.il Natural Gas Authority annual reviews: supply-by-source and export rows")
    for y, u in [(2022, "ng-2022/he/ng-2022.pdf"), (2023, "ng-2023/he/ng-2023.pdf"), (2024, "ng-2024/he/ng-2024.pdf"),
                 (2025, "ng-2025/he/ng-2025.pdf")]:
        rr = get("https://www.gov.il/BlobFolder/reports/" + u, "cffi")
        if rr is None or rr.status_code != 200:
            continue
        t = re.sub(r"\s+", " ", fast_text(rr.content, 120))
        print("   ", y, f"{len(t):,} chars")
        for pat in [r".{0,80}(?:סה\"כ|סה״כ).{0,20}(?:אספקת|הפקת).{0,200}", r".{0,150}(?:למצרים|לירדן).{0,40}BCM.{0,150}",
                    r".{0,150}BCM.{0,40}(?:למצרים|לירדן).{0,150}", r".{0,100}(?:כריש).{0,30}\d+\.\d+ ?BCM.{0,100}"]:
            for mm in list(re.finditer(pat, t))[:4]:
                print("        |", mm.group(0)[:330])

    print("\n##### C. Energean: text after 'Israel (kboed)' / 'Israel average W.I.'")
    for i in ["trading-statement-operational-update/9576511", "trading-statement-operational-update/9388069",
              "trading-statement-operational-update/9257244", "trading-statement-operational-update/8891249"]:
        r = get("https://www.investegate.co.uk/announcement/rns/energean--enog/" + i)
        if r is None:
            continue
        txt = re.sub(r"<(?:td|th)[^>]*>", " | ", r.text)
        txt = re.sub(r"<[^>]+>", " ", txt)
        txt = re.sub(r"&nbsp;|&#160;", " ", txt)
        txt = re.sub(r"\s+", " ", txt)
        for mm in list(re.finditer(r"(?:Israel average W\.I\. production|Israel \(kboed\))", txt))[:2]:
            print("   ", i, "|", txt[max(0, mm.start() - 250): mm.end() + 250])

    print("\n##### D. MAYA via Playwright: all tase.co.il requests for Tamar Petroleum / Energean pages")
    from playwright.sync_api import sync_playwright
    with sync_playwright() as p:
        b = p.chromium.launch()
        pg = b.new_page(locale="he-IL")
        seen = []
        pg.on("response", lambda resp: seen.append((resp.status, resp.request.method, resp.url)))
        for u in ["https://maya.tase.co.il/he/reports/companies?q=%7B%22Page%22:1,%22Q%22:%22%D7%AA%D7%9E%D7%A8%20%D7%A4%D7%98%D7%A8%D7%95%D7%9C%D7%99%D7%95%D7%9D%22%7D",
                  "https://maya.tase.co.il/he/companies/1919/reports"]:
            try:
                pg.goto(u, timeout=60000, wait_until="networkidle")
                pg.wait_for_timeout(5000)
            except Exception as e:  # noqa: BLE001
                print("   goto", e)
            print("   page", u[:90], "|", pg.inner_text("body")[:800].replace("\n", " | "))
            hrefs = pg.eval_on_selector_all("a", "els => els.map(e => e.href)")
            print("   report links:", [h for h in hrefs if "reports/details" in h][:20])
        for st, me, url in seen:
            if "tase.co.il" in url and not re.search(r"\.(js|css|woff2?|png|svg|ico|jpg)(\?|$)", url):
                print("    ", st, me, url[:220])
        b.close()


def round7():
    """MAYA through Playwright: company ids, the reports POST body and a replay for Tamar Petroleum / Energean."""
    from playwright.sync_api import sync_playwright
    want = re.compile(r"תמר|ניו-?מד|אנרג'?יאן|Energean|רציו|ישראמקו|Tamar|NewMed|Ratio|Isramco", re.I)
    with sync_playwright() as p:
        b = p.chromium.launch()
        pg = b.new_page(locale="he-IL")
        got = {}

        def on_resp(resp):
            u = resp.url
            if "/api/v1/" in u:
                try:
                    got[u + " " + resp.request.method] = (resp.request.post_data, resp.json())
                except Exception:  # noqa: BLE001
                    pass
        pg.on("response", on_resp)
        pg.goto("https://maya.tase.co.il/he/companies/475/reports", timeout=60000, wait_until="networkidle")
        pg.wait_for_timeout(5000)
        post_body = None
        for k, (pd, js) in got.items():
            print("   ", k, "| post:", (pd or "")[:400])
            sj = json.dumps(js, ensure_ascii=False)
            print("      json head:", sj[:600])
            if "metadata/companies" in k:
                items = js if isinstance(js, list) else next((v for v in js.values() if isinstance(v, list)), [])
                print(f"      {len(items)} companies; matches:")
                for it in items:
                    if want.search(json.dumps(it, ensure_ascii=False)):
                        print("        ", json.dumps(it, ensure_ascii=False)[:200])
            if k.endswith("reports/companies POST"):
                post_body = pd
        if post_body:
            body = json.loads(post_body)
            print("   reports POST body keys:", list(body))
            for cid in ["475", "1983", "1141", "2109", "1100"]:
                b2 = json.loads(post_body)
                txt = json.dumps(b2).replace('"475"', f'"{cid}"').replace(": 475", f": {cid}")
                res = pg.evaluate("""async (body) => { const r = await fetch('/api/v1/reports/companies',
                    {method: 'POST', headers: {'Content-Type': 'application/json'}, body: body});
                    return [r.status, (await r.text()).slice(0, 2500)]; }""", txt)
                print(f"   replay company {cid}: HTTP {res[0]} {res[1][:1500]}")
        b.close()


if __name__ == "__main__":
    if ROUND == 7:
        round7()
    elif ROUND == 6:
        round6()
    elif ROUND == 5:
        round5()
    elif ROUND == 4:
        round4()
    elif ROUND == 3:
        round3()
    elif ROUND == 2:
        round2()
        maya_playwright()
    else:
        {1: round1}.get(ROUND, round1)()
