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


if __name__ == "__main__":
    if ROUND == 2:
        round2()
        maya_playwright()
    else:
        {1: round1}.get(ROUND, round1)()
