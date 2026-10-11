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
SEC_UA = {"User-Agent": "Energy data research thomasdalton9 github (probe)", "Accept-Encoding": "gzip, deflate"}
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


if __name__ == "__main__":
    {1: round1}.get(ROUND, round1)()
