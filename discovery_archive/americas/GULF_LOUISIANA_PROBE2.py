"""Probe 2 (manual, Actions) for americas/GULF_COAST_BALANCE.py: EIA 'U.S. liquefaction capacity' workbook (every sheet, rows that
mention Louisiana projects), EIA Today in Energy LNG articles, FERC 'North American LNG export terminals' attachments, and
company/press pages for Louisiana LNG start dates."""
import io
import re

import pandas as pd
import requests

H = {"User-Agent": "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 Chrome/120 Safari/537.36"}
KW = re.compile(r"Louisiana|, LA|Sabine|Cameron|Calcasieu|Plaquemines|CP2|Commonwealth|Delfin|Lake Charles|Woodside|Venture|Cheniere|Driftwood|Magnolia|Gulf LNG|Argent|Texas LNG|Port Arthur|Golden Pass|Rio Grande|Corpus", re.I)


def get(u, **kw):
    r = requests.get(u, headers=H, timeout=60, **kw)
    print(f"\n## {u} -> {r.status_code} {r.headers.get('content-type', '')[:40]} {len(r.content)}")
    return r


def text(r):
    return re.sub(r"\s+", " ", re.sub(r"<script.*?</script>|<style.*?</style>|<[^>]+>", " ", r.text, flags=re.S))


for u in ["https://www.eia.gov/naturalgas/importsexports/liquefactioncapacity/U.S.liquefactioncapacity_2026_Q2.xlsx",
          "https://www.eia.gov/naturalgas/U.S.liquefactioncapacity.xlsx"]:
    try:
        r = get(u)
        if r.status_code == 200:
            for n, d in pd.read_excel(io.BytesIO(r.content), sheet_name=None, header=None).items():
                print("-- sheet", n, d.shape)
                for _, row in d.iterrows():
                    t = " | ".join(str(v) for v in row if str(v) != "nan")
                    if t and (d.shape[0] < 12 or KW.search(t) or _ < 6):
                        print("   ", t[:700])
            break
    except Exception as e:  # noqa: BLE001
        print("fail", e)

for u in ["https://www.eia.gov/todayinenergy/detail.php?id=68064", "https://www.eia.gov/todayinenergy/detail.php?id=68144"]:
    try:
        r = get(u)
        t = text(r)
        i = t.find("Today in Energy")
        print(t[:6000] if r.status_code == 200 else "")
    except Exception as e:  # noqa: BLE001
        print("fail", e)

u = "https://www.ferc.gov/media/north-american-lng-export-terminals-existing-approved-not-yet-built-and-proposed-8"
r = get(u)
for l in sorted(set(re.findall(r'href="([^"]+)"', r.text))):
    if re.search(r"\.(pdf|xlsx|xls|csv)|/media/|lng", l, re.I):
        print("   link", l[:200])
txt = text(r)
for m in re.finditer(r"Louisiana|Plaquemines|CP2|Commonwealth|Lake Charles|Woodside|Delfin", txt):
    print("   FERCtxt:", txt[max(0, m.start() - 100):m.end() + 200])
    break

# follow FERC attachment links
for l in sorted(set(re.findall(r'href="([^"]+\.(?:pdf|xlsx|xls))"', r.text, re.I)))[:8]:
    if l.startswith("/"):
        l = "https://www.ferc.gov" + l
    try:
        rr = get(l)
        if rr.status_code != 200:
            continue
        if l.lower().endswith(".pdf"):
            from pypdf import PdfReader
            for pi, pg in enumerate(PdfReader(io.BytesIO(rr.content)).pages[:60]):
                for ln in (pg.extract_text() or "").splitlines():
                    if KW.search(ln):
                        print(f"   p{pi + 1}: {ln[:300]}")
        else:
            for n, d in pd.read_excel(io.BytesIO(rr.content), sheet_name=None, header=None).items():
                print("-- sheet", n, d.shape)
                for _, row in d.iterrows():
                    t = " | ".join(str(v) for v in row if str(v) != "nan")
                    if KW.search(t):
                        print("   ", t[:500])
    except Exception as e:  # noqa: BLE001
        print("fail", l, e)

for u in ["https://investors.venturegloballng.com/news-releases", "https://venturegloballng.com/", "https://www.woodsideenergy.com/",
          "https://www.woodside.com/", "https://www.commonwealthlng.com/", "https://www.commonwealthlng.com/project/",
          "https://lakecharleslng.com/", "https://delfinmidstream.com/delfin-lng/", "https://www.cheniere.com/",
          "https://lngir.cheniere.com/news-releases", "https://www.energytransfer.com/media/press-releases/"]:
    try:
        r = get(u)
        if r.status_code != 200:
            continue
        t = text(r)
        for m in list(re.finditer(r"first LNG|first cargo|Plaquemines|CP2|Louisiana LNG|Lake Charles|Commonwealth|Stage 5|Phase 2|FID|final investment", t))[:12]:
            print("   txt:", t[max(0, m.start() - 150):m.end() + 250])
        for l in sorted(set(re.findall(r'href="([^"]*(?:cp2|plaquemines|louisiana-lng|lake-charles|commonwealth|stage-5|expansion|news)[^"]*)"', r.text, re.I)))[:15]:
            print("   link", l[:200])
    except Exception as e:  # noqa: BLE001
        print("fail", u, type(e).__name__)
