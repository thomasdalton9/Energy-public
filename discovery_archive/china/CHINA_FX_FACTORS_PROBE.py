"""Manual probe (7-8 Oct 2026): what can GitHub Actions reach for (1) the yuan/US$ rate, (2) the wording of NBS's
10-day price table (is kcal/kg stated, net or gross?), (3) published conversion factors (BP/EI approximate conversion
factors, EIA heat contents, GIIGNL). Writes everything it opens to discovery_archive/results/china_fx/.
"""
import os, re, sys, io
import requests

OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "results", "china_fx")
os.makedirs(OUT, exist_ok=True)
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "..", "asia"))
UA = {"User-Agent": "Mozilla/5.0 (research script)"}
summary = []


def say(*a):
    s = " ".join(str(x) for x in a)
    print(s, flush=True)
    summary.append(s)


def get(name, url, **kw):
    try:
        r = requests.get(url, headers=UA, timeout=40, **kw)
    except Exception as e:  # noqa: BLE001
        say(f"[{name}] {url} -> {type(e).__name__}: {str(e)[:150]}")
        return None
    say(f"[{name}] {url} -> {r.status_code} {r.headers.get('content-type')} {len(r.content)} bytes")
    return r


def text_of(r):
    if r.content[:4] == b"%PDF":
        try:
            import fitz
            doc = fitz.open(stream=r.content, filetype="pdf")
            return "\n".join(f"=== page {i+1} ===\n" + p.get_text() for i, p in enumerate(doc))
        except Exception as e:  # noqa: BLE001
            return f"(pdf text failed {e})"
    return r.text


# 1. FX
r = get("fred_csv", "https://fred.stlouisfed.org/graph/fredgraph.csv?id=DEXCHUS")
if r is not None and r.status_code == 200:
    lines = r.text.splitlines()
    say("  fred head", lines[:3], "tail", lines[-4:], "rows", len(lines))
    open(os.path.join(OUT, "fred_dexchus_sample.csv"), "w").write("\n".join(lines[:5] + lines[-30:]))
r = get("fred_csv_cosd", "https://fred.stlouisfed.org/graph/fredgraph.csv?id=DEXCHUS&cosd=2021-01-01")
if r is not None and r.status_code == 200:
    say("  fred cosd rows", len(r.text.splitlines()))
r = get("ecb_hist", "https://www.ecb.europa.eu/stats/eurofxref/eurofxref-hist.csv")
if r is not None and r.status_code == 200:
    l = r.text.splitlines(); say("  ecb", l[0][:200], l[1][:100])
r = get("ecb_sdmx", "https://data-api.ecb.europa.eu/service/data/EXR/D.CNY.EUR.SP00.A?lastNObservations=3&format=csvdata")
if r is not None: say("  ", r.text[:600])
r = get("ecb_hist2", "https://www.ecb.europa.eu/stats/eurofxref/eurofxref-hist.csv")
if r is not None and r.status_code == 200:
    l = r.text.splitlines(); say("  ecb hist last/first rows", l[1][:60], l[-1][:60])
r = get("h10", "https://www.federalreserve.gov/releases/h10/hist/dat00_ch.htm")
if r is not None and r.status_code == 200:
    t = r.text
    open(os.path.join(OUT, "h10_dat00_ch_head.txt"), "w").write(t[:6000] + "\n.....\n" + t[-6000:])
    say("  h10 len", len(t))
for yy in ("dat00_ch", "dat25_ch", "dat26_ch"):
    get("h10_" + yy, f"https://www.federalreserve.gov/releases/h10/hist/{yy}.htm")
get("pboc", "http://www.pbc.gov.cn/en/3688006/index.html")

# 2. NBS wording
try:
    import china_nbs_common as nbs
    import CHINA_NBS_MARKET_PRICES as mp
    rel = nbs.crawl_index(mp.TITLE_RE, lambda t: True, False, mp.EXTRA_RELEASES, stop_after_known=100)[:1]
    rel += [x for x in mp.EXTRA_RELEASES if "2021年8月下旬" in x[0] or "2021年1月上旬" in x[0]]
    # one release per year around mid-2024 and Dec 2025 (still carried the 4500/5000/5800 blends)
    full = nbs.crawl_index(mp.TITLE_RE, lambda t: ("2025年12月" in t or "2024年6月" in t), True, (), stop_after_known=10**6)
    rel += full[:3]
    for title, url in rel:
        html = nbs.fetch(url)
        say(f"NBS {title} {url} len={len(html or '')}")
        rows = nbs.table_rows(html or "")
        keep = [c for c in rows if re.search(r"煤|焦|液化|汽油|柴油|石蜡|天然气|LNG|LPG", "".join(c))]
        for c in keep:
            say("   ", c)
        for m in re.finditer(r"[^。<>]{0,60}(大卡|发热量|千卡|kcal|注[:：])[^。<>]{0,80}", html or ""):
            say("    NOTE:", m.group(0))
        open(os.path.join(OUT, f"nbs_{url.rsplit('/',1)[-1]}"), "w").write(html or "")
except Exception as e:  # noqa: BLE001
    say("NBS part failed", type(e).__name__, e)

# 3. conversion-factor documents
DOCS = {
 "un_yearbook_2023_conv": "https://unstats.un.org/unsd/energystats/pubs/yearbook/2023/09ii.pdf",
 "un_yearbook_2022_conv": "https://unstats.un.org/unsd/energystats/pubs/yearbook/2022/09ii.pdf",
 "cer_conv": "https://apps.cer-rec.gc.ca/Conversion/conversion-tables.aspx",
 "giignl_annual": "https://giignl.org/annual-report/",
 "giignl_2026": "https://www.giignl.org/annual-report/2026",
 "gasgrid_conv": "https://gasgrid.fi/wp-content/uploads/Annex-13-Conversion-table-2.pdf",
 "eia_mer_pdf": "https://www.eia.gov/totalenergy/data/monthly/pdf/mer.pdf",
 "eia_psm_append": "http://www.eia.gov/petroleum/supply/monthly/pdf/append.pdf",
}
for k, u in DOCS.items():
    r = get(k, u)
    if r is not None and r.status_code == 200:
        t = text_of(r)
        open(os.path.join(OUT, k + ".txt"), "w").write(t[:400000])
        if "html" in (r.headers.get("content-type") or ""):
            for m in re.finditer(r'href="([^"]+\.(?:pdf|xlsx?))"', r.text):
                say("    link", m.group(1))
open(os.path.join(OUT, "probe_log.txt"), "w").write("\n".join(summary))
