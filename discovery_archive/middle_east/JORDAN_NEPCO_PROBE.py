"""
Probe (Oct 2026): Jordan power and gas data - NEPCO, EMRC, MEMR, DOS, opendata.gov.jo - reachable from GitHub Actions?

Usage: python3 JORDAN_NEPCO_PROBE.py <round>
Workflow (manual): discovery_archive/workflows/jordan_nepco_probe.yml (input `round`).

(findings below, per RUN)
"""
import io, json, re, subprocess, sys, time
import requests
requests.packages.urllib3.disable_warnings()
H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36",
     "Accept-Language": "en-US,en;q=0.9"}
try:
    from curl_cffi import requests as creq
except Exception:
    creq = None


def fetch(url, method="GET", data=None, cffi_first=False, timeout=40, **kw):
    """requests, then curl_cffi chrome. Returns (response, how) or (None, err)."""
    tries = ["cffi", "req"] if cffi_first else ["req", "cffi"]
    last = None
    for how in tries:
        try:
            if how == "req":
                r = requests.request(method, url, data=data, headers=H, timeout=(10, timeout), verify=False, **kw)
            else:
                if creq is None:
                    continue
                r = creq.request(method, url, data=data, impersonate="chrome", timeout=timeout, verify=False, **kw)
            if r.status_code < 400 and len(r.content) > 300:
                return r, how
            last = f"{how}:{r.status_code}/{len(r.content)}B"
        except Exception as e:
            last = f"{how}:ERR {type(e).__name__} {str(e)[:100]}"
    return None, last


def text_of(html):
    body = re.sub(r'<script.*?</script>|<style.*?</style>', ' ', html, flags=re.S | re.I)
    return re.sub(r'\s+', ' ', re.sub(r'<[^>]+>', ' ', body))


def links(html, base=""):
    out = []
    for l in re.findall(r'(?:href|src|action)=["\']([^"\']+)["\']', html):
        if re.search(r'\.(xlsx?|csv|pdf|json|asmx|ashx|aspx)|api|stat|load|generat|download|report|energy|electric|gas|px', l, re.I) \
                and not re.search(r'\.(png|jpe?g|gif|css|js|svg|woff2?)(\?|$)', l, re.I):
            out.append(l)
    return sorted(set(out))


def page(url, n=1500, nl=60, **kw):
    r, how = fetch(url, **kw)
    if r is None:
        print(f"\n=== {url}\n  FAIL {how}", flush=True)
        return None
    ct = r.headers.get("content-type", "")
    print(f"\n=== {url}\n  OK via {how} {r.status_code} {len(r.content)}B {ct} final={r.url}", flush=True)
    if "html" in ct or "text" in ct or "json" in ct:
        print("  TEXT:", text_of(r.text)[:n])
        for l in links(r.text)[:nl]:
            print("   L", l[:180])
    return r


def round1():
    for u in [
        "https://www.nepco.com.jo/en/Default.aspx", "https://www.nepco.com.jo/en/Stats.aspx",
        "https://www.nepco.com.jo/en/MaxLoadArchive.aspx", "https://www.nepco.com.jo/en/AnnualReports.aspx",
        "https://www.nepco.com.jo/ar/Stats.aspx", "https://www.nepco.com.jo/ar/MaxLoadArchive.aspx",
        "https://nepco.com.jo/", "https://www.nepco.com.jo/en/DailyLoad.aspx",
        "https://www.emrc.gov.jo/Default/En", "https://www.emrc.gov.jo/", "https://emrc.gov.jo/En/List/Statistics",
        "https://www.memr.gov.jo/Default/En", "https://www.memr.gov.jo/En/List/Statistics",
        "https://www.memr.gov.jo/En/Pages/Energy_Facts_and_Figures", "https://www.memr.gov.jo/En/List/Studies_and_Statistics",
        "https://www.memr.gov.jo/En/List/Annual_Reports",
        "https://dosweb.dos.gov.jo/", "https://dosweb.dos.gov.jo/economic/energy/",
        "https://jorinfo.dos.gov.jo/Databank/pxweb/en/", "https://jorinfo.dos.gov.jo/PXWeb2014R2/Menu.aspx?px_language=en",
        "https://opendata.gov.jo/en/dataset/nepco-2170-2023", "https://opendata.gov.jo/dataset/nepco-2170-2023",
        "https://portal.jordan.gov.jo/wps/portal/OpenData",
    ]:
        page(u)
    for path in ["api/3/action/package_show?id=nepco-2170-2023", "api/3/action/package_search?q=nepco&rows=30",
                 "api/3/action/package_search?q=electricity&rows=30", "api/3/action/organization_list"]:
        r, how = fetch("https://opendata.gov.jo/" + path)
        print(f"\n=== CKAN {path} -> {how if r is None else 'OK '+how}")
        if r is not None:
            try:
                j = r.json()["result"]
                if isinstance(j, dict) and "results" in j:
                    print("  count", j.get("count"))
                    for p in j["results"]:
                        print("  PKG", p["name"], "|", (p.get("organization") or {}).get("name"), "|",
                              [(x.get("format"), x.get("url")) for x in p.get("resources", [])][:5])
                elif isinstance(j, dict):
                    for x in j.get("resources", []):
                        print("  RES", x.get("name"), x.get("format"), x.get("url"))
                else:
                    print(" ", str(j)[:1500])
            except Exception as e:
                print("  not json", r.text[:300])


def forms(html):
    """Print ASP.NET form inputs (names only for hidden), selects with first/last options."""
    for m in re.finditer(r'<input[^>]+>', html):
        tag = m.group(0)
        nm = re.search(r'name=["\']([^"\']+)', tag)
        tp = re.search(r'type=["\']([^"\']+)', tag)
        vl = re.search(r'value=["\']([^"\']*)', tag)
        if nm:
            v = vl.group(1) if vl else ""
            print("   INPUT", tp.group(1) if tp else "?", nm.group(1), (v[:60] + f"..({len(v)})") if len(v) > 60 else v)
    for m in re.finditer(r'<select[^>]*name=["\']([^"\']+)["\'][^>]*>(.*?)</select>', html, re.S):
        opts = re.findall(r'<option[^>]*value=["\']([^"\']*)["\'][^>]*>([^<]*)', m.group(2))
        print("   SELECT", m.group(1), len(opts), opts[:4], opts[-3:])
    for m in re.findall(r'<table.*?</table>', html, re.S)[:4]:
        print("   TABLE", text_of(m)[:1200])


def pdf_text(url, kw=r'GWh|MW|generat|توليد|MMSCF|gas|غاز|import|export|peak|load', maxpages=400, n=1500, showpages=12):
    r, how = fetch(url, timeout=120)
    if r is None:
        print(f"\n=== PDF {url}\n  FAIL {how}"); return None
    print(f"\n=== PDF {url}\n  OK via {how} {len(r.content)}B {r.headers.get('content-type')} last-mod={r.headers.get('last-modified')}", flush=True)
    import pdfplumber
    shown = 0
    try:
        with pdfplumber.open(io.BytesIO(r.content)) as pdf:
            print("  pages", len(pdf.pages))
            for i, pg in enumerate(pdf.pages[:maxpages]):
                t = pg.extract_text() or ""
                if i == 0:
                    print("  P1:", re.sub(r'\s+', ' ', t)[:400])
                hits = len(re.findall(kw, t, re.I))
                if hits >= 3 and shown < showpages:
                    shown += 1
                    print(f"  --- p{i+1} ({hits} hits)\n" + t[:n])
    except Exception as e:
        print("  PDF ERR", e)
    return r


def round2():
    for u in ["https://nepco.com.jo/Stats.aspx", "https://nepco.com.jo/MaxLoadArchive.aspx", "https://nepco.com.jo/AnnualReports.aspx",
              "https://nepco.com.jo/Interconnection.aspx", "https://nepco.com.jo/JordanElecticSystem.aspx",
              "https://www.nepco.com.jo/en/Stats.aspx", "https://www.nepco.com.jo/en/MaxLoadArchive.aspx",
              "https://www.nepco.com.jo/en/AnnualReports.aspx", "https://www.nepco.com.jo/en/Interconnection.aspx",
              "https://nepco.com.jo/en/Stats.aspx", "https://nepco.com.jo/en/MaxLoadArchive.aspx"]:
        r = page(u, n=2500, nl=80)
        if r is not None:
            forms(r.text)
    r = page("https://nepco.com.jo/", n=200, nl=0)
    if r is not None:
        for m in re.finditer(r'الحمل[^<]{0,160}', r.text):
            print("   PEAKNEWS", m.group(0)[:200])
    for u in ["https://nepco.com.jo/assets/doc/peak.pdf", "https://nepco.com.jo/assets/doc/elec10-2025-0.pdf"]:
        pdf_text(u, showpages=6)


if __name__ == "__main__":
    rnd = int(sys.argv[1]) if len(sys.argv) > 1 else 1
    globals()[f"round{rnd}"]()
