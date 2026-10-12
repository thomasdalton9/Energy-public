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


ML = "https://www.nepco.com.jo/en/MaxLoadArchive.aspx"
P = "ctl00$ContentPlaceHolder1$"


def hidden(html):
    return {k: v for k, v in re.findall(r'<input[^>]+type="hidden"[^>]+name="([^"]+)"[^>]*value="([^"]*)"', html)}


def rows(html):
    return re.findall(r'(\d{2}/\d{2}/\d{4})\s*</[^>]+>\s*<[^>]+>\s*(\d+(?:\.\d+)?)', html) or \
        re.findall(r'(\d{2}/\d{2}/\d{4})\D{1,200}?(\d{3,4})\b', text_of(html))


def maxload_query(sess, d1, m1, y1, d2, m2, y2):
    r = sess.get(ML, headers=H, timeout=40, verify=False)
    f = hidden(r.text)
    f.update({P + "txtDayFrom": str(d1), P + "fromMonth": str(m1), P + "txtYearFrom": str(y1),
              P + "txtDayTo": str(d2), P + "toMonth": str(m2), P + "txtYearTo": str(y2), P + "btnSearch": "Search"})
    r2 = sess.post(ML, data=f, headers=H, timeout=60, verify=False)
    return r2


def round3():
    s = requests.Session()
    r = s.get(ML, headers=H, timeout=40, verify=False)
    print("GET rows", rows(r.text)[:12])
    pagers = sorted(set(re.findall(r"__doPostBack\(&#39;([^&]+)&#39;", r.text)))
    print("pager targets", pagers)
    for (y) in [1995, 2000, 2003, 2005, 2008, 2010, 2012, 2015, 2018, 2020, 2022, 2024, 2025]:
        try:
            r2 = maxload_query(s, 1, 7, y, 31, 7, y)
            rr = rows(r2.text)
            print(f"Jul {y}: status {r2.status_code} {len(r2.content)}B rows {len(rr)} {rr[:10]}", flush=True)
        except Exception as e:
            print("ERR", y, e)
    # whole range, then follow 'Last' pager to the oldest
    r2 = maxload_query(s, 1, 1, 1990, 31, 12, 2026)
    rr = rows(r2.text)
    print("ALL first page", rr[:10])
    txt = text_of(r2.text)
    m = re.search(r'Start.{0,300}', txt)
    print("pager text", m.group(0) if m else None)
    last = [t for t in re.findall(r"__doPostBack\(&#39;([^&]+)&#39;", r2.text)]
    print("pager targets after search", sorted(set(last))[-6:])
    if last:
        f = hidden(r2.text)
        f.update({"__EVENTTARGET": sorted(set(last))[-1], "__EVENTARGUMENT": ""})
        for k in ["txtDayFrom", "fromMonth", "txtYearFrom", "txtDayTo", "toMonth", "txtYearTo"]:
            m2 = re.search(r'name="' + re.escape(P + k) + r'"[^>]*value="([^"]*)"', r2.text)
            if m2: f[P + k] = m2.group(1)
        f[P + "fromMonth"] = "1"; f[P + "toMonth"] = "12"
        r3 = s.post(ML, data=f, headers=H, timeout=60, verify=False)
        print("LAST page rows", rows(r3.text)[:12])
    # Stats page raw content
    r = requests.get("https://www.nepco.com.jo/en/Stats.aspx", headers=H, timeout=40, verify=False)
    i = r.text.find("Statistical Info", r.text.find("<body"))
    seg = r.text[i:i + 6000]
    print("STATS RAW:", re.sub(r'\s+', ' ', seg)[:4000])
    for m in re.findall(r'<(?:iframe|img|embed|object|canvas)[^>]+>', r.text):
        print("  EMBED", m[:200])


def month_rows(sess, y, m, maxpages=6):
    """All rows of one month via the search form plus DataPager page-number postbacks."""
    r = maxload_query(sess, 1, m, y, 31, m, y)
    got = dict(rows(r.text))
    html = r.text
    for pg in range(2, maxpages + 1):
        targets = re.findall(r"__doPostBack\(&#39;([^&]+)&#39;,&#39;&#39;\)\">\s*" + str(pg) + r"\s*<", html)
        if not targets:
            break
        f = hidden(html)
        f.update({"__EVENTTARGET": targets[0], "__EVENTARGUMENT": ""})
        for k, v in [("txtDayFrom", 1), ("fromMonth", m), ("txtYearFrom", y), ("txtDayTo", 31), ("toMonth", m), ("txtYearTo", y)]:
            f[P + k] = str(v)
        r = sess.post(ML, data=f, headers=H, timeout=60, verify=False)
        html = r.text
        new = dict(rows(html))
        if not set(new) - set(got):
            break
        got.update(new)
    return got


def round4():
    s = requests.Session()
    r = s.get(ML, headers=H, timeout=40, verify=False)
    i = r.text.find("DataPager1")
    print("PAGER HTML:", re.sub(r'\s+', ' ', r.text[i - 200:i + 1500]))
    for y, m in [(2019, 1), (2019, 6), (2019, 12), (2020, 1), (2020, 2), (2020, 3), (2020, 6), (2021, 1), (2023, 8), (2025, 8), (2026, 9), (2026, 10)]:
        try:
            g = month_rows(s, y, m)
            ks = sorted(g, key=lambda d: d[6:] + d[3:5] + d[:2])
            print(f"{y}-{m:02d}: {len(g)} days; first {ks[:1]} last {ks[-1:]}; max {max(map(float, g.values())) if g else None}; days {[k[:2] for k in ks]}", flush=True)
        except Exception as e:
            print("ERR", y, m, e)
    # Stats images: save for OCR / artifact
    import os
    os.makedirs("probe_out", exist_ok=True)
    for lang in ["en", "ar"]:
        for k in range(1, 7):
            u = f"https://www.nepco.com.jo/assets/img/stats/{lang}/stat-{k}.png?V=2026.2"
            rr, how = fetch(u)
            if rr is None:
                print("IMG", u, how); continue
            fn = f"probe_out/stat_{lang}_{k}.png"
            open(fn, "wb").write(rr.content)
            print("IMG", u, len(rr.content), rr.headers.get("last-modified"))
            try:
                o = subprocess.run(["tesseract", fn, "-", "--psm", "6"], capture_output=True, text=True, timeout=120).stdout
                print("  OCR:", re.sub(r'\s+', ' ', o)[:1500])
            except Exception as e:
                print("  OCR ERR", e)
    # NEPCO annual reports: size / pages, and the statistics pages of 2024
    for y in ["2024_EN", "2023_EN", "2019_en"]:
        u = f"https://www.nepco.com.jo/store/DOCS/web/{y}.pdf"
        rr, how = fetch(u, timeout=180)
        if rr is not None:
            open(f"probe_out/nepco_{y}.pdf", "wb").write(rr.content)
    for y in ["2010_en", "2015_en"]:
        u = f"https://www.nepco.com.jo/store/docs/web/{y}.pdf"
        rr, how = fetch(u, timeout=180)
        print("AR", y, how if rr is None else (len(rr.content), rr.headers.get("last-modified")))
        if rr is not None:
            open(f"probe_out/nepco_{y}.pdf", "wb").write(rr.content)
    import pdfplumber
    kw = r'GWh|Generated|Generation|Natural Gas|Imported|Exported|Peak Load|Fuel|MMSCF|Egypt|Interconnect'
    for fn in sorted(os.listdir("probe_out")):
        if not fn.endswith(".pdf"):
            continue
        with pdfplumber.open("probe_out/" + fn) as pdf:
            print(f"\n##### {fn} pages {len(pdf.pages)}")
            for i, pg in enumerate(pdf.pages):
                t = pg.extract_text() or ""
                h = len(re.findall(kw, t))
                if h >= 4:
                    print(f"  p{i+1} hits {h}: " + re.sub(r'\s+', ' ', t)[:160])


def page_targets(html):
    return {int(n): t for t, n in re.findall(r"__doPostBack\(&#39;([^&]+)&#39;,&#39;&#39;\)[\"']>\s*(\d+)\s*<", html)}


def month_rows2(sess, y, m, verbose=False):
    r = maxload_query(sess, 1, m, y, 31, m, y)
    html = r.text
    got = dict(rows(html))
    pg = 1
    while True:
        t = page_targets(html)
        if verbose:
            print("   targets", sorted(t)[:12])
        if pg + 1 not in t:
            break
        f = hidden(html)
        f.update({"__EVENTTARGET": t[pg + 1], "__EVENTARGUMENT": ""})
        for k, v in [("txtDayFrom", 1), ("fromMonth", m), ("txtYearFrom", y), ("txtDayTo", 31), ("toMonth", m), ("txtYearTo", y)]:
            f[P + k] = str(v)
        r = sess.post(ML, data=f, headers=H, timeout=60, verify=False)
        html = r.text
        new = dict(rows(html))
        if verbose:
            print("   page", pg + 1, list(new.items())[:3])
        if not set(new) - set(got):
            break
        got.update(new)
        pg += 1
    return got


def pages_text(fn, pagelist, n=3500):
    import pdfplumber
    with pdfplumber.open(fn) as pdf:
        for i in pagelist:
            if i - 1 < len(pdf.pages):
                print(f"\n----- {fn} p{i}\n" + (pdf.pages[i - 1].extract_text() or "")[:n])


def kw_pages(fn, kw, minhits=3, maxshow=8, n=2500):
    import pdfplumber
    shown = 0
    with pdfplumber.open(fn) as pdf:
        print(f"\n##### {fn} pages {len(pdf.pages)}")
        for i, pg in enumerate(pdf.pages):
            t = pg.extract_text() or ""
            h = len(re.findall(kw, t, re.I))
            if h >= minhits and shown < maxshow:
                shown += 1
                print(f"\n----- p{i+1} hits {h}\n" + t[:n])


def save(url, fn, timeout=180):
    import os
    os.makedirs("probe_out", exist_ok=True)
    r, how = fetch(url, timeout=timeout)
    print("SAVE", url, how if r is None else (len(r.content), r.headers.get("content-type"), r.headers.get("last-modified")), flush=True)
    if r is None:
        return None
    open("probe_out/" + fn, "wb").write(r.content)
    return "probe_out/" + fn


def round5():
    s = requests.Session()
    g = month_rows2(s, 2026, 9, verbose=True)
    print("2026-09", len(g), sorted(g)[:40])
    for y, m in [(2019, 11), (2019, 12), (2020, 1), (2023, 8), (2025, 8)]:
        g = month_rows2(s, y, m)
        ks = sorted(g, key=lambda d: d[6:] + d[3:5] + d[:2])
        print(f"{y}-{m:02d}: {len(g)} days; {[k[:2] for k in ks]}; max {max(map(float, g.values())) if g else None}", flush=True)
    # NEPCO annual report tables
    f24 = save("https://www.nepco.com.jo/store/DOCS/web/2024_EN.pdf", "n24.pdf")
    pages_text(f24, [25, 29, 30, 37])
    kw_pages(f24, r'January|February|Monthly|monthly', minhits=3, maxshow=4, n=1500)
    f10 = save("https://www.nepco.com.jo/store/docs/web/2010_en.pdf", "n10.pdf")
    pages_text(f10, [19, 20, 21], n=2000)
    r = page("https://nepco.com.jo/AnnualReports.aspx", n=100, nl=0)
    if r is not None:
        print("  AR links", sorted(set(re.findall(r'href="([^"]*store[^"]+)"', r.text, re.I))))
    # MEMR
    base = "https://www.memr.gov.jo/ebv4.0/root_storage/en/eb_list_page/"
    for fn in ["jordan_energy_balance.pdf", "electricity_consumption.pdf", "general_indecators.pdf"]:
        f = save(base + fn, fn)
        if f:
            kw_pages(f, r'GWh|ktoe|gas|electric|Egypt|LNG|generat', minhits=2, maxshow=4, n=2500)
    f = save(base + "annual_report_2024.pdf", "memr24.pdf")
    if f:
        kw_pages(f, r'natural gas|LNG|Egypt|Aqaba|MMSCF|BCM|GWh|generat|Leviathan|Israel|Noble', minhits=4, maxshow=8, n=2200)
    for u in ["https://www.memr.gov.jo/En/List/Energy_Balance", "https://www.memr.gov.jo/En/Pages/Energy_Balance",
              "https://www.memr.gov.jo/En/List/Open_Data", "https://www.memr.gov.jo/Ar/List/الدراسات_والاحصائيات"]:
        page(u, n=300, nl=40)
    # EMRC
    for u in ["https://www.emrc.gov.jo/EN/Pages/Electricity_Sector", "https://www.emrc.gov.jo/En/List/Open_Data",
              "https://www.emrc.gov.jo/En/List/Annual_Reports", "https://www.emrc.gov.jo/En/Pages/Open_Data",
              "https://www.emrc.gov.jo/Ar/List/التقارير_السنوية"]:
        page(u, n=300, nl=50)
    # DOS
    for u in ["https://jorinfo.dos.gov.jo/Databank/pxweb/en/DOS_Database/", "https://dosweb.dos.gov.jo/industry/industry-and-energy/",
              "https://jorinfo.dos.gov.jo/View_ind/ind_en.aspx", "https://jorinfo.dos.gov.jo/Databank/pxweb/en/Environment/"]:
        page(u, n=800, nl=50)


def round6():
    import calendar, collections
    s = requests.Session()
    # daily peak via <=9-day windows (each answer fits the 10-row first page)
    peryear = collections.defaultdict(dict)
    t0 = time.time()
    for y in range(2019, 2027):
        for m in range(1, 13):
            if (y, m) < (2019, 10) or (y, m) > (2026, 10):
                continue
            nd = calendar.monthrange(y, m)[1]
            for d1, d2 in [(1, 9), (10, 18), (19, 27), (28, nd)]:
                try:
                    r = maxload_query(s, d1, m, y, d2, m, y)
                    rr = dict(rows(r.text))
                    if len(rr) >= 10:
                        print("  WARN window full", y, m, d1, d2)
                    peryear[y].update(rr)
                except Exception as e:
                    print("  ERR", y, m, d1, e)
        if peryear[y]:
            ks = sorted(peryear[y], key=lambda d: d[6:] + d[3:5] + d[:2])
            vals = [float(peryear[y][k]) for k in ks]
            bym = collections.Counter(k[3:5] for k in ks)
            print(f"{y}: {len(ks)} days, first {ks[0]} last {ks[-1]}, max {max(vals)} on {ks[vals.index(max(vals))]}, min {min(vals)}; days/month {dict(sorted(bym.items()))} ({time.time()-t0:.0f}s)", flush=True)
    # weekday of missing days in 2025
    import datetime as dt
    have = set(peryear[2025])
    miss = [dt.date(2025, 1, 1) + dt.timedelta(i) for i in range(365)]
    miss = [d for d in miss if d.strftime("%d/%m/%Y") not in have]
    print("2025 missing", len(miss), collections.Counter(d.strftime("%a") for d in miss))
    round6b()


def round6b():
    # NEPCO annual reports: 2025, Arabic
    for u in ["https://www.nepco.com.jo/store/DOCS/web/2025_EN.pdf", "https://www.nepco.com.jo/store/DOCS/web/2025_AR.pdf",
              "https://www.nepco.com.jo/store/DOCS/web/2024_AR.pdf", "https://www.nepco.com.jo/store/docs/web/2024_ar.pdf"]:
        r, how = fetch(u, timeout=60)
        print("AR?", u, how if r is None else (len(r.content), r.headers.get("content-type"), r.headers.get("last-modified")))
    f24 = save("https://www.nepco.com.jo/store/DOCS/web/2024_EN.pdf", "n24.pdf")
    kw_pages(f24, r'Export|Exported|Lebanon|Palestin|Iraq|Syria', minhits=3, maxshow=5, n=1800)
    kw_pages(f24, r'Peak Load|peak load|MW', minhits=4, maxshow=3, n=1500)
    # MEMR annual report: gas volumes by source
    f = save("https://www.memr.gov.jo/ebv4.0/root_storage/en/eb_list_page/annual_report_2024.pdf", "memr24.pdf")
    kw_pages(f, r'cubic feet|MMSCF|mmscf|Egyptian gas|Northern gas|LNG|Risha', minhits=3, maxshow=8, n=2500)
    f = save("https://www.memr.gov.jo/ebv4.0/root_storage/en/eb_list_page/annual_report_2019.pdf", "memr19.pdf")
    kw_pages(f, r'cubic feet|MMSCF|mmscf|Egyptian gas|LNG|Risha|GWh', minhits=3, maxshow=5, n=2000)
    for fn in ["abc1-0.pdf", "aghihj2025_v2-0.pdf", "import_of_crude_oil_and_petroleum_products.pdf"]:
        f = save("https://www.memr.gov.jo/ebv4.0/root_storage/en/eb_list_page/" + fn, fn)
        if f:
            pages_text(f, [1, 2], n=1200)
    # MEMR Arabic statistics list (energy facts and figures)
    for u in ["https://www.memr.gov.jo/Ar/List/الدراسات_والإحصائيات", "https://www.memr.gov.jo/Ar/List/Studies_and_Statistics",
              "https://www.memr.gov.jo/Ar/List/التقارير_السنوية", "https://www.memr.gov.jo/Ar/List/Annual_Reports"]:
        r = page(u, n=200, nl=0)
        if r is not None:
            for l in sorted(set(re.findall(r'href="([^"]+\.pdf)"', r.text, re.I))):
                if "eb_list_page" in l.lower():
                    print("   PDF", l)
    # EMRC electricity sector list
    r = page("https://www.emrc.gov.jo/EN/List/Electricity_Sector", n=200, nl=0)
    if r is not None:
        for l in sorted(set(re.findall(r'href="([^"]+\.(?:pdf|xlsx?))"', r.text, re.I))):
            print("   DOC", l)
    r = page("https://www.emrc.gov.jo/Ar/List/قطاع_الكهرباء", n=200, nl=0)
    # DOS
    page("https://jorinfo.dos.gov.jo/Databank/pxweb/en/DOS_Database/START__10__1001__1101/", n=1500, nl=40)
    for u in ["https://jorinfo.dos.gov.jo/Databank/api/v1/en/DOS_Database", "https://jorinfo.dos.gov.jo/Databank/api/v1/en/",
              "https://jorinfo.dos.gov.jo/PXWeb/api/v1/en/"]:
        r, how = fetch(u)
        print("PXAPI", u, how if r is None else (r.status_code, r.text[:600]))
    f = save("https://dosweb.dos.gov.jo/Databank/Metadata/Industry _Energy .xlsx", "dos_ind_energy.xlsx")
    if f:
        import openpyxl
        wb = openpyxl.load_workbook(f, read_only=True)
        for ws in wb.worksheets[:3]:
            print("SHEET", ws.title)
            for i, row in enumerate(ws.iter_rows(values_only=True)):
                if i > 25: break
                print("  ", [c for c in row if c is not None][:8])
    # opendata NEPCO 'Jordanian electrical system 2023'
    f = save("https://opendata.gov.jo/dataset/MIGRATED-2170-2023/resource/6e1f8110-cb08-46ac-98fd-1200b5754d04/download/-2023.pdf", "od_system.pdf")
    if f:
        pages_text(f, [1, 2, 3], n=1200)


def round7():
    """Run the document checks first, then a THROTTLED daily-peak scan (3 s between requests)."""
    import calendar, collections
    try:
        round6b()
    except Exception as e:
        print("round6b ERR", e)
    s = requests.Session()
    allrows = {}
    t0 = time.time()
    for y in range(2020, 2027):
        for m in range(1, 13):
            if (y, m) > (2026, 10) or time.time() - t0 > 2400:
                continue
            nd = calendar.monthrange(y, m)[1]
            got = {}
            note = ""
            for d1, d2 in [(1, 9), (10, 18), (19, 27), (28, nd)]:
                time.sleep(3)
                try:
                    r = maxload_query(s, d1, m, y, d2, m, y)
                    rr = dict(rows(r.text))
                    if "Max Load" not in r.text:
                        note += " NOTABLE"
                    if len(rr) >= 10:
                        note += " FULL"
                    got.update(rr)
                except Exception as e:
                    note += " ERR:" + type(e).__name__
                    time.sleep(20)
            allrows.update(got)
            ks = sorted(got)
            vals = [float(v) for v in got.values()]
            print(f"{y}-{m:02d}: {len(got)}/{nd} days, max {max(vals) if vals else None} min {min(vals) if vals else None}{note} ({time.time()-t0:.0f}s)", flush=True)
    import datetime as dt
    for y in (2024, 2025):
        days = [dt.date(y, 1, 1) + dt.timedelta(i) for i in range(366 if y == 2024 else 365)]
        miss = [d for d in days if d.strftime("%d/%m/%Y") not in allrows]
        print(y, "missing", len(miss), collections.Counter(d.strftime("%a") for d in miss))
    big = {k: v for k, v in allrows.items() if float(v) > 5000 or float(v) < 1000}
    print("outliers", big)


def round8():
    """Slow scan (10 s between form queries) of 2021-2026 samples; NEPCO 2025 Arabic annual report gas tables."""
    import calendar, collections, datetime as dt
    s = requests.Session()
    allrows = {}
    n = 0
    t0 = time.time()
    months = [(y, m) for y in (2021, 2022, 2023, 2024) for m in (1, 7)] + [(2025, m) for m in range(1, 13)] + [(2026, m) for m in range(1, 11)]
    for y, m in months:
        nd = calendar.monthrange(y, m)[1]
        got, note = {}, ""
        for d1, d2 in [(1, 9), (10, 18), (19, 27), (28, nd)]:
            time.sleep(10)
            n += 1
            try:
                r = maxload_query(s, d1, m, y, d2, m, y)
                if "Max Load" not in r.text:
                    note += f" NOTABLE@q{n}"
                got.update(dict(rows(r.text)))
            except Exception as e:
                note += f" ERR@q{n}:" + type(e).__name__
        allrows.update(got)
        vals = [float(v) for v in got.values()]
        print(f"{y}-{m:02d}: {len(got)}/{nd} days, max {max(vals) if vals else None} min {min(vals) if vals else None}{note} ({time.time()-t0:.0f}s)", flush=True)
    days = [dt.date(2025, 1, 1) + dt.timedelta(i) for i in range(365)]
    miss = [d for d in days if d.strftime("%d/%m/%Y") not in allrows]
    print("2025 missing", len(miss), collections.Counter(d.strftime("%a") for d in miss), [d.isoformat() for d in miss][:60])
    v25 = {k: float(v) for k, v in allrows.items() if k.endswith("2025")}
    if v25:
        k = max(v25, key=v25.get)
        print("2025 max", k, v25[k], "mean", sum(v25.values()) / len(v25))
    print("outliers", {k: v for k, v in allrows.items() if float(v) > 5000 or float(v) < 1000})
    # NEPCO 2025 Arabic annual report
    f = save("https://www.nepco.com.jo/store/DOCS/web/2025_AR.pdf", "n25ar.pdf")
    if f:
        import pdfplumber
        with pdfplumber.open(f) as pdf:
            print("pages", len(pdf.pages))
            shown = 0
            for i, pg in enumerate(pdf.pages):
                t = pg.extract_text() or ""
                # Arabic glyph order is often reversed by pdfplumber: match both orders
                if re.search(r'غاز|زاغ|GWh|ك.و.س|س.و.ك', t) and len(re.findall(r'\d{3,}', t)) > 15 and shown < 10:
                    shown += 1
                    print(f"\n----- p{i+1}\n" + t[:1800])


if __name__ == "__main__":
    rnd = int(sys.argv[1]) if len(sys.argv) > 1 else 1
    globals()[f"round{rnd}"]()
