"""
India daily renewable generation discovery, round 2.

Round 1 (INDIA_DISCOVERY1): NPP publishes no RE report (dgr1..dgr17 conventional; dgr17 = all-India by fuel incl.
lignite, naphtha, Bhutan import); CEA admin-ajax lists 2,380 daily RE PDFs, 2019-05-01..2025-11-18, then stops;
CEA monthly RE generation PDFs (Renewable Project Monitoring Division) continue to Aug 2026; Grid-India, RLDCs,
Vidyut Pravah, MERIT all reset / time out; India Data Portal CKAN has CEA daily RE csv (to 2024).
This round:
  CEA   daily RE PDF layout (2019, 2022, 2025 samples); guessed daily_reports file names after 2025-11-18;
        monthly RE PDF page 3 (all-India by source) full text; the listing page's resd links
  ICED  (NITI Aayog energy dashboard) API endpoints from its JS bundles
  NPP   landing-page dashboard data (web-db.js etc.)
  IDP   daily RE csv head / tail / yearly sums
  Wayback CDX of Grid-India PSP files; powermin.gov.in; grid-india over http
"""
import io
import re
from datetime import date, timedelta

import pandas as pd
import requests
import urllib3

urllib3.disable_warnings()
H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 "
                   "Safari/537.36", "Accept": "*/*"}
T = (15, 90)


def out(*a):
    print(*a, flush=True)


def get(u, quiet=False, **kw):
    try:
        r = requests.get(u, headers=H, timeout=T, verify=False, **kw)
        if not quiet or r.ok:
            out(f"  GET {u} -> {r.status_code} {len(r.content)} {r.headers.get('content-type', '')[:40]}")
        return r
    except Exception as e:  # noqa: BLE001
        out(f"  GET {u}: {type(e).__name__}: {str(e)[:150]}")
        return None


def pdf_pages(content, pages, n=2500):
    import pdfplumber
    with pdfplumber.open(io.BytesIO(content)) as p:
        out(f"    pdf pages: {len(p.pages)}")
        for i in pages:
            if i < len(p.pages):
                pg = p.pages[i]
                out(f"    --- page {i + 1} text:\n" + (pg.extract_text() or "")[:n])
                tb = pg.extract_tables()
                for t in tb[:2]:
                    out(f"    table ({len(t)} rows): " + " || ".join(" | ".join(str(c) for c in r) for r in t[:30])[:n])


def cea_daily():
    out("\n#### CEA daily RE PDFs")
    r = get("https://cea.nic.in/wp-admin/admin-ajax.php?action=getpostsfordatatables&code=renewable")
    rows = sorted(r.json().get("data", []), key=lambda x: x.get("date", "")) if r is not None and r.ok else []
    bad = [x for x in rows if x.get("link") == "file_not_found"]
    out(f"  {len(rows)} rows, {len(bad)} file_not_found; first ok: "
        f"{next((x['date'] for x in rows if x.get('link') != 'file_not_found'), None)}")
    by_year = {}
    for x in rows:
        by_year.setdefault(x["date"][:4], [0, 0])[x.get("link") == "file_not_found"] += 1
    out(f"  per year [ok, missing]: {by_year}")
    exts = {}
    for x in rows:
        e = x.get("link", "").split("^")[0].rsplit(".", 1)[-1].lower()
        exts[e] = exts.get(e, 0) + 1
    out(f"  link extensions: {exts}")
    for want in ("2019-06-15", "2021-03-10", "2023-08-01", "2025-11-18"):
        x = next((x for x in rows if x["date"] >= want and x.get("link") != "file_not_found"), None)
        if not x:
            continue
        link = x["link"].split("^")[0]
        out(f"\n  sample {x['date']}: {link}")
        c = get(link)
        if c is None or not c.ok:
            continue
        if link.lower().endswith(".pdf"):
            pdf_pages(c.content, [0, 1], 3000)
        else:
            try:
                df = pd.read_excel(io.BytesIO(c.content), header=None)
                out(df.head(60).to_string(max_colwidth=25)[:5000])
            except Exception as e:  # noqa: BLE001
                out(f"    {e}")
    out("\n  guessed names after 2025-11-18:")
    hits = 0
    for d in [date(2025, 11, 19) + timedelta(days=k) for k in range(0, 320, 7)] + [date(2026, 9, 30), date(2026, 10, 1)]:
        mon = d.strftime("%b")
        for name in (f"{d.day:02d}_{mon}_{d.year}_Daily_RE_Generation_Report.pdf",
                     f"{d.day:02d}_{mon}_{d.year}_Daily_RE_Generation.pdf", f"{d.day:02d}_{mon}_{d.year}_merged.pdf"):
            r = get(f"https://cea.nic.in/wp-content/uploads/daily_reports/{name}", quiet=True)
            if r is not None and r.ok and r.content[:4] == b"%PDF":
                hits += 1
    out(f"  guessed hits: {hits}")


def cea_monthly():
    out("\n#### CEA monthly RE report")
    r = get("https://cea.nic.in/renewable-generation-report/?lang=en")
    if r is not None and r.ok:
        ls = sorted(set(re.findall(r'href=["\']([^"\']+)["\']', r.text)))
        out("  resd/pdf links: " + str([x for x in ls if re.search(r"resd|upload.*(pdf|xls)", x, re.I)][:80]))
        out("  ajax/code: " + str(sorted(set(re.findall(r"(?:action|code)\W{1,4}([A-Za-z_\-]{3,40})", r.text)))[:40]))
    for u in ("https://cea.nic.in/wp-content/uploads/resd/2026/08/Monthly_RE_Generation_report_August_2026.pdf",
              "https://cea.nic.in/wp-content/uploads/resd/2025/12/Monthly_RE_Generation_report_December_2025.pdf"):
        c = get(u)
        if c is not None and c.ok:
            pdf_pages(c.content, [1, 2], 4000)


def iced():
    out("\n#### ICED (NITI Aayog)")
    for u in ("https://iced.niti.gov.in/", "https://iced.niti.gov.in/energy/electricity/generation",
              "https://iced.niti.gov.in/energy"):
        r = get(u)
        if r is None or not r.ok:
            continue
        js = sorted(set(re.findall(r'src=["\']([^"\']+\.js[^"\']*)["\']', r.text)))
        out("    js: " + str(js[:30]))
        for j in js[:25]:
            ju = j if j.startswith("http") else "https://iced.niti.gov.in" + ("" if j.startswith("/") else "/") + j
            jr = get(ju, quiet=True)
            if jr is not None and jr.ok:
                api = sorted(set(re.findall(r'["\'`]((?:https?://[^"\'`\s]*)?/?api[^"\'`\s]{2,120})["\'`]', jr.text)))
                gen = sorted(set(re.findall(r'["\'`]([^"\'`\s]{0,80}(?:generation|daily|wind|solar)[^"\'`\s]{0,80})["\'`]',
                                            jr.text, re.I)))
                if api or gen:
                    out(f"    {ju}: api {api[:40]}\n      gen {gen[:40]}")
        break


def npp_home():
    out("\n#### NPP landing dashboard")
    r = get("https://npp.gov.in/")
    if r is None or not r.ok:
        return
    js = sorted(set(re.findall(r'src=["\']([^"\']+\.js[^"\']*)["\']', r.text)))
    out("    js: " + str(js))
    for m in re.findall(r'(?:url|ajax|fetch|getJSON)\s*[:(]\s*["\']([^"\']+)', r.text)[:40]:
        out(f"    endpoint in page: {m}")
    for j in js:
        if "jquery" in j or "bootstrap" in j:
            continue
        jr = get("https://npp.gov.in" + j if j.startswith("/") else j, quiet=True)
        if jr is not None and jr.ok:
            eps = sorted(set(re.findall(r'(?:url|ajax|fetch|getJSON|get|post)\s*[:(]\s*["\']([^"\']{3,150})', jr.text)))
            out(f"    {j}: {eps[:40]}")
    for u in ("https://npp.gov.in/dashBoard/rs-map", "https://npp.gov.in/powerGenerationDashboard",
              "https://npp.gov.in/getAllIndiaGeneration", "https://npp.gov.in/dashBoard"):
        rr = get(u)
        if rr is not None and rr.ok:
            out("    " + re.sub(r"\s+", " ", rr.text)[:600])


def mirrors():
    out("\n#### India Data Portal daily RE csv")
    r = get("https://ckandev.indiadataportal.com/dataset/6b0ff86f-1a92-40a6-93fa-086be0334c8e/resource/"
            "f009766a-c8b1-4322-91dd-f13dd653b45b/download/daily-renewable-energy-generation.csv")
    if r is not None and r.ok:
        df = pd.read_csv(io.BytesIO(r.content))
        out(df.head(5).to_string())
        out(df.tail(3).to_string())
        df["date"] = pd.to_datetime(df["date"], errors="coerce")
        num = df.select_dtypes("number").columns
        out(df.groupby(df["date"].dt.year)[list(num)].sum().round(0).to_string())
        out(f"  days per year: {df.groupby(df['date'].dt.year)['date'].nunique().to_dict()}")
    out("\n#### Wayback / other")
    for u in ("http://web.archive.org/cdx/search/cdx?url=webcdn.grid-india.in/files/grdw/2026/*&limit=15&output=json",
              "http://web.archive.org/cdx/search/cdx?url=grid-india.in/*psp*&limit=15&from=2025",
              "http://grid-india.in/", "https://powermin.gov.in/", "https://powermin.gov.in/en/content/power-sector-glance-all-india",
              "https://dataful.in/datasets/1222"):
        r = get(u)
        if r is not None and r.ok:
            out("    " + re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", r.text))[:1200])
            out("    links: " + str([x for x in sorted(set(re.findall(r'href=["\']([^"\']+)', r.text)))
                                     if re.search(r"psp|generation|supply|daily|renew", x, re.I)][:30]))


if __name__ == "__main__":
    for f in (cea_daily, cea_monthly, iced, npp_home, mirrors):
        try:
            f()
        except Exception as e:  # noqa: BLE001
            out(f"{f.__name__} failed: {type(e).__name__}: {e}")
