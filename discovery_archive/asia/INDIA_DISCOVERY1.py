"""
India daily renewable generation discovery, round 1 (run from a GitHub Actions runner).

Known: NPP dgr2/dgr6 xls (conventional only); CEA admin-ajax 'renewable' daily RE xlsx listing stopped 2025-11-18;
Grid-India hosts reset GitHub runners. This round:
  NPP     dgr1..dgr30 xls + pdf for two dates: which report (if any) carries RES / wind / solar
  NPP     publishedReports / dgrReports page links
  CEA     admin-ajax 'renewable' listing (resumed?), other codes; monthly RE generation PDF (daily wind/solar table?)
  RLDCs   nrldc / wrldc / srldc / erldc / nerldc reachability (regional PSP reports incl. wind/solar)
  Mirrors India Data Portal CKAN (CEA daily RE csv), Grid-India hosts, Vidyut Pravah, MERIT
"""
import io
import re

import pandas as pd
import requests
import urllib3

urllib3.disable_warnings()
H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 "
                   "Safari/537.36", "Accept": "*/*"}
T = (15, 60)


def out(*a):
    print(*a, flush=True)


def get(u, **kw):
    try:
        r = requests.get(u, headers=H, timeout=T, verify=False, **kw)
        out(f"  GET {u} -> {r.status_code} {len(r.content)} {r.headers.get('content-type', '')[:40]}")
        return r
    except Exception as e:  # noqa: BLE001
        out(f"  GET {u}: {type(e).__name__}: {str(e)[:150]}")
        return None


def pdf_text(content, pages=3, n=3000):
    import pdfplumber
    with pdfplumber.open(io.BytesIO(content)) as p:
        out(f"    pdf pages: {len(p.pages)}")
        t = "\n".join((pg.extract_text() or "") for pg in p.pages[:pages])
    return t[:n]


def show_xls(content, rows=25):
    try:
        xl = pd.ExcelFile(io.BytesIO(content))
        for s in xl.sheet_names[:2]:
            df = xl.parse(s, header=None)
            out(f"    sheet {s!r} {df.shape}")
            out(df.head(rows).to_string(max_colwidth=28)[:4000])
            hits = df[df.apply(lambda r: r.astype(str).str.contains("WIND|SOLAR|RES|RENEW", case=False).any(), axis=1)]
            if len(hits):
                out("    RE-ish rows:\n" + hits.head(15).to_string(max_colwidth=28)[:3000])
    except Exception as e:  # noqa: BLE001
        out(f"    excel read failed: {e}")


def links(html, pat=None, n=80):
    ls = sorted(set(re.findall(r'(?:href|src)=["\']([^"\']+)["\']', html)))
    if pat:
        ls = [x for x in ls if re.search(pat, x, re.I)]
    return ls[:n]


def npp():
    out("\n#### NPP dgr1..dgr30")
    for d, iso in (("01-10-2026", "2026-10-01"), ("13-06-2026", "2026-06-13")):
        for k in range(1, 31):
            for ext in ("xls", "pdf"):
                u = f"https://npp.gov.in/public-reports/cea/daily/dgr/{d}/dgr{k}-{iso}.{ext}"
                r = get(u)
                if r is None or not r.ok or len(r.content) < 500:
                    continue
                if d != "01-10-2026" and k not in (1, 17):
                    continue
                if ext == "xls":
                    show_xls(r.content, rows=12)
                else:
                    try:
                        t = pdf_text(r.content, pages=2, n=1500)
                        out("    " + t.replace("\n", "\n    "))
                        hit = [ln for ln in t.splitlines() if re.search(r"wind|solar|R\.?E\.?S|renew", ln, re.I)]
                        out(f"    RE lines: {hit[:15]}")
                    except Exception as e:  # noqa: BLE001
                        out(f"    pdf failed {e}")
    for u in ("https://npp.gov.in/publishedReports", "https://npp.gov.in/dgrReports", "https://npp.gov.in/",
              "https://npp.gov.in/public-reports/cea/daily/", "https://npp.gov.in/public-reports/cea/daily/re/"):
        r = get(u)
        if r is not None and r.ok:
            out("    links: " + str(links(r.text, r"report|dgr|renew|api|json|\.js|daily|re", 120)))
            out("    text: " + re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", r.text))[:1500])


def cea():
    out("\n#### CEA")
    for code in ("renewable", "resd", "re-generation", "daily-renewable", "renewable-generation", "re"):
        r = get(f"https://cea.nic.in/wp-admin/admin-ajax.php?action=getpostsfordatatables&code={code}")
        if r is None or not r.ok:
            continue
        try:
            rows = r.json().get("data", [])
        except Exception:  # noqa: BLE001
            out("    " + r.text[:300])
            continue
        rows = sorted(rows, key=lambda x: str(x.get("date", "")))
        out(f"    code={code}: {len(rows)} rows")
        for x in rows[:2] + rows[-8:]:
            out(f"      {str(x)[:300]}")
    for u in ("https://cea.nic.in/daily-renewable-generation-report/?lang=en",
              "https://cea.nic.in/renewable-generation-report/?lang=en",
              "https://cea.nic.in/re-generation-report/?lang=en"):
        r = get(u)
        if r is not None and r.ok:
            out("    links: " + str(links(r.text, r"upload|resd|xls|pdf|ajax|code=", 60)))
            out("    codes: " + str(sorted(set(re.findall(r"code[\"'=: ]+([a-z_\-]+)", r.text)))[:30]))
    for u in ("https://cea.nic.in/wp-content/uploads/resd/2026/08/Monthly_RE_Generation_report_August_2026.pdf",
              "https://cea.nic.in/wp-content/uploads/resd/2026/07/Monthly_RE_Generation_report_July_2026.pdf",
              "https://cea.nic.in/wp-content/uploads/resd/2026/08/Broad_Overview_of_RE_Generation_July_2026.pdf"):
        r = get(u)
        if r is not None and r.ok:
            import pdfplumber
            with pdfplumber.open(io.BytesIO(r.content)) as p:
                out(f"    pages {len(p.pages)}")
                for i, pg in enumerate(p.pages[:12]):
                    t = pg.extract_text() or ""
                    out(f"    --- page {i + 1}: " + t[:900].replace("\n", " | "))


def rldc():
    out("\n#### RLDCs / Grid-India / mirrors")
    for u in ("https://nrldc.in/", "https://www.nrldc.in/", "https://wrldc.in/", "https://www.wrldc.in/",
              "https://srldc.in/", "https://www.srldc.in/", "https://erldc.in/", "https://www.erldc.in/",
              "https://nerldc.in/", "https://www.nerldc.in/", "https://grid-india.in/", "https://webapi.grid-india.in/",
              "https://report.grid-india.in/", "https://posoco.in/", "https://vidyutpravah.in/", "https://meritindia.in/",
              "https://www.nldc.in/"):
        r = get(u)
        if r is not None and r.ok:
            out("    links: " + str(links(r.text, r"psp|report|daily|api|wind|solar|renew", 40)))
    r = get("https://ckandev.indiadataportal.com/api/3/action/resource_show?id=f009766a-c8b1-4322-91dd-f13dd653b45b")
    if r is not None and r.ok:
        out("    " + r.text[:1500])
    r = get("https://ckandev.indiadataportal.com/api/3/action/package_show?id=power")
    if r is not None and r.ok:
        try:
            for x in r.json()["result"]["resources"]:
                out(f"    res {x.get('name')} | {x.get('url')} | {x.get('last_modified')}")
        except Exception as e:  # noqa: BLE001
            out(f"    {e}")


if __name__ == "__main__":
    for f in (cea, npp, rldc):
        try:
            f()
        except Exception as e:  # noqa: BLE001
            out(f"{f.__name__} failed: {e}")
