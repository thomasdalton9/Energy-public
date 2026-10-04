"""
Philippines discovery, round 1: a fuel split for Philippine generation.
  DOE    power statistics pages (doe.gov.ph / legacy.doe.gov.ph): data-file links (xlsx / pdf), monthly or annual
         gross generation by source and grid; open the first few files
  IEMOP  every market-data page (slug, post_id, first file names) - looking for a resource / registered-generator list
         with fuel types; one DIPC Energy Results zip: columns and resource names (do names encode the fuel?)
  IEMOP  market reports (monthly market assessment): generation mix by plant type?
"""
import io
import json
import re
import sys
import zipfile
import base64

import pandas as pd
import requests
import urllib3

urllib3.disable_warnings()
H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0 Safari/537.36",
     "Accept": "*/*", "Accept-Language": "en-US,en;q=0.9"}
T = (20, 90)
AJAX = "https://www.iemop.ph/wp-admin/admin-ajax.php"


def out(*a):
    print(*a, flush=True)


def get(u, **kw):
    try:
        r = requests.get(u, headers=H, timeout=T, verify=False, **kw)
        out(f"GET {r.url} -> {r.status_code} {r.headers.get('content-type')} {len(r.content)}")
        return r
    except Exception as e:  # noqa: BLE001
        out(f"GET {u} ERROR {type(e).__name__}: {str(e)[:150]}")
        return None


def show_file(r, name):
    c = r.content
    try:
        if c[:4] == b"%PDF":
            import pdfplumber
            with pdfplumber.open(io.BytesIO(c)) as pdf:
                out(f"    PDF {len(pdf.pages)} pages")
                for p in pdf.pages[:3]:
                    out("    | " + (p.extract_text() or "")[:1800].replace("\n", " | "))
        else:
            x = pd.ExcelFile(io.BytesIO(c))
            out(f"    sheets: {x.sheet_names[:30]}")
            for sh in x.sheet_names[:3]:
                d = x.parse(sh, header=None, nrows=30)
                out(f"    --- {sh} {d.shape}\n" + d.to_string(max_cols=14, max_colwidth=16)[:3000])
    except Exception as e:  # noqa: BLE001
        out(f"    {name}: {e!r}")


def doe():
    seen = set()
    for page in ("https://doe.gov.ph/energy-statistics/power-statistics",
                 "https://doe.gov.ph/energy-information-resources?q=power-statistics",
                 "https://legacy.doe.gov.ph/energy-statistics?q=power-statistics",
                 "https://legacy.doe.gov.ph/power-statistics",
                 "https://doe.gov.ph/power-statistics",
                 "https://www.doe.gov.ph/energy-statistics",
                 "https://doe.gov.ph/articles/group/energy-statistics?category=Power%20Statistics"):
        r = get(page)
        if r is None or r.status_code != 200:
            continue
        links = sorted(set(re.findall(r'href="([^"]+)"', r.text)))
        files = [l for l in links if re.search(r"\.(xlsx?|pdf|csv)(\?|$)", l, re.I)]
        rel = [l for l in links if re.search(r"power|generation|statistic", l, re.I)]
        out(f"  {len(links)} links; files {len(files)}: {files[:40]}")
        out(f"  power/statistics links: {rel[:60]}")
        text = re.sub(r"<script.*?</script>|<style.*?</style>", " ", r.text, flags=re.S)
        text = re.sub(r"<[^>]+>|\s+", " ", text)
        for m in re.finditer(r".{0,80}(Gross Generation|Power Statistics|Generation Mix|Power Situation).{0,160}", text):
            out("    txt: " + m.group(0)[:240])
        for f in files:
            if f in seen or len(seen) >= 6 or not re.search(r"gen|power|stat", f, re.I):
                continue
            seen.add(f)
            u = f if f.startswith("http") else requests.compat.urljoin(r.url, f)
            x = get(u)
            if x is not None and x.status_code == 200:
                show_file(x, f)


def iemop_pages():
    r = get("https://www.iemop.ph/market-data/")
    slugs = sorted(set(re.findall(r'https://www\.iemop\.ph/market-data/([a-z0-9\-]+)/', r.text))) if r is not None else []
    out(f"  {len(slugs)} market-data slugs: {slugs}")
    for slug in slugs:
        p = get(f"https://www.iemop.ph/market-data/{slug}/")
        if p is None or not p.ok:
            continue
        m = re.search(r'"post_id":"(\d+)","min_date":"([^"]+)"', p.text)
        if not m:
            out(f"  {slug}: no post_id")
            continue
        try:
            x = requests.post(AJAX, data={"action": "display_filtered_market_data_files", "sort": "", "datefilter": "",
                                          "page": 1, "post_id": m.group(1)}, headers=dict(H, Referer=p.url), timeout=T)
            j = x.json()
            names = [base64.b64decode(s).decode().rsplit("/", 1)[-1] for s in (j.get("source") or [])[:3]]
            out(f"  {slug}: post_id {m.group(1)} min_date {m.group(2)} files {names}")
        except Exception as e:  # noqa: BLE001
            out(f"  {slug}: post_id {m.group(1)} ajax {e!r}")


def iemop_dipc():
    x = requests.post(AJAX, data={"action": "display_filtered_market_data_files", "sort": "", "datefilter": "",
                                  "page": 1, "post_id": 5754}, headers=H, timeout=T)
    src = x.json().get("source") or []
    path = base64.b64decode(src[0]).decode()
    u = path if path.startswith("http") else "https://www.iemop.ph" + path
    z = get(u)
    with zipfile.ZipFile(io.BytesIO(z.content)) as zf:
        n = zf.namelist()[0]
        d = pd.read_csv(zf.open(n))
    out(f"  {n}: {d.shape}; columns {list(d.columns)}")
    out(d.head(5).to_string()[:1500])
    col = next((c for c in d.columns if "RESOURCE" in c.upper() and "TYPE" not in c.upper()), None)
    if col:
        g = d[d.get("SCHED_MW", 0) > 0] if "SCHED_MW" in d else d
        names = sorted(g[col].astype(str).unique())
        out(f"  {len(names)} resources with output: {names}")
    for c in d.columns:
        if d[c].dtype == object and d[c].nunique() < 30:
            out(f"  {c}: {sorted(d[c].astype(str).unique())}")


def iemop_reports():
    for page in ("https://www.iemop.ph/market-reports/", "https://www.iemop.ph/market-reports/monthly-market-assessment-report/",
                 "https://www.iemop.ph/news-and-reports/", "https://www.iemop.ph/registered-participants/",
                 "https://www.iemop.ph/market-participants/"):
        r = get(page)
        if r is None or r.status_code != 200:
            continue
        links = sorted(set(re.findall(r'href="([^"]+)"', r.text)))
        keep = [l for l in links if re.search(r"report|assessment|participant|registr|\.pdf|\.xls", l, re.I)]
        out(f"  links (report/participant/pdf/xls): {keep[:60]}")
        pdfs = [l for l in links if re.search(r"\.(pdf|xlsx?)$", l, re.I)]
        for f in pdfs[:2]:
            x = get(f)
            if x is not None and x.status_code == 200:
                show_file(x, f)


if __name__ == "__main__":
    for f in sys.argv[1:] or ["doe", "iemop_pages", "iemop_dipc", "iemop_reports"]:
        out(f"\n==================== {f}")
        try:
            globals()[f]()
        except Exception as e:  # noqa: BLE001
            out(f"!! {f}: {e!r}")
