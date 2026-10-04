"""
Philippines discovery, round 2 (after PH_DISCOVERY1.py):
  IEMOP  registered-capacity-generation (CAPEG daily csv): columns, does it give each resource's fuel / plant type?
         DIPC Energy Results raw zip (fixed url): columns, resource names; join to CAPEG
         Market operations highlights 2026, monthly summary report, weekly market operations summary: PDF / xlsx
         links, first file's text (generation mix by plant type?)
  DOE    /data-and-prices/energy-statistics (JS site): data links, script bundles, any api / json endpoints
"""
import base64
import io
import re
import sys
import zipfile

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


def iemop_file(post_id, k=0):
    x = requests.post(AJAX, data={"action": "display_filtered_market_data_files", "sort": "", "datefilter": "",
                                  "page": 1, "post_id": post_id}, headers=H, timeout=T)
    path = base64.b64decode(x.json()["source"][k]).decode()
    path = path[path.find("/wp-content/"):]
    return get("https://www.iemop.ph" + path)


def show_pdf(c, pages=4, n=2500):
    import pdfplumber
    with pdfplumber.open(io.BytesIO(c)) as pdf:
        out(f"    PDF {len(pdf.pages)} pages")
        for i, p in enumerate(pdf.pages[:pages]):
            out(f"    --- p{i + 1}: " + (p.extract_text() or "")[:n].replace("\n", " | "))


def capeg():
    r = iemop_file(302634)
    d = pd.read_csv(io.BytesIO(r.content))
    out(f"  CAPEG {d.shape}; columns {list(d.columns)}")
    out(d.head(15).to_string()[:3000])
    for c in d.columns:
        if d[c].dtype == object and d[c].nunique() < 60:
            out(f"  {c}: {d[c].value_counts().to_dict()}")
    return d


def dipc(cap):
    r = iemop_file(5754, 2)
    with zipfile.ZipFile(io.BytesIO(r.content)) as zf:
        n = zf.namelist()[0]
        d = pd.read_csv(zf.open(n))
    out(f"  {n}: {d.shape}; columns {list(d.columns)}")
    out(d.head(6).to_string()[:1500])
    rc = next((c for c in d.columns if "RESOURCE_NAME" in c.upper()), None)
    if rc and cap is not None:
        kc = next((c for c in cap.columns if "RESOURCE" in c.upper() and "NAME" in c.upper()), None) or \
            next((c for c in cap.columns if "RESOURCE" in c.upper()), None)
        if kc:
            names = set(d.loc[d.get("SCHED_MW", 1) > 0, rc].astype(str))
            have = set(cap[kc].astype(str))
            out(f"  DIPC resources with output {len(names)}; in CAPEG {len(names & have)}; "
                f"missing {sorted(names - have)[:40]}")


def reports():
    for page in ("https://www.iemop.ph/market-reports/2026-market-operations-highlights/",
                 "https://www.iemop.ph/market-reports/monthly-summary-and-significant-variations-report/",
                 "https://www.iemop.ph/market-reports/weekly-market-operations-summary/",
                 "https://www.iemop.ph/market-reports/2023-market-operations-highlights/"):
        r = get(page)
        if r is None or r.status_code != 200:
            continue
        files = sorted(set(re.findall(r'(https?://[^"\'\s]+\.(?:pdf|xlsx?|csv))', r.text, re.I)))
        out(f"  {len(files)} files: {files[:25]}")
        m = re.search(r'"post_id":"(\d+)"', r.text)
        if m:
            out(f"  post_id {m.group(1)}")
            try:
                x = requests.post(AJAX, data={"action": "display_filtered_market_data_files", "sort": "",
                                              "datefilter": "", "page": 1, "post_id": m.group(1)}, headers=H, timeout=T)
                src = x.json().get("source") or []
                names = [base64.b64decode(s).decode() for s in src[:6]]
                out(f"  ajax files: {names}")
                if names:
                    p = names[0][names[0].find("/wp-content/"):]
                    f = get("https://www.iemop.ph" + p)
                    if f is not None and f.content[:4] == b"%PDF":
                        show_pdf(f.content)
                    elif f is not None and f.ok:
                        x2 = pd.ExcelFile(io.BytesIO(f.content))
                        out(f"  sheets {x2.sheet_names}")
                        out(x2.parse(x2.sheet_names[0], header=None, nrows=30).to_string()[:3000])
            except Exception as e:  # noqa: BLE001
                out(f"  ajax {e!r}")
        for f in files[:1]:
            x = get(f)
            if x is not None and x.content[:4] == b"%PDF":
                show_pdf(x.content)


def doe():
    r = get("https://doe.gov.ph/data-and-prices/energy-statistics")
    if r is None:
        return
    links = sorted(set(re.findall(r'(?:href|src)="([^"]+)"', r.text)))
    out(f"  links: {links[:80]}")
    for m in re.finditer(r'["\'](/api/[^"\']+|https?://[^"\']*(?:api|strapi|cms|storage|uploads)[^"\']*)["\']', r.text):
        out(f"  api-ish: {m.group(1)[:200]}")
    for js in [l for l in links if l.endswith(".js")][:12]:
        u = js if js.startswith("http") else "https://doe.gov.ph" + js
        x = get(u)
        if x is None or not x.ok:
            continue
        for m in set(re.findall(r'["\'`](https?://[^"\'`]{8,120}|/api/[^"\'`]{2,120})["\'`]', x.text)):
            if re.search(r"api|cms|storage|upload|statistic", m, re.I):
                out(f"    {js[-40:]}: {m}")
    text = re.sub(r"<script.*?</script>|<style.*?</style>", " ", r.text, flags=re.S)
    out("  text: " + re.sub(r"<[^>]+>|\s+", " ", text)[:2500])


if __name__ == "__main__":
    cap = None
    for f in sys.argv[1:] or ["capeg", "dipc", "reports", "doe"]:
        out(f"\n==================== {f}")
        try:
            if f == "capeg":
                cap = capeg()
            elif f == "dipc":
                dipc(cap)
            else:
                globals()[f]()
        except Exception as e:  # noqa: BLE001
            out(f"!! {f}: {e!r}")
