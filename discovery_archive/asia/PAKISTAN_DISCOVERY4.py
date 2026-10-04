"""
Pakistan power discovery, round 4: CPPA-G downloads pages (cppa.gov.pk/downloads/xwdiscos-energy-purchase-data/<id>,
one per fiscal year). List the monthly files (PDF / Excel?) and dump a sample of each type.
"""
import io
import re
from urllib.parse import urljoin, unquote

import pandas as pd
import pdfplumber
import requests
import urllib3

urllib3.disable_warnings()
H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 "
                   "Safari/537.36"}
T = (15, 120)
s = requests.Session()
s.headers.update(H)


def out(*a):
    print(*a, flush=True)


def get(u):
    try:
        return s.get(u, timeout=T, verify=False)
    except Exception as e:  # noqa: BLE001
        out(f"  ERR {u}: {e}")
        return None


def content(t):
    k = t.find('<div class="col-md-9">')
    f = t.find("<footer")
    return t[k if k > 0 else 0:f if f > k else len(t)]


def dump_file(u, r):
    ct = r.headers.get("content-type", "")
    cd = r.headers.get("content-disposition", "")
    out(f"    -> {r.status_code} {ct} {cd} {len(r.content)}b final={r.url}")
    b = r.content
    if b[:4] == b"%PDF":
        with pdfplumber.open(io.BytesIO(b)) as p:
            for i, pg in enumerate(p.pages):
                t = pg.extract_text() or ""
                if "Summary" in t or i == len(p.pages) - 1:
                    out(f"    page {i + 1}/{len(p.pages)} text {len(t)}:")
                    out("    " + t[-3000:].replace("\n", "\n    "))
                    break
    elif b[:2] == b"PK" or b[:4] == b"\xd0\xcf\x11\xe0":
        x = pd.ExcelFile(io.BytesIO(b))
        out(f"    sheets: {x.sheet_names}")
        for sh in x.sheet_names[:6]:
            df = x.parse(sh, header=None)
            out(f"    == {sh} {df.shape}")
            out(df.head(12).to_string(max_cols=14, max_colwidth=28)[:2500])
            m = df.astype(str).apply(lambda c: c.str.contains("Summary|Coal-Local|Grand Total", case=False)).any(axis=1)
            if m.any():
                i0 = m[m].index[0]
                out(f"    -- rows from {i0}:")
                out(df.iloc[i0:i0 + 25].to_string(max_cols=8, max_colwidth=22)[:3500])


def main():
    seen_types = set()
    for page in ("https://cppa.gov.pk/downloads/xwdiscos-energy-purchase-data",
                 "https://cppa.gov.pk/downloads/fuel-adjustment-notifications",
                 "https://cppa.gov.pk/downloads/others", "https://cppa.gov.pk/downloads/policies-reports"):
        r = get(page)
        if r is None:
            continue
        c = content(r.text)
        subs = sorted(set(re.findall(r'href="(https://cppa\.gov\.pk/downloads/[a-z-]+/\d+)"', c)))
        titles = re.findall(r"<h3[^>]*>(.*?)</h3>", c, re.S)
        out(f"\n==== {page}: {r.status_code} sub-pages {subs}; titles {[t.strip() for t in titles][:40]}")
        for sp in subs if "xwdiscos" in page else subs[:3]:
            q = get(sp)
            if q is None:
                continue
            ct = q.headers.get("content-type", "")
            out(f"\n  -- {sp}: {q.status_code} {ct} {len(q.content)}b")
            if "html" not in ct:
                dump_file(sp, q)
                continue
            cc = content(q.text)
            L = re.findall(r'<a[^>]+href="([^"]+)"[^>]*>(.*?)</a>', cc, re.S)
            out(f"     {len(L)} links")
            for h, t in L[:80]:
                tt = re.sub(r"<[^>]+>|\s+", " ", t).strip()[:70]
                out(f"     {tt!r} -> {h}")
            for h, t in L:
                u = urljoin(sp, h)
                ext = re.sub(r".*\.", "", unquote(u).lower())[:5]
                if ext in seen_types or "downloads" not in u and "storage" not in u:
                    continue
                f = get(u)
                if f is None:
                    continue
                seen_types.add(ext)
                out(f"   sample {u}")
                try:
                    dump_file(u, f)
                except Exception as e:  # noqa: BLE001
                    out(f"    dump error {e}")
    # older fiscal-year ids?
    for n in list(range(150, 230)):
        u = f"https://cppa.gov.pk/downloads/xwdiscos-energy-purchase-data/{n}"
        q = get(u)
        if q is not None and q.status_code == 200:
            h3 = re.findall(r"<h1[^>]*>(.*?)</h1>|<h3[^>]*>(.*?)</h3>", content(q.text), re.S)
            out(f"  id {n}: {q.status_code} {q.headers.get('content-type')} {len(q.content)}b {h3[:3]}")


if __name__ == "__main__":
    main()
