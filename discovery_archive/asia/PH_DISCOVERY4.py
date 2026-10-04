"""
Philippines discovery, round 4: DOE 'List of Existing Power Plants' (per grid: facility, fuel, capacity) for the
IEMOP resource -> fuel map. DOE's site is Next.js on a Payload CMS whose files sit on
d24qbtp4vooyzi.cloudfront.net/api/media/file/<name> (round 3); try the Payload REST API (/api/media?where=...)
on the CDN and on doe.gov.ph to list the media files, then print the plant tables of the newest per-grid lists.
"""
import io
import json
import re
import urllib.parse

import requests
import urllib3

urllib3.disable_warnings()
H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0 Safari/537.36",
     "Accept": "*/*"}
T = (20, 120)
HOSTS = ["https://d24qbtp4vooyzi.cloudfront.net", "https://doe.gov.ph", "https://cms.doe.gov.ph",
         "https://prod-cms.doe.gov.ph"]


def out(*a):
    print(*a, flush=True)


def get(u):
    try:
        r = requests.get(u, headers=H, timeout=T, verify=False)
        out(f"GET {u[:200]} -> {r.status_code} {r.headers.get('content-type')} {len(r.content)}")
        return r
    except Exception as e:  # noqa: BLE001
        out(f"GET {u[:200]} ERROR {type(e).__name__}: {str(e)[:120]}")
        return None


def media_search():
    found = {}
    for host in HOSTS:
        for term in ("Existing", "Grid", "LoEPP", "Power Plants"):
            q = urllib.parse.urlencode({"limit": 200, "depth": 0, "where[filename][like]": term})
            r = get(f"{host}/api/media?{q}")
            if r is None or not r.ok or "json" not in (r.headers.get("content-type") or ""):
                if r is not None and r.ok:
                    out("   body: " + r.text[:300].replace("\n", " "))
                continue
            j = r.json()
            docs = j.get("docs") or []
            out(f"   totalDocs {j.get('totalDocs')}; keys {list(docs[0].keys()) if docs else None}")
            for d in docs:
                fn = d.get("filename") or ""
                found[fn] = (d.get("url") or "", d.get("updatedAt") or d.get("createdAt") or "", host)
        if found:
            break
    for fn, (u, t, host) in sorted(found.items(), key=lambda x: x[1][1]):
        out(f"  MEDIA {t[:10]} {fn} {u[:160]}")
    return found


def pages_search():
    for host in HOSTS[:2]:
        for coll in ("pages", "posts", "documents", "files", "downloads"):
            q = urllib.parse.urlencode({"limit": 5, "depth": 1, "where[title][like]": "Existing"})
            r = get(f"{host}/api/{coll}?{q}")
            if r is not None and r.ok and "json" in (r.headers.get("content-type") or ""):
                out("   " + json.dumps(r.json())[:3000])


def pdf_rows(content, label, maxpages=80):
    import pdfplumber
    with pdfplumber.open(io.BytesIO(content)) as pdf:
        out(f"  PDF {label}: {len(pdf.pages)} pages")
        for i, p in enumerate(pdf.pages[:maxpages]):
            tabs = p.extract_tables()
            if not tabs:
                out(f"  p{i + 1} text: " + (p.extract_text() or "")[:2000].replace("\n", " | "))
            for t in tabs:
                for row in t:
                    cells = ["" if c is None else str(c).replace("\n", " ").strip() for c in row]
                    if any(cells):
                        out("DOE;" + label + ";" + " | ".join(cells))


def main():
    found = media_search()
    if not found:
        pages_search()
    want = [fn for fn in found if re.search(r"exist|loepp|grid", fn, re.I)
            and re.search(r"luzon|visayas|mindanao|lvm|loepp|exist", fn, re.I)
            and not re.search(r"generation|demand|sales|consumption|peak", fn, re.I)]
    # newest per grid
    by = {}
    for fn in want:
        g = next((k for k in ("luzon", "visayas", "mindanao") if k in fn.lower()), "other")
        if g not in by or found[fn][1] > found[by[g]][1]:
            by[g] = fn
    out(f"  chosen {by}")
    for g, fn in by.items():
        u, _, host = found[fn]
        u = u if u.startswith("http") else host + u
        r = get(u)
        if r is None or not r.ok:
            r = get("https://d24qbtp4vooyzi.cloudfront.net/api/media/file/" + urllib.parse.quote(fn) +
                    "?prefix=dev%2Fmedia")
        if r is not None and r.ok and r.content[:4] == b"%PDF":
            pdf_rows(r.content, g)
        elif r is not None and r.ok and r.content[:2] == b"PK":
            import pandas as pd
            xl = pd.ExcelFile(io.BytesIO(r.content))
            for sh in xl.sheet_names:
                for row in xl.parse(sh, header=None).itertuples(index=False):
                    out("DOE;" + g + ";" + " | ".join("" if pd.isna(c) else str(c) for c in row))


if __name__ == "__main__":
    main()
