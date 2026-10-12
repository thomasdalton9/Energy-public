"""
Morocco gas probe, pass 3: the leads pass 2 did not finish.

Targets:
  stats     HCP statistical yearbook (Excel) energy tables, Office des Changes 'produits remarquables' / series, data.gov.ma MEF
            sector indicators, mem.gov.ma static mirror, one.org.ma (no certificate check)
  comtrade  UN Comtrade public preview, Morocco monthly HS 271121 imports by partner, 2022-2025 (one period per call)

Usage: python3 MOROCCO_GAS_PROBE3.py <target>
"""
import io
import os
import re
import sys
import time
import zipfile

import pandas as pd
import requests
import urllib3

urllib3.disable_warnings()
UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36"
H = {"User-Agent": UA}
OUT = "probe_files"
os.makedirs(OUT, exist_ok=True)


def get(url, tries=2, timeout=(15, 120), **kw):
    kw.setdefault("headers", H)
    last = None
    for _ in range(tries):
        try:
            return requests.get(url, timeout=timeout, **kw)
        except requests.RequestException as e:
            last = e
            time.sleep(3)
    print(f"  ERR {url[:140]} {type(last).__name__}: {str(last)[:160]}")
    return None


def page_links(url, pat=r".", n=60, **kw):
    r = get(url, **kw)
    print(f"\n=== {url} -> {r.status_code if r is not None else None} {len(r.content) if r is not None else ''} {r.url if r is not None else ''}")
    if r is None or r.status_code != 200:
        return None, []
    out, seen = [], set()
    for h, t in re.findall(r'href=["\']([^"\'#]+)["\'][^>]*>(.{0,300}?)</a', r.text, re.S):
        t = re.sub(r"<[^>]+>|\s+", " ", t).strip()
        if re.search(pat, h + " " + t, re.I) and h not in seen:
            seen.add(h)
            out.append((requests.compat.urljoin(r.url, h), t))
    for h, t in out[:n]:
        print("   ", h[:170], "|", t[:90])
    return r, out


def grep_book(content, pat, label, maxrows=15):
    try:
        x = pd.read_excel(io.BytesIO(content), sheet_name=None, header=None)
    except Exception as e:  # noqa: BLE001
        print("   xlsx error", label, type(e).__name__, str(e)[:80])
        return
    print(f"   {label}: {len(x)} sheets: {list(x)[:40]}")
    for s, d in x.items():
        txt = d.astype(str)
        m = txt.apply(lambda col: col.str.contains(pat, case=False, regex=True)).any(axis=1)
        if m.any():
            print(f"   --- sheet {s!r} {d.shape}: {int(m.sum())} matching rows")
            for i in list(d.index[m])[:maxrows]:
                print("     ", " | ".join(str(v) for v in txt.loc[i].tolist() if str(v) != "nan")[:320])


PAT = r"gaz naturel|gaz nat|thermique|cycle combin|Tahaddart|Beni Mathar|production.*électri|production.*electri|énergie électrique|energie electrique"


def stats():
    # HCP yearbook: list the Excel files of the newest yearbook
    r, out = page_links("https://www.hcp.ma/downloads/Annuaires-statistiques-du-Maroc-Format-Excel_t22392.html", r"/file/", 40)
    if r is not None:
        # titles are in the surrounding HTML, print them around each file link
        for m in re.finditer(r'(/file/\d+/)', r.text):
            ctx = re.sub(r"<[^>]+>|\s+", " ", r.text[max(0, m.start() - 400):m.end() + 200])
            print("   ctx", m.group(1), ctx[-300:][:300])
            if m.start() > 0 and len(re.findall(r"/file/", r.text[:m.start()])) > 12:
                break
    for fid in ("248151", "248150"):
        rr = get(f"https://www.hcp.ma/file/{fid}/", timeout=(15, 240))
        if rr is None or not rr.ok:
            continue
        cd = rr.headers.get("content-disposition", "")
        print(f"\n  HCP file {fid}: {len(rr.content)} bytes {rr.headers.get('content-type')} {cd} magic={rr.content[:4]}")
        if rr.content[:2] == b"PK":
            z = zipfile.ZipFile(io.BytesIO(rr.content))
            names = z.namelist()
            if any(n.startswith("xl/") for n in names):
                grep_book(rr.content, PAT, f"HCP {fid}")
            else:
                print("   entries:", names[:80])
                for n in names:
                    if n.lower().endswith((".xlsx", ".xls")) and re.search(r"nerg|lectr|mine|ndus|10|11|12", n, re.I):
                        grep_book(z.read(n), PAT, n)
        elif rr.content[:4] == b"\xd0\xcf\x11\xe0":
            grep_book(rr.content, PAT, f"HCP {fid} xls")
    # Office des Changes
    for u in ["https://www.oc.gov.ma/sites/default/files/Produits%20remarquables%20site%201_18.xlsx"]:
        rr = get(u)
        if rr is not None and rr.ok:
            grep_book(rr.content, r"gaz|hydrocarb|énerg|energ", "OC produits remarquables", 30)
    page_links("https://www.oc.gov.ma/fr/etudes-et-statistiques/series-statistiques", r"xls|pdf|s[ée]rie|commerce|annuel|mensuel|energ", 60)
    page_links("https://services.oc.gov.ma/DataBase/CommerceExterieur/", r".", 40)
    # data.gov.ma MEF sector indicators (energy)
    j = get("https://data.gov.ma/data/fr/api/3/action/package_search", params={"q": "energie", "rows": 20})
    if j is not None and j.ok:
        for d in j.json()["result"]["results"]:
            for x in d.get("resources", []):
                u = x.get("url") or ""
                if u.lower().split("?")[0].endswith((".xls", ".xlsx")):
                    rr = get(u)
                    if rr is not None and rr.ok:
                        grep_book(rr.content, r"gaz|thermique|électri|electri|GWh", f"{d['name']} {u.split('/')[-1][:60]}", 25)
    # mem.gov.ma static mirror and ONEE
    for u in ["https://www.mem.gov.ma/Pages/index.html"]:
        r, out = page_links(u, r"stat|chiffre|bilan|tableau|publication|rapport|xls|pdf|energie|gaz|electr", 80)
    rr = get("https://www.one.org.ma/", verify=False, timeout=(25, 60))
    if rr is not None:
        print("\none.org.ma body:", rr.status_code, rr.text[:2000])
    for u in ["https://www.one.org.ma/fr/pages/index.asp", "https://www.one.org.ma/FR/pages/interne.asp?esp=1&id1=3&id2=114&t2=1"]:
        page_links(u, r"rapport|annuel|activit|chiffre|statist|pdf|production", 50, verify=False, timeout=(25, 60))


def comtrade():
    base = "https://comtradeapi.un.org/public/v1/preview/C/M/HS"
    rows = []
    for y in (2022, 2023, 2024, 2025):
        for m in range(1, 13):
            per = f"{y}{m:02d}"
            for a in range(4):
                r = get(base, params={"reporterCode": 504, "cmdCode": "271121", "flowCode": "M", "period": per}, timeout=(15, 60))
                if r is not None and r.status_code == 429:
                    time.sleep(6 + 4 * a)
                    continue
                break
            time.sleep(4)
            if r is None or r.status_code != 200:
                print(per, r.status_code if r is not None else None)
                continue
            d = pd.DataFrame(r.json().get("data", []))
            if d.empty:
                print(per, "no rows")
                continue
            rows.append(d)
            for _, x in d.iterrows():
                print(f"  {per} partner={x.get('partnerCode')} qty={x.get('qty')} unit={x.get('qtyUnitAbbr')} netWgt={x.get('netWgt')} "
                      f"alt={x.get('altQty')} {x.get('altQtyUnitAbbr')} US$={x.get('primaryValue')}")
    if rows:
        pd.concat(rows).to_csv(f"{OUT}/comtrade_morocco_271121_monthly.csv", index=False)


if __name__ == "__main__":
    {"stats": stats, "comtrade": comtrade}[sys.argv[1]]()
