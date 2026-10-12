"""
Morocco gas probe, pass 2 (after MOROCCO_GAS_PROBE.py).

Targets:
  enagas2   every Enagas monthly bulletin 2016-today: the Tarifa / Almeria rows of the interconnection tables (text saved per bulletin),
            plus the Enagas 'technical capacity / physical flows' page (any JSON endpoint)
  entsog2   ENTSOG before 2021: does the API give Tarifa history (periodType day/month, aggregatedData)?
  morocco2  mem.gov.ma (what the 745-byte home page is), ONHYM / ONEE without certificate check, EET (Tahaddart) key figures,
            ANRE annual reports (electricity by source, gas), HCP statistical yearbook (Excel), Office des Changes monthly indicators
            and 'produits remarquables', data.gov.ma MEF energy indicators
  trade2    UN Comtrade public preview, one period per call (the preview allows only one), with pauses for the rate limit

Usage: python3 MOROCCO_GAS_PROBE2.py <target>
"""
import io
import json
import os
import re
import sys
import time
from datetime import date

import pandas as pd
import requests
import urllib3

urllib3.disable_warnings()
UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36"
H = {"User-Agent": UA}
OUT = "probe_files"
os.makedirs(OUT, exist_ok=True)


def get(url, tries=2, timeout=(15, 90), **kw):
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


def links(url, pat=r".", n=80, **kw):
    r = get(url, **kw)
    print(f"\n=== {url} -> {r.status_code if r is not None else None} {len(r.content) if r is not None else ''} {r.url if r is not None else ''}")
    if r is None or r.status_code != 200:
        if r is not None:
            print("   body:", r.text[:400].replace("\n", " "))
        return None, []
    out, seen = [], set()
    for h, t in re.findall(r'href=["\']([^"\'#]+)["\'][^>]*>(.{0,200}?)</a', r.text, re.S):
        t = re.sub(r"<[^>]+>|\s+", " ", t).strip()
        if re.search(pat, h + " " + t, re.I) and h not in seen:
            seen.add(h)
            out.append((requests.compat.urljoin(r.url, h), t))
    for h, t in out[:n]:
        print("   ", h[:170], "|", t[:80])
    return r, out


def pdf_text(content, maxpages=60):
    import pdfplumber
    with pdfplumber.open(io.BytesIO(content)) as pdf:
        return [(pg.extract_text() or "") for pg in pdf.pages[:maxpages]]


# ------------------------------------------------------------------------------------------------ Enagas
ENAGAS_PAGE = "https://www.enagas.es/en/technical-management-system/energy-data/publications/gas-statistical-bulletin/"


def enagas2():
    r, out = links("https://www.enagas.es/en/technical-management-system/energy-data/physical-parameters/technical-capacity-physical-flows/",
                   r"json|xls|csv|flow|capacity", 60)
    if r is not None:
        for m in sorted(set(re.findall(r'["\']([^"\']+\.(?:json|xlsx?|csv)[^"\']*)["\']', r.text)))[:40]:
            print("   endpoint:", m[:200])
        for m in sorted(set(re.findall(r'data-[a-z-]*url=["\']([^"\']+)["\']', r.text)))[:40]:
            print("   data-url:", m[:200])
    files = {}
    today = date.today()
    for y in range(2015, today.year + 1):
        for m in range(1, 13):
            if (y, m) > (today.year, today.month):
                continue
            rr = get(ENAGAS_PAGE, params={"category": "", "month": m, "year": y})
            if rr is None:
                continue
            for x in re.findall(r'href="(/content/dam[^"]+\.pdf)"', rr.text):
                files.setdefault(x, (y, m))
    print("bulletins listed:", len(files))
    os.makedirs(f"{OUT}/enagas_text", exist_ok=True)
    for f, ym in sorted(files.items(), key=lambda kv: kv[1]):
        rr = get("https://www.enagas.es" + f, timeout=(15, 120))
        if rr is None or rr.status_code != 200:
            print("  ", ym, f.split("/")[-1], "download failed")
            continue
        try:
            pages = pdf_text(rr.content)
        except Exception as e:  # noqa: BLE001
            print("  ", ym, f.split("/")[-1], "pdf error", type(e).__name__)
            continue
        name = re.sub(r"[^A-Za-z0-9_.-]+", "_", requests.utils.unquote(f.split("/")[-1]))
        keep = [f"##### PAGE {i + 1}\n{t}" for i, t in enumerate(pages)
                if re.search(r"tarifa|evolution of gas demand|evoluci.n de la demanda|international|internacional", t, re.I)]
        open(f"{OUT}/enagas_text/{ym[0]}_{ym[1]:02d}_{name}.txt", "w").write(f"{f}\n" + "\n".join(keep))
        tl = [ln for t in pages for ln in t.split("\n") if re.search(r"^\s*Tarifa", ln, re.I)]
        title = re.sub(r"\s+", " ", pages[0][:120]) if pages else ""
        print(f"  {ym} {name[:60]} | {title[:60]} | {' || '.join(tl)[:300]}")


# ------------------------------------------------------------------------------------------------ ENTSOG before 2021
API = "https://transparency.entsog.eu/api/v1"


def eget(path, params):
    for a in range(4):
        try:
            r = requests.get(f"{API}/{path}", params=params, headers={"User-Agent": "Mozilla/5.0", "Accept": "application/json"}, timeout=(15, 300))
        except requests.RequestException as e:
            print("   ", type(e).__name__)
            time.sleep(10)
            continue
        if r.status_code == 200:
            return r.json()
        if r.status_code == 404:
            return {"_status": 404, "_text": r.text[:200]}
        print("   ", path, r.status_code, r.text[:150])
        time.sleep(15 * (a + 1))
    return {}


def entsog2():
    pd_ = "ES-TSO-0006ITP-00082entry"
    for y in (2015, 2017, 2019, 2020):
        for per in ("day", "month", "year"):
            j = eget("operationalData", {"indicator": "Physical Flow", "periodType": per, "from": f"{y}-01-01", "to": f"{y}-12-31",
                                          "pointDirection": pd_, "limit": -1})
            rows = j.get("operationalData", [])
            vals = [float(r["value"]) for r in rows if r.get("value") not in (None, "")]
            print(f"operationalData {per} {y}: {len(rows)} rows, {sum(vals) / 1e9:.2f} TWh  {j.get('_text', '')[:120]}")
        for ind in ("Allocation", "Nomination"):
            j = eget("operationalData", {"indicator": ind, "periodType": "day", "from": f"{y}-01-01", "to": f"{y}-12-31",
                                          "pointDirection": pd_, "limit": -1})
            rows = j.get("operationalData", [])
            vals = [float(r["value"]) for r in rows if r.get("value") not in (None, "")]
            print(f"  {ind} day {y}: {len(rows)} rows, {sum(vals) / 1e9:.2f} TWh")
        j = eget("aggregatedData", {"periodType": "day", "from": f"{y}-01-01", "to": f"{y}-01-31", "countryKey": "ES", "limit": -1})
        rows = j.get("aggregatedData", [])
        print(f"  aggregatedData ES Jan {y}: {len(rows)} rows", json.dumps(rows[:1])[:300])
    # Tarifa: first date with data in the API for each direction
    for pdk in ("ES-TSO-0006ITP-00082entry", "ES-TSO-0006ITP-00082exit"):
        j = eget("operationalData", {"indicator": "Physical Flow", "periodType": "day", "from": "2010-01-01", "to": "2021-02-01",
                                      "pointDirection": pdk, "limit": 5, "sort": "periodFrom"})
        print(pdk, [(r.get("periodFrom"), r.get("value")) for r in j.get("operationalData", [])])


# ------------------------------------------------------------------------------------------------ Morocco
def grep_pdf(content, pat, maxpages=120):
    try:
        pages = pdf_text(content, maxpages)
    except Exception as e:  # noqa: BLE001
        print("   pdf error", type(e).__name__)
        return []
    print(f"   pages={len(pages)}")
    hits = []
    for i, t in enumerate(pages):
        for ln in t.split("\n"):
            if re.search(pat, ln, re.I):
                hits.append((i + 1, ln.strip()))
    return hits


def grep_xlsx(content, pat, label):
    try:
        x = pd.read_excel(io.BytesIO(content), sheet_name=None, header=None)
    except Exception as e:  # noqa: BLE001
        print("   xlsx error", type(e).__name__, str(e)[:80])
        return
    print(f"   {label}: {len(x)} sheets")
    for s, d in x.items():
        txt = d.astype(str)
        m = txt.apply(lambda col: col.str.contains(pat, case=False, regex=True)).any(axis=1)
        if m.any():
            print(f"   --- sheet {s!r} ({d.shape}) rows matching: {int(m.sum())}")
            for i in list(d.index[m])[:12]:
                print("     ", " | ".join(v for v in txt.loc[i].tolist() if v != "nan")[:300])


GASPAT = r"tahaddart|beni mathar|ain beni|gaz naturel|natural gas|cycle combin|combined cycle|gazoduc|GME"


def morocco2():
    r = get("https://www.mem.gov.ma/")
    if r is not None:
        print("mem.gov.ma home body:", r.status_code, r.text[:745])
    for u in ["https://www.mem.gov.ma/fr", "https://www.mem.gov.ma/Pages/index.aspx", "https://www.mem.gov.ma/fr/Pages/index.aspx",
              "https://www.mem.gov.ma/Pages/secteur.aspx?e=1", "https://www.mem.gov.ma/Pages/publications.aspx?e=6",
              "https://www.mem.gov.ma/en", "https://www.mem.gov.ma/fr/content/chiffres-cles"]:
        links(u, r"stat|chiffre|bilan|tableau|publication|rapport|xls|pdf|energie|gaz", 50)
    for u in ["https://www.onhym.com/", "https://www.onhym.com/fr/", "http://www.onhym.com/", "https://www.one.org.ma/", "http://www.one.org.ma/"]:
        links(u, r"rapport|annuel|activit|chiffre|statist|pdf|production|gaz", 50, verify=False, timeout=(25, 60))
    # EET (Tahaddart operator)
    for u in ["https://eet.ma/nous-connaitre/chiffres-cles/", "https://eet.ma/decouvrir-notre-activite/"]:
        rr = get(u)
        if rr is not None and rr.ok:
            t = re.sub(r"<script.*?</script>|<style.*?</style>", " ", rr.text, flags=re.S)
            t = re.sub(r"<[^>]+>", " ", t)
            t = re.sub(r"\s+", " ", t)
            print(f"\n=== EET {u}\n", t[:4000])
    # ANRE annual reports
    _, out = links("https://anre.ma/rapports-annuels/", r"\.pdf|rapport", 40)
    for h, t in out:
        if not h.lower().endswith(".pdf"):
            continue
        rr = get(h, timeout=(15, 180))
        if rr is None or not rr.ok:
            continue
        print(f"\n  ANRE DOC {h} {len(rr.content)}")
        for pg, ln in grep_pdf(rr.content, GASPAT + r"|production nationale|Thermique|GWh"):
            print(f"     p{pg}: {ln[:220]}")
    # full text of the 2024 report's electricity-mix and gas pages
    rr = get("https://anre.ma/wp-content/uploads/2025/12/ANRE-Rapport-VF-web.pdf", timeout=(15, 180))
    if rr is not None and rr.ok:
        pages = pdf_text(rr.content, 40)
        for i in (17, 18, 19, 37, 38):
            if i < len(pages):
                print(f"\n##### ANRE 2024 report page {i + 1}\n{pages[i][:5000]}")
    # HCP yearbook (Excel), energy chapter
    _, out = links("https://www.hcp.ma/downloads/Annuaires-statistiques-du-Maroc-Format-Excel_t22392.html", r"file/|Annuaire", 30)
    for h, t in out[:4]:
        if "/file/" not in h:
            continue
        rr = get(h, timeout=(15, 180))
        if rr is None or not rr.ok:
            continue
        print(f"\n  HCP {t[:70]} {h} {len(rr.content)} {rr.headers.get('content-type')} {rr.content[:4]}")
        if rr.content[:2] == b"PK":
            import zipfile
            z = zipfile.ZipFile(io.BytesIO(rr.content))
            names = z.namelist()
            print("   zip/xlsx entries:", len(names), [n for n in names if re.search(r"energ|electr|mine", n, re.I)][:20] or names[:20])
            for n in names:
                if n.lower().endswith((".xlsx", ".xls")) and re.search(r"energ|electr|mine|indus", n, re.I):
                    grep_xlsx(z.read(n), r"gaz|thermique|électri|electri|GWh", n)
            if any(n.startswith("xl/") for n in names):
                grep_xlsx(rr.content, r"gaz naturel|thermique|énergie électrique|Production d.énergie|GWh", "workbook")
        break
    # Office des Changes: monthly indicators PDF and 'produits remarquables'
    for u in ["https://www.oc.gov.ma/sites/default/files/2026-09/IMEE_Ao%C3%BBt_2026_0.pdf"]:
        rr = get(u, timeout=(15, 120))
        if rr is not None and rr.ok:
            print("\n  OC IMEE", len(rr.content))
            for pg, ln in grep_pdf(rr.content, r"gaz|énergétique|energetique|hydrocarb"):
                print(f"     p{pg}: {ln[:220]}")
    for u in ["https://www.oc.gov.ma/sites/default/files/Produits%20remarquables%20site%201_18.xlsx"]:
        rr = get(u, timeout=(15, 120))
        if rr is not None and rr.ok:
            grep_xlsx(rr.content, r"gaz|hydrocarb|énerg", "OC produits remarquables")
    links("https://www.oc.gov.ma/fr/etudes-et-statistiques/series-statistiques", r"xls|pdf|serie|série|commerce|annuel|mensuel", 60)
    links("https://services.oc.gov.ma/DataBase/CommerceExterieur/", r".", 40)
    # data.gov.ma MEF energy indicators
    j = get("https://data.gov.ma/data/fr/api/3/action/package_search", params={"q": "energie", "rows": 20})
    if j is not None and j.ok:
        for d in j.json()["result"]["results"]:
            for x in d.get("resources", []):
                u = x.get("url") or ""
                if not u.lower().split("?")[0].endswith((".xls", ".xlsx")):
                    continue
                rr = get(u, timeout=(15, 120))
                if rr is not None and rr.ok:
                    grep_xlsx(rr.content, r"gaz|thermique|électri|electri|GWh|production", f"{d['name']} {u.split('/')[-1]}")


# ------------------------------------------------------------------------------------------------ Comtrade
def trade2():
    base = "https://comtradeapi.un.org/public/v1/preview/C"
    rows = []
    for rep, partner, flow, label in ((504, None, "M", "Morocco imports"), (724, 504, "X", "Spain exports to Morocco")):
        for y in range(2015, 2026):
            p = {"reporterCode": rep, "cmdCode": "271121,271111", "flowCode": flow, "period": y}
            if partner:
                p["partnerCode"] = partner
            for a in range(4):
                r = get(f"{base}/A/HS", params=p, timeout=(15, 60))
                if r is not None and r.status_code == 429:
                    time.sleep(6 + 4 * a)
                    continue
                break
            time.sleep(4)
            if r is None or r.status_code != 200:
                print(label, y, r.status_code if r is not None else None, r.text[:150] if r is not None else "")
                continue
            d = pd.DataFrame(r.json().get("data", []))
            if d.empty:
                print(label, y, "no rows")
                continue
            d["label"] = label
            rows.append(d)
            for _, x in d.iterrows():
                if partner is None and x.get("partnerCode") not in (0,) and x.get("primaryValue", 0) < 1e7:
                    continue
                print(f"  {label} {y} HS{x.get('cmdCode')} partner={x.get('partnerDesc')} netWgt_kg={x.get('netWgt')} "
                      f"qty={x.get('qty')} {x.get('qtyUnitAbbr')} US$={x.get('primaryValue')}")
    if rows:
        pd.concat(rows).to_csv(f"{OUT}/comtrade_morocco_gas_annual.csv", index=False)
    # one monthly period to see whether Morocco reports monthly
    for per in ("202406", "202412"):
        r = get(f"{base}/M/HS", params={"reporterCode": 504, "cmdCode": "271121", "flowCode": "M", "period": per}, timeout=(15, 60))
        time.sleep(5)
        print("monthly", per, r.status_code if r is not None else None, (r.text[:300] if r is not None else ""))


if __name__ == "__main__":
    {"enagas2": enagas2, "entsog2": entsog2, "morocco2": morocco2, "trade2": trade2}[sys.argv[1]]()
