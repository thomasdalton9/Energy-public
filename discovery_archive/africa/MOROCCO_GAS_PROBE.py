"""
Probe: what can be built for Morocco's gas, and gas burn at Tahaddart (CCGT, ~384 MW) and Ain Beni Mathar (ISCC, ~472 MW).

Targets (one per Actions job; output files go to ./probe_files/):
  entsog   ENTSOG transparency: Tarifa (GME, Spain-Morocco) and Almeria (Medgaz) points, operators, daily physical flow 2014-today
  enagas   Enagas monthly statistical bulletin (international connections table) and the energy-data pages
  morocco  ONEE, Ministry of Energy Transition, ONHYM, ANRE, Office des Changes, HCP, Bank Al-Maghrib, data.gov.ma, plant operators
  trade    UN Comtrade public preview (Morocco HS 271121 natural gas imports; Spain's exports to Morocco)
  ember    Ember monthly and yearly releases: Morocco rows

Reachability/structure only: nothing here is an assumption carried in; numbers are printed as the source gives them.
Usage: python3 MOROCCO_GAS_PROBE.py <target>
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

UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36"
H = {"User-Agent": UA}
OUT = "probe_files"
os.makedirs(OUT, exist_ok=True)


def get(url, tries=2, timeout=(15, 90), **kw):
    kw.setdefault("headers", H)
    last = None
    for a in range(tries):
        try:
            r = requests.get(url, timeout=timeout, **kw)
            return r
        except requests.RequestException as e:
            last = e
            time.sleep(3)
    print(f"  ERR {url[:140]} {type(last).__name__}: {str(last)[:150]}")
    return None


def links(url, pat=r".", n=80, show=True):
    r = get(url)
    print(f"\n=== {url} -> {r.status_code if r is not None else None} {len(r.content) if r is not None else ''} "
          f"{r.url if r is not None else ''}")
    if r is None or r.status_code != 200:
        if r is not None:
            print("   body:", r.text[:300].replace("\n", " "))
        return None, []
    out, seen = [], set()
    for h, t in re.findall(r'href=["\']([^"\'#]+)["\'][^>]*>(.{0,200}?)</a', r.text, re.S):
        t = re.sub(r"<[^>]+>|\s+", " ", t).strip()
        if re.search(pat, h + " " + t, re.I) and h not in seen:
            seen.add(h)
            out.append((requests.compat.urljoin(r.url, h), t))
    if show:
        for h, t in out[:n]:
            print("   ", h[:160], "|", t[:80])
    return r, out


def pdf_grep(content, pats, maxpages=80, ctx=0):
    import pdfplumber
    hits = []
    try:
        with pdfplumber.open(io.BytesIO(content)) as pdf:
            print(f"   pages={len(pdf.pages)}")
            for i, pg in enumerate(pdf.pages[:maxpages]):
                t = pg.extract_text() or ""
                lines = t.split("\n")
                for k, ln in enumerate(lines):
                    if re.search(pats, ln, re.I):
                        seg = lines[max(0, k - ctx):k + ctx + 1]
                        hits.append((i + 1, " || ".join(s.strip() for s in seg)))
    except Exception as e:  # noqa: BLE001
        print("   pdf error", type(e).__name__, str(e)[:100])
    return hits


# ------------------------------------------------------------------------------------------------ ENTSOG
API = "https://transparency.entsog.eu/api/v1"


def eget(path, params, tries=4):
    for a in range(tries):
        try:
            r = requests.get(f"{API}/{path}", params=params, headers={"User-Agent": "Mozilla/5.0", "Accept": "application/json"},
                             timeout=(15, 300))
        except requests.RequestException as e:
            print("   ", path, type(e).__name__)
            time.sleep(10 * (a + 1))
            continue
        if r.status_code == 404:
            return {}
        if r.status_code == 200:
            return r.json()
        print("   ", path, params, r.status_code, r.text[:150])
        time.sleep(15 * (a + 1))
    return {}


PAT = r"tarifa|almer|medgaz|maghreb|magreb|gme|morocc|marruec|algeri|beni|hassi"


def entsog():
    ics = eget("interconnections", {"limit": -1}).get("interconnections", [])
    print("interconnections:", len(ics))
    sel = []
    for i in ics:
        txt = json.dumps(i, ensure_ascii=False)
        if re.search(PAT, txt, re.I) or i.get("fromCountryKey") in ("MA", "DZ") or i.get("toCountryKey") in ("MA", "DZ"):
            sel.append(i)
    keep = ["pointKey", "pointLabel", "fromSystemLabel", "fromOperatorKey", "fromOperatorLabel", "fromCountryKey",
            "fromInfrastructureTypeLabel", "toSystemLabel", "toOperatorKey", "toOperatorLabel", "toCountryKey",
            "toInfrastructureTypeLabel", "fromDirectionKey", "toDirectionKey", "validFrom", "validTo"]
    df = pd.DataFrame(sel)
    if len(df):
        df = df[[c for c in keep if c in df.columns]]
        print(df.to_string())
        df.to_csv(f"{OUT}/entsog_interconnections_morocco_algeria.csv", index=False)
    ops = eget("operators", {"limit": -1}).get("operators", [])
    print("\noperators:", len(ops))
    for o in ops:
        if (o.get("operatorCountryKey") or o.get("operatorKey", "")[:2]) in ("MA", "DZ") or re.search(PAT, json.dumps(o, ensure_ascii=False), re.I):
            print("  OP", {k: o.get(k) for k in ("operatorKey", "operatorLabel", "operatorCountryKey", "operatorTypeLabel")})
    cps = eget("connectionpoints", {"limit": -1}).get("connectionpoints", [])
    print("\nconnection points:", len(cps))
    for c in cps:
        if re.search(PAT, json.dumps(c, ensure_ascii=False), re.I):
            print("  CP", {k: c.get(k) for k in ("pointKey", "pointLabel", "pointType", "infrastructureLabel", "isCrossBorder", "isInvalid",
                                                   "pointEicCode", "tpMapX", "tpMapY", "commercialType", "validFrom", "validTo")})
    # every operator/point/direction pair at the selected points
    pairs = set()
    for i in sel:
        if i.get("fromOperatorKey"):
            pairs.add((i["fromOperatorKey"], i["pointKey"], "exit", i.get("pointLabel")))
        if i.get("toOperatorKey"):
            pairs.add((i["toOperatorKey"], i["pointKey"], "entry", i.get("pointLabel")))
    # also ask the operationalData endpoint by point only (any operator), to find rows from operators not in the list
    points = sorted({p[1] for p in pairs})
    print("\npoints to pull:", points)
    for pk in points:
        rows = eget("operationalData", {"indicator": "Physical Flow", "periodType": "day", "from": "2025-09-01", "to": "2025-09-03",
                                        "pointKey": pk, "limit": -1}).get("operationalData", [])
        for r in rows:
            pairs.add((r.get("operatorKey"), r.get("pointKey"), r.get("directionKey"), r.get("pointLabel")))
            print("   sample", pk, r.get("operatorKey"), r.get("operatorLabel"), r.get("directionKey"), r.get("periodFrom"),
                  r.get("value"), r.get("unit"), r.get("flowStatus"), r.get("isUnderMaintenance"))
    series = {}
    meta = []
    for opk, pk, dr, lab in sorted(pairs, key=lambda x: (x[1], x[0], x[2])):
        name = f"{pk}|{opk}|{dr}"
        rec, units, first = {}, set(), None
        for y in range(2010, date.today().year + 1):
            rows = eget("operationalData", {"indicator": "Physical Flow", "periodType": "day", "from": f"{y}-01-01",
                                            "to": f"{y}-12-31" if y < date.today().year else date.today().isoformat(),
                                            "pointDirection": f"{opk}{pk}{dr}", "limit": -1}).get("operationalData", [])
            for r in rows:
                units.add(r.get("unit"))
                try:
                    rec[pd.Timestamp(r["periodFrom"][:10])] = float(r["value"]) / 1e6
                except (TypeError, ValueError, KeyError):
                    pass
        if rec:
            s = pd.Series(rec).sort_index()
            series[name] = s
            meta.append((name, lab, len(s), s.index.min().date(), s.index.max().date(), units))
            print(f"\n{name} ({lab}) units={units} days={len(s)} {s.index.min().date()}..{s.index.max().date()}")
            print("  TWh/yr:", (s.groupby(s.index.year).sum() / 1000).round(3).to_dict())
            print("  nonzero days/yr:", (s[s.abs() > 0.001].groupby(s[s.abs() > 0.001].index.year).count()).to_dict())
        else:
            print(f"\n{name} ({lab}): no physical flow rows")
    if series:
        d = pd.DataFrame(series).sort_index()
        d.index.name = "gas_day"
        d.round(4).to_csv(f"{OUT}/entsog_tarifa_almeria_physical_flow_daily_GWh.csv")
        print("\nfirst nonzero / last nonzero per series:")
        for c in d:
            nz = d[c][d[c].abs() > 0.001]
            if len(nz):
                print(" ", c, nz.index.min().date(), nz.index.max().date())
    # allocations / nominations at Tarifa for a cross-check (one recent year)
    for ind in ("Allocation", "Nomination", "Firm Technical"):
        for opk, pk, dr, lab in sorted(pairs):
            if not re.search(r"tarifa", str(lab), re.I):
                continue
            rows = eget("operationalData", {"indicator": ind, "periodType": "day", "from": "2025-01-01", "to": "2025-12-31",
                                            "pointDirection": f"{opk}{pk}{dr}", "limit": -1}).get("operationalData", [])
            vals = [float(r["value"]) for r in rows if r.get("value") not in (None, "")]
            if vals:
                print(f"  {ind} {opk} {pk} {dr} 2025: {len(vals)} rows, sum {sum(vals) / 1e9:.3f} TWh, mean {sum(vals) / len(vals) / 1e6:.1f} GWh/d")


# ------------------------------------------------------------------------------------------------ Enagas
ENAGAS_PAGE = "https://www.enagas.es/en/technical-management-system/energy-data/publications/gas-statistical-bulletin/"


def enagas():
    links("https://www.enagas.es/en/technical-management-system/energy-data/", r"energy-data|json|xls|csv", 120)
    links("https://www.enagas.es/es/gestion-tecnica-sistema/datos-energeticos/", r"datos-energeticos|json|xls|csv", 120)
    files = {}
    today = date.today()
    for y in range(2015, today.year + 1):
        for m in (1, 4, 7, 10):
            if (y, m) > (today.year, today.month):
                continue
            r = get(ENAGAS_PAGE, params={"category": "", "month": m, "year": y})
            if r is None:
                continue
            for x in re.findall(r'href="(/content/dam[^"]+\.pdf)"', r.text):
                files.setdefault(x, (y, m))
    print("bulletins listed:", len(files))
    for f, ym in sorted(files.items(), key=lambda kv: kv[1]):
        print("  ", ym, f)
    # one bulletin per year (the July listing), grep for the international-connections table
    by_year = {}
    for f, ym in sorted(files.items(), key=lambda kv: kv[1]):
        if ym[1] == 7 or ym[0] not in by_year:
            by_year[ym[0]] = f
    for y, f in sorted(by_year.items()):
        r = get("https://www.enagas.es" + f)
        if r is None or r.status_code != 200:
            print(y, f, "download failed")
            continue
        print(f"\n##### {y} {f.split('/')[-1]} {len(r.content)} bytes")
        open(f"{OUT}/enagas_bulletin_{y}.pdf", "wb").write(r.content) if y in (2019, 2021, 2023, today.year) else None
        for pg, ln in pdf_grep(r.content, r"tarifa|almer|marruec|morocc|maghreb|magreb|medgaz|international conn|conexiones intern|exports|exportac", ctx=1):
            print(f"   p{pg}: {ln[:300]}")


# ------------------------------------------------------------------------------------------------ Morocco official sites
KW = r"rapport|annuel|statist|chiffre|bilan|activit|production|tableau|bord|donn|open|pdf|xls|gaz|energie|énergie|electric|électric|publication"


def morocco():
    import urllib3
    urllib3.disable_warnings()
    sites = ["https://www.one.org.ma/", "https://www.one.org.ma/fr/pages/interne.asp?esp=1&id1=3&t1=1",
             "https://www.one.org.ma/FR/pages/chiffres.asp", "https://www.one.org.ma/fr/pages/chiffrescles.asp",
             "https://www.mem.gov.ma/", "https://www.mem.gov.ma/Pages/secteur.aspx?e=2", "https://www.mem.gov.ma/Pages/publications.aspx",
             "https://www.mem.gov.ma/Lists/Lst_rapports/AllItems.aspx",
             "https://www.onhym.com/", "https://www.onhym.com/fr/", "https://www.anre.ma/", "https://anre.ma/publications/",
             "https://www.oc.gov.ma/", "https://www.oc.gov.ma/fr/etudes-et-statistiques", "https://www.oc.gov.ma/fr/statistiques",
             "https://www.hcp.ma/", "https://www.hcp.ma/downloads/", "https://www.bkam.ma/", "https://www.bkam.ma/Statistiques",
             "https://www.finances.gov.ma/fr/Pages/statistiques.aspx", "https://depf.finances.gov.ma/",
             "https://www.masen.ma/fr", "https://www.eet.ma/", "https://www.tahaddart.ma/", "https://www.amee.ma/",
             "https://www.data.gov.ma/", "https://data.gov.ma/data/fr/dataset"]
    found = {}
    for u in sites:
        r, out = links(u, KW, 60)
        for h, t in out:
            found[h] = t
    # data.gov.ma CKAN search
    for q in ["gaz", "gaz naturel", "energie", "electricite", "onee", "importations", "production electricite", "hydrocarbures"]:
        for base in ("https://data.gov.ma/data/fr/api/3/action/package_search", "https://www.data.gov.ma/data/fr/api/3/action/package_search"):
            r = get(base, params={"q": q, "rows": 30})
            if r is not None and r.ok and r.text.strip().startswith("{"):
                j = r.json()["result"]
                print(f"\nCKAN {base.split('/')[2]} q={q}: {j['count']}")
                for d in j["results"]:
                    org = (d.get("organization") or {}).get("title", "")
                    print("   DS", d["name"][:70], d.get("metadata_modified", "")[:10], org[:40])
                    for x in d.get("resources", [])[:6]:
                        print("      ", x.get("format"), (x.get("url") or "")[:140])
                break
            else:
                print("CKAN", base, q, r.status_code if r is not None else None)
    # PDFs and spreadsheets that look like annual reports / statistics: download and grep
    cand = [h for h, t in found.items() if re.search(r"\.(pdf|xlsx?)(\?|$)", h, re.I)
            and re.search(r"rapport|annuel|activit|statist|chiffre|bilan|energ|gaz|electr|tableau", h + " " + t, re.I)]
    print(f"\n{len(cand)} candidate documents")
    for h in cand[:40]:
        r = get(h, timeout=(15, 120))
        if r is None or r.status_code != 200:
            print("  doc", h[:150], r.status_code if r is not None else None)
            continue
        print(f"\n  DOC {h[:160]} {len(r.content)} bytes {r.headers.get('content-type')}")
        if h.lower().split("?")[0].endswith(".pdf") or r.content[:4] == b"%PDF":
            for pg, ln in pdf_grep(r.content, r"tahaddart|beni mathar|ain beni|gaz naturel|natural gas|cycle combin|combined cycle|GME|gazoduc|thermique|GWh"):
                print(f"     p{pg}: {ln[:220]}")
    # ONEE / plant search terms on the ONEE site search, if any
    for u in ["https://www.one.org.ma/fr/pages/recherche.asp?q=Tahaddart", "https://www.one.org.ma/FR/pages/interne.asp?esp=2&id1=7&id2=50&t2=1"]:
        links(u, r"tahaddart|mathar|rapport|pdf", 40)
    # SDX / Gharb gas operator (domestic supply) and Sound Energy (Tendrara)
    for u in ["https://www.sdxenergy.com/", "https://www.sdxenergy.com/investors/reports-and-presentations",
              "https://www.soundenergyplc.com/", "https://www.chariotenergygroup.com/"]:
        links(u, r"report|annual|result|operational|update|pdf", 40)


# ------------------------------------------------------------------------------------------------ trade (UN Comtrade public preview)
def trade():
    base = "https://comtradeapi.un.org/public/v1/preview/C"
    for freq in ("A", "M"):
        for rep, partner, flow, label in ((504, None, "M", "Morocco imports"), (724, 504, "X", "Spain exports to Morocco"),
                                          (504, 724, "M", "Morocco imports from Spain")):
            for cmd in ("271121", "271111", "2711"):
                periods = ",".join(str(y) for y in range(2015, 2026)) if freq == "A" else ",".join(f"{y}{m:02d}" for y in (2023, 2024) for m in range(1, 13))
                p = {"reporterCode": rep, "cmdCode": cmd, "flowCode": flow, "period": periods}
                if partner:
                    p["partnerCode"] = partner
                r = get(f"{base}/{freq}/HS", params=p, timeout=(15, 60))
                if r is None:
                    continue
                print(f"\n{freq} {label} HS{cmd}: HTTP {r.status_code} {r.text[:200] if r.status_code != 200 else ''}")
                if r.status_code != 200:
                    continue
                d = pd.DataFrame(r.json().get("data", []))
                if d.empty:
                    print("   no rows")
                    continue
                cols = [c for c in ("period", "reporterDesc", "partnerDesc", "partnerCode", "flowDesc", "cmdCode", "netWgt", "qty", "qtyUnitAbbr",
                                    "primaryValue", "isAggregate") if c in d.columns]
                if partner is None and "partnerCode" in d.columns:
                    d = d.sort_values(["period", "primaryValue"], ascending=[True, False])
                print(d[cols].head(80).to_string())
                d[cols].to_csv(f"{OUT}/comtrade_{freq}_{rep}_{partner}_{flow}_{cmd}.csv", index=False)
                time.sleep(2)
    # Office des Changes: try its statistics search pages (trade by product)
    for u in ["https://www.oc.gov.ma/fr/publications", "https://www.oc.gov.ma/fr/etudes-et-publications",
              "https://www.oc.gov.ma/fr/donnees-statistiques"]:
        links(u, r"commerce|echange|xls|pdf|statist|annuel|mensuel|indicateurs", 60)


# ------------------------------------------------------------------------------------------------ Ember
def ember():
    b = "https://storage.googleapis.com/emb-prod-bkt-publicdata/public-downloads/"
    for f in ("monthly_full_release_long_format.csv", "yearly_full_release_long_format.csv"):
        r = get(b + f, timeout=(15, 300))
        if r is None or r.status_code != 200:
            print(f, r.status_code if r is not None else None)
            continue
        d = pd.read_csv(io.BytesIO(r.content), low_memory=False)
        print(f"\n{f}: {len(d)} rows, columns {list(d.columns)}")
        a = d[d["Area"].astype(str).str.contains("Morocco", case=False)]
        print(f"  Morocco rows: {len(a)}")
        if a.empty:
            print("  Morocco not in", f)
            continue
        g = a[(a["Variable"].isin(["Gas", "Coal", "Total Generation", "Demand", "Net Imports", "Wind", "Solar", "Hydro", "Other Fossil"]))
              & (a["Unit"] == "TWh")]
        piv = g.pivot_table(index="Date", columns="Variable", values="Value", aggfunc="first")
        print(piv.tail(40).to_string())
        piv.to_csv(f"{OUT}/ember_{f.split('_')[0]}_morocco_TWh.csv")


if __name__ == "__main__":
    {"entsog": entsog, "enagas": enagas, "morocco": morocco, "trade": trade, "ember": ember}[sys.argv[1]]()
