"""
Europe FUTURE gas infrastructure (LNG regasification, pipelines/interconnectors, storage, upstream supply additions)
for the master future-capacity table that feeds the monthly gas/power model to 2030.

  output/Data and Chart Outputs/europe_future_gas_infrastructure.xlsx
    Projects        shared row schema (region ... notes), one row per dated change (new / expansion / closure / removal)
    History         every change of status / expected date / capacity between runs (old -> new, date seen)
    Regas capacity  one row per terminal (current ceiling + months to 2030) and the monthly totals by group
    Storage         AGSI+ facilities (current working gas) and the GIE storage map by country and status
    Summary         totals by country / status
    Sources, Tried, Releases, Doc checks, GIE LNG cards, GIE storage table, TYNDP rows  (the committed workbook is
                    the history store: heavy files are re-read only when Last-Modified / ETag / hash changes)

Sources (all official): GIE ALSI + AGSI+ APIs (terminal dtrs, send-out, facility working gas; key GIE_API_KEY),
GIE LNG Map and Storage Map PDFs (cards read with pdftotext -bbox-layout), ENTSOG TYNDP 2024 Annex A (project list,
capacity increments / decrements at interconnection points), operator / TSO / authority documents (each curated fact
is written only when the opened document contains the checked text), Norwegian Offshore Directorate fact pages,
National Gas site storage data already pulled by this repo (Rough).

Units: Bcf/d for gas assets (bcm/y x 35.3147 / 365.25; GWh/d / 11.2 GWh per mcm x 0.0353147; GIE map bcm are
normal-m3 and are NOT corrected to standard m3), TWh working gas for storage.
Date rule for the monthly totals: a stated year without a month counts from January of that year (upper bound, the
ceiling use); 'by YYYY' the same. A row is counted in the month it starts and removed in the month it ends.

Usage: python3 EUROPE_GAS_INFRA.py [--out-dir DIR] [--png-dir DIR] [--no-network]
"""
import argparse
import hashlib
import io
import os
import re
import subprocess
import sys
import tempfile
import time
import unicodedata
from datetime import date, datetime, timedelta

import numpy as np
import pandas as pd
import requests

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(os.path.dirname(HERE))
sys.path.insert(0, ROOT)
import xlsx_charts  # noqa: E402
import xlsx_notes  # noqa: E402

try:
    import api_keys  # type: ignore  # noqa: F401
except Exception:  # noqa: BLE001
    api_keys = None

OUT_DEFAULT = os.path.join(ROOT, "output", "Data and Chart Outputs")
PNG_DEFAULT = os.path.join(ROOT, "output", "PNG Charts")
OUT_NAME = "europe_future_gas_infrastructure.xlsx"
UA = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36"}
TODAY = date.today()
TODAY_S = TODAY.isoformat()

GWH_PER_MCM = 11.2
BCF_PER_MCM = 0.0353147
SCHEMA = ["region", "country", "asset_type", "project", "operator", "fuel_or_type", "capacity_value", "capacity_unit",
          "capacity_mw", "capacity_bcfd", "status", "expected_online", "online_date_basis", "first_seen", "last_seen",
          "source_publisher", "source_document", "source_url", "as_of_date", "evidence", "notes"]
MONTHS = pd.date_range("2025-01-01", "2030-12-01", freq="MS")

TRIED = []   # (what, url, result)


def tried(what, url, result):
    TRIED.append({"what": what, "url": url, "result": result, "date": TODAY_S})
    print("TRIED", what, "->", result)


def gwhd_to_bcfd(g):
    return None if g is None or pd.isna(g) else float(g) / GWH_PER_MCM * BCF_PER_MCM


def bcmy_to_bcfd(b):
    return None if b is None or pd.isna(b) else float(b) * 35.3147 / 365.25


def twh_to_bcm(t):
    return None if t is None or pd.isna(t) else float(t) / GWH_PER_MCM


def norm(s):
    s = unicodedata.normalize("NFKD", str(s)).encode("ascii", "ignore").decode().lower()
    return re.sub(r"[^a-z0-9 ]+", " ", s).strip()


def month_start(x):
    return pd.Timestamp(x).to_period("M").to_timestamp()


# ------------------------------------------------------------------ http / releases
def http(url, stream=False, headers=None, tries=3, timeout=(15, 240)):
    last = None
    for a in range(tries):
        try:
            r = requests.get(url, headers=dict(UA, **(headers or {})), timeout=timeout, stream=stream)
            if r.status_code in (429, 500, 502, 503, 504):
                last = f"HTTP {r.status_code}"
                time.sleep(4 * (a + 1))
                continue
            return r
        except requests.RequestException as e:
            last = type(e).__name__
            time.sleep(4 * (a + 1))
    raise RuntimeError(f"{url}: {last}")


def head_sig(url):
    """(last_modified, etag, length) from a HEAD (falls back to a streamed GET without reading the body)."""
    try:
        r = requests.head(url, headers=UA, timeout=(15, 60), allow_redirects=True)
        if r.status_code >= 400 or not (r.headers.get("Last-Modified") or r.headers.get("ETag") or r.headers.get("Content-Length")):
            raise RuntimeError("no usable HEAD")
    except Exception:  # noqa: BLE001
        r = requests.get(url, headers=UA, timeout=(15, 60), stream=True)
        r.close()
    return (r.headers.get("Last-Modified", ""), r.headers.get("ETag", ""), r.headers.get("Content-Length", ""))


class Releases:
    """Per-URL signature store (sheet 'Releases'): heavy files are downloaded only when the signature changes."""
    cols = ["url", "label", "last_modified", "etag", "length", "sha256", "first_fetched", "last_fetched", "last_changed", "note"]

    def __init__(self, df):
        self.df = df if df is not None and len(df) else pd.DataFrame(columns=self.cols)

    def get(self, url):
        m = self.df[self.df["url"] == url]
        return None if m.empty else m.iloc[-1].to_dict()

    def unchanged(self, url):
        old = self.get(url)
        if not old:
            return False
        try:
            lm, et, ln = head_sig(url)
        except Exception:  # noqa: BLE001
            return False
        if not (lm or et or ln):
            return False
        return (str(old.get("last_modified") or "") == lm and str(old.get("etag") or "") == et and str(old.get("length") or "") == ln)

    def record(self, url, label, content=None, note="", sig=None):
        lm, et, ln = sig or ("", "", "")
        sha = hashlib.sha256(content).hexdigest() if content is not None else ""
        old = self.get(url)
        changed = TODAY_S if (not old or str(old.get("sha256") or "") != sha) else old.get("last_changed")
        row = {"url": url, "label": label, "last_modified": lm, "etag": et, "length": ln, "sha256": sha,
               "first_fetched": (old or {}).get("first_fetched") or TODAY_S, "last_fetched": TODAY_S, "last_changed": changed, "note": note}
        self.df = pd.concat([self.df[self.df["url"] != url], pd.DataFrame([row])], ignore_index=True)


# ------------------------------------------------------------------ GIE ALSI / AGSI
def gie_key():
    k = os.environ.get("GIE_API_KEY") or (getattr(api_keys, "GIE_API_KEY", "") if api_keys else "")
    return k


def gie_json(base, params, key):
    for a in range(4):
        try:
            r = requests.get(base, params=params, headers=dict(UA, **{"x-key": key}), timeout=(10, 90))
        except requests.RequestException:
            time.sleep(4 * (a + 1))
            continue
        if r.status_code == 200:
            return r.json()
        if r.status_code in (429, 500, 502, 503):
            time.sleep(6 * (a + 1))
            continue
        return None
    return None


def gie_facilities(base, key, kind):
    """About endpoint -> flat list of facilities with company, country, type and operating dates."""
    j = gie_json(base + "/about", {}, key)
    out = []
    if not j:
        return out
    for typ, regs in j.items():
        for reg, countries in regs.items():
            for cname, companies in countries.items():
                for co in companies:
                    for f in co.get("facilities", []):
                        out.append({"kind": kind, "region": reg, "country": cname, "company": co.get("short_name"),
                                    "company_eic": co.get("eic"), "fac_eic": f.get("eic"), "name": f.get("name"),
                                    "ftype": f.get("type"), "start": f.get("operational_start_date"),
                                    "end": f.get("operational_end_date"), "cc": (f.get("country") or {}).get("code")})
    return out


def gie_latest(base, key, f, days=12):
    frm = (TODAY - timedelta(days=days)).isoformat()
    for params in ({"country": f["cc"], "company": f["company_eic"], "facility": f["fac_eic"], "from": frm, "to": TODAY_S, "size": 30},
                   {"country": f["cc"], "facility": f["fac_eic"], "from": frm, "to": TODAY_S, "size": 30}):
        j = gie_json(base, params, key)
        if j and j.get("data"):
            return j["data"][0]
    return None


def num(x):
    try:
        v = float(x)
        return v
    except (TypeError, ValueError):
        return np.nan


# ------------------------------------------------------------------ GIE map PDFs (pdftotext -bbox-layout)
def pdf_words(pdf_path):
    out = os.path.join(tempfile.mkdtemp(), "m.html")
    subprocess.run(["pdftotext", "-bbox-layout", pdf_path, out], check=True)
    t = open(out, encoding="utf-8", errors="ignore").read()
    return [(float(a), float(b), float(c), float(d), w) for a, b, c, d, w in
            re.findall(r'<word xMin="([\d.]+)" yMin="([\d.]+)" xMax="([\d.]+)" yMax="([\d.]+)">([^<]*)</word>', t)]


# name on the map -> (country, canonical key)
KNOWN_LNG = {
    "Zeebrugge": "Belgium", "Gdańsk 2": "Poland", "Gdańsk": "Poland", "Stade": "Germany", "Brunsbüttel": "Germany",
    "Wilhelmshaven 1": "Germany", "Wilhelmshaven 2": "Germany", "Wilhelmshaven": "Germany", "Lubmin 2": "Germany",
    "Mukran": "Germany", "Krk Island": "Croatia", "Inkoo": "Finland", "Dunkerque LNG": "France", "Fos Cavaou": "France",
    "Fos Tonkin": "France", "Montoir de Bretagne": "France", "Le Havre": "France", "OLT Offshore LNG": "Italy",
    "Klaipėda": "Lithuania", "Alexandroupolis": "Greece", "Revithoussa": "Greece", "Ravenna": "Italy", "Rovigo": "Italy",
    "Delimara": "Malta", "Piombino": "Italy", "Panigaglia": "Italy", "Mugardos": "Spain", "Sagunto": "Spain",
    "Świnoujście": "Poland", "Sines": "Portugal", "Huelva": "Spain", "El Musel": "Spain", "EemsEnergyTerminal": "Netherlands",
    "Gate terminal": "Netherlands", "Cartagena": "Spain", "Barcelona": "Spain", "Bilbao": "Spain", "Dörtyol": "Turkey",
    "Gulf of Saros": "Turkey", "Marmara Ereglisi": "Turkey", "Isle of Grain": "United Kingdom", "Aliaga Etki": "Turkey",
    "Aliaga Izmir": "Turkey", "Dragon LNG": "United Kingdom", "South Hook": "United Kingdom", "Teesside Gasport": "United Kingdom",
    "Vlorë": "Albania", "Vasiliko": "Cyprus", "Paldiski": "Estonia", "Tallinn (Muuga)": "Estonia", "Argo": "Greece",
    "Dioriga Gas": "Greece", "Thessaloniki": "Greece", "Thrace LNG": "Greece", "Shannon LNG": "Ireland", "Mag Mell": "Ireland",
    "Porto Empedocle": "Italy", "Skulte": "Latvia", "Black Sea LNG": "Romania", "Hamina": "Finland", "Kaliningrad": "Russia",
   
}
KNOWN_ORDER = sorted(KNOWN_LNG, key=lambda s: -len(s))   # longest first ('Wilhelmshaven 2' before 'Wilhelmshaven')


def parse_lng_cards(words):
    """GIE LNG map -> one row per terminal card: name, country, start-up year, vessel/type, status (by map column),
    nominal annual capacity and the expansion lines, all in bcm(N)/y as printed."""
    anchors = [w for w in words if w[4] == "Start-up" and w[0] > 300]   # x<300: legend sample card
    rows, bad = [], []
    for a in anchors:
        ax, ay = a[0], a[1]
        planned_col = ax >= 3100 and ay >= 1250
        compact = planned_col
        y_hi = ay + (46 if compact else 100)
        box = [w for w in words if ax - 12 < w[0] < ax + 218 and ay - (14 if compact else 26) < w[1] < y_hi]
        box.sort(key=lambda w: (round(w[1] / 4), w[0]))
        band = sorted([w for w in box if ay - 16 < w[1] < ay + 2], key=lambda w: w[0])
        band_txt = " ".join(w[4] for w in band)
        name = next((k for k in KNOWN_ORDER if re.search(r"(?<!\w)" + re.escape(k) + r"(?!\w)", band_txt)), None)
        if name is None:
            band_txt2 = " ".join(w[4] for w in box)
            name = next((k for k in KNOWN_ORDER if re.search(r"(?<!\w)" + re.escape(k) + r"(?!\w)", band_txt2)), None)
        txt = " ".join(w[4] for w in box)
        if name == "Gdańsk" and planned_col:
            name = "Gdańsk 2"
        if name == "Wilhelmshaven" and not planned_col:
            name = "Wilhelmshaven 1"
        m = re.search(r"Start-up\s*:\s*(\d{4}|-)", txt)
        start = int(m.group(1)) if m and m.group(1).isdigit() else None
        if name == "Wilhelmshaven 1" and start == 2025:
            name = "Wilhelmshaven 2"
        if name is None:
            bad.append((round(ax), round(ay), band_txt[:60]))
            continue
        vessel = "FSRU" if re.search(r"\bFSRU\b|\bFRU\b", txt) else ("FSU" if "FSU" in txt else ("GBS" if "GBS" in txt else "onshore"))
        if ax < 2300:
            status = "UNDER_CONSTRUCTION"
        elif planned_col:
            status = "SUSPENDED" if ay >= 2195 else "PLANNED"
        else:
            status = "OPERATIONAL"
        caps = []
        for mm in re.finditer(r"(?:(current|future|additional|Possible|interruptible|by\s*(\d{4}))\s*:?\s*)?(\+?)\s*(\d+(?:,\d+)?)\s*billion", txt):
            label = (mm.group(1) or "").strip()
            caps.append((label, mm.group(2), mm.group(3) == "+", float(mm.group(4).replace(",", "."))))
        nominal, exp = None, []
        for label, yr, plus, v in caps:
            if plus or label.startswith("by") or label in ("future", "additional", "Possible"):
                exp.append({"label": label or "additional", "year": int(yr) if yr else None, "bcm": v})
            elif label == "interruptible":
                exp.append({"label": "interruptible", "year": None, "bcm": v, "info": True})
            elif nominal is None:
                nominal = v
        rows.append({"name": name, "country": KNOWN_LNG[name], "start_up_year": start, "vessel": vessel, "status": status,
                     "nominal_bcm_n": nominal, "expansions": "; ".join(f"{e['label']}{' ' + str(e['year']) if e.get('year') else ''}: +{e['bcm']}" for e in exp),
                     "exp_list": exp, "x": round(ax), "y": round(ay)})
    # suspended cards (below the SUSPENDED heading) carry no 'Start-up' line
    head = [w for w in words if w[4] == "SUSPENDED" and w[0] > 3100 and w[1] > 2150]
    if head:
        y0 = min(w[1] for w in head)
        for nm, first in (("Le Havre", "Le"), ("Teesside Gasport", "Teesside")):
            cand = [w for w in words if w[4] == first and w[0] > 3100 and w[1] > y0]
            if not cand:
                continue
            c = min(cand, key=lambda w: w[1])
            box = sorted([w for w in words if c[0] - 12 < w[0] < c[0] + 218 and c[1] - 4 < w[1] < c[1] + 50], key=lambda w: (round(w[1] / 4), w[0]))
            txt = " ".join(w[4] for w in box)
            mm = re.search(r"(\d+(?:,\d+)?)\s*billion", txt)
            rows.append({"name": nm, "country": KNOWN_LNG[nm], "start_up_year": None, "vessel": "FSRU" if "FSRU" in txt else "onshore", "status": "SUSPENDED",
                         "nominal_bcm_n": float(mm.group(1).replace(",", ".")) if mm else None, "expansions": "", "exp_list": [], "x": round(c[0]), "y": round(c[1])})
    return rows, bad


PARSER = "2"


def parse_storage_table(words):
    """GIE storage map summary table: working gas (TWh) and number of facilities by status and country."""
    hdr = sorted([w for w in words if 95 < w[1] < 120 and w[4] == "operational"], key=lambda w: w[0])
    heads = [w for w in hdr if w[0] > 2000]
    if len(heads) < 2:
        return []
    out = {}
    names = set(["Austria", "Belgium", "Bulgaria", "Croatia", "Czech", "Denmark", "France", "Germany", "Greece", "Hungary", "Ireland", "Italy",
                 "Latvia", "Netherlands", "Poland", "Portugal", "Romania", "Slovakia", "Spain", "Sweden", "EU27", "Serbia", "Ukraine",
                 "Turkey", "Türkiye", "United", "Albania", "Slovenia", "Lithuania", "Estonia", "Finland", "Switzerland"])
    for table, h in (("twh", heads[0]), ("n", heads[1])):
        c0 = (h[0] + h[2]) / 2
        cols = [c0, c0 + 50, c0 + 99, c0 + 148, c0 + 201]
        rows = {}
        inreg = [w for w in words if 115 < w[1] < 800 and c0 - 90 < (w[0] + w[2]) / 2 < c0 + 235]
        namew = [w for w in inreg if (w[0] + w[2]) / 2 < c0 - 40 and w[4] in names]
        seen = set()
        for nw in sorted(namew, key=lambda w: w[1]):
            if any(abs(nw[1] - y) < 4 for y in seen):
                continue
            seen.add(nw[1])
            ws = sorted([w for w in inreg if abs(w[1] - nw[1]) < 4], key=lambda w: w[0])
            left = [w[4] for w in ws if (w[0] + w[2]) / 2 < c0 - 40]
            country = " ".join(left)
            if country.startswith("Czech"):
                country = "Czechia"
            vals = [None] * 5
            for w in ws:
                xc = (w[0] + w[2]) / 2
                if xc >= c0 - 40 and re.fullmatch(r"\d+(?:,\d+)?", w[4]):
                    k = int(np.argmin([abs(xc - c) for c in cols]))
                    if abs(xc - cols[k]) < 20:
                        vals[k] = float(w[4].replace(",", "."))
            out.setdefault(country, {})[table] = vals
    rows = []
    for c, d in out.items():
        t, n = d.get("twh", [None] * 5), d.get("n", [None] * 5)
        rows.append({"country": c, "operational_twh": t[0], "under_construction_twh": t[1], "planned_twh": t[2], "mothballed_twh": t[3],
                     "total_twh": t[4], "operational_n": n[0], "under_construction_n": n[1], "planned_n": n[2], "mothballed_n": n[3], "total_n": n[4]})
    return rows


def map_version(words):
    txt = " ".join(w[4] for w in words)
    m = re.search(r"Version\s*:\s*([A-Za-z]+ \d{4})", txt) or re.search(r"status of ([A-Za-z]+ \d{4})", txt)
    return m.group(1) if m else ""


def find_pdf(page, pat):
    r = http(page)
    m = re.findall(r'href="([^"]*' + pat + r'[^"]*\.pdf)"', r.text, re.I)
    return m[0] if m else None


# ------------------------------------------------------------------ TYNDP 2024 Annex A
TYNDP_A = "https://www.entsog.eu/sites/default/files/2024-07/TYNDP%202024%20Annex%20A%20-%20List%20of%20projects.xlsx"
TYNDP_PAGE = "https://tyndp2024.entsog.eu/downloads/"


def parse_tyndp(content):
    x = pd.read_excel(io.BytesIO(content), sheet_name=None, header=None)
    m = x["Project Main Info"].copy()
    m.columns = [str(v) for v in m.iloc[2]]
    m = m.iloc[3:]
    m = m[m["Code"].astype(str).str.contains("-")]
    main = m[["Code", "Project Name", "Country", "Promoter", "Maturity Status", "Project Commissioning Year First", "Project Commissioning Year Last"]].copy()
    main.columns = ["code", "name", "country", "promoter", "maturity", "year_first", "year_last"]
    ci = x["Capacity Increments"].copy()
    hi = next(i for i in range(15) if str(ci.iloc[i, 1]).strip() == "Point")
    ci.columns = [str(v).strip() for v in ci.iloc[hi]]
    ci = ci.iloc[hi + 1:].dropna(subset=["Code"])
    ci = ci.rename(columns={"Capacity (GWh/d)*": "cap_gwhd", "Commissioning Year": "year", "Project Name": "name", "Status": "maturity"})
    ci["cap_gwhd"] = pd.to_numeric(ci["cap_gwhd"], errors="coerce")
    ci["year"] = pd.to_numeric(ci["year"], errors="coerce")
    ci = ci[["Point", "From System", "To System", "Code", "name", "Promoter", "maturity", "Operator", "year", "cap_gwhd"]]
    ci.columns = ["point", "from_system", "to_system", "code", "name", "promoter", "maturity", "operator", "year", "cap_gwhd"]
    lng = x["LNG"].copy()
    lng.columns = [str(v) for v in lng.iloc[3]]
    lng = lng.iloc[4:].dropna(subset=["Code"])
    lng = lng[["Code", "LNG Facility", "Maturity Status", "Project Phase", "Project Yearly Volume (bcm/y)"]]
    lng.columns = ["code", "facility", "maturity", "phase", "bcm_y"]
    ugs = x["UGS"].copy()
    ugs.columns = [str(v) for v in ugs.iloc[3]]
    ugs = ugs.iloc[4:].dropna(subset=["Code"])
    keep = {"Code": "code", "UGS Facility": "facility", "Type": "utype", "Project Phase": "phase", "WGV (GWh)": "wgv_gwh",
            "Withdrawal Capacity (GWh/d)": "wd_gwhd", "Injection Capacity (GWh/d)": "inj_gwhd"}
    ugs = ugs[[c for c in keep if c in ugs.columns]].rename(columns=keep)
    return main, ci, lng, ugs


# ------------------------------------------------------------------ curated operator / authority documents
DOCS = {
    "det_stade": ("https://energy-terminal.de/en/terminals/stade", "Deutsche Energy Terminal", "Stade terminal page",
                  {"cap17": r"feed-in capacity of up to 1\.7 billion cubic meters through the end of 2026", "docked": r"Energos Force has docked in Stade"}),
    "det_diary": ("https://energy-terminal.de/en/construction-diary-stade", "Deutsche Energy Terminal", "Construction diary Stade",
                  {"cap32": r"up to 3\.2 billion cubic meters", "nov": r"first LNG tanker to arrive in November"}),
    "gaz_gdansk": ("https://www.gaz-system.pl/en/for-media/press-releases/2026/august/11-08-2026-the-offshore-pipeline-microtunnel-for-the-fsru-has-reached-its-target-length.html",
                   "GAZ-SYSTEM", "Press release 11 Aug 2026 (Gdansk FSRU)", {"early2028": r"scheduled to commence in early 2028"}),
    "gaz_fsru": ("https://www.gaz-system.pl/en/for-media/press-releases/2025/december/11-12-2025-strategic-project-for-real-independence-10-years-of-the-lng-terminal-in-swinoujscie.html",
                 "GAZ-SYSTEM", "Press release 11 Dec 2025 (Swinoujscie 10 years / FSRU)", {"cap61": r"regasification capacity: 6\.1 billion"}),
    "gaz_fsru2": ("https://www.gaz-system.pl/en/for-media/press-releases/2025/december/23-12-2025-market-consultations-for-the-binding-fsru-2-open-season-procedure.html",
                  "GAZ-SYSTEM", "Press release 23 Dec 2025 (FSRU 2 open season)", {"cap45": r"additional 4\.5 bcma"}),
    "gasunie_ee": ("https://www.gasunie.nl/en/newsletter/gasunie-update-2-2026/seven-questions-about-lng-and-the-continued-use-of-eemsenergyterminal",
                   "Gasunie", "Gasunie update 2-2026: EemsEnergyTerminal", {"to2036": r"continued operations planned until 2036", "cfid": r"conditional investment decision"}),
    "gasunie_ee2": ("https://www.gasunie.nl/en/news/important-new-phase-for-the-future-of-eemsenergyterminal-at-the-port-of-eemshaven",
                    "Gasunie", "Gasunie news: EemsEnergyTerminal new phase", {"orig2027": r"originally planned to run up to the end of 2027"}),
    "lng_hr": ("https://lng.hr/en/2026/01/01/5-years-of-operation-of-the-lng-terminal-croatia/", "LNG Hrvatska", "5 years of operation (1 Jan 2026)",
               {"exp": r"maximum capacity increased from 3\.9 to 6\.1 billion"}),
    "gasgrid_inkoo": ("https://gasgrid.fi/en/lng-terminal/inkoo-lng-terminal/", "Gasgrid Finland", "Inkoo LNG terminal",
                      {"cap": r"140 ?GWh per day and over 40 TWh per year"}),
    "gasgrid_lease": ("https://gasgrid.fi/en/bulletins/gasgrid-finland-oy-and-excelerate-energy-inc-have-agreed-to-lease-an-lng-terminal-vessel-the-lng-terminal-vessel-exemplar-will-ensure-finlands-security-of-supply-while-also-serving-the-needs-of/",
                      "Gasgrid Finland", "Bulletin 20 May 2022: Exemplar lease", {"ten": r"leasing of an LNG floating terminal vessel for the duration of ten years"}),
    "total_lehavre": ("https://totalenergies.com/news/press-releases/france-totalenergies-demobilizes-its-floating-lng-terminal-le-havre",
                      "TotalEnergies", "Press release 25 Nov 2025: Le Havre FSRU demobilised", {"demob": r"decided to demobilize its LNG FSRU in Le Havre"}),
    "snam_piombino": ("https://www.snam.it/content/dam/snam/pages-attachments/it/i-nostri-business/rigassificazione/documents/codice-tariffe-area-comitato/codice-di-rigassificazione-fsru-italia/eng-regas-code-piombino/02_DES~1.PDF",
                      "Snam", "Piombino regasification code: description of the terminal", {"reloc": r"shutdown, relocation and recommissioning of the Terminal will last for a maximum of twelve months, starting from October 1, 2026"}),
    "gastrade_cap": ("https://www.gastrade.gr/further-increase-in-maximum-regasification-capacity/", "Gastrade", "Announcement 21 Oct 2025",
                     {"step": r"increased to 136\.2 GWh/day, up from 90\.8 GWh/day"}),
    "gastrade_proj": ("https://www.gastrade.gr/the-project/", "Gastrade", "Alexandroupolis INGS / Thrace INGS", {"cap": r"5\.5 bcm Annual Capacity", "thrace": r"Thrace Independent Natural Gas System"}),
    "tap_first": ("https://www.tap-ag.com/news/news-stories/tap-triggers-the-first-level-of-capacity-expansion", "Trans Adriatic Pipeline",
                  "News 30 Jan 2023: first level of capacity expansion", {"fullop": r"reconfirm full operability of the first level of capacity expansion, triggered during the 2021 Market Test, with 1\.2 billion cubic meters per year",
                   "dbl": r"doubling the pipeline.s capacity by 2027", "bcm": r"1\.2 billion cubic meters \(bcm\) of incremental capacity per year"}),
    "omv_neptun": ("https://www.omvpetrom.com/en/media/latest-news/2026/neptun-deep-advances-towards-first-gas-expected-in-the-first-half-of-2027",
                   "OMV Petrom", "Press release 28 Sep 2026: Neptun Deep", {"h1": r"first gas production to enter the National Transmission System \(NTS\) in the first half of 2027",
                                                                             "bcm": r"around 8 bcm \(100%\)"}),
    "np_irpa": ("https://www.norskpetroleum.no/en/facts/field/irpa/", "Norwegian Offshore Directorate / Norwegianpetroleum.no", "Field page IRPA",
                {"start": r"production is planned to start in 2026"}),
    "np_fenris": ("https://www.norskpetroleum.no/en/facts/field/fenris/", "Norwegian Offshore Directorate / Norwegianpetroleum.no", "Field page FENRIS",
                  {"start": r"production is planned to start in 2027"}),
    "np_berling": ("https://www.norskpetroleum.no/en/facts/field/berling/", "Norwegian Offshore Directorate / Norwegianpetroleum.no", "Field page BERLING",
                   {"start": r"production is planned to start in 2028"}),
    "np_hugin": ("https://www.norskpetroleum.no/en/facts/field/hugin/", "Norwegian Offshore Directorate / Norwegianpetroleum.no", "Field page HUGIN",
                 {"status": r"Status Approved for production"}),
    "np_munin": ("https://www.norskpetroleum.no/en/facts/field/munin/", "Norwegian Offshore Directorate / Norwegianpetroleum.no", "Field page MUNIN", {"status": r"Status Approved for production"}),
    "np_fulla": ("https://www.norskpetroleum.no/en/facts/field/fulla/", "Norwegian Offshore Directorate / Norwegianpetroleum.no", "Field page FULLA", {"status": r"Status Approved for production"}),
    "btg": ("https://bulgartransgaz.bg/public/files/useruploads/%D0%BD%D0%BE%D0%B2%D0%B8%D0%BD%D0%B8_%D1%81%D0%B0%D0%B9%D1%82/27042026/TYNDP%20BTG%202026-2035%20%D0%B3.%20-%20public%20consultation%20EN.pdf",
            "Bulgartransgaz", "Ten-year network development plan 2026-2035 (public consultation)",
            {"negru": r"increase of gas transmission capacity by 137\.2 GWh/d from Bulgaria to Romania at IP Negru Voda 1/Kardam in Q3 2026",
             "kulata": r"capacity on its side of IP Kulata/Sidirokastro by 35\.4 GWh/d in direction from",
             "depo": r"increase in storage capacity up to 400 mcm, with estimated commissioning in 2026"}),
    "transgaz": ("https://www.transgaz.ro/sites/default/files/DEVELOPMENT%20PLAN%20FOR%20THE%20NATIONAL%20GAS%20TRANSMISSION%20SYSTEM%202026-2035.pdf",
                 "Transgaz", "Development plan for the NTS 2026-2035", {"vc": r"Vertical Corridor.{0,400}2029.{0,200}4\.38.{0,100}5\.32 bcm/year"}),
}


def fetch_doc(url):
    r = http(url, timeout=(15, 120))
    if r.status_code != 200:
        return None, f"HTTP {r.status_code}"
    c = r.content
    if c[:4] == b"%PDF":
        p = os.path.join(tempfile.mkdtemp(), "d.pdf")
        open(p, "wb").write(c)
        q = p + ".txt"
        subprocess.run(["pdftotext", "-layout", p, q], check=False)
        t = open(q, encoding="utf-8", errors="ignore").read() if os.path.exists(q) else ""
    else:
        t = re.sub(r"<(script|style)[^>]*>.*?</\1>", " ", r.text, flags=re.S)
        t = re.sub(r"<[^>]+>", " ", t)
        import html as _h
        t = _h.unescape(t)
    return re.sub(r"\s+", " ", t).replace(" ", " ").replace("’", "'"), hashlib.sha256(c).hexdigest()[:16]


def run_doc_checks(prev_checks):
    """-> dict (doc, check) -> bool, plus the Doc checks table. Documents are fetched every run (small pages); a PDF
    larger than 1 MB is re-parsed only when its hash changes (previous result is reused)."""
    res, rows = {}, []
    prev = {(r["doc"], r["check"]): r for _, r in prev_checks.iterrows()} if prev_checks is not None and len(prev_checks) else {}
    for did, (url, pub, title, checks) in DOCS.items():
        try:
            t, h = fetch_doc(url)
        except Exception as e:  # noqa: BLE001
            t, h = None, f"{type(e).__name__}"
        if t is None:
            tried(f"document {did}", url, f"not read ({h}); facts that rely on it are not written")
            for c in checks:
                old = prev.get((did, c))
                res[(did, c)] = False
                rows.append({"doc": did, "check": c, "passed": False, "snippet": "", "url": url, "doc_hash": h or "", "checked": TODAY_S})
            continue
        for c, rx in checks.items():
            m = re.search(rx, t, re.I | re.S)
            res[(did, c)] = bool(m)
            snip = t[max(0, m.start() - 60): m.end() + 60][:260] if m else ""
            rows.append({"doc": did, "check": c, "passed": bool(m), "snippet": snip, "url": url, "doc_hash": h, "checked": TODAY_S})
            if not m:
                tried(f"check {did}.{c}", url, "text not found in the opened document (wording changed?)")
    return res, pd.DataFrame(rows)


# ------------------------------------------------------------------ build
def row(**k):
    r = {c: None for c in SCHEMA}
    r.update(region="Europe", first_seen=TODAY_S, last_seen=TODAY_S)
    r.update(k)
    return r


def mkpub(did):
    u, pub, title, _ = DOCS[did]
    return pub, title, u


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out-dir", default=OUT_DEFAULT)
    ap.add_argument("--png-dir", default=PNG_DEFAULT)
    ap.add_argument("--no-network", action="store_true")
    ap.add_argument("--no-charts", action="store_true")
    a = ap.parse_args()
    os.makedirs(a.out_dir, exist_ok=True)
    path = os.path.join(a.out_dir, OUT_NAME)

    def prev_sheet(name):
        try:
            return pd.read_excel(path, sheet_name=name)
        except Exception:  # noqa: BLE001
            return None
    prev_projects, prev_hist = prev_sheet("Projects"), prev_sheet("History")
    rel = Releases(prev_sheet("Releases"))
    prev_checks = prev_sheet("Doc checks")
    key = gie_key()

    # ---- 1. ALSI / AGSI facility data
    alsi_f, agsi_f, alsi_now, agsi_now = [], [], {}, {}
    if key and not a.no_network:
        alsi_f = gie_facilities("https://alsi.gie.eu/api", key, "LNG")
        agsi_f = gie_facilities("https://agsi.gie.eu/api", key, "UGS")
        for f in alsi_f:
            if f["end"] and f["end"] < "2025-01-01":
                continue
            d = gie_latest("https://alsi.gie.eu/api", key, f)
            if d:
                alsi_now[f["fac_eic"] + "|" + f["name"]] = d
            elif not f["end"]:
                tried("ALSI latest data", f["name"], "no row in the last 12 days (no declared capacity read)")
            time.sleep(0.15)
        for f in agsi_f:
            if f["end"] and f["end"] < "2025-01-01":
                continue
            if f["start"] and f["start"] > TODAY_S:
                continue
            d = gie_latest("https://agsi.gie.eu/api", key, f)
            if d:
                agsi_now[f["fac_eic"] + "|" + f["name"]] = d
            time.sleep(0.15)
        print("ALSI facilities", len(alsi_f), "with data", len(alsi_now), "| AGSI", len(agsi_f), "with data", len(agsi_now))
    else:
        tried("GIE ALSI/AGSI", "https://alsi.gie.eu/api", "GIE_API_KEY missing or --no-network: no declared capacities read")

    # ---- 2. GIE maps (conditional download)
    lng_cards = prev_sheet("GIE LNG cards")
    stor_tab = prev_sheet("GIE storage table")
    map_ver = {}
    if not a.no_network:
        for tag, page, pat in (("lng", "https://www.gie.eu/publications/maps/gie-lng-map/", r"filr/\d+/GIE_LNG"),
                               ("stor", "https://www.gie.eu/publications/maps/gie-storage-map/", r"filr/\d+/GIE_STOR")):
            try:
                url = find_pdf(page, pat)
                if not url:
                    tried(f"GIE {tag} map link", page, "no PDF link found on the page")
                    continue
                have = lng_cards if tag == "lng" else stor_tab
                if have is not None and len(have) and "parser" in have.columns and str(have["parser"].iloc[0]) == PARSER and rel.unchanged(url):
                    print("GIE", tag, "map unchanged, parsed rows reused")
                    continue
                sig = head_sig(url)
                r = http(url, timeout=(20, 600))
                tmp = os.path.join(tempfile.mkdtemp(), "m.pdf")
                open(tmp, "wb").write(r.content)
                w = pdf_words(tmp)
                map_ver[tag] = map_version(w)
                if tag == "lng":
                    cards, bad = parse_lng_cards(w)
                    for b in bad:
                        tried("GIE LNG card not recognised", url, f"at x,y={b[0]},{b[1]} text '{b[2]}'")
                    lng_cards = pd.DataFrame(cards)
                    lng_cards["map_version"] = map_ver[tag]
                    lng_cards["parser"] = PARSER
                    lng_cards["url"] = url
                else:
                    stor_tab = pd.DataFrame(parse_storage_table(w))
                    stor_tab["map_version"] = map_ver[tag]
                    stor_tab["parser"] = PARSER
                    stor_tab["url"] = url
                rel.record(url, f"GIE {tag} map", r.content, f"version {map_ver[tag]}", sig)
            except Exception as e:  # noqa: BLE001
                tried(f"GIE {tag} map", page, f"{type(e).__name__}: {str(e)[:120]}")
    if lng_cards is not None and "exp_list" not in lng_cards.columns:
        lng_cards["exp_list"] = None
    lng_ver = (lng_cards["map_version"].iloc[0] if lng_cards is not None and len(lng_cards) else "")
    stor_ver = (stor_tab["map_version"].iloc[0] if stor_tab is not None and len(stor_tab) else "")

    # ---- 3. TYNDP 2024 Annex A
    tyndp = None
    cache = {k: prev_sheet(f"TYNDP {k}") for k in ("main", "capinc", "lng", "ugs")}
    if not a.no_network:
        try:
            if all(v is not None for v in cache.values()) and rel.unchanged(TYNDP_A):
                tyndp = tuple(cache[k] for k in ("main", "capinc", "lng", "ugs"))
                print("TYNDP Annex A unchanged, parsed rows reused")
            else:
                sig = head_sig(TYNDP_A)
                r = http(TYNDP_A, timeout=(20, 300))
                tyndp = parse_tyndp(r.content)
                rel.record(TYNDP_A, "ENTSOG TYNDP 2024 Annex A - List of projects", r.content, "edition July 2024", sig)
        except Exception as e:  # noqa: BLE001
            tried("TYNDP 2024 Annex A", TYNDP_A, f"{type(e).__name__}: {str(e)[:120]}")
    if tyndp is None and all(v is not None for v in cache.values()):
        tyndp = tuple(cache[k] for k in ("main", "capinc", "lng", "ugs"))

    # ---- 4. curated documents
    checks, checks_df = ({}, pd.DataFrame())
    if not a.no_network:
        checks, checks_df = run_doc_checks(prev_checks)
    ok = lambda d, c: bool(checks.get((d, c)))  # noqa: E731

    # ---- 5. terminal master table
    terms = build_terminals(alsi_f, alsi_now, lng_cards, lng_ver)
    if ok("lng_hr", "exp"):   # operator: 6.1 bcm/y since Nov 2025; ALSI dtrs / map lag
        for t in terms:
            if t["key"] == "krk":
                t["ceiling_bcfd"], t["ceiling_basis"] = bcmy_to_bcfd(6.1), "operator (LNG Hrvatska: maximum capacity 6.1 bcm/y since Nov 2025)"

    # ---- 6. Projects rows
    rows = []
    rows += regas_rows(terms, lng_cards, lng_ver, ok, tyndp)
    rows += pipeline_rows(tyndp, ok, path)
    rows += storage_rows(tyndp, stor_tab, stor_ver, agsi_f, ok)
    rows += upstream_rows(ok, a)
    proj = pd.DataFrame(rows, columns=SCHEMA)
    proj = raw_data_only(proj)
    proj = merge_first_seen(proj, prev_projects)
    hist = build_history(proj, prev_projects, prev_hist)

    # ---- 7. timelines
    reg_tab, reg_det, reg_month = regas_timeline(terms, proj)
    stor_sheet = storage_sheet(agsi_f, agsi_now, stor_tab)

    summ = summary(proj, terms, reg_month)
    sources = sources_table(lng_ver, stor_ver)
    tried_df = pd.DataFrame(TRIED) if TRIED else pd.DataFrame(columns=["what", "url", "result", "date"])

    sheets = {
        "Projects": proj, "History": hist, "Regas capacity": reg_tab, "Regas by month": reg_det, "Regas monthly": reg_month, "Storage": stor_sheet, "Summary": summ,
        "Sources": sources, "Tried": tried_df, "Releases": rel.df, "Doc checks": checks_df,
        "GIE LNG cards": (lng_cards.drop(columns=["exp_list"], errors="ignore") if lng_cards is not None else pd.DataFrame()),
        "GIE storage table": stor_tab if stor_tab is not None else pd.DataFrame(),
    }
    if tyndp is not None:
        for k, v in zip(("main", "capinc", "lng", "ugs"), tyndp):
            sheets[f"TYNDP {k}"] = v
    write_all(path, sheets, reg_month, proj, stor_sheet, a)


# ------------------------------------------------------------------ terminals
def canon(name, vessel=None):
    n = norm(name)
    n = re.sub(r"\b(lng|terminal|fsru|gas|the|energy|offshore)\b", " ", n)
    n = re.sub(r"\s+", " ", n).strip()
    return n


ALSI_KEY = [  # (regex on ALSI facility name, canonical key, map name)
    (r"Zeebrugge", "zeebrugge", "Zeebrugge"), (r"Krk", "krk", "Krk Island"), (r"Inkoo", "inkoo", "Inkoo"), (r"Hamina", "hamina", "Hamina"),
    (r"Fos Tonkin", "fos tonkin", "Fos Tonkin"), (r"Montoir", "montoir", "Montoir de Bretagne"), (r"Dunkerque", "dunkerque", "Dunkerque LNG"),
    (r"Fos Cavaou", "fos cavaou", "Fos Cavaou"), (r"Le Havre", "le havre", "Le Havre"),
    (r"Wilhelmshaven LNG Terminal 1", "wilhelmshaven 1", "Wilhelmshaven 1"), (r"Brunsbuettel", "brunsbuettel fsru", "Brunsbüttel"),
    (r"Stade", "stade fsru", "Stade"), (r"Wilhelmshaven LNG Terminal 2", "wilhelmshaven 2", "Wilhelmshaven 2"),
    (r"Lubmin", "lubmin", "Lubmin 2"), (r"Mukran", "mukran", "Mukran"), (r"Alexandroupolis", "alexandroupolis", "Alexandroupolis"),
    (r"Revythoussa", "revithoussa", "Revithoussa"), (r"Panigaglia", "panigaglia", "Panigaglia"), (r"OLT", "olt", "OLT Offshore LNG"),
    (r"Piombino", "piombino", "Piombino"), (r"Ravenna", "ravenna", "Ravenna"), (r"Rovigo", "rovigo", "Rovigo"),
    (r"Klaip", "klaipeda", "Klaipėda"), (r"Eemshaven|EemsEnergy", "eemshaven", "EemsEnergyTerminal"), (r"Gate", "gate", "Gate terminal"),
    (r"winouj", "swinoujscie", "Świnoujście"), (r"Sines", "sines", "Sines"), (r"Bilbao", "bilbao", "Bilbao"), (r"Barcelona", "barcelona", "Barcelona"),
    (r"Huelva", "huelva", "Huelva"), (r"Cartagena", "cartagena", "Cartagena"), (r"El Musel", "el musel", "El Musel"), (r"Sagunto", "sagunto", "Sagunto"),
    (r"Mugardos", "mugardos", "Mugardos"), (r"Grain", "isle of grain", "Isle of Grain"), (r"South Hook", "south hook", "South Hook"),
]
GROUP = {"Turkey": "Turkey", "United Kingdom": "United Kingdom", "Albania": "Balkans", "Russia": "other"}
EU27 = {"Austria", "Belgium", "Bulgaria", "Croatia", "Cyprus", "Czechia", "Denmark", "Estonia", "Finland", "France", "Germany", "Greece", "Hungary",
        "Ireland", "Italy", "Latvia", "Lithuania", "Luxembourg", "Malta", "Netherlands", "Poland", "Portugal", "Romania", "Slovakia", "Slovenia",
        "Spain", "Sweden"}


def group_of(country):
    return "EU27" if country in EU27 else GROUP.get(country, "other Europe")


MAPKEY = {mn: k for rx, k, mn in ALSI_KEY}


def map_key(name, vessel, status):
    if name == "Brunsbüttel":
        return "brunsbuettel fsru" if vessel == "FSRU" else "brunsbuettel onshore"
    if name == "Stade":
        return "stade fsru" if vessel == "FSRU" else "stade onshore"
    if name == "Wilhelmshaven":
        return "wilhelmshaven onshore"
    return MAPKEY.get(name) or canon(name)


def build_terminals(alsi_f, alsi_now, cards, ver):
    """One row per terminal: ALSI facility (declared capacity, dates) joined to the GIE map card."""
    terms = {}
    for f in alsi_f:
        if f["end"] and f["end"] < "2025-01-01":
            continue
        if f["name"].startswith("PRIOR") or "virtual" in f["name"].lower():
            continue
        hit = next((k for rx, k, mn in ALSI_KEY if re.search(rx, f["name"], re.I)), None)
        if hit is None:
            hit = canon(f["name"])
            tried("ALSI facility without map-key alias", f["name"], f"keyed as '{hit}'")
        d = alsi_now.get(f["fac_eic"] + "|" + f["name"]) or {}
        dtrs = num(d.get("dtrs"))
        start = f["start"] if f["start"] and f["start"] < "2999" else None
        not_started = bool(f["start"] and f["start"] >= "2999")
        end = f["end"] if f["end"] and f["end"] < "2999" else None
        # Panigaglia has two ALSI entries (old and new reporting entity): keep the open one
        if hit in terms and terms[hit].get("end") is None and end is not None:
            continue
        terms[hit] = {"key": hit, "terminal": re.sub(r"\s*\(.*\)$", "", f["name"]).replace(" LNG Terminal", "").replace(" Terminal", ""),
                      "full_name": f["name"], "country": f["country"], "operator": f["company"], "type": {"FSRU": "FSRU", "LAND": "onshore regas"}.get(f["ftype"], f["ftype"]),
                      "alsi_start": start, "alsi_end": end, "alsi_dtrs_gwhd": dtrs, "alsi_sendout_gwhd": num(d.get("sendOut")),
                      "alsi_gas_day": d.get("gasDayStart"), "has_alsi": True, "not_started": not_started}
    if cards is not None and len(cards):
        for _, c in cards.iterrows():
            k = map_key(c["name"], c["vessel"], c["status"])
            t = terms.setdefault(k, {"key": k, "terminal": c["name"], "full_name": c["name"], "country": c["country"], "operator": None,
                                     "type": c["vessel"], "has_alsi": False, "alsi_start": None, "alsi_end": None,
                                     "alsi_dtrs_gwhd": np.nan, "alsi_sendout_gwhd": np.nan, "alsi_gas_day": None})
            if c["status"] in ("UNDER_CONSTRUCTION", "PLANNED", "SUSPENDED") and t["has_alsi"] and t.get("alsi_start") and c["status"] != "SUSPENDED":
                k2 = k + " (map " + c["status"].lower() + ")"
                terms[k2] = {"key": k2, "terminal": c["name"] + " (" + c["status"].lower().replace("_", " ") + ")", "full_name": c["name"], "country": c["country"],
                             "operator": None, "type": c["vessel"], "has_alsi": False, "alsi_start": None, "alsi_end": None,
                             "alsi_dtrs_gwhd": np.nan, "alsi_sendout_gwhd": np.nan, "alsi_gas_day": None,
                             "map_status": c["status"], "map_start": c["start_up_year"], "map_nominal": c["nominal_bcm_n"], "map_exp": c["exp_list"], "map_name": c["name"]}
                continue
            t.update(map_status=c["status"], map_start=c["start_up_year"], map_nominal=c["nominal_bcm_n"], map_exp=c["exp_list"], map_name=c["name"])
    out = []
    for k, t in terms.items():
        t.setdefault("map_status", None)
        t.setdefault("map_nominal", np.nan)
        t["group"] = group_of(t["country"])
        alsi_b = gwhd_to_bcfd(t["alsi_dtrs_gwhd"])
        map_b = bcmy_to_bcfd(t["map_nominal"])
        t["alsi_bcfd"], t["map_bcfd"] = alsi_b, map_b
        if alsi_b is not None and not pd.isna(alsi_b) and alsi_b > 0:
            t["ceiling_bcfd"], t["ceiling_basis"] = alsi_b, "ALSI declared total reference send-out (dtrs)"
        elif map_b is not None and not pd.isna(map_b):
            t["ceiling_bcfd"], t["ceiling_basis"] = map_b, f"GIE LNG map nominal annual capacity ({ver})"
        else:
            t["ceiling_bcfd"], t["ceiling_basis"] = np.nan, "none published"
        out.append(t)
    return out


def regas_rows(terms, cards, ver, ok, tyndp):
    R = []
    pubmap = ("Gas LNG Europe / GIE", f"GIE LNG Map ({ver})" if ver else "GIE LNG Map", "https://www.gie.eu/publications/maps/gie-lng-map/")

    def add(country, project, operator, ftype, value, unit, bcfd, status, online, basis, pub, doc, url, asof, evidence, notes, asset="regas_terminal"):
        R.append(row(country=country, asset_type=asset, project=project, operator=operator, fuel_or_type=ftype, capacity_value=value,
                     capacity_unit=unit, capacity_bcfd=bcfd, status=status, expected_online=online, online_date_basis=basis, source_publisher=pub,
                     source_document=doc, source_url=url, as_of_date=asof, evidence=evidence, notes=notes))
    tk = {t["key"]: t for t in terms}
    asof_map = ver or ""

    # ---- GIE map cards: new terminals, expansions, suspended
    if cards is not None and len(cards):
        for _, c in cards.iterrows():
            nm, ctry, st = c["name"], c["country"], c["status"]
            exps = c["exp_list"] if isinstance(c["exp_list"], list) else []
            kk = map_key(nm, c["vessel"], st)
            if st == "SUSPENDED" and tk.get(kk, {}).get("alsi_end"):
                continue   # covered by the ALSI end-date row below
            if st in ("UNDER_CONSTRUCTION", "PLANNED", "SUSPENDED") and c["nominal_bcm_n"] and not pd.isna(c["nominal_bcm_n"]):
                yr = c["start_up_year"]
                nt = ""
                if st == "SUSPENDED":
                    nt = "Shown as suspended on the GIE map."
                add(ctry, f"{nm} ({c['vessel']}) - new terminal", None, c["vessel"], c["nominal_bcm_n"], "bcm(N)/y", bcmy_to_bcfd(c["nominal_bcm_n"]), st,
                    str(int(yr)) if yr and not pd.isna(yr) else None, "SOURCED_YEAR" if yr and not pd.isna(yr) else "NOT_STATED", *pubmap[:2], pubmap[2],
                    asof_map, "SOURCED", f"Nominal annual capacity as printed on the map (normal m3); map column '{st}'. {nt} bcm->Bcf/d x35.3147/365.25 (DERIVED, normal m3 not corrected).")
            if st == "OPERATIONAL":
                for e in exps:
                    if e.get("info"):
                        continue
                    yr = e.get("year")
                    add(ctry, f"{nm}: expansion ({e['label']})", None, c["vessel"], e["bcm"], "bcm(N)/y", bcmy_to_bcfd(e["bcm"]), "PLANNED",
                        str(yr) if yr else None, "SOURCED_YEAR" if yr else "NOT_STATED", *pubmap[:2], pubmap[2], asof_map, "SOURCED",
                        "Expansion line on the operating terminal's map card; the map does not label its FID status, so PLANNED unless an operator document says otherwise. "
                        + ("'by YYYY' is counted from January of that year in the monthly totals." if yr else "No year printed: not counted in the monthly totals."))
    # ---- ALSI end dates / recent starts
    for t in terms:
        if not t["has_alsi"]:
            continue
        if t["alsi_start"] and t["alsi_start"] >= "2025-01-01" and t["key"] != "panigaglia":   # Panigaglia: new ALSI reporting entity, not a commissioning
            add(t["country"], f"{t['terminal']} ({t['type']}) - started", t["operator"], t["type"], t["alsi_dtrs_gwhd"], "GWh/d (ALSI dtrs)",
                t["alsi_bcfd"], "COMMISSIONED_RECENT", t["alsi_start"][:7], "SOURCED_MONTH", "GIE ALSI", "ALSI facility start date and declared total reference send-out",
                "https://alsi.gie.eu/api/about", TODAY_S, "SOURCED", "Start date from the ALSI facility list; capacity = current dtrs.")
        if t["alsi_end"] and t["alsi_end"] >= "2025-01-01":
            doc = ok("total_lehavre", "demob") and "le havre" in t["key"]
            pub, title, url = mkpub("total_lehavre") if doc else ("GIE ALSI", "ALSI facility end date", "https://alsi.gie.eu/api/about")
            add(t["country"], f"{t['terminal']} ({t['type']}) - ended", t["operator"], t["type"],
                -(t["map_nominal"] if not pd.isna(t.get("map_nominal", np.nan)) else 0) or None, "bcm(N)/y" if not pd.isna(t.get("map_nominal", np.nan)) else None,
                -(t["map_bcfd"] or 0) or None, "COMMISSIONED_RECENT", t["alsi_end"][:7], "SOURCED_MONTH", pub, title, url, TODAY_S, "SOURCED",
                "REMOVAL: ALSI facility end date" + ("; TotalEnergies press release 25 Nov 2025 confirms the FSRU was demobilised (Rouen court decision 16 Oct 2025)." if doc else ".")
                + " Capacity (negative) = GIE map nominal.")
    # ---- operator-document rows
    p = lambda d: mkpub(d)  # noqa: E731
    if ok("det_diary", "nov") and ok("det_diary", "cap32"):
        pub, title, url = p("det_diary")
        t = tk.get("stade fsru")
        add("Germany", "Stade LNG Terminal 1 (FSRU Energos Force) - first LNG cargo", "Deutsche Energy Terminal", "FSRU", 3.2, "bcm/y (2027 expected feed-in)", bcmy_to_bcfd(3.2),
            "UNDER_CONSTRUCTION", "2026-11", "SOURCED_MONTH", pub, title, url, TODAY_S, "SOURCED",
            "DET: Energos Force docked 15 Sep 2026; first LNG tanker expected in November 2026; up to 3.2 bcm feed-in expected for 2027; feed-in capacity 1.7 bcm to end-2026 "
            "(Stade page). GIE map nominal 4.7 bcm(N)/y.")
    if ok("gaz_gdansk", "early2028"):
        pub, title, url = p("gaz_gdansk")
        add("Poland", "Gdansk FSRU Terminal (FSRU Solidarity)", "GAZ-SYSTEM", "FSRU", 6.1 if ok("gaz_fsru", "cap61") else None, "bcm/y",
            bcmy_to_bcfd(6.1) if ok("gaz_fsru", "cap61") else None, "UNDER_CONSTRUCTION", "2028", "SOURCED_YEAR", pub, title, url, TODAY_S, "SOURCED",
            "Commercial regasification 'scheduled to commence in early 2028' (Aug 2026); design capacity 6.1 bcm/y (Dec 2025 release). GIE map still shows 2027.")
    if ok("gaz_fsru2", "cap45"):
        pub, title, url = p("gaz_fsru2")
        add("Poland", "Gdansk FSRU 2 (second vessel)", "GAZ-SYSTEM", "FSRU", 4.5, "bcm/y", bcmy_to_bcfd(4.5), "PLANNED", None, "NOT_STATED", pub, title, url, TODAY_S, "SOURCED",
            "Binding open season; GAZ-SYSTEM proceeds only if offers cover at least 50% of slots. No date stated.")
    if ok("gasunie_ee", "to2036") and ok("gasunie_ee", "cfid"):
        pub, title, url = p("gasunie_ee")
        add("Netherlands", "EemsEnergyTerminal: two FSRU charters end Sep 2027, extension to 2036 (conditional FID)", "Gasunie / Vopak", "FSRU", 8.0, "bcm/y", bcmy_to_bcfd(8.0),
            "FID", "2027-09", "SOURCED_MONTH", pub, title, url, TODAY_S, "SOURCED",
            "REMOVAL RISK that has been mitigated: original contracts run to the end of Sep 2027; conditional investment decision (June 2026) keeps the terminal in operation "
            "until 2036 (contracts 2028-2036 signed for most of 8 bcm/y, about 2.4 bcm/y still open). Capacity kept in the monthly totals.")
    if ok("lng_hr", "exp"):
        pub, title, url = p("lng_hr")
        add("Croatia", "Krk LNG (FSRU LNG Croatia): regasification module added", "LNG Hrvatska", "FSRU", 2.2, "bcm/y (from 3.9 to 6.1)", bcmy_to_bcfd(2.2),
            "COMMISSIONED_RECENT", "2025-11", "SOURCED_MONTH", pub, title, url, "2026-01-01", "DERIVED",
            "Operator: maximum capacity increased from 3.9 to 6.1 bcm/y after the FSRU returned from the yard (commissioning completed 7 Nov 2025 per the operator's news page). "
            "Increment = 6.1 - 3.9 (DERIVED). The GIE map (Aug 2026) still prints 3.5 bcm(N) + 2.6 by 2029.")
    if ok("snam_piombino", "reloc"):
        pub, title, url = p("snam_piombino")
        t = tk.get("piombino")
        add("Italy", "Piombino FSRU (Italis LNG): shutdown / relocation window", "Snam", "FSRU", -5.0, "bcm/y", bcmy_to_bcfd(-5.0), "PLANNED", "2026-10", "SOURCED_MONTH", pub, title, url,
            "2025-06-11", "SOURCED",
            "REMOVAL RISK: Snam's regasification code says shutdown, relocation and recommissioning (to Vado Ligure) last at most twelve months from 1 Oct 2026 or the 36th month after "
            "start-up, whichever is earlier. Relocation authorisation was unresolved in 2025 (ARERA resolution 88/2025 called the move inefficient); the GIE map (Aug 2026) still lists "
            "Piombino as operating, so it stays in the monthly totals. Ravenna FSRU BW Singapore (5 bcm/y) started June 2025.")
    if ok("gastrade_cap", "step"):
        pub, title, url = p("gastrade_cap")
        add("Greece", "Alexandroupolis FSRU: maximum regasification capacity raised", "Gastrade", "FSRU", 136.2, "GWh/d (up from 90.8 on 12 Sep 2025)", gwhd_to_bcfd(136.2),
            "COMMISSIONED_RECENT", "2025-10", "SOURCED_MONTH", pub, title, url, "2025-10-21", "SOURCED", "Operator announcement 21 Oct 2025; terminal nominal 5.5 bcm/y (project page).")
    if ok("gasgrid_inkoo", "cap") and ok("gasgrid_lease", "ten"):
        pub, title, url = p("gasgrid_inkoo")
        add("Finland", "Inkoo FSRU (Exemplar): lease term", "Gasgrid Finland", "FSRU", 140, "GWh/d", gwhd_to_bcfd(140), "COMMISSIONED_RECENT", None, "NOT_STATED", pub, title, url, TODAY_S, "SOURCED",
            "Information row: regasification capacity 140 GWh/d (over 40 TWh/y); the Exemplar was leased on 20 May 2022 for ten years (use from Jan 2023), so the charter ends about end-2032 (DERIVED, "
            "exact end date not stated). Not a change before 2030.")
    # ---- TYNDP LNG projects not on the map
    if tyndp is not None:
        main, ci, lng, ugs = tyndp
        have = " ".join(norm(r["project"]) for r in R)
        for _, r in lng.iterrows():
            m = main[main["code"] == r["code"]]
            if m.empty:
                continue
            m = m.iloc[0]
            fac = str(r["facility"])
            skip = any(w in norm(fac) for w in ("alexandroupolis", "ravenna")) or pd.isna(r["bcm_y"]) or float(r["bcm_y"]) == 0
            yr = pd.to_numeric(m["year_first"], errors="coerce")
            if skip or pd.isna(yr) or yr < 2025:
                continue
            # already covered by a GIE map line for the same terminal and year?
            kmap = {"gate": "gate", "gdansk": "gdansk", "thrace": "thrace", "krk": "krk", "winoujscie": "swinoujscie"}
            if any(k in norm(fac) for k in ("gdansk", "thrace")) and any(k in have for k in ("gdansk", "thrace")):
                continue
            st = {"FID": "FID", "Advanced": "PLANNED", "Less-Advanced": "TARGET"}.get(r["maturity"], "PLANNED")
            add(m["country"], f"{r['facility']} - {m['name']} [{r['code']}]", m["promoter"], "regas expansion/terminal", float(r["bcm_y"]), "bcm/y", bcmy_to_bcfd(float(r["bcm_y"])), st,
                str(int(yr)), "SOURCED_YEAR", "ENTSOG", "TYNDP 2024 Annex A - List of projects (LNG sheet; July 2024 edition)", TYNDP_PAGE, "2024-07-03", "SOURCED",
                f"TYNDP maturity '{r['maturity']}'; phase '{r['phase']}'. Edition is two years old: dates may have moved (compare the GIE map lines above).")
    # operator documents supersede the map line for the same project
    if any(r["project"].startswith("Stade LNG Terminal 1") for r in R):
        R = [r for r in R if not r["project"].startswith("Stade (FSRU)")]
    if any(r["project"].startswith("Gdansk FSRU Terminal") for r in R):
        R = [r for r in R if not (r["project"].startswith("Gdańsk (FSRU) - new terminal") and r["status"] == "UNDER_CONSTRUCTION")]
    if any(r["project"].startswith("Krk LNG") for r in R):
        R = [r for r in R if not r["project"].startswith("Krk Island: expansion")]
    return R


def pipeline_rows(tyndp, ok, path):
    R = []
    if tyndp is None:
        return R
    main, ci, lng, ugs = tyndp
    mm = main.set_index("code")
    c = ci.copy()
    c = c[c["year"] >= 2025]
    c = c[c["cap_gwhd"].notna() & (c["cap_gwhd"] != 0)]
    c["typ"] = c["code"].str.split("-").str[0]
    sel = c[(c["typ"].isin(["TRA", "RET"])) | ((c["cap_gwhd"] < 0) & c["typ"].isin(["H2T", "H2S", "H2L", "BIO", "OTH", "SYN", "CO2"]))]
    pub = "ENTSOG"
    doc = "TYNDP 2024 Annex A - List of projects (Capacity Increments sheet; July 2024 edition)"
    for _, r in sel.iterrows():
        mrow = mm.loc[r["code"]] if r["code"] in mm.index else None
        if isinstance(mrow, pd.DataFrame):
            mrow = mrow.iloc[0]
        ctry = str(mrow["country"]).split(",")[0].strip() if mrow is not None else ""
        st = {"FID": "FID", "Advanced": "PLANNED", "Less-Advanced": "TARGET"}.get(str(r["maturity"]), "PLANNED")
        neg = r["cap_gwhd"] < 0
        h2 = r["typ"].startswith("H2")
        note = ("GAS CAPACITY REMOVED (TYNDP 'natural gas decrement': hydrogen repurposing, assumes long-term gas capacity contracts are fulfilled). " if neg and h2 else
                "Capacity decrement at the point. " if neg else "") + f"TYNDP maturity '{r['maturity']}'. Edition July 2024: check dates against operator plans."
        R.append(row(country=ctry, asset_type="gas_pipeline",
                     project=f"{r['name']} [{r['code']}] @ {r['point']}: {str(r['from_system']).strip()} -> {str(r['to_system']).strip()}",
                     operator=r["operator"] if isinstance(r["operator"], str) else r["promoter"], fuel_or_type="pipeline",
                     capacity_value=float(r["cap_gwhd"]), capacity_unit="GWh/d", capacity_bcfd=gwhd_to_bcfd(r["cap_gwhd"]), status=st,
                     expected_online=str(int(r["year"])), online_date_basis="SOURCED_YEAR", source_publisher=pub, source_document=doc,
                     source_url=TYNDP_PAGE, as_of_date="2024-07-03", evidence="SOURCED", notes=note))
    # TAP first expansion level: done (operator)
    if ok("tap_first", "fullop") and ok("tap_first", "bcm"):
        p1 = mkpub("tap_first")
        R.append(row(country="Greece", asset_type="gas_pipeline", project="TAP first level of capacity expansion (1.2 bcm/y) - full operability reconfirmed", operator="Trans Adriatic Pipeline",
                     fuel_or_type="pipeline", capacity_value=1.2, capacity_unit="bcm/y", capacity_bcfd=bcmy_to_bcfd(1.2), status="COMMISSIONED_RECENT", expected_online="2026-01",
                     online_date_basis="SOURCED_YEAR", source_publisher=p1[0], source_document=p1[1], source_url=p1[2], as_of_date=TODAY_S, evidence="SOURCED",
                     notes="TAP: 1.2 bcm/y of incremental capacity allocated 'starting from 2026' (30 Jan 2023 release); the news list reconfirms full operability of the first level. "
                           "Aim to double capacity by 2027 depends on further market tests (TYNDP row TRA-A-810: 205 GWh/d in 2029)."))
    # Bulgartransgaz / Transgaz Vertical Corridor
    if ok("btg", "negru"):
        p1 = mkpub("btg")
        R.append(row(country="Bulgaria", asset_type="gas_pipeline", project="Negru Voda 1 / Kardam (BG -> RO) capacity increase (Vertical Corridor)", operator="Bulgartransgaz / Transgaz",
                     fuel_or_type="pipeline", capacity_value=137.2, capacity_unit="GWh/d", capacity_bcfd=gwhd_to_bcfd(137.2), status="PLANNED", expected_online="2026-09",
                     online_date_basis="SOURCED_QUARTER", source_publisher=p1[0], source_document=p1[1], source_url=p1[2], as_of_date="2026-04-27", evidence="SOURCED",
                     notes="Bulgartransgaz 10-year plan 2026-2035 (public consultation, 27 Apr 2026): project adds 137.2 GWh/d Bulgaria -> Romania in Q3 2026 (Q3 counted as Sep 2026); Transgaz declared technical readiness."))
    if ok("btg", "kulata"):
        p1 = mkpub("btg")
        R.append(row(country="Bulgaria", asset_type="gas_pipeline", project="Kulata / Sidirokastro (GR -> BG) capacity increase (Bulgarian side +35.4 GWh/d, total 102 GWh/d)", operator="Bulgartransgaz",
                     fuel_or_type="pipeline", capacity_value=35.4, capacity_unit="GWh/d", capacity_bcfd=gwhd_to_bcfd(35.4), status="PLANNED", expected_online="2026-06",
                     online_date_basis="SOURCED_QUARTER", source_publisher=p1[0], source_document=p1[1], source_url=p1[2], as_of_date="2026-04-27", evidence="SOURCED",
                     notes="Bulgartransgaz 10-year plan 2026-2035: commissioning expected Q2 2026 (counted as Jun 2026); status after Q2 not confirmed in the document."))
    return R


def storage_rows(tyndp, stor_tab, ver, agsi_f, ok):
    R = []
    pubm = ("Gas Infrastructure Europe", f"GIE Storage Map ({ver})" if ver else "GIE Storage Map", "https://www.gie.eu/publications/maps/gie-storage-map/")
    if stor_tab is not None and len(stor_tab):
        for _, s in stor_tab.iterrows():
            if s["country"] == "EU27":
                continue
            for col, st, lab in (("under_construction_twh", "UNDER_CONSTRUCTION", "under construction"), ("planned_twh", "PLANNED", "planned")):
                v = s.get(col)
                if v is None or pd.isna(v) or v == 0:
                    continue
                R.append(row(country=s["country"], asset_type="gas_storage", project=f"{s['country']}: new storage facilities {lab} (GIE map country total)", operator=None,
                             fuel_or_type="underground storage", capacity_value=float(v), capacity_unit="TWh working gas", capacity_bcfd=None, status=st, expected_online=None,
                             online_date_basis="NOT_STATED", source_publisher=pubm[0], source_document=pubm[1], source_url=pubm[2], as_of_date=ver, evidence="SOURCED",
                             notes=f"Country total of facilities in this status on the map (new facilities only, not extensions); no dates on the map; "
                                   f"{twh_to_bcm(v):.2f} bcm at 11.2 GWh/mcm (DERIVED)."))
            v = s.get("mothballed_twh")
            if v is not None and not pd.isna(v) and v > 0:
                R.append(row(country=s["country"], asset_type="gas_storage", project=f"{s['country']}: mothballed storage (GIE map country total)", operator=None, fuel_or_type="underground storage",
                             capacity_value=float(v), capacity_unit="TWh working gas", status="SUSPENDED", online_date_basis="NOT_STATED", source_publisher=pubm[0], source_document=pubm[1],
                             source_url=pubm[2], as_of_date=ver, evidence="SOURCED", notes="Mothballed on the map: not available to the market; shown as SUSPENDED."))
    if tyndp is not None:
        main, ci, lng, ugs = tyndp
        mm = main.set_index("code")
        g = ugs.copy()
        for c_ in ("wgv_gwh", "wd_gwhd", "inj_gwhd"):
            g[c_] = pd.to_numeric(g.get(c_), errors="coerce")
        for code, d in g.groupby("code"):
            if code not in mm.index:
                continue
            m = mm.loc[code]
            if isinstance(m, pd.DataFrame):
                m = m.iloc[0]
            yr = pd.to_numeric(m["year_first"], errors="coerce")
            if pd.isna(yr) or yr < 2025:
                continue
            wgv, wd = d["wgv_gwh"].sum(min_count=1), d["wd_gwhd"].max()
            st = {"FID": "FID", "Advanced": "PLANNED", "Less-Advanced": "TARGET"}.get(m["maturity"], "PLANNED")
            R.append(row(country=m["country"], asset_type="gas_storage", project=f"{m['name']} [{code}]", operator=m["promoter"], fuel_or_type="underground storage",
                         capacity_value=(wgv / 1000.0 if not pd.isna(wgv) else None), capacity_unit="TWh working gas (added)",
                         capacity_bcfd=gwhd_to_bcfd(wd) if not pd.isna(wd) else None, status=st, expected_online=str(int(yr)),
                         online_date_basis="SOURCED_YEAR", source_publisher="ENTSOG", source_document="TYNDP 2024 Annex A - List of projects (UGS sheet; July 2024 edition)",
                         source_url=TYNDP_PAGE, as_of_date="2024-07-03", evidence="SOURCED",
                         notes=f"TYNDP maturity '{m['maturity']}'. capacity_bcfd = maximum withdrawal capacity {wd if not pd.isna(wd) else 'n/a'} GWh/d (DERIVED conversion); "
                               f"working gas volume in GWh summed over the project's rows."))
    if ok("btg", "depo"):
        p1 = mkpub("btg")
        R.append(row(country="Romania", asset_type="gas_storage", project="Depomures UGS expansion, phase 1 (to 400 mcm; phase 2 to 600 mcm)", operator="Depogaz / Romgaz", fuel_or_type="underground storage",
                     capacity_value=400, capacity_unit="mcm working gas (after phase 1)", capacity_bcfd=None, status="FID", expected_online="2026", online_date_basis="SOURCED_YEAR",
                     source_publisher=p1[0], source_document=p1[1], source_url=p1[2], as_of_date="2026-04-27", evidence="SOURCED",
                     notes="Bulgartransgaz plan, Romania section: first phase to 400 mcm, estimated commissioning 2026; second phase to 600 mcm (no date). 400 mcm = 4.5 TWh at 11.2 GWh/mcm (DERIVED)."))
    # closures: AGSI end dates with 'decommissioned'
    for f in agsi_f:
        if f["end"] and "2024-01-01" <= f["end"] < "2999" and "decommissioned" in f["name"].lower():
            R.append(row(country=f["country"], asset_type="gas_storage", project=f"{f['name']} - closed", operator=f["company"], fuel_or_type="underground storage",
                         status="COMMISSIONED_RECENT", expected_online=f["end"][:7], online_date_basis="SOURCED_MONTH", source_publisher="GIE AGSI+", source_document="AGSI+ facility list (operational_end_date)",
                         source_url="https://agsi.gie.eu/api/about", as_of_date=TODAY_S, evidence="SOURCED", notes="CLOSURE: AGSI+ marks the facility decommissioned from this date; no volume given in the facility list."))
    # Rough (derived from this repo's National Gas site data)
    try:
        d = pd.read_excel(os.path.join(os.path.dirname(os.path.abspath(path_out())), "gb_storage_sites_daily.xlsx"), sheet_name="Daily")
        d["date"] = pd.to_datetime(d["date"])
        s = d.set_index("date")["Rough_stock"].astype(float)
        inflow = d.set_index("date")["Rough_inflow"].astype(float)
        last_stock = s[s > 0].index.max()
        last_in = inflow[inflow > 0].index.max()
        peak = s.max()
        R.append(row(country="United Kingdom", asset_type="gas_storage", project="Rough (Centrica) - no injection, zero stock", operator="Centrica Energy Storage", fuel_or_type="underground storage",
                     capacity_value=-round(float(peak) / 1000, 2), capacity_unit="TWh (peak stock in the stored daily series, GWh/1000)", status="COMMISSIONED_RECENT",
                     expected_online=last_stock.strftime("%Y-%m"), online_date_basis="SOURCED_MONTH", source_publisher="National Gas Transmission (data portal, via this repo's gb_storage_sites_daily.xlsx)",
                     source_document="Daily storage stock / inflow by site (Rough_stock, Rough_inflow)", source_url="https://data.nationalgas.com/", as_of_date=TODAY_S, evidence="DERIVED",
                     notes=f"CLOSURE in effect, not an announcement: Rough stock last above zero on {last_stock.date()}, last injection {last_in.date()}, peak stock in the series {peak:.0f} GWh; "
                           "stock is zero in the latest data. No official closure statement was found (Centrica talks of redevelopment under a cap-and-floor model), so treat as out of service, not as retired."))
    except Exception as e:  # noqa: BLE001
        tried("Rough closure from gb_storage_sites_daily.xlsx", "output/Data and Chart Outputs/gb_storage_sites_daily.xlsx", f"{type(e).__name__}")
    return R


_OUT = [None]


def path_out():
    return _OUT[0] or os.path.join(OUT_DEFAULT, OUT_NAME)


def upstream_rows(ok, a):
    R = []
    if ok("omv_neptun", "h1"):
        p1 = mkpub("omv_neptun")
        R.append(row(country="Romania", asset_type="gas_field", project="Neptun Deep (Black Sea) - first gas to the NTS", operator="OMV Petrom / Romgaz", fuel_or_type="gas field",
                     capacity_value=8.0, capacity_unit="bcm/y (plateau, 100%)", capacity_bcfd=bcmy_to_bcfd(8.0), status="UNDER_CONSTRUCTION", expected_online="2027",
                     online_date_basis="SOURCED_YEAR", source_publisher=p1[0], source_document=p1[1], source_url=p1[2], as_of_date="2026-09-28", evidence="SOURCED",
                     notes="First gas into the National Transmission System expected in H1 2027; plateau expected by end of Q3 2027 at around 8 bcm/y (100%), about 100 bcm recoverable; "
                           "50% OMV Petrom (operator), 50% Romgaz. 'gas_field' = upstream supply addition (see Sources)."))
    # Norwegian fields approved for production (Sodir fact pages)
    for did, nm, hc in (("np_irpa", "Irpa (tie-in to Aasta Hansteen, gas)", "gas"), ("np_fenris", "Fenris (with Valhall)", "oil/gas"), ("np_berling", "Berling (tie-in to Asgard B, gas)", "gas")):
        if ok(did, "start"):
            p1 = mkpub(did)
            yr = re.search(r"start in (\d{4})", DOCS[did][3]["start"]).group(1)
            R.append(row(country="Norway", asset_type="gas_field", project=f"{nm} - production start", operator=None, fuel_or_type="gas field", capacity_value=None,
                         capacity_unit="n/a (no plateau rate published on the field page)", capacity_bcfd=None, status="UNDER_CONSTRUCTION", expected_online=yr, online_date_basis="SOURCED_YEAR",
                         source_publisher=p1[0], source_document=p1[1], source_url=p1[2], as_of_date=TODAY_S, evidence="SOURCED",
                         notes=f"Status 'Approved for production'; 'production is planned to start in {yr}'. Hydrocarbon type {hc}. Recoverable volumes are on the recent-activity page (million Sm3 o.e.), not a rate."))
    if ok("np_hugin", "status"):
        p1 = mkpub("np_hugin")
        R.append(row(country="Norway", asset_type="gas_field", project="Yggdrasil area (Hugin, Munin, Fulla) - approved for production", operator="Aker BP", fuel_or_type="oil/gas fields",
                     capacity_value=None, capacity_unit="n/a", status="UNDER_CONSTRUCTION", expected_online=None, online_date_basis="NOT_STATED", source_publisher=p1[0], source_document=p1[1],
                     source_url=p1[2], as_of_date=TODAY_S, evidence="SOURCED", notes="PDO approved 2023; field pages give no start-up year; Sodir lists Yggdrasil as an ongoing coordinated development at end-2025. Mainly oil with associated/gas export via Statpipe-Statfjord/Frigg route (not quantified)."))
    return R


def raw_data_only(proj):
    """Owner rule (Oct 2026): raw data only - physical projects / facilities with a sourced capacity AND a sourced date.
    Drops undated or capacity-less rows. TARGET rows (TYNDP 'Less-Advanced') stay, labelled, and are never summed with project rows; every drop is listed on 'Tried'."""
    ok_ = (proj["capacity_value"].notna() | (proj["asset_type"] == "gas_field")) & proj["expected_online"].notna()
    for _, r in proj[~ok_].iterrows():
        why = ("no sourced date" if pd.isna(r["expected_online"]) else "no sourced capacity")
        tried("row excluded (raw data only)", r["source_url"], f"{r['country']}: {str(r['project'])[:90]} - {why}")
    return proj[ok_].reset_index(drop=True)


# ------------------------------------------------------------------ history / first seen
def pkey(df):
    return (df["asset_type"].astype(str) + "|" + df["country"].astype(str) + "|" + df["project"].astype(str))


def merge_first_seen(cur, prev):
    if prev is None or not len(prev):
        return cur
    p = prev.copy()
    p["_k"] = pkey(p)
    fs = p.drop_duplicates("_k").set_index("_k")["first_seen"].to_dict()
    cur = cur.copy()
    k = pkey(cur)
    cur["first_seen"] = [fs.get(x, TODAY_S) for x in k]
    return cur


def build_history(cur, prev, prev_hist):
    cols = ["date_seen", "asset_type", "country", "project", "field", "old", "new"]
    h = prev_hist.copy() if prev_hist is not None and len(prev_hist) else pd.DataFrame(columns=cols)
    if prev is None or not len(prev):
        return h
    p = prev.copy()
    p["_k"] = pkey(p)
    c = cur.copy()
    c["_k"] = pkey(c)
    pi = p.drop_duplicates("_k").set_index("_k")
    new = []
    for _, r in c.iterrows():
        if r["_k"] in pi.index:
            o = pi.loc[r["_k"]]
            for f in ("expected_online", "status", "capacity_value"):
                ov, nv = o[f], r[f]
                if (pd.isna(ov) and pd.isna(nv)) or str(ov) == str(nv):
                    continue
                new.append({"date_seen": TODAY_S, "asset_type": r["asset_type"], "country": r["country"], "project": r["project"], "field": f, "old": ov, "new": nv})
        else:
            new.append({"date_seen": TODAY_S, "asset_type": r["asset_type"], "country": r["country"], "project": r["project"], "field": "new entry", "old": None, "new": r["status"]})
    for k in set(pi.index) - set(c["_k"]):
        o = pi.loc[k]
        new.append({"date_seen": TODAY_S, "asset_type": o["asset_type"], "country": o["country"], "project": o["project"], "field": "dropped", "old": o["status"], "new": None})
    return pd.concat([h, pd.DataFrame(new, columns=cols)], ignore_index=True) if new else h


# ------------------------------------------------------------------ regas timeline
def regas_timeline(terms, proj):
    """Per-terminal capacity (Bcf/d) by month, committed vs incl. planned, plus totals by group.
    Capacity(m) = base + sum of dated changes up to month m. Base = current ceiling less recent step-ups already realised."""
    ev_committed = {"UNDER_CONSTRUCTION", "FID", "COMMISSIONED_RECENT"}
    ev_planned = {"PLANNED", "PERMITTED"}
    rows = []
    cols = {}
    P = proj[proj["asset_type"] == "regas_terminal"].copy()
    for t in terms:
        if t.get("map_status") in ("UNDER_CONSTRUCTION", "PLANNED") and not t["has_alsi"] or "(map " in t["key"]:
            continue
    # build monthly series per (terminal key, scenario)
    ser = {}
    def addser(lbl, scen, group, country, start, end, bcfd):
        if bcfd is None or pd.isna(bcfd) or start is None:
            return
        s = pd.Series(0.0, index=MONTHS)
        s[(MONTHS >= month_start(start)) & ((MONTHS < month_start(end)) if end is not None else True)] = bcfd
        ser.setdefault((lbl, country, group), {})[scen] = ser.get((lbl, country, group), {}).get(scen, 0) + s
    START = pd.Timestamp("2025-01-01")
    for t in terms:
        k = t["key"]
        base = t["ceiling_bcfd"]
        if pd.isna(base):
            continue
        status = t.get("map_status")
        if "(map " in k:   # a second map card (new onshore unit at an existing site)
            yr = t.get("map_start")
            if yr and not pd.isna(yr) and status in ("UNDER_CONSTRUCTION", "PLANNED"):
                st = pd.Timestamp(f"{int(yr)}-01-01")
                addser(t["terminal"], "committed" if status == "UNDER_CONSTRUCTION" else "planned", t["group"], t["country"], st, None, base)
            continue
        if t["has_alsi"]:
            s0 = pd.Timestamp(t["alsi_start"]) if t["alsi_start"] else None
            start = START if (s0 is None or s0 < START) else s0
            end = pd.Timestamp(t["alsi_end"]) if t["alsi_end"] else None
            if end is not None and end < START:
                continue
            if t.get("not_started") or (s0 is not None and s0 > pd.Timestamp(TODAY_S) and not (status in ("OPERATIONAL",))):
                # ALSI lists it, not started: use the map date
                yr = t.get("map_start")
                st = pd.Timestamp(f"{int(yr)}-01-01") if yr and not pd.isna(yr) else None
                override = proj[(proj["project"].str.contains("Stade LNG Terminal 1", na=False)) & (proj["expected_online"].notna())] if "stade" in k else pd.DataFrame()
                if len(override):
                    st = pd.Timestamp(str(override.iloc[0]["expected_online"]) + ("-01" if len(str(override.iloc[0]["expected_online"])) == 7 else "-01-01"))
                addser(t["terminal"], "committed", t["group"], t["country"], st, None, base)
                continue
            for sc in ("committed", "planned"):
                addser(t["terminal"], sc, t["group"], t["country"], start, end, base)
        else:
            yr = t.get("map_start")
            if status == "OPERATIONAL":
                for sc in ("committed", "planned"):
                    addser(t["terminal"], sc, t["group"], t["country"], START, None, base)
            elif status in ("UNDER_CONSTRUCTION", "PLANNED") and yr and not pd.isna(yr):
                st = pd.Timestamp(f"{int(yr)}-01-01")
                if t.get("map_name") == "Gdańsk" and proj["project"].str.startswith("Gdansk FSRU Terminal").any():
                    st = pd.Timestamp("2028-01-01")
                addser(t["terminal"], "committed" if status == "UNDER_CONSTRUCTION" else "planned", t["group"], t["country"], st, None, base)
    # expansions / step changes from Projects
    for _, r in P.iterrows():
        v = r["capacity_bcfd"]
        if v is None or pd.isna(v) or r["expected_online"] is None or pd.isna(r["expected_online"]):
            continue
        proj_name = str(r["project"])
        if "charters end" in proj_name or "shutdown / relocation" in proj_name or "new terminal" in proj_name or "started" in proj_name or "ended" in proj_name or "first LNG cargo" in proj_name or "Gdansk FSRU Terminal" in proj_name:
            continue
        if "expansion (" in proj_name or "extension" in proj_name.lower() or "regasification module" in proj_name or "[LNG-" in proj_name:
            eo = str(r["expected_online"])
            st = pd.Timestamp(eo + "-01") if len(eo) == 7 else pd.Timestamp(eo + "-01-01")
            if st < START:
                continue
            if r["status"] in ev_committed:
                scen = "committed"
            elif r["status"] in ev_planned:
                scen = "planned"
            else:
                continue
            ctry = r["country"]
            if r["status"] == "COMMISSIONED_RECENT":
                continue   # realised steps are inside the current ceiling; kept for the base adjustment below
            addser(proj_name[:70], scen, group_of(ctry), ctry, st, None, v)
    # realised recent steps: base before the step is lower (Krk Nov 2025, Alexandroupolis Oct 2025)
    for _, r in P[P["status"] == "COMMISSIONED_RECENT"].iterrows():
        nm = str(r["project"])
        if ("Krk" in nm or "Alexandroupolis" in nm) and r["expected_online"]:
            eo = str(r["expected_online"])
            st = pd.Timestamp(eo + "-01")
            pre = r["capacity_bcfd"] if "Krk" in nm else (gwhd_to_bcfd(136.2 - 90.8))
            if st > START:
                for key_, d in ser.items():
                    if (("Krk" in nm and "krk" in norm(key_[0])) or ("Alexandroupolis" in nm and "alexandroupolis" in norm(key_[0]))):
                        for sc in d:
                            s = d[sc].copy()
                            s[(MONTHS < st) & (s > 0)] = (s[(MONTHS < st)] - pre).clip(lower=0)[s[(MONTHS < st)] > 0]
                            d[sc] = s
    # assemble
    tab = []
    tot = {sc: {} for sc in ("committed", "planned")}
    for (lbl, ctry, grp), d in ser.items():
        rec = {"terminal / project": lbl, "country": ctry, "group": grp}
        for sc in ("committed", "planned"):
            s = d.get(sc)
            if s is None:
                s = d.get("committed", pd.Series(0.0, index=MONTHS)) if sc == "planned" else pd.Series(0.0, index=MONTHS)
            rec[f"_{sc}"] = s
        tab.append(rec)
    # planned scenario = committed + planned rows: make planned series cumulative
    for rec in tab:
        rec["_planned"] = rec["_planned"] if rec["_planned"].sum() else rec["_committed"]
    months_lbl = [m.strftime("%Y-%m") for m in MONTHS]
    detail = []
    for rec in sorted(tab, key=lambda r: (r["group"], r["country"], r["terminal / project"])):
        row_ = {"terminal / project": rec["terminal / project"], "country": rec["country"], "group": rec["group"], "scenario": "committed (operating + under construction + FID)"}
        row_.update(dict(zip(months_lbl, rec["_committed"].round(4).tolist())))
        detail.append(row_)
    for rec in sorted(tab, key=lambda r: (r["group"], r["country"], r["terminal / project"])):
        s_pl = rec["_planned"]
        row_ = {"terminal / project": rec["terminal / project"], "country": rec["country"], "group": rec["group"], "scenario": "incl. dated planned"}
        # planned-only rows add on top of committed rows: the planned scenario of a row that has no committed part is the row itself
        tot_pl = s_pl if (rec["_committed"].sum() == 0 or (s_pl != rec["_committed"]).any()) else rec["_committed"]
        row_.update(dict(zip(months_lbl, tot_pl.round(4).tolist())))
        detail.append(row_)
    det = pd.DataFrame(detail, columns=["terminal / project", "country", "group", "scenario"] + months_lbl)
    # totals
    mcols = months_lbl
    sc_c = det[det["scenario"].str.startswith("committed")]
    sc_p = det[det["scenario"].str.startswith("incl.")]
    month = pd.DataFrame(index=MONTHS)
    for grp in ("EU27", "United Kingdom", "Turkey", "Balkans", "other Europe"):
        month[f"{grp} committed Bcf/d"] = sc_c[sc_c["group"] == grp][mcols].sum().values
        month[f"{grp} incl. planned Bcf/d"] = sc_p[sc_p["group"] == grp][mcols].sum().values
    month["Europe committed Bcf/d"] = sc_c[mcols].sum().values
    month["Europe incl. planned Bcf/d"] = sc_p[mcols].sum().values
    month["EU27 committed bcm/y"] = month["EU27 committed Bcf/d"] * 365.25 / 35.3147
    month["EU27 incl. planned bcm/y"] = month["EU27 incl. planned Bcf/d"] * 365.25 / 35.3147
    month["Europe committed bcm/y"] = month["Europe committed Bcf/d"] * 365.25 / 35.3147
    month["Europe incl. planned bcm/y"] = month["Europe incl. planned Bcf/d"] * 365.25 / 35.3147
    month.index.name = "month"
    # per-terminal table
    tt = []
    for t in terms:
        if "(map " in t["key"]:
            continue
        tt.append({"terminal": t["terminal"], "country": t["country"], "group": t["group"], "type": t["type"], "operator": t.get("operator"),
                   "status_on_map": t.get("map_status"), "alsi_start": t.get("alsi_start"), "alsi_end": t.get("alsi_end"),
                   "alsi_dtrs_gwhd": t["alsi_dtrs_gwhd"], "alsi_dtrs_bcfd": t["alsi_bcfd"], "alsi_dtrs_bcmy": (t["alsi_bcfd"] * 365.25 / 35.3147 if t["alsi_bcfd"] and not pd.isna(t["alsi_bcfd"]) else np.nan),
                   "alsi_latest_sendout_gwhd": t["alsi_sendout_gwhd"], "alsi_gas_day": t["alsi_gas_day"],
                   "gie_map_nominal_bcm_n": t.get("map_nominal"), "gie_map_nominal_bcfd": t["map_bcfd"],
                   "ceiling_bcfd": t["ceiling_bcfd"], "ceiling_bcmy": (t["ceiling_bcfd"] * 365.25 / 35.3147 if not pd.isna(t["ceiling_bcfd"]) else np.nan), "ceiling_basis": t["ceiling_basis"]})
    ttab = pd.DataFrame(tt).sort_values(["group", "country", "terminal"])
    ttab["alsi_vs_map_ratio"] = (ttab["alsi_dtrs_bcfd"] / ttab["gie_map_nominal_bcfd"]).round(2)
    # monthly detail goes under the terminal table in the same sheet, separated by a blank row
    return ttab.reset_index(drop=True), det, month


def storage_sheet(agsi_f, agsi_now, stor_tab):
    rows = []
    for f in agsi_f:
        d = agsi_now.get(f["fac_eic"] + "|" + f["name"])
        if not d:
            continue
        rows.append({"country": f["country"], "facility": f["name"], "company": f["company"], "type": f["ftype"], "start": f["start"], "end": f["end"],
                     "working_gas_volume_twh": num(d.get("workingGasVolume")), "injection_capacity_gwhd": num(d.get("injectionCapacity")),
                     "withdrawal_capacity_gwhd": num(d.get("withdrawalCapacity")), "gas_in_storage_twh": num(d.get("gasInStorage")), "full_pct": num(d.get("full")),
                     "gas_day": d.get("gasDayStart")})
    return pd.DataFrame(rows).sort_values(["country", "facility"]) if rows else pd.DataFrame(columns=["country", "facility"])


def summary(proj, terms, month):
    s = proj.groupby(["asset_type", "status"]).agg(rows=("project", "count"), bcfd_sum=("capacity_bcfd", "sum")).reset_index()
    c = proj.groupby(["country", "asset_type", "status"]).agg(rows=("project", "count"), bcfd_sum=("capacity_bcfd", "sum")).reset_index()
    s.insert(0, "level", "asset_type x status")
    c.insert(0, "level", "country x asset_type x status")
    t = []
    for g in ("EU27", "United Kingdom", "Turkey", "Balkans", "other Europe", "Europe"):
        cc = f"{g} committed Bcf/d"
        pc = f"{g} incl. planned Bcf/d"
        if cc in month.columns:
            t.append({"level": "regas capacity Bcf/d", "asset_type": g, "status": "Jan 2025 / Dec 2026 / Dec 2030 committed", "rows": None,
                      "bcfd_sum": f"{month[cc].iloc[0]:.2f} / {month[cc].loc['2026-12-01']:.2f} / {month[cc].iloc[-1]:.2f}"})
            t.append({"level": "regas capacity Bcf/d", "asset_type": g, "status": "Jan 2025 / Dec 2026 / Dec 2030 incl. planned", "rows": None,
                      "bcfd_sum": f"{month[pc].iloc[0]:.2f} / {month[pc].loc['2026-12-01']:.2f} / {month[pc].iloc[-1]:.2f}"})
    return pd.concat([s, c, pd.DataFrame(t)], ignore_index=True)


def sources_table(lng_ver, stor_ver):
    rows = [
        ("GIE ALSI API (terminal list, declared total reference send-out dtrs, send-out)", "Gas Infrastructure Europe", "https://alsi.gie.eu/api", "current technical capacity per terminal; key GIE_API_KEY"),
        ("GIE AGSI+ API (facility working gas, injection / withdrawal capacity, operating dates)", "Gas Infrastructure Europe", "https://agsi.gie.eu/api", "current storage per facility; closures from end dates"),
        (f"GIE LNG Map (PDF, version {lng_ver}) - terminal cards", "Gas LNG Europe / GIE", "https://www.gie.eu/publications/maps/gie-lng-map/", "nominal annual capacity (bcm normal m3), start-up year, status by map column, expansion lines"),
        (f"GIE Storage Map (PDF, version {stor_ver}) - country summary table", "Gas Infrastructure Europe", "https://www.gie.eu/publications/maps/gie-storage-map/", "working gas TWh by status per country"),
        ("ENTSOG TYNDP 2024 Annex A - List of projects (xlsx, July 2024)", "ENTSOG", TYNDP_PAGE, "project maturity, commissioning year, capacity increments / decrements at interconnection points (GWh/d), LNG and UGS project volumes"),
    ]
    for did, (u, pub, title, _) in DOCS.items():
        rows.append((title, pub, u, "curated fact, written only when the opened text contains the checked wording (sheet 'Doc checks')"))
    rows.append(("gb_storage_sites_daily.xlsx (this repo, National Gas data portal)", "National Gas Transmission", "https://data.nationalgas.com/", "Rough stock / injections (closure evidence, DERIVED)"))
    return pd.DataFrame(rows, columns=["source", "publisher", "url", "used_for"])


# ------------------------------------------------------------------ workbook + charts
NOTES = [
    "Europe FUTURE gas infrastructure (EUROPE_GAS_INFRA.py). Shared row schema on 'Projects'; one row per dated change.",
    "",
    "Units",
    "Gas assets: capacity_value + capacity_unit as published (bcm(N)/y, GWh/d, TWh working gas); capacity_bcfd = Bcf/d. bcm/y x 35.3147 / 365.25; GWh/d / 11.2 GWh per mcm x 0.0353147; "
    "GIE map bcm are normal m3 and are not corrected to standard m3. capacity_mw is blank (gas assets).",
    "asset_type: regas_terminal | gas_pipeline | gas_storage | gas_field (NEW: upstream supply addition, e.g. Neptun Deep, Norwegian fields; capacity_value is a plateau rate when published, else blank).",
    "Negative capacity_value = capacity REMOVED (ended FSRU, closed storage, gas capacity lost to hydrogen repurposing at an interconnection point).",
    "status: GIE map column (operating / under construction / planned / suspended); TYNDP maturity FID -> FID, Advanced -> PLANNED, Less-Advanced -> TARGET; COMMISSIONED_RECENT also marks closures after 1 Jan 2025.",
    "",
    "Monthly totals ('Regas monthly', 'Regas capacity')",
    "Ceiling per terminal = ALSI declared total reference send-out (dtrs) where published, else the GIE map nominal capacity. A year without a month counts from January of that year; "
    "'committed' = operating + under construction + FID, 'incl. planned' adds dated PLANNED / PERMITTED rows; undated rows, TARGET, SUSPENDED and CANCELLED are not counted. "
    "A start counts in its month, an end is removed in its month. Capacity before the current ceiling was reached is rebuilt only for the dated recent steps (Krk Nov 2025, Alexandroupolis Oct 2025).",
    "ALSI send-out history in the chart is the sum of the EU countries' monthly mean of daily ALSI send-out (GWh/d -> Bcf/d); GB stopped reporting to ALSI on 31 Dec 2020, Turkey never did.",
    "",
    "Checks",
    "'Doc checks': each curated fact needs its checked wording found in the freshly opened document; otherwise it is not written (and 'Tried' says so). 'Releases': Last-Modified / ETag / length / sha256 "
    "of the heavy files (GIE map PDFs, TYNDP workbook); they are re-read only when that signature changes. The TYNDP 2024 edition is July 2024: dates may be stale, newer operator documents are listed separately.",
]
NOTE_TITLES = {"Units", "Monthly totals ('Regas monthly', 'Regas capacity')", "Checks"}


def write_all(path, sheets, reg_month, proj, stor_sheet, a):
    out = {}
    for k, v in sheets.items():
        out[k] = v.copy()
    # monthly detail rows go on 'Regas capacity' after the per-terminal table: keep two sheets
    rc = out["Regas capacity"]
    out["Regas capacity"] = rc
    out["Regas monthly"] = pd.concat([reg_month.reset_index().assign(month=lambda d: d["month"].dt.strftime("%Y-%m"))], axis=0)
    xlsx_notes.write_workbook(path, {k: v.reset_index(drop=True) for k, v in out.items()}, NOTES, NOTE_TITLES)
    # drop pandas' index column on every sheet
    from openpyxl import load_workbook
    wb = load_workbook(path)
    for ws in wb.worksheets:
        if ws.title == "Units":
            continue
        ws.delete_cols(1)
        ws.freeze_panes = "A2"
        for col in ws.columns:
            w = max(len(str(c.value)) if c.value is not None else 0 for c in list(col)[:60])
            ws.column_dimensions[col[0].column_letter].width = min(max(10, w + 2), 60)
    if not a.no_charts:
        add_all_charts(wb, reg_month, proj, stor_sheet, a)
    xlsx_charts.save_atomic(wb, path)
    print("wrote", path)


def alsi_history(alsi_path):
    try:
        d = pd.read_excel(alsi_path, sheet_name="Daily")
    except Exception:  # noqa: BLE001
        return None
    d["date"] = pd.to_datetime(d["date"])
    cols = [c for c in d.columns if c.endswith("_sendout_GWhd") and not c.startswith(("EU_", "GB_"))]
    d = d.set_index("date")[cols].sum(axis=1, min_count=1)
    m = d.resample("MS").mean()
    # the month in progress is partial: keep it (mean of days) but drop if fewer than 15 days
    n = d.resample("MS").count()
    m = m[n >= 15]
    return m / GWH_PER_MCM * BCF_PER_MCM


def add_all_charts(wb, reg_month, proj, stor_sheet, a):
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    pngdir = a.png_dir
    os.makedirs(pngdir, exist_ok=True)
    alsi = alsi_history(os.path.join(a.out_dir, "eu_lng_terminals_daily.xlsx"))
    # chart 1: regas capacity vs actual send-out
    df = pd.DataFrame(index=MONTHS)
    df["EU27 capacity: operating + under construction / FID"] = reg_month["EU27 committed Bcf/d"]
    df["EU27 capacity incl. dated planned"] = reg_month["EU27 incl. planned Bcf/d"]
    if alsi is not None:
        df["EU27 ALSI send-out (monthly mean)"] = alsi.reindex(MONTHS)
    t1 = "European LNG regasification capacity to 2030 vs ALSI send-out (EU27)"
    xlsx_charts.add_chart_sheet(None, df, t1, "Bcf/d", kind="line", sheet_name="Chart - Regas capacity", wb=wb)
    df2 = pd.DataFrame(index=MONTHS)
    for g, lab in (("EU27", "EU27"), ("United Kingdom", "United Kingdom"), ("Turkey", "Turkey"), ("Balkans", "Balkans (Albania)"), ("other Europe", "other (Cyprus etc.)")):
        c = f"{g} committed Bcf/d"
        if c in reg_month.columns and reg_month[c].abs().sum() > 0:
            df2[lab] = reg_month[c]
    xlsx_charts.add_chart_sheet(None, df2, "European regasification capacity by country group, operating + under construction / FID", "Bcf/d", kind="stacked_bar", sheet_name="Chart - Regas by group", wb=wb)
    # PNG 1
    fig, ax = plt.subplots(figsize=(11, 5.5))
    ax.fill_between(df.index, df.iloc[:, 0], df.iloc[:, 1], color="#EB6834", alpha=0.18, label="dated planned additions")
    ax.plot(df.index, df.iloc[:, 0], color="#2A78D6", lw=2.2, label=df.columns[0])
    ax.plot(df.index, df.iloc[:, 1], color="#EB6834", lw=1.6, ls="--", label=df.columns[1])
    if alsi is not None:
        ax.plot(df.index, df.iloc[:, 2], color="#1BAF7A", lw=2, label=df.columns[2])
    ax.set_ylabel("Bcf/d"); ax.set_title(t1, fontsize=12, fontweight="bold"); ax.grid(alpha=.25); ax.set_ylim(bottom=0)
    ax.legend(loc="lower right", frameon=False, fontsize=9)
    for s in ("top", "right"):
        ax.spines[s].set_visible(False)
    import matplotlib.dates as mdates
    ax.xaxis.set_major_formatter(mdates.DateFormatter("%b/%y")); fig.autofmt_xdate()
    fig.text(0.01, 0.01, "Source: GIE ALSI (declared capacity, send-out), GIE LNG map, operator documents. EU27 only; UK and Turkey on the second chart.", fontsize=8, color="#555")
    fig.tight_layout(rect=(0, 0.03, 1, 1)); fig.savefig(os.path.join(pngdir, "europe_future_gas_infrastructure__regas_capacity.png"), dpi=130); plt.close(fig)
    # PNG 1b
    fig, ax = plt.subplots(figsize=(11, 5.5))
    bottom = np.zeros(len(df2))
    pal = ["#2A78D6", "#EB6834", "#1BAF7A", "#EDA100", "#E87BA4"]
    for i, c in enumerate(df2.columns):
        ax.bar(df2.index, df2[c], width=25, bottom=bottom, color=pal[i % 5], label=c); bottom += df2[c].fillna(0).values
    ax.set_ylabel("Bcf/d"); ax.set_title("European regasification capacity by country group (operating + under construction / FID)", fontsize=12, fontweight="bold")
    ax.legend(frameon=False, ncol=3, loc="upper left", fontsize=9); ax.grid(alpha=.25, axis="y")
    for s in ("top", "right"):
        ax.spines[s].set_visible(False)
    ax.xaxis.set_major_formatter(mdates.DateFormatter("%b/%y")); fig.autofmt_xdate()
    fig.tight_layout(); fig.savefig(os.path.join(pngdir, "europe_future_gas_infrastructure__regas_by_group.png"), dpi=130); plt.close(fig)
    # chart 2: storage by country and status (TWh) from the GIE map table
    stor = None
    try:
        stor = pd.read_excel(os.path.join(a.out_dir, OUT_NAME), sheet_name="GIE storage table") if False else None
    except Exception:  # noqa: BLE001
        stor = None
    ws_src = wb["GIE storage table"] if "GIE storage table" in wb.sheetnames else None
    if ws_src is not None and ws_src.max_row > 1:
        hdr = [c.value for c in ws_src[1]]
        data = pd.DataFrame([[c.value for c in r] for r in ws_src.iter_rows(min_row=2)], columns=hdr)
        data = data[~data["country"].isin(["EU27", "Ukraine"])]
        data = data[["country", "operational_twh", "under_construction_twh", "planned_twh", "mothballed_twh"]].fillna(0)
        data = data.sort_values("operational_twh", ascending=False)
        add_category_chart(wb, data.set_index("country"), "Underground gas storage working gas by country and status (GIE storage map)", "TWh", "Chart - Storage", stacked=True,
                           note="Source: GIE Storage Map. Planned and under construction are new facilities only, no dates on the map.")
        fig, ax = plt.subplots(figsize=(11, 5.5))
        b = np.zeros(len(data))
        for col, colr, lab in (("operational_twh", "#2A78D6", "operational"), ("under_construction_twh", "#1BAF7A", "under construction"), ("planned_twh", "#EDA100", "planned"), ("mothballed_twh", "#9A9A9A", "mothballed")):
            ax.bar(data["country"], data[col], bottom=b, color=colr, label=lab); b += data[col].values
        ax.set_ylabel("TWh working gas"); ax.set_title("Underground gas storage working gas by country and status", fontsize=12, fontweight="bold")
        plt.setp(ax.get_xticklabels(), rotation=60, ha="right"); ax.legend(frameon=False); ax.grid(alpha=.25, axis="y")
        for s in ("top", "right"):
            ax.spines[s].set_visible(False)
        fig.text(0.01, 0.01, "Source: GIE Storage Map (planned / under construction = new facilities, undated).", fontsize=8, color="#555")
        fig.tight_layout(rect=(0, 0.03, 1, 1)); fig.savefig(os.path.join(pngdir, "europe_future_gas_infrastructure__storage.png"), dpi=130); plt.close(fig)
    # chart 3: pipeline capacity additions by commissioning year and maturity (GWh/d -> Bcf/d)
    pl = proj[(proj["asset_type"] == "gas_pipeline") & proj["capacity_bcfd"].notna() & proj["expected_online"].notna()].copy()
    if len(pl):
        pl["year"] = pl["expected_online"].astype(str).str[:4].astype(int)
        pl = pl[(pl["year"] >= 2025) & (pl["year"] <= 2035) & (pl["status"] != "TARGET")]   # TARGET rows are never stacked with projects
        pl["kind"] = np.where(pl["capacity_bcfd"] < 0, "removed (incl. H2 repurposing)", pl["status"].map({"FID": "FID", "PLANNED": "planned", "TARGET": "target", "COMMISSIONED_RECENT": "commissioned"}).fillna("planned"))
        pv = pl.pivot_table(index="year", columns="kind", values="capacity_bcfd", aggfunc="sum").fillna(0)
        order = [c for c in ("commissioned", "FID", "planned", "target", "removed (incl. H2 repurposing)") if c in pv.columns]
        pv = pv[order]
        pvd = pv.copy()
        pvd.index = pd.to_datetime(pvd.index.astype(str) + "-01-01")
        xlsx_charts.add_chart_sheet(None, pvd, "Pipeline / interconnector capacity added and removed at border points by commissioning year (Bcf/d, directional sum)", "Bcf/d",
                                    kind="stacked_bar", sheet_name="Chart - Pipelines", date_format="%Y", wb=wb)
        fig, ax = plt.subplots(figsize=(11, 5.5))
        pos = np.zeros(len(pv)); neg = np.zeros(len(pv))
        cols_ = {"commissioned": "#1BAF7A", "FID": "#2A78D6", "planned": "#EDA100", "target": "#E87BA4", "removed (incl. H2 repurposing)": "#E34948"}
        for c in order:
            v = pv[c].values
            if c.startswith("removed"):
                ax.bar(pv.index, v, bottom=neg, color=cols_[c], label=c); neg += v
            else:
                ax.bar(pv.index, v, bottom=pos, color=cols_[c], label=c); pos += v
        ax.axhline(0, color="#333", lw=.8)
        ax.set_ylabel("Bcf/d (sum of directional IP entries)"); ax.set_title("Gas interconnector capacity added / removed by commissioning year (TYNDP 2024 + operator documents)", fontsize=12, fontweight="bold")
        ax.legend(frameon=False, ncol=3, fontsize=9); ax.grid(alpha=.25, axis="y")
        for s in ("top", "right"):
            ax.spines[s].set_visible(False)
        fig.text(0.01, 0.01, "Source: ENTSOG TYNDP 2024 Annex A (July 2024; a bidirectional upgrade counts in both directions), Bulgartransgaz, TAP. Not netted against the same-point reverse flows.", fontsize=7.5, color="#555")
        fig.tight_layout(rect=(0, 0.03, 1, 1)); fig.savefig(os.path.join(pngdir, "europe_future_gas_infrastructure__pipelines.png"), dpi=130); plt.close(fig)


def add_category_chart(wb, df, title, y_title, sheet_name, stacked=True, note=""):
    from openpyxl.chart import BarChart, Reference
    from openpyxl.chart.shapes import GraphicalProperties
    from openpyxl.drawing.line import LineProperties
    if sheet_name in wb.sheetnames:
        del wb[sheet_name]
    ws = wb.create_sheet(sheet_name, 1)
    ws.cell(row=1, column=1, value="Country")
    for j, c in enumerate(df.columns):
        ws.cell(row=1, column=2 + j, value=str(c).replace("_twh", "").replace("_", " "))
    for i, (idx, r) in enumerate(df.iterrows(), start=2):
        ws.cell(row=i, column=1, value=str(idx))
        for j, v in enumerate(r.values):
            ws.cell(row=i, column=2 + j, value=float(v))
    ch = BarChart()
    ch.type = "col"
    ch.grouping = "stacked" if stacked else "clustered"
    ch.overlap = 100 if stacked else None
    ch.add_data(Reference(ws, min_col=2, max_col=1 + df.shape[1], min_row=1, max_row=1 + len(df)), titles_from_data=True)
    ch.set_categories(Reference(ws, min_col=1, min_row=2, max_row=1 + len(df)))
    cols = ["2A78D6", "1BAF7A", "EDA100", "9A9A9A"]
    for i, s in enumerate(ch.series):
        s.graphicalProperties = GraphicalProperties(solidFill=cols[i % 4], ln=LineProperties(noFill=True))
    ch.title = title
    ch.y_axis.title = y_title
    ch.x_axis.delete = False
    ch.y_axis.delete = False
    xlsx_charts.rotated_labels(ch.x_axis, -45)
    ch.legend.position = "b"
    ch.graphical_properties = GraphicalProperties(ln=LineProperties(noFill=True))
    ch.plot_area.graphicalProperties = GraphicalProperties(ln=LineProperties(noFill=True))
    ch.width, ch.height = 28, 13
    xlsx_charts.tidy_layout(ch)
    ws.add_chart(ch, "H2")
    if note:
        ws.cell(row=len(df) + 4, column=1, value=note)


if __name__ == "__main__":
    main()
