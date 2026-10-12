"""
Egypt natural gas balance: MONTHLY from JODI Gas (labelled fallback) plus an ANNUAL official layer (CAPMAS, MOEE).

OWNER'S DECISIONS (Oct 2026)
  1. JODI monthly is used and labelled "JODI (fallback)" on every series, sheet and chart: no official Egyptian source
     publishes monthly gas statistics that GitHub Actions can reach (EGAS, EGPC and the Ministry of Petroleum publish no
     tables; probes discovery_archive/middle_east/EGYPT_GAS_HISTORY_PROBE.py, EGYPT_ISRAEL_GAS_POWER_PROBE.py).
  2. The official annual layer is CAPMAS (JSON API) and the Ministry of Electricity (MOEE statistics pages).
  3. No Comtrade / customs layer.
  4. LNG imports and exports are IMPLIED from the balance, not taken from JODI's reported LNG rows:
        Implied_net_LNG_imports = gross inland deliveries (observed) + pipeline exports - production - pipeline imports
     positive = net LNG imports, negative = net LNG exports (the owner checks it against Kpler). JODI's own LNG rows are
     kept on the raw sheet for reference only (not charted, not used in the balance; 'Checks' shows the difference).
     Stock change and receipts from other sources enter the identity only if JODI reports a non-zero value for them
     (JODI: GID calculated = production + receipts + imports - exports - stock change, sign tested on the whole JODI
     table). For Egypt both are 0 to Dec 2023 and not reported ('-') from Jan 2024, so neither is in the identity.

SOURCES
  JODI Gas World database, Beyond 20/20 table viewer at jodidb.org (ReportId 38673; method of
      discovery_archive/rest_of_world/JODI_LNG_LAG_PULL4.py): Country = Egypt, Product = natural gas, every BALANCE
      row, unit Million Standard Cubic Metres, monthly from Jan 2011 (Egypt's first month), ~2.5 months lag. The static
      GAS_world_NewFormat.zip is stale (Jun 2025) and is not used. JODI data are the country's own submission (here
      Egypt's), compiled by JODI; they are not an Egyptian official publication, hence 'fallback'.
  CAPMAS (Central Agency for Public Mobilization and Statistics) JSON API
      https://www.capmas.gov.eg:8080/api/Indicator/IndicatorFilter?IndicatorId=<id>&SubSubjectId=30, units from
      api/Indicator/IndicatorDetails (measureUnit, as stated; 'unit not stated' where the API gives none; never
      converted). Energy balance (publication 57): 3112 primary production, 3102 gross domestic consumption, 3108 final
      consumption, 3121 imports, 3124 exports of natural gas (Million cubic meter), 3155 final energy consumption by
      sector (1000 tons of oil equivalent, all fuels). Electricity and energy statistics (publication 77): 9543 natural
      gas sold by use, 9545 extracted crude oil / condensates / natural gas / LPG, 5937 petroleum-sector export
      quantities, 5939 import quantities (all 'Quantity in 1000 tons' per the API - an indicator-wide unit). The API
      gives yearTypeId=2 and does not say whether its years are calendar or fiscal (Jul-Jun) years.
  MOEE (Ministry of Electricity and Renewable Energy) http://www.moee.gov.eg/english_new/ST_fuel.aspx: fuel used in
      power stations by Egyptian FISCAL year (Jul-Jun; the page prints '2023/2022' for FY2022/23), natural gas in
      million m3 (header 'N.G. Million m3'), HFO etc.; latest FY2022/23.

INCREMENTAL: the committed workbook is the history store. Each run re-fetches JODI from 12 months before the last
stored month (revisions) plus any gap, and keeps older stored months; a value JODI now shows as '-' does not erase a
stored one. CAPMAS and MOEE are fetched every run (a few small requests) but parsed and replaced only when the response
hash differs from the one on the 'Releases' sheet.

Schedule: 1st and 15th (.github/workflows/egypt_gas.yml). Charts: add_charts.py REGISTRY 'egypt_gas_monthly.xlsx'.
Usage: python3 middle_east/EGYPT_GAS.py [--out <xlsx>] [--full]
"""

import argparse
import csv
import hashlib
import html
import io
import json
import os
import re
import sys
import time
from datetime import datetime, timezone

import pandas as pd
import requests

REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, REPO_ROOT)
import xlsx_notes  # noqa: E402

DEFAULT_OUT = os.path.join(REPO_ROOT, "output", "Data and Chart Outputs", "egypt_gas_monthly.xlsx")
BCF_PER_MCM = 0.0353147
UA = {"User-Agent": "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36",
      "Accept-Language": "en-US,en;q=0.9"}
JODI = "http://www.jodidb.org"
JODI_REPORT = "38673"
JODI_T0 = pd.Timestamp("2009-01-01")      # TIME handle 0 of the viewer
FIRST_MONTH = pd.Timestamp("2011-01-01")  # Egypt's first JODI month
REVISION_MONTHS = 12
CAPMAS = "https://www.capmas.gov.eg:8080/api/"
CAPMAS_IDS = [  # (indicatorId, publication id)
    (3112, 57), (3102, 57), (3108, 57), (3121, 57), (3124, 57), (3155, 57),
    (9543, 77), (9545, 77), (5937, 77), (5939, 77)]
MOEE_FUEL = "http://www.moee.gov.eg/english_new/ST_fuel.aspx"

S_MONTHLY = "Monthly JODI (fallback)"
S_RAW = "JODI (fallback) all rows"
S_ANNUAL = "Annual official"
S_CHECKS = "Checks"
S_CHECKS_FY = "Checks FY"
S_REL = "Releases"

# JODI BALANCE rows in viewer order -> column stem. Sub-rows ('   LNG', '   Pipeline') follow their parent.
BAL = {
    "Production": "Production",
    "Receipts from Other Sources": "Receipts_other_sources",
    "Imports": "Imports_total",
    "Imports/LNG": "Imports_LNG_reported",
    "Imports/Pipeline": "Imports_pipeline",
    "Exports": "Exports_total",
    "Exports/LNG": "Exports_LNG_reported",
    "Exports/Pipeline": "Exports_pipeline",
    "Stock Change": "Stock_change",
    "Gross Inland Deliveries (Calculated)": "GID_calculated",
    "Statistical Difference (Calculated)": "Statistical_difference",
    "Gross Inland Deliveries (Observed)": "GID_observed",
    "of which: Electricity and Heat Generation": "Electricity_and_heat",
    "Closing Stocks": "Closing_stocks",
}
RAW_COLS = [f"{v}_JODI_fallback_mcm" for v in BAL.values()]


def log(msg):
    print(msg, file=sys.stderr, flush=True)


def col(stem):
    return f"{stem}_JODI_fallback_mcm"


# ------------------------------------------------------------------ JODI viewer

class Viewer:
    def __init__(self):
        self.s = requests.Session()
        self.s.headers.update(UA)
        self.ref = {"Referer": f"{JODI}/TableViewer/tableView.aspx?ReportId={JODI_REPORT}"}
        r = self.req("GET", f"{JODI}/TableViewer/tableView.aspx?ReportId={JODI_REPORT}")
        self.base = self.inputs(r.text)
        self.rv0 = self.base["sWD_ReportView"]
        counts = [int(x) for x in self.base.get("sWD_SelectedItemsCount", "").split(",") if x.strip().isdigit()]
        self.n_time = counts[-1] if len(counts) == 5 else None   # order Country, Product, BALANCE, Unit, TIME
        log(f"JODI viewer: item counts {counts}")

    def req(self, method, url, **kw):
        for i in range(5):
            try:
                r = self.s.request(method, url, timeout=300, **kw)
                r.raise_for_status()
                return r
            except requests.RequestException as e:
                log(f"  retry {i + 1} {url[:70]} {type(e).__name__}")
                time.sleep(5 * (i + 1))
        r = self.s.request(method, url, timeout=300, **kw)
        r.raise_for_status()
        return r

    @staticmethod
    def inputs(t):
        out = {}
        for m in re.finditer(r"<(input|select|textarea)\b((?:[^>\"]|\"[^\"]*\")*)>", t, flags=re.I | re.S):
            a = dict(re.findall(r'(\w+)\s*=\s*"([^"]*)"', m.group(2)))
            if "name" in a:
                out[a["name"]] = html.unescape(a.get("value", ""))
        return out

    def build_rv(self, sel):
        rv = self.rv0
        for name, handles in sel.items():
            if handles is None:
                continue
            d = '<Definition><Items type="handles"><String value="%s"/></Items></Definition>' % ",".join(map(str, handles))
            rv, n = re.subn(r'(<Dim name="%s">.*?)<Definition>.*?</Definition>' % re.escape(name),
                            lambda m: m.group(1) + d, rv, count=1, flags=re.S)
            if n != 1:
                raise RuntimeError(f"JODI viewer: dimension {name} not in the report view")
        return rv

    def download(self, sel):
        data = dict(self.base)
        data["sWD_ReportView"] = self.build_rv(sel)
        data["WD_Command"] = ""
        f = self.inputs(self.req("POST", f"{JODI}/TableViewer/tableView.aspx", data=data, headers=self.ref).text)
        if "sWD_ReportView" not in f:
            raise RuntimeError("JODI viewer: selection not applied")
        f.update({"WD_DownloadFormat": "CSV", "sCS_SpawnWindow": "True", "WD_PageBeforeSummary": "/TableViewer/tableView.aspx"})
        pt = self.req("POST", f"{JODI}/TableViewer/downloadPrompt.aspx", data=f, headers=self.ref).text
        urls = re.findall(r"OnDownload\(\s*['\"]([^'\"]+)['\"]\s*\)", pt)
        ivt = re.findall(r"OnDownloadIvt\(\s*['\"]([^'\"]*)['\"]\s*,\s*['\"]([^'\"]*)['\"]\s*,\s*['\"]([^'\"]*)['\"]\s*\)", pt)
        hdr = {"Referer": f"{JODI}/TableViewer/downloadPrompt.aspx"}
        if urls:
            u = urls[0]
            u = u if u.startswith("http") else JODI + (u if u.startswith("/") else "/TableViewer/" + u)
            r = self.req("GET", u, headers=hdr)
        elif ivt:
            mime, name, path = ivt[0]
            pf = self.inputs(pt)
            pf.update({"WD_Command": "Download", "WD_DownloadMimeType": mime, "WD_DownloadFilename": name,
                       "WD_DownloadFilePath": path})
            r = self.req("POST", f"{JODI}/TableViewer/download.aspx", data=pf, headers=hdr)
        else:
            raise RuntimeError("JODI viewer: download prompt gave no link")
        body = r.content.decode("utf-8-sig", "replace")
        if body.lstrip()[:1] == "<":
            raise RuntimeError("JODI viewer: download is not CSV")
        return body

    def egypt_handle(self):
        """Country handles follow the viewer's country list order (handle 0 = Albania); checked by label."""
        listing = self.download({"Country": None, "Product": [0], "BALANCE": [0], "Unit": [0], "TIME": [0]})
        names = [r[0] for r in csv.reader(io.StringIO(listing)) if len(r) > 2 and r[0] not in ("", "Country")]
        guess = next((i for i, n in enumerate(names) if n.startswith("Egypt")), None)
        if guess is None:
            raise RuntimeError("JODI viewer: Egypt not in the country list")
        for h in [guess] + [guess + d for d in (-1, 1, -2, 2, -3, 3)]:
            body = self.download({"Country": [h], "Product": [0], "BALANCE": [0], "Unit": [0], "TIME": [0]})
            rows = [r for r in csv.reader(io.StringIO(body)) if len(r) > 1 and r[0] not in ("", "Country")]
            if rows and rows[0][0].startswith("Egypt"):
                return h, rows[0][0]
        raise RuntimeError("JODI viewer: no handle near the list position returns Egypt")


def parse_jodi(body):
    """Viewer CSV (BALANCE rows x TIME columns, one unit) -> wide frame of the RAW_COLS, monthly index."""
    rows = list(csv.reader(io.StringIO(body)))
    months = next(r for r in rows if len(r) > 1 and r[1] == "TIME")[2:]
    units = next((r for r in rows if len(r) > 1 and r[1] == "Unit"), [None, None, ""])[2:]
    if any(u and u != "Million Standard Cubic Metres" for u in units):
        raise RuntimeError(f"JODI: unexpected unit {set(units)}")
    parent, out = None, {}
    for r in rows:
        if len(r) < 3 or r[0] in ("", "Country") or not r[0].startswith("Egypt"):
            continue
        label = r[1].strip()
        if label in ("Imports", "Exports"):
            parent = label
        key = f"{parent}/{label}" if label in ("LNG", "Pipeline") else label
        if key not in BAL:
            raise RuntimeError(f"JODI: unknown balance row {r[1]!r}")
        vals = {}
        for m, v in zip(months, r[2:]):
            try:
                vals[pd.Timestamp(datetime.strptime(m, "%b%Y"))] = float(v)
            except ValueError:
                vals[pd.Timestamp(datetime.strptime(m, "%b%Y"))] = float("nan")
        out[col(BAL[key])] = pd.Series(vals)
    df = pd.DataFrame(out)
    df.index.name = "date"
    return df.reindex(columns=RAW_COLS)


def fetch_jodi(stored, full=False):
    v = Viewer()
    handle, label = v.egypt_handle()
    log(f"JODI: Egypt = country handle {handle} ({label})")
    n_time = v.n_time or 0
    last_handle = n_time - 1
    start = FIRST_MONTH
    if not full and stored is not None and len(stored):
        have = stored.dropna(how="all")
        if len(have):
            all_m = pd.date_range(FIRST_MONTH, have.index.max(), freq="MS")
            gaps = [m for m in all_m if m not in have.index]
            start = min([have.index.max() - pd.DateOffset(months=REVISION_MONTHS)] + gaps)
    h0 = (start.year - JODI_T0.year) * 12 + start.month - 1
    handles = list(range(h0, last_handle + 1)) if n_time else None
    log(f"JODI: requesting {start:%b %Y} .. handle {last_handle} ({len(handles) if handles else 'all'} months)")
    body = v.download({"Country": [handle], "Product": [0], "BALANCE": None, "Unit": [0], "TIME": handles})
    new = parse_jodi(body)
    return new, f"Egypt country handle {handle}, TIME handles {h0}-{last_handle}"


def merge(stored, new):
    if stored is None or not len(stored):
        return new
    stored = stored.reindex(columns=RAW_COLS)
    out = stored.reindex(stored.index.union(new.index))
    for c in RAW_COLS:
        nv = new[c].dropna()
        lost = stored[c].dropna().index.intersection(new.index).difference(nv.index)
        if len(lost):
            log(f"  {c}: JODI now shows '-' for {len(lost)} stored month(s); stored values kept")
        out.loc[nv.index, c] = nv
    return out.sort_index()


# ------------------------------------------------------------------ balance

def monthly_balance(raw):
    d = raw[raw.index >= FIRST_MONTH].copy()
    gid, prod = d[col("GID_observed")], d[col("Production")]
    pipe_in, pipe_out = d[col("Imports_pipeline")], d[col("Exports_pipeline")]
    implied = gid + pipe_out - prod - pipe_in
    notes = []
    for stem, sign, what in (("Stock_change", +1, "stock change"), ("Receipts_other_sources", -1, "receipts from other sources")):
        s = d[col(stem)]
        if (s.fillna(0).abs() > 0).any():
            implied = implied + sign * s.fillna(0)
            notes.append(f"{what} is reported non-zero in {int((s.fillna(0).abs() > 0).sum())} month(s) and is IN the identity")
        else:
            notes.append(f"{what}: JODI reports only 0 or '-' (not reported), so it is NOT in the identity")
    m = pd.DataFrame(index=d.index)
    keep = [("Production", "Production"), ("Imports_pipeline", "Pipeline_imports"), ("Exports_pipeline", "Pipeline_exports"),
            ("GID_observed", "Consumption_GID_observed"), ("GID_calculated", "Consumption_GID_calculated"),
            ("Electricity_and_heat", "Power_electricity_and_heat"), ("Stock_change", "Stock_change")]
    for src, name in keep:
        m[f"{name}_JODI_fallback_mcm"] = d[col(src)]
    m["Consumption_plus_pipeline_exports_JODI_fallback_mcm"] = gid + pipe_out
    m["Implied_net_LNG_imports_mcm"] = implied.round(4)
    days = m.index.days_in_month
    for c in list(m.columns):
        m[c.replace("_mcm", "_Bcf_d")] = (m[c] * BCF_PER_MCM / days).round(4)
    m = m.dropna(how="all")
    return m, notes


# ------------------------------------------------------------------ official annual layer

def capmas_fetch():
    rows, blob = [], []
    for iid, pub in CAPMAS_IDS:
        r = requests.get(f"{CAPMAS}Indicator/IndicatorFilter?IndicatorId={iid}&SubSubjectId=30", headers=UA, timeout=90)
        r.raise_for_status()
        data = r.json().get("data") or []
        det = {}
        try:
            rd = requests.get(f"{CAPMAS}Indicator/IndicatorDetails?indicatorId={iid}&publicationId={pub}", headers=UA, timeout=90)
            det = (rd.json() or {}).get("data") or {}
        except Exception as e:  # noqa: BLE001
            log(f"  CAPMAS {iid} details: {type(e).__name__}")
        blob.append(json.dumps([data, {k: det.get(k) for k in ("measureUnit", "yearTypeId", "endPeriod", "name")}],
                               sort_keys=True, ensure_ascii=False))
        mu = det.get("measureUnit") or {}
        unit_en = next((t.get("name") for t in mu.get("measureUnitTranslations") or [] if t.get("locale") == "en"), None)
        unit = (unit_en or mu.get("name") or "unit not stated").strip()
        ind_en = next((t.get("name") for t in det.get("indicatorTranslations") or [] if t.get("locale") == "en"), "")
        notes_ar = re.sub(r"<[^>]+>", " ", str(det.get("notes") or "")).replace("</n>", " ")
        notes_ar = re.sub(r"\s+", " ", notes_ar).strip()
        for f in data:
            fen = next((t.get("name") for t in f.get("filterTranslations") or [] if t.get("locale") == "en"), "")
            for p in f.get("data") or []:
                if p.get("year") is None:
                    continue
                period = str(p["year"]) + (f" Q{p['quarter']}" if p.get("quarter") else "") + (f" M{p['month']}" if p.get("month") else "")
                rows.append({"Source": "CAPMAS", "Publication": {57: "Energy Balance", 77: "Electricity and Energy Statistics"}[pub],
                             "Indicator_id": iid, "Indicator_EN": ind_en, "Indicator_AR": det.get("name") or "",
                             "Item_EN": fen, "Item_AR": f.get("name") or "", "Period_as_published": period,
                             "Year": int(p["year"]), "Period_type": f"year as given by the API (yearTypeId={det.get('yearTypeId')}; "
                             "calendar or fiscal not stated)", "Value": p.get("value"), "Value_as_published": str(p.get("value")),
                             "Unit_as_stated": unit, "Notes_AR": notes_ar[:500],
                             "Link": f"{CAPMAS}Indicator/IndicatorFilter?IndicatorId={iid}&SubSubjectId=30"})
        time.sleep(0.3)
    return rows, hashlib.sha256("\n".join(blob).encode()).hexdigest()


def moee_fetch():
    r = requests.get(MOEE_FUEL, headers=UA, timeout=90)
    r.raise_for_status()
    tabs = pd.read_html(io.StringIO(r.text))
    t = next(t for t in tabs if "Year" in [str(c) for c in t.columns])
    t = t.drop(columns=[c for c in t.columns if str(c) in ("Chart",) or str(c).startswith("Unnamed")])
    rows = []
    for _, rr in t.iterrows():
        yr = str(rr["Year"]).strip()
        m = re.match(r"^(\d{4})\s*/\s*(\d{4})$", yr)
        if not m:
            continue
        end = max(int(m.group(1)), int(m.group(2)))
        fy = f"FY{end - 1}/{str(end)[2:]} (Jul {end - 1} - Jun {end})"
        for c in t.columns:
            if str(c) == "Year":
                continue
            raw = str(rr[c]).strip()
            txt = raw.replace("٫", ".").replace(",", "")
            try:
                val = float(txt)
            except ValueError:
                val = None    # '-' or blank as published
            head = str(c).strip()
            parts = head.split()
            unit = " ".join(parts[-2:]) if len(parts) >= 3 and parts[-2].lower() == "million" else parts[-1]
            rows.append({"Source": "MOEE", "Publication": "Statistics - fuel consumption (ST_fuel)", "Indicator_id": "ST_fuel",
                         "Indicator_EN": "Fuel used in power stations", "Indicator_AR": "", "Item_EN": head, "Item_AR": "",
                         "Period_as_published": yr, "Year": end, "Period_type": f"Egyptian fiscal year {fy}; Year = FY end",
                         "Value": val, "Value_as_published": raw, "Unit_as_stated": unit + " (from the column header)",
                         "Notes_AR": "", "Link": MOEE_FUEL})
    return rows, hashlib.sha256(t.to_csv(index=False).encode()).hexdigest()


# ------------------------------------------------------------------ checks

def checks(raw, monthly, annual):
    r = raw[raw.index >= FIRST_MONTH]
    yr = r.index.year
    months = r[col("Production")].notna().groupby(yr).sum()
    full = months[months == 12].index
    s = r.groupby(yr).sum(min_count=1)
    imp = monthly["Implied_net_LNG_imports_mcm"].groupby(monthly.index.year).sum(min_count=1)
    c = pd.DataFrame(index=s.index)
    c["JODI_months_reported"] = months
    for stem, name in (("Production", "Production"), ("Imports_pipeline", "Pipeline_imports"),
                       ("Exports_pipeline", "Pipeline_exports"), ("GID_observed", "GID_observed"),
                       ("GID_calculated", "GID_calculated"), ("Electricity_and_heat", "Power"),
                       ("Imports_LNG_reported", "LNG_imports_reported"), ("Exports_LNG_reported", "LNG_exports_reported"),
                       ("Imports_total", "Imports_total"), ("Exports_total", "Exports_total"),
                       ("Stock_change", "Stock_change")):
        c[f"JODI_{name}_mcm"] = s[col(stem)]
    c["JODI_reported_net_LNG_mcm"] = c["JODI_LNG_imports_reported_mcm"] - c["JODI_LNG_exports_reported_mcm"]
    c["Implied_net_LNG_imports_mcm"] = imp
    c["Implied_minus_reported_net_LNG_mcm"] = (c["Implied_net_LNG_imports_mcm"] - c["JODI_reported_net_LNG_mcm"]).round(1)
    a = pd.DataFrame(annual)
    cap = a[(a.get("Source") == "CAPMAS") & (a["Unit_as_stated"] == "Million cubic meter")] if len(a) else a
    for iid, name, jodi in ((3112, "production", "JODI_Production_mcm"), (3102, "gross_domestic_consumption", "JODI_GID_observed_mcm"),
                            (3121, "imports", "JODI_Imports_total_mcm"), (3124, "exports", "JODI_Exports_total_mcm"),
                            (3108, "final_consumption", None)):
        if not len(cap):
            break
        v = cap[cap["Indicator_id"] == iid].groupby("Year")["Value"].sum(min_count=1)
        c[f"CAPMAS_{iid}_{name}_million_m3"] = v
        if jodi:
            c[f"Ratio_JODI_to_CAPMAS_{name}"] = (c[jodi].where(c.index.isin(full)) / c[f"CAPMAS_{iid}_{name}_million_m3"]).round(3)
    c = c.round({k: (3 if k.startswith('Ratio') else 1) for k in c.columns})
    c.index = pd.to_datetime(c.index.astype(int).astype(str) + "-01-01")
    c.index.name = "date"
    # fiscal years (Jul-Jun) for MOEE gas for power
    fy_end = r.index.year + (r.index.month >= 7)
    pw = r[col("Electricity_and_heat")]
    f = pd.DataFrame({"JODI_months": pw.notna().groupby(fy_end).sum(),
                      "JODI_power_electricity_and_heat_mcm": pw.groupby(fy_end).sum(min_count=1).round(1)})
    f.loc[f["JODI_months"] < 12, "JODI_power_electricity_and_heat_mcm"] = float("nan")
    if len(a):
        mo = a[(a["Source"] == "MOEE") & a["Item_EN"].str.startswith("N.G.")].groupby("Year")["Value"].sum(min_count=1)
        f = f.join(mo.rename("MOEE_gas_for_power_million_m3"), how="outer")
        f["Ratio_JODI_to_MOEE"] = (f["JODI_power_electricity_and_heat_mcm"] / f["MOEE_gas_for_power_million_m3"]).round(3)
    f.index = pd.to_datetime(f.index.astype(int).astype(str) + "-01-01")
    f.index.name = "date"
    f.insert(0, "Fiscal_year", [f"FY{y - 1}/{str(y)[2:]} (Jul {y - 1} - Jun {y})" for y in f.index.year])
    return c, f[f.index.year >= 2000]


# ------------------------------------------------------------------ main

def read_sheet(path, sheet, index="date"):
    try:
        d = pd.read_excel(path, sheet_name=sheet)
    except (FileNotFoundError, ValueError):
        return None
    d = d.drop(columns=[c for c in d.columns if str(c).startswith("Unnamed")])
    if index and index in d.columns:
        d[index] = pd.to_datetime(d[index])
        d = d.set_index(index)
    return d


NOTES_TITLES = {"UNITS", "OWNER'S DECISIONS", "IMPLIED LNG", "SOURCES", "SHEETS", "COVERAGE", "CAVEATS", "IDENTITY TERMS"}


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", default=DEFAULT_OUT)
    ap.add_argument("--full", action="store_true", help="re-fetch every JODI month")
    args = ap.parse_args()
    now = datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M UTC")

    stored = read_sheet(args.out, S_RAW)
    if stored is not None:
        stored = stored.reindex(columns=RAW_COLS)
    rel = read_sheet(args.out, S_REL, index=None)
    rel = rel if rel is not None else pd.DataFrame(columns=["Source", "What", "Hash", "Checked_utc", "Changed_utc", "Rows"])
    annual_old = read_sheet(args.out, S_ANNUAL, index=None)
    annual_old = annual_old if annual_old is not None else pd.DataFrame()
    failures = []

    try:
        new, what = fetch_jodi(stored, args.full)
        raw = merge(stored, new)
        log(f"JODI: {new.dropna(how='all').shape[0]} month(s) received; stored {raw.dropna(how='all').index.min():%b %Y}"
            f" .. {raw.dropna(how='all').index.max():%b %Y}")
        rel = upsert(rel, "JODI", what, "", now, now, int(new.notna().sum().sum()))
    except Exception as e:  # noqa: BLE001
        failures.append(f"JODI: {type(e).__name__}: {e}")
        if stored is None:
            sys.exit(failures[-1])
        raw = stored

    annual = annual_old
    for src, fn, what in (("CAPMAS", capmas_fetch, "IndicatorFilter " + ",".join(str(i) for i, _ in CAPMAS_IDS)),
                          ("MOEE", moee_fetch, MOEE_FUEL)):
        try:
            rows, h = fn()
            prev = rel[rel["Source"] == src]
            if len(prev) and str(prev["Hash"].iloc[0]) == h and len(annual_old) and (annual_old["Source"] == src).any():
                log(f"{src}: unchanged (hash {h[:12]}), stored rows kept")
                rel = upsert(rel, src, what, h, now, prev["Changed_utc"].iloc[0], int(prev["Rows"].iloc[0]))
            else:
                log(f"{src}: new or changed release, {len(rows)} rows")
                keep = annual[annual["Source"] != src] if len(annual) else annual
                annual = pd.concat([keep, pd.DataFrame(rows)], ignore_index=True)
                rel = upsert(rel, src, what, h, now, now, len(rows))
        except Exception as e:  # noqa: BLE001
            failures.append(f"{src}: {type(e).__name__}: {str(e)[:200]} (stored rows kept)")
    if len(annual):
        annual = annual.sort_values(["Source", "Indicator_id", "Item_EN", "Year"], key=lambda s: s.astype(str)).reset_index(drop=True)

    monthly, id_notes = monthly_balance(raw)
    chk, chk_fy = checks(raw, monthly, annual.to_dict("records") if len(annual) else [])
    last = monthly["Production_JODI_fallback_mcm"].dropna().index.max()
    notes = [
        "UNITS",
        "JODI (fallback): million m3 per month (JODI 'Million Standard Cubic Metres'), as submitted by Egypt to JODI. "
        "'_Bcf_d' columns = mcm x 0.0353147 / days in the month (average daily rate).",
        "Annual official: values exactly as published; 'Unit_as_stated' is the unit the source states (CAPMAS API "
        "IndicatorDetails measureUnit; MOEE column header); where none is stated it reads 'unit not stated'. No unit "
        "is assumed and nothing on that sheet is converted.",
        "",
        "OWNER'S DECISIONS",
        "1. JODI monthly is used and labelled 'JODI (fallback)' on every series, sheet and chart: no official Egyptian "
        "monthly gas statistics are reachable (EGAS, EGPC, Ministry of Petroleum publish no tables).",
        "2. Official annual layer = CAPMAS (energy balance; petroleum-sector statistics) and MOEE (gas used in power "
        "stations, fiscal years).",
        "3. No Comtrade / customs layer.",
        "4. LNG is IMPLIED from the balance (below), not taken from JODI's reported LNG rows; the owner verifies it "
        "against Kpler. JODI's reported LNG imports/exports are on the raw sheet for reference only (not charted, not "
        "used in the balance).",
        "",
        "IMPLIED LNG",
        "Implied_net_LNG_imports = Gross Inland Deliveries (Observed) + pipeline exports - production - pipeline imports "
        "(JODI fallback rows). Positive = net LNG imports, negative = net LNG exports. It carries JODI's statistical "
        "difference (GID calculated - observed): large in 2011-13 (-3 to -15 bcm a year), when it differs strongly "
        "from JODI's reported net LNG (see 'Checks').",
        "",
        "IDENTITY TERMS",
        "JODI's identity is GID calculated = production + receipts from other sources + imports - exports - stock "
        "change (sign tested on the whole JODI table). Stock change and receipts enter the implied-LNG identity only "
        "if JODI reports non-zero values:",
        *[f"- {n}" for n in id_notes],
        "Egypt has no significant gas storage; JODI's stock change is 0 to Dec 2023 and not reported ('-') from Jan 2024; "
        "closing stocks are never reported.",
        "",
        "SOURCES",
        f"JODI Gas World database, jodidb.org Beyond 20/20 viewer (ReportId {JODI_REPORT}): "
        f"{JODI}/TableViewer/tableView.aspx?ReportId={JODI_REPORT} - Egypt, natural gas, all balance rows.",
        "CAPMAS JSON API: " + CAPMAS + "Indicator/IndicatorFilter?IndicatorId=<id>&SubSubjectId=30 (ids "
        + ", ".join(f"{i} (pub {p})" for i, p in CAPMAS_IDS) + "); units from api/Indicator/IndicatorDetails.",
        "MOEE statistics, fuel used in power stations: " + MOEE_FUEL + " (Egyptian fiscal years Jul-Jun).",
        "",
        "SHEETS",
        f"'{S_MONTHLY}': the balance used by the charts (JODI fallback rows in mcm and Bcf/d, implied net LNG).",
        f"'{S_RAW}': every JODI Egypt balance row in mcm, incl. the reported LNG rows (reference only).",
        f"'{S_ANNUAL}': CAPMAS and MOEE rows as published (long format, one row per indicator, item and year).",
        f"'{S_CHECKS}': calendar-year JODI sums vs CAPMAS energy balance (Million cubic meter as stated; ratios only for "
        "complete JODI years; CAPMAS does not state whether its years are calendar or fiscal), implied vs JODI-reported "
        f"net LNG (shown, not used). '{S_CHECKS_FY}': JODI electricity and heat use summed over Jul-Jun fiscal years vs "
        "MOEE gas for power. CAPMAS petroleum-sector series (9543, 9545, 5937, 5939) are in 'Quantity in 1000 tons' and "
        "are not compared (no conversion is made).",
        f"'{S_REL}': JODI request window, CAPMAS/MOEE response hashes (annual rows are replaced only when the hash changes).",
        "",
        "COVERAGE",
        f"JODI Egypt from Jan 2011; latest month with production: {last:%b %Y}" if pd.notna(last) else "JODI: no data",
        "",
        "CAVEATS",
        "JODI is a compiled submission, not an Egyptian official publication; months can be revised (the last 12 "
        "stored months are re-fetched every run).",
        "CAPMAS energy balance 2016 values (168) look like a publication error and are kept as published.",
        f"Last run {now}." + (" Failures: " + "; ".join(failures) if failures else ""),
    ]
    xlsx_notes.write_workbook(args.out, {
        S_MONTHLY: monthly,
        S_RAW: raw[raw.index >= FIRST_MONTH].dropna(how="all"),
        S_ANNUAL: annual if len(annual) else pd.DataFrame({"Source": []}),
        S_CHECKS: chk,
        S_CHECKS_FY: chk_fy,
        S_REL: rel.set_index("Source"),
    }, notes, NOTES_TITLES)
    print(f"Saved {args.out}: JODI {monthly.index.min():%b %Y} .. {last:%b %Y}; annual official rows {len(annual)}")
    for f in failures:
        log("FAILED: " + f)
    if failures:
        sys.exit(f"{len(failures)} source(s) failed (listed above); everything else is saved")


def upsert(rel, src, what, h, checked, changed, n):
    rel = rel[rel["Source"] != src] if len(rel) else rel
    return pd.concat([rel, pd.DataFrame([{"Source": src, "What": what, "Hash": h, "Checked_utc": checked,
                                          "Changed_utc": changed, "Rows": n}])], ignore_index=True)


if __name__ == "__main__":
    main()
