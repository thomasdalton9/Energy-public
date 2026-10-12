"""
Morocco natural gas, focused on gas burn at the two gas-fired plants: Tahaddart (EET, combined cycle, 384 MW) and
Ain Beni Mathar (ONEE, integrated solar combined cycle, 472 MW).

WHAT THE DATA SAY (probes discovery_archive/africa/MOROCCO_GAS_PROBE*.py, results discovery_archive/results/morocco/)
  - Since 2023 every m3 Morocco imports arrives through the GME pipeline at Tarifa (Spain -> Morocco, regasified LNG
    bought on the market) and goes to the two plants; Morocco's own small production (Gharb/Essaouira) goes to industry.
    So the Tarifa EXIT flow published by Enagas on ENTSOG = the two plants' gas burn (to within the checks below).
  - Before 31 Oct 2021 Tarifa carried Algerian transit gas Morocco -> Spain (Enagas ENTRY; Morocco took its royalty in
    kind upstream, which is not visible at Tarifa). No flow Nov 2021 - 28 Jun 2022. From 29 Jun 2022 the flow is
    reversed, Spain -> Morocco.

SOURCES
  ENTSOG transparency platform, point ITP-00082 Tarifa, operator Enagas (ES-TSO-0006), indicator Physical Flow, daily,
      kWh/d (stored as GWh/d), exit (Spain -> Morocco) and entry (Morocco -> Spain). Morocco's ONHYM publishes no rows.
      ENTSOG keeps a rolling five years (first day 2021-01-01 when built), so the COMMITTED WORKBOOK IS THE HISTORY STORE.
  Enagas monthly statistical bulletin PDFs (www.enagas.es, 'gas statistical bulletin', Jan 2017 on): the Tarifa row of
      the interconnection-points table (net balance, GWh: month, same month a year earlier, year to date, moving
      annual total; positive = entry into Spain). A bulletin is downloaded once, when its file first appears.
  HCP (Haut-Commissariat au Plan) 'Annuaire statistique du Maroc', chapter 'Energie' (Excel), ONEE tables: thermal
      generation by plant (million kWh), fuel use by plant (natural gas in thousand Nm3), natural gas activity (million
      Nm3: production, royalty, imports, use by power plants and industry). Each yearbook covers three years; the newest
      yearbook wins for each year. A yearbook (file id) is downloaded once.
  Ministry of Energy Transition and Sustainable Development (www.mem.gov.ma/Pages/rapports.html, decoded as UTF-8):
      monthly 'Chiffres cles' (year-to-date fuel burned for power, natural gas in MNm3, source ONEE, about two months
      behind; only recent editions stay online, so EVERY edition seen is stored) and the annual key figures (gas supply:
      imports, GME royalty in kind, national production; use: power plants, industry; 2007 on).
  UN Comtrade public preview API (keyless, one period per call, rate-limited): Morocco imports HS 271121 (natural gas
      in gaseous state), monthly from Jul 2022, by partner (= the LNG origin). The quantity Comtrade labels kg is in fact
      Enagas's kWh (it equals the Tarifa exit), value US$.
  JODI Gas World database (Morocco's own submission, jodidb.org viewer, reusing middle_east/EGYPT_GAS.py): FALLBACK.
  Ember monthly electricity data, Morocco gas generation: FALLBACK only (14-18% below the official figures).

LABELS: every figure computed here is marked DERIVED (volume factor, heat rate, Nm3/MWh, mcm/d, Bcf/d) or ESTIMATE
(gas-fired generation from Tarifa gas); JODI and Ember are FALLBACK. No gap is filled.

Schedule: 1st and 15th (.github/workflows/morocco_gas.yml). Charts: add_charts.py REGISTRY 'morocco_gas_daily.xlsx'
(africa/morocco_charts.py). Usage: python3 africa/MOROCCO_GAS.py [--out <xlsx>] [--only entsog,enagas,...]
"""
import argparse
import glob
import hashlib
import io
import os
import re
import shutil
import subprocess
import sys
import tempfile
import time
import zipfile
from datetime import date, datetime, timezone
from urllib.parse import unquote, urljoin

import pandas as pd
import requests

REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, REPO_ROOT)
sys.path.insert(0, os.path.join(REPO_ROOT, "middle_east"))
import xlsx_notes  # noqa: E402

DEFAULT_OUT = os.path.join(REPO_ROOT, "output", "Data and Chart Outputs", "morocco_gas_daily.xlsx")
UA = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36"}
BCF_PER_MCM = 0.0353147
GWH_PER_MNM3 = 11.7          # DERIVED: Enagas Tarifa GWh / ministry gas-to-power MNm3, 11.68-11.77 in 2023-25 ('Checks')
KWH_TO_MMBTU = 0.00341214
GAS_TO_POWER_FROM = pd.Timestamp("2023-01-01")

ENTSOG = "https://transparency.entsog.eu/api/v1/operationalData"
ENTSOG_KEYS = {"exit": "ES-TSO-0006ITP-00082exit", "entry": "ES-TSO-0006ITP-00082entry"}
ENTSOG_FIRST = pd.Timestamp("2021-01-01")
ENAGAS = "https://www.enagas.es"
ENAGAS_LIST = ENAGAS + "/en/technical-management-system/energy-data/publications/gas-statistical-bulletin/"
ENAGAS_FIRST = (2017, 1)
HCP_LIST = "https://www.hcp.ma/downloads/Annuaires-statistiques-du-Maroc-Format-Excel_t22392.html"
HCP_FIRST_YEARBOOK = 2012
MEM_PAGE = "https://www.mem.gov.ma/Pages/rapports.html"
COMTRADE = "https://comtradeapi.un.org/public/v1/preview/C/M/HS"
COMTRADE_FIRST = pd.Timestamp("2022-07-01")
COMTRADE_MAX_CALLS = 45
EMBER = "https://storage.googleapis.com/emb-prod-bkt-publicdata/public-downloads/monthly_full_release_long_format.csv"
JODI_FIRST = pd.Timestamp("2014-01-01")
PARTNERS = {0: "World", 842: "United States", 724: "Spain", 566: "Nigeria", 12: "Algeria", 634: "Qatar", 643: "Russia",
            624: "Guinea-Bissau", 226: "Equatorial Guinea", 818: "Egypt", 578: "Norway", 780: "Trinidad and Tobago",
            24: "Angola", 251: "France", 528: "Netherlands", 620: "Portugal", 899: "Areas, nes"}

S_DAILY, S_MONTHLY, S_BULLETIN = "Tarifa daily", "Tarifa monthly", "Enagas bulletin"
S_PLANTS, S_HCP = "Plants annual", "HCP by yearbook"
S_MIN_A, S_MIN_Y = "Ministry annual", "Ministry YTD"
S_COMTRADE, S_JODI, S_EMBER = "Comtrade monthly", "JODI (fallback)", "Ember (fallback)"
S_EST, S_EST_A, S_CHECKS, S_CHECKS_M, S_REL = "Estimate", "Estimate annual", "Checks", "Checks monthly", "Releases"
REL_COLS = ["Source", "Key", "What", "Hash", "Last_Modified", "Size", "Checked_utc", "Changed_utc", "Rows", "Note"]

MONTHS = {"jan": 1, "ene": 1, "janv": 1, "janvier": 1, "feb": 2, "fev": 2, "fevrier": 2, "mar": 3, "mars": 3,
          "apr": 4, "abr": 4, "avr": 4, "avril": 4, "may": 5, "mai": 5, "jun": 6, "juin": 6, "jul": 7, "juil": 7,
          "juillet": 7, "aug": 8, "ago": 8, "aout": 8, "sep": 9, "sept": 9, "septembre": 9, "oct": 10, "octobre": 10,
          "nov": 11, "novembre": 11, "dec": 12, "dic": 12, "decembre": 12}


def log(msg):
    print(msg, file=sys.stderr, flush=True)


def now_utc():
    return datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M UTC")


def get(url, tries=3, timeout=(20, 240), **kw):
    kw.setdefault("headers", UA)
    last = None
    for a in range(tries):
        try:
            r = requests.get(url, timeout=timeout, **kw)
            if r.status_code in (429, 500, 502, 503, 504):
                last = f"HTTP {r.status_code}"
                time.sleep(6 * (a + 1))
                continue
            return r
        except requests.RequestException as e:
            last = f"{type(e).__name__}: {str(e)[:120]}"
            time.sleep(4 * (a + 1))
    raise RuntimeError(f"{url[:120]}: {last}")


def deaccent(s):
    import unicodedata
    return "".join(c for c in unicodedata.normalize("NFKD", str(s)) if not unicodedata.combining(c))


def read_sheet(path, sheet, index=None):
    try:
        d = pd.read_excel(path, sheet_name=sheet)
    except (FileNotFoundError, ValueError):
        return None
    d = d.drop(columns=[c for c in d.columns if str(c).startswith("Unnamed")])
    if not len(d):
        return None
    if index:
        if index not in d.columns:
            return None
        d[index] = pd.to_datetime(d[index])
        d = d.set_index(index).sort_index()
    return d


class Releases:
    """'Releases' sheet: one row per (source, key) with hash / Last-Modified, so official files are re-parsed only when
    they change."""

    def __init__(self, df):
        self.df = df if df is not None and len(df) else pd.DataFrame(columns=REL_COLS)
        self.df = self.df.reindex(columns=REL_COLS)

    def get(self, src, key):
        m = self.df[(self.df["Source"] == src) & (self.df["Key"].astype(str) == str(key))]
        return m.iloc[0].to_dict() if len(m) else None

    def keys(self, src):
        return set(self.df.loc[self.df["Source"] == src, "Key"].astype(str))

    def put(self, src, key, changed, **kw):
        old = self.get(src, key) or {}
        row = {**{c: old.get(c) for c in REL_COLS}, **kw, "Source": src, "Key": str(key), "Checked_utc": now_utc()}
        if changed or not old.get("Changed_utc"):
            row["Changed_utc"] = now_utc()
        self.df = self.df[~((self.df["Source"] == src) & (self.df["Key"].astype(str) == str(key)))]
        self.df = pd.concat([self.df, pd.DataFrame([row])], ignore_index=True)


def sha(b):
    return hashlib.sha256(b).hexdigest()[:16]


# ------------------------------------------------------------------ ENTSOG (Tarifa daily)

def entsog_fetch(stored):
    """Physical flow at Tarifa, both directions, from the last stored day minus 30 (first run: 2021-01-01)."""
    start = ENTSOG_FIRST
    if stored is not None and len(stored):
        start = max(ENTSOG_FIRST, stored.index.max() - pd.Timedelta(days=30))
    end = pd.Timestamp(date.today())
    out = {}
    for d, key in ENTSOG_KEYS.items():
        vals = {}
        y0 = start
        while y0 <= end:
            y1 = min(end, pd.Timestamp(f"{y0.year}-12-31"))
            for a in range(4):
                try:
                    r = requests.get(ENTSOG, params={"indicator": "Physical Flow", "periodType": "day",
                                                     "from": y0.strftime("%Y-%m-%d"), "to": y1.strftime("%Y-%m-%d"),
                                                     "pointDirection": key, "limit": -1, "timezone": "CET"},
                                     headers={"User-Agent": "Mozilla/5.0", "Accept": "application/json"}, timeout=(20, 300))
                except requests.RequestException as e:
                    log(f"  ENTSOG {d} {y0:%Y}: {type(e).__name__}, retry")
                    time.sleep(10 * (a + 1))
                    continue
                if r.status_code == 404:      # no rows in the window
                    break
                if r.status_code != 200:
                    log(f"  ENTSOG {d} {y0:%Y}: HTTP {r.status_code}, retry")
                    time.sleep(15 * (a + 1))
                    continue
                for row in r.json().get("operationalData", []):
                    if row.get("unit") and str(row["unit"]).lower() != "kwh/d":
                        raise RuntimeError(f"ENTSOG: unexpected unit {row['unit']}")
                    try:
                        vals[pd.Timestamp(row["periodFrom"][:10])] = float(row["value"]) / 1e6   # kWh/d -> GWh/d
                    except (TypeError, ValueError, KeyError):
                        pass
                break
            else:
                raise RuntimeError(f"ENTSOG {d} {y0:%Y}: no answer after retries")
            y0 = y1 + pd.Timedelta(days=1)
        out[d] = pd.Series(vals, dtype=float)
    new = pd.DataFrame({"Exit_ES_to_MA_GWh_d": out["exit"], "Entry_MA_to_ES_GWh_d": out["entry"]}).sort_index()
    new.index.name = "date"
    return new, f"Physical Flow {start:%Y-%m-%d} .. {end:%Y-%m-%d}"


def tarifa_daily(raw):
    d = raw[["Exit_ES_to_MA_GWh_d", "Entry_MA_to_ES_GWh_d"]].copy()
    for a, b in (("Exit_ES_to_MA", "exit"), ("Entry_MA_to_ES", "entry")):
        d[f"{a}_mcm_d_DERIVED"] = (d[f"{a}_GWh_d"] / GWH_PER_MNM3).round(4)
        d[f"{a}_Bcf_d_DERIVED"] = (d[f"{a}_mcm_d_DERIVED"] * BCF_PER_MCM).round(5)
    ex, en = d["Exit_ES_to_MA_GWh_d"].fillna(0), d["Entry_MA_to_ES_GWh_d"].fillna(0)
    d["Direction"] = "no flow"
    d.loc[(ex > 0.5) & (ex >= en), "Direction"] = "Spain -> Morocco"
    d.loc[(en > 0.5) & (en > ex), "Direction"] = "Morocco -> Spain"
    d.loc[d[["Exit_ES_to_MA_GWh_d", "Entry_MA_to_ES_GWh_d"]].isna().all(axis=1), "Direction"] = "not reported"
    lab = pd.Series("", index=d.index)
    lab[d.index < "2021-11-01"] = "Algerian transit gas to Spain (GME, to 31 Oct 2021)"
    lab[(d.index >= "2021-11-01") & (d.index < "2022-06-29")] = "GME shut (no flow)"
    lab[(d.index >= "2022-06-29") & (d.index < GAS_TO_POWER_FROM)] = "regasified LNG to Morocco, plant restart (2022)"
    lab[d.index >= GAS_TO_POWER_FROM] = "gas to power (from 2023)"
    d["Label"] = lab
    return d


# ------------------------------------------------------------------ Enagas bulletins

def enagas_listing(rel):
    """Bulletin PDFs on the listing page per (year, month) filter; the listing month is the publication month."""
    seen = sorted(tuple(int(x) for x in str(k).split("|")[0].split("-")) for k in rel.keys("Enagas")
                  if re.match(r"^\d{4}-\d{2}\|", str(k)))
    today = date.today()
    y, m = ENAGAS_FIRST
    if seen:
        y, m = seen[-1]
        y, m = (y, m - 2) if m > 2 else (y - 1, m + 10)
    out = []
    while (y, m) <= (today.year, today.month):
        try:
            r = get(ENAGAS_LIST, params={"category": "", "month": m, "year": y}, tries=2, timeout=(20, 90))
            for x in re.findall(r'href="(/content/dam[^"]+\.pdf)"', r.text):
                out.append(((y, m), x))
        except RuntimeError as e:
            log(f"  Enagas listing {y}-{m:02d}: {e}")
        y, m = (y, m + 1) if m < 12 else (y + 1, 1)
    uniq, keep = set(), []
    for ym, x in out:
        if x not in uniq:
            uniq.add(x)
            keep.append((ym, x))
    return keep


MON_RE = re.compile(r"\b([A-Za-z]{3,4})[- ]?(\d{4}|\d{2})\b")


def _month_tokens(line):
    out = []
    for mo, yr in MON_RE.findall(deaccent(line)):
        k = mo.lower()
        if k not in MONTHS:
            continue
        yy = int(yr) + (2000 if len(yr) == 2 else 0)
        if 2010 <= yy <= 2040:
            out.append(pd.Timestamp(yy, MONTHS[k], 1))
    return out


def _num(tok):
    t = tok.strip()
    if not re.fullmatch(r"-?[\d.,]+", t):
        return None
    t = t.replace(".", "").replace(",", "")
    try:
        return float(t)
    except ValueError:
        return None


def _pct(tok):
    t = tok.strip().replace("%", "").replace(",", ".")
    try:
        return float(t)
    except ValueError:
        return None


def parse_bulletin_text(pages, listed=None):
    """The Tarifa row of the interconnection-points table: returns dict(month, current, previous, ytd, mat) or None.
    The two month columns are read from the header line above (their order changes between editions). listed =
    (year, month) of the listing page the file was found on: a bulletin covers the month 0-3 months before it, so a
    header printed with the wrong year (Jan 2020 bulletin: 'Jan-18 Jan-19') is moved to the year that fits."""
    for t in pages:
        lines = t.split("\n")
        for i, ln in enumerate(lines):
            m = re.match(r"^\s*Tarifa(?:\s+GME)?\s+(-?[\d.,]+\s.*)$", ln)
            if not m:
                continue
            hdr, months = None, []
            for j in range(i - 1, max(-1, i - 9), -1):
                mt = _month_tokens(lines[j])
                if "GWh" in lines[j] and len(mt) >= 2:
                    hdr, months = lines[j], mt[:2]
                    break
            if not months:
                continue
            toks = m.group(1).split()
            vals = [_num(x) for x in toks if "%" not in x and _num(x) is not None]
            pcts = [x for x in toks if "%" in x]
            if len(vals) < 3:
                continue
            cur_i = 0 if months[0] > months[1] else 1
            cur, prev = vals[cur_i], vals[1 - cur_i]
            month, shift, note = months[cur_i], 0, ""
            if listed:
                lm = pd.Timestamp(listed[0], listed[1], 1)
                for k in (0, 1, -1, 2):
                    lag = (lm.year - (month.year + k)) * 12 + lm.month - month.month
                    if -1 <= lag <= 3:
                        shift = k
                        break
                if shift:
                    note = f"header year {month.year} corrected to {month.year + shift} (listing {lm:%b %Y})"
                    month = pd.Timestamp(month.year + shift, month.month, 1)
            ytd = None
            yr = re.search(r"(?:Year|A.o)\s+(\d{4})", deaccent(hdr))
            if len(vals) >= 4 and not (yr and int(yr.group(1)) + shift < month.year):
                ytd = vals[2]
            check = ""
            p = _pct(pcts[0]) if pcts else None
            if p is not None and prev not in (0, None):
                calc = (cur - prev) / abs(prev) * 100
                check = "ok" if abs(calc - p) < 1.5 or abs(-calc - p) < 1.5 else f"pct mismatch ({calc:.1f} vs {p})"
            if note:
                check = (check + "; " if check else "") + note
            return {"month": month, "Tarifa_net_GWh": cur, "Same_month_prev_year_GWh": prev,
                    "Year_to_date_GWh": ytd, "Moving_annual_total_GWh": vals[-1], "Header": hdr.strip()[:120],
                    "Row_as_printed": ln.strip()[:160], "Check": check}
    return None


def pdf_pages(content, order=None, maxpages=70):
    import pdfplumber
    with pdfplumber.open(io.BytesIO(content)) as pdf:
        n = min(len(pdf.pages), maxpages)
        idx = order(n) if order else range(n)
        for i in idx:
            yield pdf.pages[i].extract_text() or ""


def enagas_fetch(rel, stored):
    rows = [] if stored is None else stored.reset_index().to_dict("records")
    listing = enagas_listing(rel)
    done = rel.keys("Enagas")
    files_done = {k.split("|", 1)[1] for k in done if "|" in k}
    new = 0
    for (y, m), path in listing:
        if path in files_done:
            continue
        key = f"{y}-{m:02d}|{path}"
        try:
            r = get(ENAGAS + path, timeout=(20, 180))
            if r.status_code != 200 or r.content[:4] != b"%PDF":
                rel.put("Enagas", key, True, What="bulletin PDF", Note=f"HTTP {r.status_code}, not a PDF")
                continue
            # the interconnection table sits around pages 9-16: read those first
            order = lambda n: list(range(8, min(n, 18))) + list(range(0, min(n, 8))) + list(range(18, n))
            res = None
            pages = []
            for t in pdf_pages(r.content, order):
                pages.append(t)
                res = parse_bulletin_text([t], (y, m))
                if res:
                    break
            note = "no Tarifa row in the PDF text" if res is None else ""
            if res:
                res.update({"File": unquote(path.split("/")[-1]), "Listed": f"{y}-{m:02d}"})
                rows.append(res)
                new += 1
            rel.put("Enagas", key, True, What="bulletin PDF", Hash=sha(r.content), Size=len(r.content),
                    Last_Modified=r.headers.get("Last-Modified"), Rows=1 if res else 0, Note=note)
        except Exception as e:  # noqa: BLE001
            log(f"  Enagas {path}: {type(e).__name__}: {str(e)[:150]}")
    log(f"Enagas: {len(listing)} bulletins listed in the window, {new} new parsed")
    if not rows:
        return None
    d = pd.DataFrame(rows)
    d["month"] = pd.to_datetime(d["month"])
    # one row per bulletin month: the latest listed file (revisions are re-posted under a new name)
    d = d.sort_values(["month", "Listed", "File"]).drop_duplicates("month", keep="last").set_index("month")
    d.index.name = "month"
    return d


# ------------------------------------------------------------------ HCP yearbooks

YEAR_CELL = re.compile(r"^\s*[^\d\s]{0,3}\s*((?:19|20)\d\d)(?:\.0)?\s*[*rRpP]?\s*$")   # '2024*', '2022R', ' م2018R'


def _year_of(v):
    if isinstance(v, (int, float)) and not pd.isna(v) and float(v).is_integer() and 1990 <= v <= 2040:
        return int(v)
    m = YEAR_CELL.match(str(v)) if isinstance(v, str) else None
    return int(m.group(1)) if m else None


def _val(v):
    if isinstance(v, (int, float)) and not pd.isna(v):
        return float(v)
    s = str(v).strip().replace(" ", "").replace(" ", "")
    s = s.replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return None


TABLES = [  # (key, title regex, [(item, label regex, unit)])
    ("gross", r"production\s+thermique", [("Tahaddart_GWh", r"tahaddart", "million kWh"),
                                           ("AinBeniMathar_GWh", r"mathar", "million kWh")]),
    ("net", r"production\s+nette", [("AinBeniMathar_net_GWh", r"mathar", "million kWh")]),
    ("fuel", r"consommation\s+et\s+des\s+stocks", [("Gas_plants_kNm3", r"^\s*gaz\s+naturel", "thousand Nm3"),
                                                  ("AinBeniMathar_kNm3", r"mathar", "thousand Nm3"),
                                                  ("Tahaddart_kNm3", r"tahaddart", "thousand Nm3")]),
    ("gas", r"activit.{0,4}\s*du\s+gaz", [("Production_MNm3", r"^\s*-?\s*production\b", "million Nm3"),
                                         ("Royalty_MNm3", r"redevance", "million Nm3"),
                                         ("Consumption_MNm3", r"^\s*-?\s*consommation\b", "million Nm3"),
                                         ("Power_plants_MNm3", r"centrales\s+thermiques", "million Nm3"),
                                         ("Industry_MNm3", r"industrie", "million Nm3"),
                                         ("Imports_MNm3", r"importation", "million Nm3")]),
]


def parse_hcp_frame(df, yearbook, workbook, sheet):
    """One sheet (header=None) -> long rows. Tables are found by their numbered title ('5 -10 Evolution de la
    production thermique ...'); year columns by the header row; rows by label."""
    rows, block, years = [], None, {}
    for _, r in df.iterrows():
        cells = list(r.values)
        texts = [(j, deaccent(c)) for j, c in enumerate(cells) if isinstance(c, str) and c.strip()]
        first = texts[0][1] if texts else ""
        if re.match(r"^\s*\d{1,2}\s*-\s*\d{1,2}\b", first):
            block = next((k for k, pat, _ in TABLES if re.search(pat, first, re.I)), None)
            title, years = first.strip(), {}
            continue
        yc = {j: _year_of(c) for j, c in enumerate(cells) if _year_of(c)}
        if len(yc) >= 2 and len(set(yc.values())) >= 2:
            years = yc
            unit = first.strip() if texts and texts[0][0] < min(yc) else ""
            continue
        if block is None or not years or not texts:
            continue
        lab_j, label = texts[0]
        if lab_j >= min(years):
            continue
        for k, pat, items in TABLES:
            if k != block:
                continue
            for item, lpat, unit_def in items:
                if re.search(lpat, label, re.I):
                    for j, yr in years.items():
                        v = _val(cells[j]) if j < len(cells) else None
                        if v is not None:
                            rows.append({"Yearbook": yearbook, "Workbook": workbook, "Sheet": str(sheet), "Table": title[:90],
                                         "Item": item, "Year": yr, "Value": v, "Unit_as_stated": unit or unit_def,
                                         "Label": label.strip()[:60]})
                    break
    return rows


def _extract(content, workdir):
    """zip or rar archive (or a bare workbook) -> list of workbook paths whose path names the energy chapter."""
    os.makedirs(workdir, exist_ok=True)
    arc = os.path.join(workdir, "a.bin")
    open(arc, "wb").write(content)
    if content[:2] == b"PK":
        z = zipfile.ZipFile(io.BytesIO(content))
        if any(n.startswith("xl/") for n in z.namelist()):
            return [arc]
        z.extractall(os.path.join(workdir, "x"))
    elif content[:4] == b"Rar!":
        ok = False
        for cmd in (["unrar", "x", "-o+", "-inul", arc, os.path.join(workdir, "x") + "/"],
                    ["7z", "x", "-y", f"-o{os.path.join(workdir, 'x')}", arc],
                    ["bsdtar", "-xf", arc, "-C", os.path.join(workdir, "x")]):
            if not shutil.which(cmd[0]):
                continue
            os.makedirs(os.path.join(workdir, "x"), exist_ok=True)
            p = subprocess.run(cmd, capture_output=True, text=True)
            files = glob.glob(os.path.join(workdir, "x", "**", "*"), recursive=True)
            if p.returncode == 0 or any(re.search(r"(?i)(energ|enreg|nergie)", f) for f in files):
                ok = True
                break
        if not ok:
            raise RuntimeError("could not extract the rar archive (unrar / 7z / bsdtar)")
    elif content[:4] == b"\xd0\xcf\x11\xe0":
        return [arc]
    else:
        raise RuntimeError(f"unknown archive type {content[:4]!r}")
    found = []
    for f in glob.glob(os.path.join(workdir, "x", "**", "*"), recursive=True):
        rel_ = os.path.relpath(f, workdir)
        if re.search(r"(?i)(energ|enreg|nergie)", rel_) and f.lower().endswith((".xls", ".xlsx")) \
                and not os.path.basename(f).startswith("~$"):
            found.append(f)
        elif f.lower().endswith(".rar") and re.search(r"(?i)(energ|enreg|nergie)", rel_):
            found += _extract(open(f, "rb").read(), f + "_x")
    return found


def hcp_fetch(rel, stored):
    r = get(HCP_LIST, timeout=(20, 90))
    items = []
    for m in re.finditer(r'href="(/file/(\d+)/)"', r.text):
        ctx = re.sub(r"<[^>]+>|\s+", " ", r.text[m.end():m.end() + 600])
        t = re.search(r"Annuaire[^|]{0,80}?ann[ée]e\s*(\d{4})", ctx)
        if t and int(t.group(1)) >= HCP_FIRST_YEARBOOK and m.group(2) not in [i[0] for i in items]:
            items.append((m.group(2), int(t.group(1))))
    log(f"HCP: {len(items)} yearbooks from {HCP_FIRST_YEARBOOK} listed")
    rows = [] if stored is None else stored.to_dict("records")
    done = rel.keys("HCP")
    for fid, yb in items:
        if fid in done and (rel.get("HCP", fid) or {}).get("Rows") not in (None, 0, "0") and \
                not str((rel.get("HCP", fid) or {}).get("Note") or "").startswith("failed"):
            continue
        try:
            rr = get(f"https://www.hcp.ma/file/{fid}/", timeout=(20, 300))
            if rr.status_code != 200:
                raise RuntimeError(f"HTTP {rr.status_code}")
            with tempfile.TemporaryDirectory() as td:
                books = _extract(rr.content, td)
                got = []
                for b in books:
                    try:
                        x = pd.read_excel(b, sheet_name=None, header=None)
                    except Exception as e:  # noqa: BLE001
                        log(f"  HCP {yb} {os.path.basename(b)}: {type(e).__name__}")
                        continue
                    for s, df in x.items():
                        got += parse_hcp_frame(df, yb, os.path.basename(b), s)
            rows = [x for x in rows if int(x["Yearbook"]) != yb] + got
            rel.put("HCP", fid, True, What=f"Annuaire statistique {yb}", Hash=sha(rr.content), Size=len(rr.content),
                    Last_Modified=rr.headers.get("Last-Modified"), Rows=len(got),
                    Note="" if got else f"no plant/gas rows found in {len(books)} energy workbook(s)")
            log(f"  HCP yearbook {yb} (file {fid}): {len(books)} energy workbook(s), {len(got)} values")
        except Exception as e:  # noqa: BLE001
            rel.put("HCP", fid, False, What=f"Annuaire statistique {yb}", Rows=0, Note=f"failed: {type(e).__name__}: {str(e)[:120]}")
            log(f"  HCP {yb} (file {fid}): {type(e).__name__}: {str(e)[:150]}")
    return pd.DataFrame(rows) if rows else None


def plants_annual(hcp):
    """Newest yearbook wins per year and item. Swap check: where a yearbook's gross table gives Tahaddart = the net
    table's Ain Beni Mathar (and not the other way), the two plant rows are swapped (the 2018 yearbook prints 2015 so)."""
    if hcp is None or not len(hcp):
        return pd.DataFrame()
    h = hcp.copy()
    h["Year"] = h["Year"].astype(int)
    h["Yearbook"] = h["Yearbook"].astype(int)
    notes = {}
    piv = h.pivot_table(index=["Yearbook", "Year"], columns="Item", values="Value", aggfunc="first")
    for (yb, yr), r in piv.iterrows():
        t, a, n = r.get("Tahaddart_GWh"), r.get("AinBeniMathar_GWh"), r.get("AinBeniMathar_net_GWh")
        if pd.notna(t) and pd.notna(a) and pd.notna(n) and n > 0 and abs(t - n) / n < 0.01 and abs(a - n) / n > 0.02:
            piv.loc[(yb, yr), ["Tahaddart_GWh", "AinBeniMathar_GWh"]] = [a, t]
            notes[(yb, yr)] = (f"{yb} yearbook prints {yr} Tahaddart/Ain Beni Mathar generation swapped; corrected "
                               f"from its net-generation table (Ain Beni Mathar {n:,.1f} million kWh)")
    out = {}
    for yr, g in piv.groupby(level="Year"):
        g = g.sort_index(level="Yearbook", ascending=False)
        rec = {}
        for c in g.columns:
            s = g[c].dropna()
            if len(s):
                rec[c] = s.iloc[0]
                rec[f"{c}__yb"] = s.index[0][0]
        rec["Note"] = "; ".join(v for (yb, y), v in notes.items() if y == yr and rec.get("Tahaddart_GWh__yb") == yb)
        out[yr] = rec
    d = pd.DataFrame(out).T.sort_index()
    p = pd.DataFrame(index=pd.to_datetime([f"{y}-01-01" for y in d.index]))
    p.index.name = "date"
    p["Year"] = d.index.values
    g = lambda c: pd.to_numeric(d[c], errors="coerce").values if c in d else float("nan")
    p["Tahaddart_GWh"] = g("Tahaddart_GWh")
    p["AinBeniMathar_GWh"] = g("AinBeniMathar_GWh")
    p["Gas_fired_GWh"] = p["Tahaddart_GWh"] + p["AinBeniMathar_GWh"]
    p["Tahaddart_gas_kNm3"] = g("Tahaddart_kNm3")
    p["AinBeniMathar_gas_kNm3"] = g("AinBeniMathar_kNm3")
    p["Plants_gas_MNm3"] = (p["Tahaddart_gas_kNm3"] + p["AinBeniMathar_gas_kNm3"]) / 1000
    p["Plants_gas_mcm_d_DERIVED"] = p["Plants_gas_MNm3"] / [366 if y % 4 == 0 else 365 for y in p["Year"]]
    p["Plants_gas_Bcf_d_DERIVED"] = p["Plants_gas_mcm_d_DERIVED"] * BCF_PER_MCM
    p["Tahaddart_Nm3_per_MWh_DERIVED"] = p["Tahaddart_gas_kNm3"] / p["Tahaddart_GWh"]
    p["AinBeniMathar_Nm3_per_MWh_DERIVED"] = p["AinBeniMathar_gas_kNm3"] / p["AinBeniMathar_GWh"]
    p["Both_plants_Nm3_per_MWh_DERIVED"] = p["Plants_gas_MNm3"] * 1000 / p["Gas_fired_GWh"]
    p["Both_plants_MWh_gas_per_MWh_DERIVED"] = p["Both_plants_Nm3_per_MWh_DERIVED"] * GWH_PER_MNM3 / 1000
    for c, src in (("Gas_activity_production_MNm3", "Production_MNm3"), ("Gas_activity_royalty_MNm3", "Royalty_MNm3"),
                   ("Gas_activity_imports_MNm3", "Imports_MNm3"), ("Gas_activity_power_plants_MNm3", "Power_plants_MNm3"),
                   ("Gas_activity_industry_MNm3", "Industry_MNm3"), ("Gas_activity_consumption_MNm3", "Consumption_MNm3")):
        p[c] = g(src)
    p["Yearbook_generation"] = d["Tahaddart_GWh__yb"].values if "Tahaddart_GWh__yb" in d else None
    p["Yearbook_gas"] = d["Tahaddart_kNm3__yb"].values if "Tahaddart_kNm3__yb" in d else None
    p["Note"] = d["Note"].values
    num = [c for c in p.columns if c not in ("Year", "Yearbook_generation", "Yearbook_gas", "Note")]
    p[num] = p[num].apply(pd.to_numeric, errors="coerce").round(4)
    return p.dropna(subset=["Tahaddart_GWh", "AinBeniMathar_GWh", "Tahaddart_gas_kNm3", "Plants_gas_MNm3",
                            "Gas_activity_power_plants_MNm3"], how="all")


# ------------------------------------------------------------------ Ministry of Energy Transition

def _edition(name):
    n = deaccent(unquote(name)).lower()
    y = re.findall(r"(20\d\d)", n)
    mo = next((MONTHS[w] for w in re.findall(r"[a-z]+", n) if w in MONTHS and len(w) > 3), None)
    if re.search(r"annuel", n):
        return "annual", (f"Edition {y[-1]}" if y else "annual (no year in name)")
    return "monthly", (f"{int(y[-1])}-{mo:02d}" if y and mo else "unknown")


def parse_mem_monthly(text):
    """'CONSOMMATION DES COMBUSTIBLES POUR LA PRODUCTION D'ELECTRICITE' table: Jan-<month> natural gas MNm3, two years."""
    lines = text.split("\n")
    for i, ln in enumerate(lines):
        if not re.search(r"CONSOMMATION DES COMBUSTIBLES POUR LA PRODUCTION", deaccent(ln), re.I) or "...." in ln:
            continue
        blk = lines[i + 1:i + 10]
        per = next((re.search(r"Janvier\s*-\s*([A-Za-z]+)", deaccent(b)) for b in blk
                    if re.search(r"Janvier\s*-\s*[A-Za-z]+", deaccent(b))), None)
        yrs = next((re.findall(r"(20\d\d)\*?", b) for b in blk if len(re.findall(r"\b20\d\d\*?", b)) == 2), None)
        gas = next((b for b in blk if re.search(r"Gaz naturel\s*\(MNm3\)", b, re.I)), None)
        if not (per and yrs and gas):
            continue
        m = re.match(r".*?\(MNm3\)\s+(-?\d{1,3}(?: \d{3})*(?:,\d+)?)\s+(-?\d{1,3}(?: \d{3})*(?:,\d+)?)\s+([-+]?[\d,]+)\s*$", gas)
        if not m:
            continue
        f = lambda s: float(s.replace(" ", "").replace(",", "."))
        a, b, pct = f(m.group(1)), f(m.group(2)), f(m.group(3))
        end = MONTHS.get(per.group(1).lower())
        check = "ok" if a and abs((b - a) / a * 100 - pct) < 1.0 else f"pct {pct} vs {(b - a) / a * 100 if a else 0:.1f}"
        return {"End_month": end, "Year_1": int(yrs[0]), "Value_1_MNm3": a, "Year_2": int(yrs[1]), "Value_2_MNm3": b,
                "Check": check, "Row_as_printed": gas.strip()[:120]}
    return None


ANNUAL_ITEMS = [("supply", r"^Importations\d*$", "Imports"), ("supply", r"^Redevance en nature\d*$", "Royalty_in_kind"),
                ("supply", r"^Production nationale\d*$", "National_production"), ("supply", r"^Total$", "Supply_total"),
                ("use", r"^Centrales Thermiques\d*$", "Power_plants"), ("use", r"^Industrie\d*$", "Industry"),
                ("use", r"^Total$", "Use_total")]


def parse_mem_annual(text):
    """Annual key figures: gas supply (2.7.2) and use (2.7.3) tables, MNm3 by year."""
    out, sec, years = [], None, []
    for ln in text.split("\n"):
        s = deaccent(ln).strip()
        if "...." in s:
            continue
        if re.search(r"Approvisionnement en gaz naturel", s, re.I) and not s.startswith(("1400", "Evolution")):
            sec = "supply"
            continue
        if re.search(r"Consommation du gaz naturel", s, re.I) and not s.startswith(("1400", "Evolution")):
            sec = "use"
            continue
        ys = re.findall(r"\b(20\d\d)\*?", s)
        if len(ys) >= 5 and re.fullmatch(r"[\d*\s]+", s):
            years = [int(y) for y in ys]
            continue
        if sec is None or not years:
            continue
        m = re.match(r"^([A-Za-z][A-Za-z ]*?\d?)\s+((?:-|[\d,]+)(?:\s+(?:-|[\d,]+))*)$", s)
        if not m:
            continue
        label, toks = m.group(1).strip(), m.group(2).split()
        item = next((it for sc, pat, it in ANNUAL_ITEMS if sc == sec and re.match(pat, label, re.I)), None)
        if not item or len(toks) != len(years):
            continue
        for y, t in zip(years, toks):
            if t != "-":
                out.append({"Table": sec, "Item": item, "Year": y, "Value_MNm3": float(t.replace(",", "."))})
        if item in ("Supply_total", "Use_total"):
            sec = None
    return out


def mem_fetch(rel, stored_a, stored_y):
    r = get(MEM_PAGE, timeout=(20, 90))
    html = r.content.decode("utf-8", errors="replace")   # the page sends no charset; r.text would give broken links
    links = []
    for h in re.findall(r'href="([^"]+\.pdf)"', html, re.I):
        u = urljoin(r.url, h)
        name = unquote(u.split("/")[-1])
        if re.search(r"chiffres", deaccent(name), re.I) and u not in links:
            links.append(u)
    log(f"Ministry: {len(links)} 'Chiffres cles' PDFs on the rapports page")
    ann = [] if stored_a is None else stored_a.to_dict("records")
    ytd = [] if stored_y is None else stored_y.to_dict("records")
    for u in links:
        name = unquote(u.split("/")[-1])
        old = rel.get("Ministry", u)
        try:
            if old:
                h = requests.head(u, headers=UA, timeout=(20, 60), allow_redirects=True)
                if h.ok and str(h.headers.get("Content-Length")) == str(old.get("Size")).split(".")[0] and \
                        str(h.headers.get("Last-Modified")) == str(old.get("Last_Modified")):
                    rel.put("Ministry", u, False, What=old.get("What"))
                    continue
            rr = get(u, timeout=(20, 180))
            if rr.content[:4] != b"%PDF":
                raise RuntimeError(f"not a PDF (HTTP {rr.status_code})")
            h = sha(rr.content)
            if old and old.get("Hash") == h:
                rel.put("Ministry", u, False, What=old.get("What"), Last_Modified=rr.headers.get("Last-Modified"),
                        Size=len(rr.content))
                continue
            kind, ed = _edition(name)
            text = "\n".join(pdf_pages(rr.content, maxpages=60))
            seen = now_utc()
            n = 0
            if kind == "annual":
                got = parse_mem_annual(text)
                ann = [x for x in ann if x.get("URL") != u] + [{"Edition": ed, "File": name, "URL": u, "First_seen_utc": seen, **g}
                                                               for g in got]
                n = len(got)
            else:
                g = parse_mem_monthly(text)
                ytd = [x for x in ytd if x.get("URL") != u]
                if g:
                    ytd.append({"Edition": ed, "File": name, "URL": u, "First_seen_utc": seen, **g})
                    n = 1
            rel.put("Ministry", u, True, What=f"{kind} key figures, {ed}", Hash=h, Size=len(rr.content),
                    Last_Modified=rr.headers.get("Last-Modified"), Rows=n, Note="" if n else "no gas table found")
            log(f"  Ministry {name}: {kind} {ed}, {n} row(s)")
        except Exception as e:  # noqa: BLE001
            log(f"  Ministry {name}: {type(e).__name__}: {str(e)[:150]}")
    a = pd.DataFrame(ann) if ann else None
    y = pd.DataFrame(ytd) if ytd else None
    return a, y


def ministry_ytd_long(y):
    """Each edition gives two year-to-date values (this year and the year before): one row per (edition, year)."""
    if y is None or not len(y):
        return pd.DataFrame()
    rows = []
    for _, r in y.iterrows():
        for k in (1, 2):
            if pd.notna(r.get(f"Year_{k}")) and pd.notna(r.get("End_month")):
                rows.append({"date": pd.Timestamp(int(r[f"Year_{k}"]), int(r["End_month"]), 1), "Edition": r["Edition"],
                             "Period": f"Jan-{pd.Timestamp(2000, int(r['End_month']), 1):%b} {int(r[f'Year_{k}'])}",
                             "Gas_to_power_YTD_MNm3": r[f"Value_{k}_MNm3"],
                             "Role": "current year" if k == 2 else "previous year (as printed)", "File": r["File"]})
    return pd.DataFrame(rows)


# ------------------------------------------------------------------ Comtrade

def comtrade_fetch(stored):
    have = set() if stored is None or not len(stored) else set(pd.to_datetime(stored["date"]).dt.strftime("%Y%m"))
    names = dict(PARTNERS)
    try:
        ref = get("https://comtradeapi.un.org/files/v1/app/reference/partnerAreas.json", tries=2, timeout=(15, 60)).json()
        names.update({int(x["id"]): x["text"] for x in ref.get("results", []) if str(x.get("id", "")).lstrip("-").isdigit()})
    except Exception as e:  # noqa: BLE001
        log(f"  Comtrade partner names: {type(e).__name__} (built-in names used)")
    last = pd.Timestamp(date.today()).to_period("M").to_timestamp() - pd.DateOffset(months=2)
    want = [p.strftime("%Y%m") for p in pd.date_range(COMTRADE_FIRST, last, freq="MS") if p.strftime("%Y%m") not in have]
    rows, calls = [], 0
    for per in want:
        if calls >= COMTRADE_MAX_CALLS:
            log(f"  Comtrade: call budget reached; the rest next run")
            break
        r = None
        for a in range(5):
            calls += 1
            try:
                r = requests.get(COMTRADE, params={"reporterCode": 504, "cmdCode": "271121", "flowCode": "M", "period": per},
                                 headers=UA, timeout=(15, 60))
            except requests.RequestException:
                time.sleep(5)
                continue
            if r.status_code == 429:
                time.sleep(8 + 5 * a)
                continue
            break
        time.sleep(4)
        if r is None or r.status_code != 200:
            log(f"  Comtrade {per}: {r.status_code if r is not None else 'no answer'}")
            continue
        for x in r.json().get("data", []):
            if str(x.get("cmdCode")) != "271121" or str(x.get("flowCode")) != "M":
                continue
            pc = int(x.get("partnerCode") or 0)
            rows.append({"date": pd.Timestamp(per[:4] + "-" + per[4:] + "-01"), "Partner_code": pc,
                         "Partner": names.get(pc, x.get("partnerDesc") or str(pc)),
                         "Qty_as_stated": x.get("qty"), "Qty_unit_as_stated": x.get("qtyUnitAbbr"),
                         "Net_weight_kg_as_stated": x.get("netWgt"), "Alt_qty_as_stated": x.get("altQty"),
                         "Value_USD": x.get("primaryValue"), "CIF_value_USD": x.get("cifvalue")})
    log(f"Comtrade: {len(want)} month(s) wanted, {len({r['date'] for r in rows})} returned data ({calls} calls)")
    new = pd.DataFrame(rows)
    if stored is not None and len(stored):
        stored = stored.copy()
        stored["date"] = pd.to_datetime(stored["date"])
        new = pd.concat([stored, new], ignore_index=True) if len(new) else stored
    if not len(new):
        return None, f"{len(want)} month(s) asked"
    return comtrade_derive(new), f"{len(want)} month(s) asked, {calls} calls"


def comtrade_derive(new):
    new = new.copy()
    new["date"] = pd.to_datetime(new["date"])
    new = new.drop_duplicates(["date", "Partner_code"], keep="last").sort_values(["date", "Partner_code"])
    q = pd.to_numeric(new["Net_weight_kg_as_stated"], errors="coerce").fillna(pd.to_numeric(new["Qty_as_stated"], errors="coerce"))
    new["Quantity_kWh_DERIVED"] = q
    new["Quantity_GWh_DERIVED"] = q / 1e6
    new["MMBtu_DERIVED"] = (q * KWH_TO_MMBTU).where(q > 0)   # a zero quantity (Nov 2025) gives no price
    new["USD_per_MMBtu_DERIVED"] = (pd.to_numeric(new["Value_USD"], errors="coerce") / new["MMBtu_DERIVED"]).round(3)
    new["Conversion"] = "quantity labelled kg = kWh (equals the Enagas Tarifa exit); MMBtu = kWh x 0.00341214"
    return new.reset_index(drop=True)


# ------------------------------------------------------------------ JODI and Ember (fallbacks)

def jodi_fetch(stored, full=False):
    import EGYPT_GAS as eg
    new, what = eg.fetch_jodi(stored, full, country="Morocco", first_month=JODI_FIRST)
    return eg.merge(stored, new), what


def ember_fetch(rel, stored):
    h = requests.head(EMBER, headers=UA, timeout=(20, 60))
    lm, size = h.headers.get("Last-Modified"), h.headers.get("Content-Length")
    old = rel.get("Ember", "monthly")
    if old and stored is not None and len(stored) and str(old.get("Last_Modified")) == str(lm):
        rel.put("Ember", "monthly", False, What=old.get("What"))
        log(f"Ember: release unchanged ({lm}), stored rows kept")
        return stored
    r = get(EMBER, timeout=(20, 600))
    d = pd.read_csv(io.BytesIO(r.content), low_memory=False)
    a = d[(d["Area"] == "Morocco") & (d["Unit"] == "TWh") & d["Variable"].isin(["Gas", "Total Generation", "Coal", "Demand"])]
    if "Category" in a:
        a = a[a["Category"].astype(str).str.contains("generation|demand", case=False)]
    p = a.pivot_table(index="Date", columns="Variable", values="Value", aggfunc="first")
    p.index = pd.to_datetime(p.index)
    p.index.name = "date"
    out = (p * 1000).rename(columns=lambda c: f"{c.replace(' ', '_')}_GWh_Ember_FALLBACK").sort_index()
    rel.put("Ember", "monthly", True, What="monthly_full_release_long_format.csv, Morocco", Last_Modified=lm, Size=size,
            Hash=sha(r.content), Rows=len(out))
    log(f"Ember: Morocco {len(out)} months to {out.index.max():%b %Y}")
    return out


# ------------------------------------------------------------------ derived tables

def tarifa_monthly(daily, bulletin):
    d = daily[["Exit_ES_to_MA_GWh_d", "Entry_MA_to_ES_GWh_d"]]
    g = d.resample("MS")
    m = pd.DataFrame({"Exit_ES_to_MA_GWh": g["Exit_ES_to_MA_GWh_d"].sum(min_count=1),
                      "Entry_MA_to_ES_GWh": g["Entry_MA_to_ES_GWh_d"].sum(min_count=1),
                      "Days_reported": d.notna().any(axis=1).resample("MS").sum()})
    m["Days_in_month"] = m.index.days_in_month
    m["Complete"] = (m["Days_reported"] >= m["Days_in_month"]).map({True: "yes", False: "no"})
    if len(m) and m.index.max() == pd.Timestamp(date.today()).to_period("M").to_timestamp():
        m.loc[m.index.max(), "Complete"] = "no (month in progress)"
    m["Net_entry_MA_to_ES_GWh"] = m["Entry_MA_to_ES_GWh"].fillna(0) - m["Exit_ES_to_MA_GWh"].fillna(0)
    m["Exit_mcm_d_DERIVED"] = m["Exit_ES_to_MA_GWh"] / GWH_PER_MNM3 / m["Days_in_month"]
    m["Exit_Bcf_d_DERIVED"] = m["Exit_mcm_d_DERIVED"] * BCF_PER_MCM
    if bulletin is not None and len(bulletin):
        b = bulletin.copy()
        b.index = pd.to_datetime(b.index)
        m = m.reindex(m.index.union(b.index))
        m["Enagas_bulletin_net_GWh"] = b["Tarifa_net_GWh"]
        # a month whose own bulletin has no readable Tarifa row: the same month as printed a year later
        nxt = b["Same_month_prev_year_GWh"].copy()
        nxt.index = nxt.index - pd.DateOffset(years=1)
        nxt = nxt[~nxt.index.duplicated()]
        fill = m["Enagas_bulletin_net_GWh"].isna() & m.index.isin(nxt.index)
        m["Enagas_bulletin_source"] = m["Enagas_bulletin_net_GWh"].notna().map({True: "own bulletin", False: ""})
        m.loc[fill, "Enagas_bulletin_net_GWh"] = nxt.reindex(m.index[fill]).values
        m.loc[fill, "Enagas_bulletin_source"] = "bulletin a year later (same-month column)"
        m["ENTSOG_minus_bulletin_GWh"] = (m["Net_entry_MA_to_ES_GWh"] - m["Enagas_bulletin_net_GWh"]).where(m["Complete"] == "yes")
    # Tarifa exit used downstream: the ENTSOG daily sum of a complete month, unless Enagas's own monthly bulletin differs
    # by more than 5% (ENTSOG shows zeros on 3-16 Dec 2024 while the bulletin, Comtrade and the ministry show the flow):
    # then the bulletin's figure for the same month is used and labelled. Nothing is filled where neither has the month.
    best = m["Exit_ES_to_MA_GWh"].where(m["Complete"] == "yes")
    srcs = pd.Series("ENTSOG daily sum", index=m.index).where(best.notna(), "")
    if "Enagas_bulletin_net_GWh" in m:
        b_exit = (-m["Enagas_bulletin_net_GWh"]).where(m["Enagas_bulletin_net_GWh"] < 0)
        diff = (m["Net_entry_MA_to_ES_GWh"] - m["Enagas_bulletin_net_GWh"]).abs()
        bad = best.notna() & b_exit.notna() & (diff > 20) & (diff > 0.05 * b_exit)
        for d_ in m.index[bad]:
            srcs[d_] = f"Enagas bulletin (ENTSOG daily sum {m.at[d_, 'Exit_ES_to_MA_GWh']:,.0f} GWh differs by more than 5%)"
        best = best.where(~bad, b_exit)
    m["Tarifa_exit_used_GWh"] = best
    m["Tarifa_exit_used_source"] = srcs
    m["Gas_to_power_MNm3_DERIVED"] = (m["Tarifa_exit_used_GWh"] / GWH_PER_MNM3).where(m.index >= GAS_TO_POWER_FROM)
    m["Days_in_month"] = m.index.days_in_month
    m.index.name = "month"
    num = m.select_dtypes("number").columns
    m[num] = m[num].round(4)
    return m


def heat_rate(plants):
    """Combined Nm3 per MWh over the HCP years since GAS_TO_POWER_FROM (DERIVED), falling back to all years."""
    if plants is None or not len(plants):
        return None, ""
    p = plants.dropna(subset=["Plants_gas_MNm3", "Gas_fired_GWh"])
    q = p[p.index >= GAS_TO_POWER_FROM]
    if not len(q):
        q = p
    if not len(q):
        return None, ""
    hr = q["Plants_gas_MNm3"].sum() * 1000 / q["Gas_fired_GWh"].sum()
    return round(hr, 1), f"HCP {int(q['Year'].min())}-{int(q['Year'].max())}: sum of gas Nm3 / sum of MWh, both plants"


def estimate(tm, plants, ember):
    hr, basis = heat_rate(plants)
    m = tm[(tm.index >= GAS_TO_POWER_FROM) & tm["Tarifa_exit_used_GWh"].notna()][
        ["Tarifa_exit_used_GWh", "Tarifa_exit_used_source", "Gas_to_power_MNm3_DERIVED"]].copy()
    m["Heat_rate_Nm3_per_MWh_DERIVED"] = hr
    m["Gas_fired_GWh_ESTIMATE"] = (m["Gas_to_power_MNm3_DERIVED"] * 1000 / hr).round(1) if hr else float("nan")
    if ember is not None and len(ember) and "Gas_GWh_Ember_FALLBACK" in ember:
        e = ember["Gas_GWh_Ember_FALLBACK"]
        m = m.reindex(m.index.union(e[e.index >= GAS_TO_POWER_FROM].index))
        m["Gas_GWh_Ember_FALLBACK"] = e
        m["ESTIMATE_over_Ember_FALLBACK"] = (m["Gas_fired_GWh_ESTIMATE"] / m["Gas_GWh_Ember_FALLBACK"]).round(3)
    m["Basis"] = f"ESTIMATE = Tarifa exit used GWh / {GWH_PER_MNM3} GWh per MNm3 / {hr} Nm3 per MWh ({basis})"
    m.index.name = "month"
    # annual: complete years only
    a = pd.DataFrame(index=sorted({y for y in m.index.year}))
    n = m["Gas_fired_GWh_ESTIMATE"].groupby(m.index.year).count()
    a["Months_estimated"] = n
    a["Gas_fired_GWh_ESTIMATE"] = m["Gas_fired_GWh_ESTIMATE"].groupby(m.index.year).sum().where(n == 12)
    if "Gas_GWh_Ember_FALLBACK" in m:
        ne = m["Gas_GWh_Ember_FALLBACK"].groupby(m.index.year).count()
        a["Gas_GWh_Ember_FALLBACK"] = m["Gas_GWh_Ember_FALLBACK"].groupby(m.index.year).sum().where(ne == 12)
    if plants is not None and len(plants):
        pl = plants.set_index("Year")
        a["Gas_fired_GWh_official_HCP"] = pl["Gas_fired_GWh"].reindex(a.index)
        a["Plants_gas_MNm3_official_HCP"] = pl["Plants_gas_MNm3"].reindex(a.index)
    if "Gas_fired_GWh_official_HCP" in a:
        a["ESTIMATE_vs_official_pct"] = ((a["Gas_fired_GWh_ESTIMATE"] / a["Gas_fired_GWh_official_HCP"] - 1) * 100).round(1)
        if "Gas_GWh_Ember_FALLBACK" in a:
            a["Ember_FALLBACK_vs_official_pct"] = ((a["Gas_GWh_Ember_FALLBACK"] / a["Gas_fired_GWh_official_HCP"] - 1) * 100).round(1)
    a.index = pd.to_datetime([f"{y}-01-01" for y in a.index])
    a.index.name = "date"
    a.insert(0, "Year", a.index.year)
    return m, a.round(1), hr, basis


def ministry_annual_latest(ma):
    """Newest edition per item and year (editions ranked by the last year their table covers)."""
    if ma is None or not len(ma):
        return pd.DataFrame()
    x = ma.copy()
    rank = x.groupby("URL")["Year"].max().rename("rank")
    x = x.join(rank, on="URL").sort_values(["rank", "First_seen_utc"])
    x = x.drop_duplicates(["Item", "Year"], keep="last")
    return x.pivot_table(index="Year", columns="Item", values="Value_MNm3", aggfunc="first")


def checks(tm, plants, ma, ytd_long, jodi, comtrade, bulletin):
    yrs = range(2014, date.today().year + 1)
    c = pd.DataFrame(index=list(yrs))
    t = tm[tm.index.notna()]
    full = t["Complete"].eq("yes").groupby(t.index.year).sum()
    c["Tarifa_exit_GWh_ENTSOG"] = t["Exit_ES_to_MA_GWh"].groupby(t.index.year).sum(min_count=1)
    c["Tarifa_entry_GWh_ENTSOG"] = t["Entry_MA_to_ES_GWh"].groupby(t.index.year).sum(min_count=1)
    c["ENTSOG_complete_months"] = full
    u = t["Tarifa_exit_used_GWh"]
    c["Tarifa_exit_used_GWh"] = u.groupby(t.index.year).sum(min_count=1).where(u.notna().groupby(t.index.year).sum() == 12)
    if bulletin is not None and len(bulletin):
        b = bulletin.copy()
        b.index = pd.to_datetime(b.index)
        dec = b[b.index.month == 12]["Moving_annual_total_GWh"]   # December's moving annual total = the calendar year
        c["Enagas_bulletin_year_net_GWh"] = pd.Series(dec.values, index=dec.index.year).reindex(c.index)
    if "Enagas_bulletin_net_GWh" in t:
        bm = t["Enagas_bulletin_net_GWh"]
        nb = bm.notna().groupby(t.index.year).sum()
        c["Enagas_bulletin_12_months_net_GWh"] = bm.groupby(t.index.year).sum(min_count=1).where(nb == 12)
    ml = ministry_annual_latest(ma)
    ytd = ytd_long if ytd_long is not None and len(ytd_long) else pd.DataFrame(columns=["date", "Gas_to_power_YTD_MNm3"])
    dec = ytd[pd.to_datetime(ytd["date"]).dt.month == 12] if len(ytd) else ytd
    gp = pd.Series(dtype=float)
    src = pd.Series(dtype=object)
    if len(ml) and "Power_plants" in ml:
        gp = ml["Power_plants"].dropna()
        src = pd.Series("Ministry annual key figures", index=gp.index)
    for _, r in dec.iterrows():
        y = pd.Timestamp(r["date"]).year
        if y not in gp.index:
            gp.loc[y] = r["Gas_to_power_YTD_MNm3"]
            src.loc[y] = f"Ministry Chiffres cles {r['Edition']} (Jan-Dec)"
    c["Ministry_gas_to_power_MNm3"] = gp.reindex(c.index)
    c["Ministry_source"] = src.reindex(c.index)
    if plants is not None and len(plants):
        c["HCP_plants_gas_MNm3"] = plants.set_index("Year")["Plants_gas_MNm3"].reindex(c.index)
        c["HCP_gas_activity_power_MNm3"] = plants.set_index("Year")["Gas_activity_power_plants_MNm3"].reindex(c.index)
    ok = c["Tarifa_exit_used_GWh"].notna()
    # Enagas's own annual figure (December bulletin) where printed, else the 12 months used
    ann = -c["Enagas_bulletin_year_net_GWh"] if "Enagas_bulletin_year_net_GWh" in c else pd.Series(float("nan"), index=c.index)
    basis = ann.where(ann > 0).fillna(c["Tarifa_exit_used_GWh"])
    c["Tarifa_GWh_for_factor"] = basis.where(c.index >= 2023)
    c["Tarifa_GWh_for_factor_source"] = pd.Series(
        ["Enagas December bulletin (annual)" if pd.notna(a) and a > 0 else ("12 months used" if pd.notna(x) else "")
         for a, x in zip(ann, c["Tarifa_exit_used_GWh"])], index=c.index).where(c.index >= 2023)
    c["Implied_GWh_per_MNm3_DERIVED"] = (c["Tarifa_GWh_for_factor"] / c["Ministry_gas_to_power_MNm3"]).round(3)
    c["Implied_GWh_per_MNm3_ENTSOG_only"] = (c["Tarifa_exit_GWh_ENTSOG"] / c["Ministry_gas_to_power_MNm3"]).where(
        (c["ENTSOG_complete_months"] == 12) & (c.index >= 2023)).round(3)
    c["Tarifa_MNm3_at_factor_DERIVED"] = (c["Tarifa_exit_used_GWh"] / GWH_PER_MNM3).where(ok).round(1)
    c["Tarifa_vs_ministry_pct"] = ((c["Tarifa_MNm3_at_factor_DERIVED"] / c["Ministry_gas_to_power_MNm3"] - 1) * 100).round(1)
    if jodi is not None and len(jodi):
        j = jodi.copy()
        j.index = pd.to_datetime(j.index)
        for colname, out in (("Imports_total_JODI_fallback_mcm", "JODI_imports_mcm_FALLBACK"),
                             ("Electricity_and_heat_JODI_fallback_mcm", "JODI_power_mcm_FALLBACK"),
                             ("Production_JODI_fallback_mcm", "JODI_production_mcm_FALLBACK")):
            if colname in j:
                s = j[colname]
                n = s.notna().groupby(s.index.year).sum()
                c[out] = s.groupby(s.index.year).sum(min_count=1).where(n == 12).reindex(c.index)
        if "JODI_imports_mcm_FALLBACK" in c:
            c["Tarifa_MNm3_over_JODI_imports"] = (c["Tarifa_MNm3_at_factor_DERIVED"] / c["JODI_imports_mcm_FALLBACK"]).where(c.index >= 2023).round(3)
    if comtrade is not None and len(comtrade):
        w = comtrade[comtrade["Partner_code"] == 0].set_index("date")["Quantity_GWh_DERIVED"]
        w.index = pd.to_datetime(w.index)
        n = w.notna().groupby(w.index.year).sum()
        c["Comtrade_qty_as_GWh"] = w.groupby(w.index.year).sum(min_count=1).where(n == 12).reindex(c.index).round(1)
        c["Comtrade_over_Tarifa_exit"] = (c["Comtrade_qty_as_GWh"] / c["Tarifa_exit_used_GWh"]).round(3)
    c.index = pd.to_datetime([f"{y}-01-01" for y in c.index])
    c.index.name = "date"
    c.insert(0, "Year", c.index.year)
    c = c.dropna(how="all", subset=[x for x in c.columns if x not in ("Year", "ENTSOG_complete_months")])
    # monthly
    cm = pd.DataFrame(index=t.index)
    cm["Tarifa_exit_GWh_ENTSOG"] = t["Exit_ES_to_MA_GWh"]
    cm["Tarifa_exit_used_GWh"] = t["Tarifa_exit_used_GWh"]
    cm["Tarifa_exit_used_source"] = t["Tarifa_exit_used_source"]
    cm["Tarifa_exit_MNm3_DERIVED"] = (t["Tarifa_exit_used_GWh"] / GWH_PER_MNM3).round(2)
    cm["Tarifa_complete"] = t["Complete"]
    if "Enagas_bulletin_net_GWh" in t:
        cm["Enagas_bulletin_net_GWh"] = t["Enagas_bulletin_net_GWh"]
    if jodi is not None and len(jodi):
        for colname, out in (("Imports_total_JODI_fallback_mcm", "JODI_imports_mcm_FALLBACK"),
                             ("Electricity_and_heat_JODI_fallback_mcm", "JODI_power_mcm_FALLBACK")):
            if colname in j:
                cm[out] = j[colname].reindex(cm.index)
        if "JODI_imports_mcm_FALLBACK" in cm:
            cm["Tarifa_MNm3_minus_JODI_imports"] = (cm["Tarifa_exit_MNm3_DERIVED"] - cm["JODI_imports_mcm_FALLBACK"]).round(2)
    if comtrade is not None and len(comtrade):
        cm["Comtrade_qty_as_GWh"] = w.reindex(cm.index)
        cm["Comtrade_minus_Tarifa_GWh"] = (cm["Comtrade_qty_as_GWh"] - cm["Tarifa_exit_used_GWh"]).round(2)
    cm = cm[cm.index >= "2017-01-01"].dropna(how="all", subset=[x for x in cm.columns if x not in ("Tarifa_complete", "Tarifa_exit_used_source")])
    cm.index.name = "month"
    return c, cm


def ytd_compare(ytd_long, tm):
    if ytd_long is None or not len(ytd_long):
        return ytd_long
    y = ytd_long.copy()
    y["date"] = pd.to_datetime(y["date"])
    t = tm["Tarifa_exit_used_GWh"]
    vals = []
    for d in y["date"]:
        part = t[(t.index.year == d.year) & (t.index.month <= d.month)].dropna()
        ok = len(part) == d.month and d >= GAS_TO_POWER_FROM
        vals.append(round(part.sum() / GWH_PER_MNM3, 1) if ok else None)
    y["Tarifa_YTD_MNm3_DERIVED"] = vals
    y["Tarifa_vs_ministry_pct"] = ((pd.to_numeric(y["Tarifa_YTD_MNm3_DERIVED"]) / y["Gas_to_power_YTD_MNm3"] - 1) * 100).round(1)
    return y.sort_values(["date", "Edition"]).set_index("date")


# ------------------------------------------------------------------ main

NOTES_TITLES = {"UNITS", "LABELS", "WHAT THE TARIFA FLOW IS", "SOURCES", "SHEETS", "DERIVED FACTORS", "COVERAGE",
                "CAVEATS", "RUN"}


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", default=DEFAULT_OUT)
    ap.add_argument("--only", default="", help="comma list of sources to fetch (entsog,enagas,hcp,ministry,comtrade,jodi,ember)")
    ap.add_argument("--full-jodi", action="store_true")
    args = ap.parse_args()
    only = set(filter(None, args.only.split(",")))
    want = lambda s: not only or s in only
    rel = Releases(read_sheet(args.out, S_REL))
    failures = []

    raw_daily = read_sheet(args.out, S_DAILY, "date")
    raw_daily = raw_daily[["Exit_ES_to_MA_GWh_d", "Entry_MA_to_ES_GWh_d"]] if raw_daily is not None else None
    if want("entsog"):
        try:
            new, what = entsog_fetch(raw_daily)
            if raw_daily is not None and len(raw_daily):
                merged = raw_daily.reindex(raw_daily.index.union(new.index))
                for c_ in new.columns:
                    nv = new[c_].dropna()
                    merged.loc[nv.index, c_] = nv
                raw_daily = merged
            else:
                raw_daily = new
            rel.put("ENTSOG", "ITP-00082", True, What=what, Rows=int(new.notna().sum().sum()))
            log(f"ENTSOG: {new.notna().any(axis=1).sum()} day(s) received; stored {raw_daily.index.min():%Y-%m-%d} .. {raw_daily.index.max():%Y-%m-%d}")
        except Exception as e:  # noqa: BLE001
            failures.append(f"ENTSOG: {type(e).__name__}: {str(e)[:200]}")
    if raw_daily is None or not len(raw_daily):
        sys.exit("ENTSOG: no Tarifa data stored or fetched")
    daily = tarifa_daily(raw_daily.sort_index())

    bulletin = read_sheet(args.out, S_BULLETIN, "month")
    if want("enagas"):
        try:
            bulletin = enagas_fetch(rel, bulletin)
        except Exception as e:  # noqa: BLE001
            failures.append(f"Enagas: {type(e).__name__}: {str(e)[:200]}")

    hcp = read_sheet(args.out, S_HCP)
    if want("hcp"):
        try:
            hcp = hcp_fetch(rel, hcp)
        except Exception as e:  # noqa: BLE001
            failures.append(f"HCP: {type(e).__name__}: {str(e)[:200]}")
    plants = plants_annual(hcp)

    ma = read_sheet(args.out, S_MIN_A)
    my_raw = read_sheet(args.out, "Ministry YTD editions")
    if want("ministry"):
        try:
            ma, my_raw = mem_fetch(rel, ma, my_raw)
        except Exception as e:  # noqa: BLE001
            failures.append(f"Ministry: {type(e).__name__}: {str(e)[:200]}")

    comtrade = read_sheet(args.out, S_COMTRADE)
    if comtrade is not None:
        comtrade = comtrade[[c for c in comtrade.columns if c in ("date", "Partner_code", "Partner", "Qty_as_stated",
                             "Qty_unit_as_stated", "Net_weight_kg_as_stated", "Alt_qty_as_stated", "Value_USD", "CIF_value_USD")]]
    if want("comtrade"):
        try:
            comtrade, what = comtrade_fetch(comtrade)
            rel.put("Comtrade", "504 M 271121", True, What=what, Rows=0 if comtrade is None else len(comtrade))
        except Exception as e:  # noqa: BLE001
            failures.append(f"Comtrade: {type(e).__name__}: {str(e)[:200]}")
    elif comtrade is not None and len(comtrade):
        comtrade = comtrade_derive(comtrade)

    jodi = read_sheet(args.out, S_JODI, "date")
    if want("jodi"):
        try:
            jodi, what = jodi_fetch(jodi, args.full_jodi)
            rel.put("JODI", "Morocco", True, What=what, Rows=int(jodi.notna().sum().sum()))
        except Exception as e:  # noqa: BLE001
            failures.append(f"JODI: {type(e).__name__}: {str(e)[:200]}")

    ember = read_sheet(args.out, S_EMBER, "date")
    if want("ember"):
        try:
            ember = ember_fetch(rel, ember)
        except Exception as e:  # noqa: BLE001
            failures.append(f"Ember: {type(e).__name__}: {str(e)[:200]}")

    tm = tarifa_monthly(daily, bulletin)
    est, est_a, hr, hr_basis = estimate(tm, plants, ember)
    ytd_long = ministry_ytd_long(my_raw)
    chk, chk_m = checks(tm, plants, ma, ytd_long, jodi, comtrade, bulletin)
    ytd_cmp = ytd_compare(ytd_long, tm)
    factors = chk["Implied_GWh_per_MNm3_DERIVED"].dropna() if "Implied_GWh_per_MNm3_DERIVED" in chk else pd.Series(dtype=float)
    if len(factors) and ((factors - GWH_PER_MNM3).abs() > 0.25).any():
        log(f"WARNING: implied GWh per MNm3 {factors.round(2).to_dict()} departs from the {GWH_PER_MNM3} used")

    flow = daily[daily[["Exit_ES_to_MA_GWh_d", "Entry_MA_to_ES_GWh_d"]].notna().any(axis=1)]
    notes = [
        "UNITS",
        "Tarifa flows: GWh/d as published by Enagas on ENTSOG (kWh/d / 1e6). mcm/d and Bcf/d are DERIVED: GWh / "
        f"{GWH_PER_MNM3} GWh per million Nm3 (DERIVED factor, below) = mcm/d (million Nm3 treated as mcm); Bcf/d = mcm/d x 0.0353147.",
        "Enagas bulletin: GWh per month, net balance as printed (positive = entry into Spain = Morocco -> Spain; negative = exit to Morocco).",
        "HCP: generation in million kWh (= GWh), gas burned in thousand Nm3, gas activity in million Nm3 (unit as stated per row).",
        "Ministry: natural gas in MNm3 (million normal m3), as printed.",
        "Comtrade: value in US$; the quantity Comtrade labels kg is kWh (it equals the Tarifa exit to within rounding), so "
        "US$/MMBtu (DERIVED) = value / (kWh x 0.00341214).",
        "JODI: million m3 (Million Standard Cubic Metres), Morocco's submission. Ember: GWh (TWh x 1000).",
        "",
        "LABELS",
        "DERIVED = computed here from published figures (volume factor, mcm/d, Bcf/d, Nm3/MWh, heat rate, US$/MMBtu). "
        "ESTIMATE = monthly gas-fired generation from Tarifa gas and the derived heat rate. FALLBACK = JODI and Ember "
        "(compiled, not Moroccan official publications; Ember Morocco gas is 14-18% below the official HCP figures). No gap is filled.",
        "",
        "WHAT THE TARIFA FLOW IS",
        "To 31 Oct 2021: Algerian gas in transit to Spain through the Maghreb-Europe pipeline (GME), Enagas ENTRY at Tarifa. "
        "Morocco's own gas then was the GME royalty taken in kind upstream plus Sonatrach contract gas, NOT visible at Tarifa.",
        "1 Nov 2021 - 28 Jun 2022: no flow (Algeria did not renew the GME contract). From 29 Jun 2022: reversed flow, Spain -> "
        "Morocco (Enagas EXIT), regasified LNG Morocco buys on the market.",
        "From 2023 the Tarifa exit = the gas burned at Tahaddart and Ain Beni Mathar: imports = power burn and Morocco's own "
        "production goes to industry (ministry and HCP balances; 'Checks'). 2022 is the restart year and is not used for the estimate.",
        "",
        "DERIVED FACTORS",
        f"Volume factor {GWH_PER_MNM3} GWh per million Nm3 = Enagas Tarifa exit GWh / ministry gas-to-power MNm3 "
        "(11.68-11.77 in 2023-25 when built; recomputed every run on 'Checks', Implied_GWh_per_MNm3_DERIVED).",
        f"Heat rate {hr} Nm3 per MWh (both plants; {hr_basis}); plant rates on 'Plants annual' (Tahaddart 177-184, "
        "Ain Beni Mathar 183-202 in 2015-24 excl. the 2022 restart). x 11.7 kWh per Nm3 = about 2.2 MWh of gas per MWh.",
        "",
        "SOURCES",
        f"ENTSOG transparency platform, {ENTSOG}?indicator=Physical Flow&periodType=day&pointDirection=ES-TSO-0006ITP-00082exit|entry "
        "(Tarifa, operator Enagas; Morocco's ONHYM publishes no rows). ENTSOG keeps a rolling 5 years: this workbook is the history store.",
        f"Enagas monthly statistical bulletin: {ENAGAS_LIST} (Tarifa row of the interconnection-points table, from Jan 2017).",
        f"HCP Annuaire statistique du Maroc (Excel), chapter Energie, ONEE tables: {HCP_LIST}",
        f"Ministry of Energy Transition and Sustainable Development, reports page: {MEM_PAGE} (monthly 'Chiffres cles', "
        "annual key figures).",
        f"UN Comtrade public API: {COMTRADE}?reporterCode=504&cmdCode=271121&flowCode=M&period=YYYYMM",
        "JODI Gas World database (jodidb.org, ReportId 38673), Morocco, all balance rows - FALLBACK.",
        f"Ember monthly electricity data: {EMBER} - FALLBACK.",
        "",
        "SHEETS",
        f"'{S_DAILY}': ENTSOG exit/entry GWh/d with DERIVED mcm/d and Bcf/d, direction and period label.",
        f"'{S_MONTHLY}': ENTSOG monthly totals (complete months flagged), gas to power MNm3 (DERIVED, from 2023) and the Enagas "
        "bulletin Tarifa net balance (own bulletin, else the same-month column of the bulletin a year later) as a cross-check.",
        f"'{S_BULLETIN}': the Tarifa row read from each bulletin (month, same month a year earlier, year to date, moving annual total).",
        f"'{S_PLANTS}': per-plant generation and gas burn (newest yearbook per year), Nm3/MWh per plant (DERIVED), HCP gas activity. "
        f"'{S_HCP}': every value read from every yearbook.",
        f"'{S_MIN_A}': annual gas supply and use, every edition kept (Edition, First_seen_utc). '{S_MIN_Y}': year-to-date gas for "
        "power from every monthly edition seen (each edition prints this year and the year before), with Tarifa year to date "
        "(DERIVED) beside it. 'Ministry YTD editions': the editions as parsed.",
        f"'{S_COMTRADE}': value, quantity as stated, DERIVED kWh / MMBtu / US$ per MMBtu, partner = LNG origin (partner 0 = World total).",
        f"'{S_JODI}': Morocco's JODI balance rows (FALLBACK). '{S_EMBER}': Ember Morocco monthly (FALLBACK).",
        f"'{S_EST}': monthly gas-fired generation ESTIMATE next to Ember (FALLBACK). '{S_EST_A}': complete-year sums beside the "
        "official HCP values.",
        f"'{S_CHECKS}': Tarifa vs ministry gas-to-power by year, implied factor, Tarifa vs JODI imports and Comtrade quantity. "
        f"'{S_CHECKS_M}': the same by month. '{S_REL}': request windows, file hashes and Last-Modified (official files are "
        "downloaded and parsed only when new or changed).",
        "",
        "COVERAGE",
        f"Tarifa daily {flow.index.min():%d %b %Y} - {flow.index.max():%d %b %Y}" if len(flow) else "Tarifa: none",
        f"Enagas bulletin months {bulletin.index.min():%b %Y} - {bulletin.index.max():%b %Y} ({len(bulletin)} months)"
        if bulletin is not None and len(bulletin) else "Enagas bulletin: none",
        f"HCP plants {int(plants['Year'].min())}-{int(plants['Year'].max())}" if len(plants) else "HCP: none",
        "",
        "CAVEATS",
        "The Tarifa exit is gas delivered into Morocco's GME section; it equals the plants' burn only on annual totals "
        "(within about 1%); monthly line-pack changes are not published.",
        "Comtrade months and origins follow Moroccan customs declarations, not the physical flow date.",
        "Ministry 'Chiffres cles' are provisional when first printed; editions are kept as printed.",
        "",
        "RUN",
        f"Last run {now_utc()}." + (" Failures: " + "; ".join(failures) if failures else ""),
    ]
    est_out = est.copy()
    sheets = {
        S_DAILY: daily,
        S_MONTHLY: tm,
        S_BULLETIN: bulletin if bulletin is not None and len(bulletin) else pd.DataFrame({"month": []}).set_index("month"),
        S_PLANTS: plants if len(plants) else pd.DataFrame({"date": []}).set_index("date"),
        S_HCP: (hcp.sort_values(["Item", "Year", "Yearbook"]).reset_index(drop=True) if hcp is not None and len(hcp)
                else pd.DataFrame({"Yearbook": []})),
        S_MIN_A: (ma.sort_values(["Item", "Year", "Edition"]).reset_index(drop=True) if ma is not None and len(ma)
                  else pd.DataFrame({"Edition": []})),
        S_MIN_Y: ytd_cmp if ytd_cmp is not None and len(ytd_cmp) else pd.DataFrame({"date": []}).set_index("date"),
        "Ministry YTD editions": (my_raw.reset_index(drop=True) if my_raw is not None and len(my_raw)
                                  else pd.DataFrame({"Edition": []})),
        S_COMTRADE: (comtrade.set_index("date") if comtrade is not None and len(comtrade) else pd.DataFrame({"date": []}).set_index("date")),
        S_JODI: (jodi[jodi.index >= JODI_FIRST].dropna(how="all") if jodi is not None and len(jodi)
                 else pd.DataFrame({"date": []}).set_index("date")),
        S_EMBER: ember if ember is not None and len(ember) else pd.DataFrame({"date": []}).set_index("date"),
        S_EST: est_out,
        S_EST_A: est_a,
        S_CHECKS: chk,
        S_CHECKS_M: chk_m,
        S_REL: rel.df.sort_values(["Source", "Key"]).set_index("Source"),
    }
    for k, v in sheets.items():
        if v.index.name in ("date", "month") and len(v):
            v.index = pd.to_datetime(v.index)
    os.makedirs(os.path.dirname(os.path.abspath(args.out)), exist_ok=True)
    xlsx_notes.write_workbook(args.out, sheets, notes, NOTES_TITLES)
    print(f"Saved {args.out}: Tarifa {flow.index.min():%Y-%m-%d} .. {flow.index.max():%Y-%m-%d}; bulletin months "
          f"{0 if bulletin is None else len(bulletin)}; HCP years {len(plants)}; heat rate {hr}")
    for f in failures:
        log("FAILED: " + f)
    if failures:
        sys.exit(f"{len(failures)} source(s) failed (listed above); everything else is saved")


if __name__ == "__main__":
    main()
