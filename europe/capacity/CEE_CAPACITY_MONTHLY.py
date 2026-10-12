"""
Monthly installed generation capacity for Hungary, Czechia, Switzerland and (quarterly, partial) Poland from the
official sources. Nothing is interpolated: a period a source does not cover is absent.

  output/Data and Chart Outputs/hungary_power_capacity_monthly.xlsx
  output/Data and Chart Outputs/czechia_power_capacity_monthly.xlsx
  output/Data and Chart Outputs/switzerland_power_capacity_monthly.xlsx
  output/Data and Chart Outputs/poland_power_capacity_quarterly.xlsx

Date convention (same as the ENTSO-E annual files, dated 1 January, and the Nordic / GB workbooks): the row dated
1 May holds the capacity installed at 00:00 on 1 May, i.e. at the END of April. Total_MW excludes pumped storage
(and batteries), as in the ENTSO-E workbooks. Sheet 'Check_ENTSOE' sets every 1 January row against ENTSO-E A68.

HUNGARY - MAVIR (the TSO), 'Energy MIX - Power Plants - Installed Capacity Data' page: the PDF 'Gross / net installed
  capacity per production type, MAVIR 2000-2020-2026' with GROSS installed capacity per production type, annual to
  2019 and MONTHLY from January 2020, current to the latest month (Aug 2026 in the file read on 12 Oct 2026). Gross
  capacity of power plants + household-size plants (HMKE) + small-scale prosumer plants (SCTE), so rooftop solar IS in
  it. Parsed from the PDF words by position; every row is validated against the PDF's own subtotals and dropped
  (and logged) if they do not add up. The PDF is whole-file: downloaded only when Last-Modified / size changes.
  Fuels: Nuclear; Coal (coal + lignite); Oil; Gas; Storage (batteries); Hydro; Wind (industrial + HMKE); Solar
  (industrial + prosumer + HMKE); Bioenergy (biogas + biomass); Other (waste, geothermal, other HMKE).

CZECHIA - ERU (energy regulator), quarterly 'Zprava o provozu elektrizacni soustavy CR', sheet 5 of the xlsx
  attachment: installed capacity by technology for EACH MONTH of the year up to the report's quarter (the Q4 report
  holds all twelve). Technologies are ERU's own: JE nuclear, PE steam (coal, lignite and biomass steam plant - kept
  as Coal_MW and also as Steam_all_fuels_MW), PPE combined cycle + PSE gas / combustion turbines and engines (summed
  into Gas_MW; includes plant on other fuels), VE hydro, PVE pumped storage, VTE wind, FVE solar. No oil,
  bioenergy, other or battery split. ERU leaves out licence-exempt plants above 100 kW not yet reported to the market
  operator (463.9 MW at 30 Jun 2026, mostly PV) so solar is a little understated. Reports from Q4 2020; the
  previous and current year are re-read every run.

SWITZERLAND - BFE / Pronovo plant register 'Elektrizitaetsproduktionsanlagen' (data.geo.admin.ch, free, updated
  monthly): every registered plant with BeginningOfOperation and power (kW). Month-end capacity is REBUILT as the
  sum of the plants in the CURRENT register whose operation began before the date. Limits: decommissioned plant is
  not in the register, so earlier months are understated by anything retired since (Muehleberg nuclear, closed Dec
  2019, is before the span); the register carries one power per plant, so later expansions (undated) are back-dated
  to the first start (Solar_initial_basis_MW is the sum of the power at first start for comparison; 6.5% below the
  current total); plants register with a lag (the last three rows are PROVISIONAL and rise when re-read). Small PV
  not registered with Pronovo is missing. ENTSO-E has no Swiss wind or solar. Pumped storage = plant category
  'Pumpspeicherkraftwerk'. Each run records the register's as-of value on 'Snapshots' (real-time view).
  The 19 MB register zip is re-downloaded only when its Last-Modified / ETag changes.

POLAND - URE (energy regulator) 'Instalacje odnawialnych zrodel energii' register, QUARTERLY (stan na 31 Mar / 30 Jun /
  30 Sep / 31 Dec, published 2-4 months later; this workbook holds every release URE still serves from Dec 2019 that
  parses): installed MW by renewable technology. PARTIAL and NOT a national total: it lists the plants in URE's
  concession / register scope and EXCLUDES prosumer micro-installations, so solar is about half of ENTSO-E's figure; wind,
  hydro and biogas are near-complete. 'Cofiring_WS_MW' (biomass co-firing in coal units) is the same plant counted
  again and is not in the totals. PSE (the TSO) publishes no installed capacity series (api.raporty.pse.pl entities
  checked); thermal plant has no monthly or quarterly source. Sheet is 'Quarterly' (one row per release, no months in
  between).

Unavailable (probed 12 Oct 2026, see discovery_archive/europe/capacity/): Slovakia (SEPS 'Installed power' is a Power BI of
ALLOCATED grid-connection capacity, not installed), Romania (ANRE / Transelectrica real-time installed power is a Power
BI whose public key no longer resolves), Bulgaria (ESO / EWRC annual only), Austria (E-Control / APG annual only).
Energy-Charts monthly installed power exists for Germany only.

Usage: python3 CEE_CAPACITY_MONTHLY.py [--out-dir DIR] [--only hu,cz,ch,pl]
"""
import argparse
import hashlib
import html
import io
import os
import re
import sys
import zipfile
from datetime import datetime, timezone

import numpy as np
import pandas as pd
import requests

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(os.path.dirname(HERE))
sys.path.insert(0, ROOT)
import xlsx_notes  # noqa: E402

OUT_DEFAULT = os.path.join(ROOT, "output", "Data and Chart Outputs")
SES = requests.Session()
SES.headers["User-Agent"] = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) "
                             "Chrome/124.0 Safari/537.36")
NOW = datetime.now(timezone.utc)
FUELS = ["Hydro", "Gas", "Wind", "Solar", "Coal", "Nuclear", "Oil", "Bioenergy", "Other"]
STD = [f"{f}_MW" for f in FUELS] + ["Total_MW", "PumpedStorage_MW", "Storage_MW"]
ENTSOE_COLS = {c: c for c in STD}


# ------------------------------------------------------------------------------------------------ helpers
def fetch(url, tries=4, **kw):
    last = None
    for i in range(tries):
        try:
            r = SES.get(url, timeout=(20, 240), **kw)
            if r.status_code in (429, 500, 502, 503, 504, 202):
                last = f"HTTP {r.status_code}"
                import time
                time.sleep(5 * (i + 1))
                continue
            return r
        except requests.RequestException as e:
            last = f"{type(e).__name__}: {e}"
            import time
            time.sleep(5 * (i + 1))
    raise RuntimeError(f"giving up on {url}: {last}")


def read_sheet(path, sheet, index_col=None):
    if not os.path.exists(path):
        return None
    try:
        return pd.read_excel(path, sheet_name=sheet, index_col=index_col)
    except Exception as e:  # noqa: BLE001
        print(f"  could not read {sheet} from {os.path.basename(path)}: {type(e).__name__}")
        return None


def load_monthly(path, sheet="Monthly"):
    d = read_sheet(path, sheet, index_col=0)
    if d is None or d.empty:
        return pd.DataFrame()
    d.index = pd.to_datetime(d.index, errors="coerce")
    d = d[d.index.notna()]
    d.index.name = "date"
    return d


def standardise(df, cols):
    """Index = 1st of month; keep `cols` (missing ones NaN), Total_MW = fuels except pumped storage / batteries."""
    out = pd.DataFrame(index=pd.DatetimeIndex(df.index, name="date"))
    for c in cols:
        out[c] = pd.to_numeric(df[c], errors="coerce") if c in df else np.nan
    fuel_cols = [f"{f}_MW" for f in FUELS if f"{f}_MW" in out]
    out["Total_MW"] = out[fuel_cols].sum(axis=1, min_count=1)
    return out.round(1).sort_index()


def merge_rows(old, new):
    """New rows replace old rows of the same date; older dates not in `new` are kept."""
    if old is None or old.empty:
        return new.sort_index()
    keep = old[~old.index.isin(new.index)]
    cols = list(dict.fromkeys(list(new.columns) + [c for c in old.columns if c not in new.columns]))
    return pd.concat([keep.reindex(columns=cols), new.reindex(columns=cols)]).sort_index()


def entsoe_check(out_dir, country, monthly, mapping):
    """1 January rows against ENTSO-E A68 ({our column: ENTSO-E column})."""
    p = os.path.join(out_dir, f"{country}_power_capacity.xlsx")
    e = read_sheet(p, "Monthly")
    if e is None or monthly.empty:
        return pd.DataFrame()
    e["date"] = pd.to_datetime(e["date"])
    e = e.set_index("date")
    rows = []
    for d in e.index:
        if d not in monthly.index:
            continue
        for ours, theirs in mapping.items():
            if ours not in monthly or theirs not in e:
                continue
            a, b = monthly.at[d, ours], e.at[d, theirs]
            if pd.isna(a) or pd.isna(b):
                continue
            rows.append({"date": d, "series": ours, "this_file_MW": round(float(a), 1), "ENTSOE_A68_MW": float(b),
                         "difference_MW": round(float(a) - float(b), 1),
                         "difference_pct": round((float(a) - float(b)) / float(b) * 100, 1) if b else np.nan})
    return pd.DataFrame(rows).set_index("date") if rows else pd.DataFrame()


def releases_frame(rows):
    return pd.DataFrame(rows, columns=["source", "checked_utc", "detail", "sha256"]).set_index("source")


def stored_release(path, source):
    r = read_sheet(path, "Releases", index_col=0)
    if r is None or source not in r.index:
        return ""
    return str(r.at[source, "detail"])


def now_s():
    return NOW.strftime("%Y-%m-%d %H:%M UTC")


def write(path, sheets, notes, titles):
    xlsx_notes.write_workbook(path, sheets, notes, titles)


# ------------------------------------------------------------------------------------------------ Hungary
HU_PAGE = "https://mavir.hu/web/mavir-en/energy-mix-power-plants-installed-capacity-data"
HU_MONTH = {"jan": 1, "feb": 2, "mar": 3, "apr": 4, "may": 5, "jun": 6, "jul": 7, "aug": 8, "sep": 9, "oct": 10,
            "nov": 11, "dec": 12}
HU_NAMES = ["Nuclear", "CoalLignite", "Oil", "Gas", "Storage", "ConvTotal", "Biogas", "Biomass", "Waste", "SolarInd",
            "WindInd", "Hydro", "Geothermal", "RenewTotal", "TotalExHMKE", "SolarSCTE", "SolarHMKE", "WindHMKE",
            "RenewOtherHMKE", "ConvOtherHMKE", "StorageHMKE", "TotalHMKE", "GrandTotal", "SolarTotal", "WeatherDep",
            "RenewGrand"]
NUM = re.compile(r"^-?\d+,\d$")


def hu_parse(content):
    import fitz
    doc = fitz.open(stream=content, filetype="pdf")
    rows, bad = {}, []
    for pi, pg in enumerate(doc):
        txt = pg.get_text()
        if "Gross Installed" not in txt or pi > 0:   # page 1 carries the gross table; later pages are net / charts
            continue
        words = pg.get_text("words")
        lines = {}
        for w in words:
            lines.setdefault(round(w[1] / 3), []).append(w)
        labels = []
        for k, ws in lines.items():
            ws = sorted(ws, key=lambda w: w[0])
            t = " ".join(w[4] for w in ws)
            m = re.match(r"^(\d{4})\.\s+\S+\s+I\s+([A-Za-z]+)\s+(\d{4})", t)
            if m and m.group(2)[:3].lower() in HU_MONTH:
                y = sum((w[1] + w[3]) / 2 for w in ws[:2]) / 2
                labels.append((y, int(m.group(3)), HU_MONTH[m.group(2)[:3].lower()]))
        if not labels:
            continue
        ys = np.array([l[0] for l in labels])
        bucket = {i: [] for i in range(len(labels))}
        for w in words:
            if NUM.match(w[4]):
                yc = (w[1] + w[3]) / 2
                i = int(np.argmin(np.abs(ys - yc)))
                if abs(ys[i] - yc) < 7:
                    bucket[i].append((w[0], float(w[4].replace(",", "."))))
        for i, (y, yr, mo) in enumerate(labels):
            vals = [v for _, v in sorted(bucket[i])]
            if len(vals) != len(HU_NAMES):
                bad.append((yr, mo, f"{len(vals)} values"))
                continue
            r = dict(zip(HU_NAMES, vals))
            tol = 0.35
            checks = [abs(sum(r[k] for k in ["Nuclear", "CoalLignite", "Oil", "Gas", "Storage"]) - r["ConvTotal"]) < tol,
                      abs(sum(r[k] for k in ["Biogas", "Biomass", "Waste", "SolarInd", "WindInd", "Hydro", "Geothermal"]) - r["RenewTotal"]) < tol,
                      abs(r["ConvTotal"] + r["RenewTotal"] - r["TotalExHMKE"]) < tol,
                      abs(r["TotalExHMKE"] + r["SolarSCTE"] + r["TotalHMKE"] - r["GrandTotal"]) < tol,
                      abs(r["SolarInd"] + r["SolarSCTE"] + r["SolarHMKE"] - r["SolarTotal"]) < tol]
            if not all(checks):
                bad.append((yr, mo, f"subtotals fail {checks}"))
                continue
            rows[pd.Timestamp(yr, mo, 1)] = r
    return rows, bad


def hu_frame(rows):
    """rows: {month start: parsed dict} -> standard layout dated the 1st of the FOLLOWING month (month-end stock)."""
    recs = {}
    for m, r in rows.items():
        d = m + pd.offsets.MonthBegin(1)
        recs[d] = {"Nuclear_MW": r["Nuclear"], "Coal_MW": r["CoalLignite"], "Oil_MW": r["Oil"], "Gas_MW": r["Gas"],
                   "Hydro_MW": r["Hydro"], "Wind_MW": r["WindInd"] + r["WindHMKE"], "Solar_MW": r["SolarTotal"],
                   "Bioenergy_MW": r["Biogas"] + r["Biomass"],
                   "Other_MW": r["Waste"] + r["Geothermal"] + r["RenewOtherHMKE"] + r["ConvOtherHMKE"],
                   "Storage_MW": r["Storage"] + r["StorageHMKE"], "PumpedStorage_MW": 0.0,
                   "Solar_industrial_MW": r["SolarInd"], "Solar_prosumer_SCTE_MW": r["SolarSCTE"],
                   "Solar_household_HMKE_MW": r["SolarHMKE"], "Gross_total_MW_MAVIR": r["GrandTotal"]}
    df = pd.DataFrame.from_dict(recs, orient="index")
    df.index.name = "date"
    return df


def run_hu(out_dir):
    path = os.path.join(out_dir, "hungary_power_capacity_monthly.xlsx")
    old = load_monthly(path)
    r = fetch(HU_PAGE)
    m = re.search(r'href="(/documents/10262/\d+/[^"]*GROSS\+NET\+INSTALLED[^"]*\.pdf[^"]*)"', r.text)
    if not m:
        raise RuntimeError("MAVIR installed-capacity PDF link not found")
    url = "https://mavir.hu" + html.unescape(m.group(1))
    head = SES.head(url, timeout=60, allow_redirects=True)
    detail = f"Last-Modified {head.headers.get('Last-Modified')}; size {head.headers.get('Content-Length')}"
    sha, content = "", None
    if old.empty or stored_release(path, "MAVIR PDF") != detail:
        content = fetch(url).content
        sha = hashlib.sha256(content).hexdigest()
    else:
        print("  MAVIR PDF unchanged since the last run; stored months kept")
    if content is not None:
        rows, bad = hu_parse(content)
        for b in bad:
            print("  dropped row", b)
        new = hu_frame(rows)
        monthly = merge_rows(old, new)
        print(f"  Hungary parsed {len(rows)} rows, {len(bad)} dropped")
    else:
        monthly = old
    monthly = monthly[monthly.index >= "2020-02-01"]
    keep_extra = [c for c in monthly.columns if c not in STD]
    monthly = pd.concat([standardise(monthly, STD), monthly[keep_extra]], axis=1)
    chk = entsoe_check(out_dir, "hungary", monthly, {c: c for c in STD if c not in ("Storage_MW", "PumpedStorage_MW")})
    if content is None:
        ro = read_sheet(path, "Releases", index_col=0)
        sha = str(ro.at["MAVIR PDF", "sha256"]) if ro is not None and "MAVIR PDF" in ro.index else ""
    rel = [("MAVIR PDF", now_s(), detail, sha)]
    sheets = {"Monthly": monthly}
    if len(chk):
        sheets["Check_ENTSOE"] = chk
    sheets["Releases"] = releases_frame([(a, b, c, d) for a, b, c, d in rel])
    notes = ["Hungary - monthly installed (gross) generation capacity by fuel, MW, MAVIR (TSO).",
             "",
             "Source", "MAVIR 'Energy MIX - Power Plants - Installed Capacity Data': PDF 'Gross / net installed capacity per production type MAVIR 2000-2020-2026' (" + HU_PAGE + ").",
             "Monthly from Jan 2020 (annual before); the row dated 1 Mar holds the capacity at the end of February (same convention as the ENTSO-E annual files).",
             "Gross capacity of power plants plus household-size plants (HMKE) plus small-scale prosumer plants (SCTE): rooftop solar is included.",
             "", "Fuels",
             "Coal_MW = coal + lignite; Gas_MW = natural gas; Storage_MW = batteries (MAVIR 'Tarolok'); Wind = industrial + HMKE; Solar = industrial + prosumer (SCTE) + HMKE;",
             "Bioenergy_MW = biogas + biomass; Other_MW = waste + geothermal + other HMKE. Total_MW excludes storage. Solar split columns kept at the right.",
             "", "Method",
             "The PDF is parsed from word positions; every month is validated against the PDF's own subtotals (conventional, renewable, totals, solar) and dropped if they do not add up.",
             "The PDF is downloaded only when its Last-Modified / size changes (sheet Releases); stored months are kept and replaced by the latest PDF figures."]
    write(path, sheets, notes, ["Source", "Fuels", "Method"])
    print(f"Hungary: {len(monthly)} months {monthly.index.min().date()} - {monthly.index.max().date()}")
    return monthly


# ------------------------------------------------------------------------------------------------ Czechia
CZ_ROM = {1: "i", 2: "ii", 3: "iii", 4: "iv"}
CZ_TECH = {"JE": "Nuclear_MW", "PE": "Steam_all_fuels_MW", "PPE": "ccgt", "PSE": "gt", "VE": "Hydro_MW", "PVE": "PumpedStorage_MW",
           "VTE": "Wind_MW", "FVE": "Solar_MW"}


def cz_slug(year, q):
    return f"ctvrtletni-zprava-o-provozu-elektrizacni-soustavy-cr-za-{CZ_ROM[q]}-ctvrtleti-{year}"


def cz_parse(content, year):
    """Sheet '5' of an ERU quarterly workbook -> {month start: {column: MW}} for months with data."""
    xl = pd.ExcelFile(io.BytesIO(content))
    sh = next((s for s in xl.sheet_names if str(s).strip() == "5"), None)
    if sh is None:
        raise RuntimeError("sheet 5 not found")
    d = xl.parse(sh, header=None)
    out = {}
    tech_rows = {}
    for i in range(len(d)):
        lab = str(d.iloc[i, 0])
        m = re.search(r"\((JE|PE|PPE|PSE|VE|PVE|VTE|FVE)\)", lab)
        if m and m.group(1) not in tech_rows and i < 25:
            tech_rows[m.group(1)] = i
    if len(tech_rows) != 8:
        raise RuntimeError(f"technology rows found: {sorted(tech_rows)}")
    for mo in range(1, 13):
        vals = {k: pd.to_numeric(d.iloc[i, mo], errors="coerce") for k, i in tech_rows.items()}
        if any(pd.isna(v) for v in vals.values()) or sum(float(v) for v in vals.values()) <= 0:
            continue
        out[pd.Timestamp(year, mo, 1)] = {CZ_TECH[k]: float(v) for k, v in vals.items()}
    return out


def cz_frame(rows):
    recs = {}
    for m, r in rows.items():
        d = m + pd.offsets.MonthBegin(1)
        recs[d] = {"Nuclear_MW": r["Nuclear_MW"], "Coal_MW": r["Steam_all_fuels_MW"], "Steam_all_fuels_MW": r["Steam_all_fuels_MW"],
                   "Gas_MW": r["ccgt"] + r["gt"], "Gas_ccgt_MW": r["ccgt"], "Gas_turbines_engines_MW": r["gt"],
                   "Hydro_MW": r["Hydro_MW"], "PumpedStorage_MW": r["PumpedStorage_MW"], "Wind_MW": r["Wind_MW"], "Solar_MW": r["Solar_MW"]}
    df = pd.DataFrame.from_dict(recs, orient="index")
    df.index.name = "date"
    return df


def run_cz(out_dir):
    path = os.path.join(out_dir, "czechia_power_capacity_monthly.xlsx")
    old = load_monthly(path)
    rel_old = read_sheet(path, "Releases", index_col=0)
    rel_rows, new_rows = [], {}
    this_year = NOW.year
    for year in range(2020, this_year + 1):
        complete = (not old.empty) and pd.Timestamp(year + 1, 1, 1) in old.index
        if complete and year < this_year - 1:
            continue
        for q in (4, 3, 2, 1):
            slug = cz_slug(year, q)
            r = SES.get("https://eru.gov.cz/" + slug, timeout=(20, 120))
            if r.status_code != 200:
                continue
            fl = sorted(set(re.findall(r'href=["\']([^"\']+\.xlsx)["\']', r.text, re.I)))
            if not fl:
                continue
            url = "https://eru.gov.cz" + fl[0] if fl[0].startswith("/") else fl[0]
            head = SES.head(url, timeout=60, allow_redirects=True)
            detail = f"Last-Modified {head.headers.get('Last-Modified')}; size {head.headers.get('Content-Length')}"
            prev = str(rel_old.at[slug, "detail"]) if rel_old is not None and slug in rel_old.index else ""
            sha = str(rel_old.at[slug, "sha256"]) if rel_old is not None and slug in rel_old.index else ""
            if prev != detail or old.empty:
                try:
                    content = fetch(url).content
                    sha = hashlib.sha256(content).hexdigest()
                    new_rows.update(cz_parse(content, year))
                    print(f"  Czechia {year} Q{q}: {url.rsplit('/', 1)[-1]} parsed")
                except Exception as e:  # noqa: BLE001
                    print(f"  Czechia {year} Q{q}: {type(e).__name__}: {e}")
                    detail = f"parse failed: {e}"
            rel_rows.append((slug, now_s(), detail, sha))
            break
    new = cz_frame(new_rows) if new_rows else pd.DataFrame()
    monthly = merge_rows(old, new) if len(new) else old
    if monthly.empty:
        raise RuntimeError("no Czech data")
    extra = [c for c in monthly.columns if c not in STD]
    monthly = pd.concat([standardise(monthly, STD), monthly[extra]], axis=1)
    monthly = monthly[monthly.index >= "2021-01-01"]
    # the series starts with the December 2020 stock in the row dated 1 Jan 2021 (from the 2020 reports)
    chk = entsoe_check(out_dir, "czechia", monthly, {"Nuclear_MW": "Nuclear_MW", "Hydro_MW": "Hydro_MW", "Wind_MW": "Wind_MW",
                                                      "Solar_MW": "Solar_MW", "Gas_MW": "Gas_MW", "Coal_MW": "Coal_MW",
                                                      "PumpedStorage_MW": "PumpedStorage_MW"})
    if rel_old is not None:
        keep = [(i, r["checked_utc"], r["detail"], r["sha256"]) for i, r in rel_old.iterrows() if i not in {x[0] for x in rel_rows}]
        rel_rows = keep + rel_rows
    sheets = {"Monthly": monthly}
    if len(chk):
        sheets["Check_ENTSOE"] = chk
    sheets["Releases"] = releases_frame(rel_rows)
    notes = ["Czechia - monthly installed generation capacity by technology, MW, ERU (energy regulator).",
             "",
             "Source", "ERU quarterly 'Zprava o provozu elektrizacni soustavy CR', sheet 5 of the xlsx attachment (https://eru.gov.cz/zpravy-o-provozu): installed capacity by technology for each month of the year up to the report's quarter.",
             "The row dated 1 Mar holds the capacity at the end of February (same convention as the ENTSO-E annual files). The Q4 report of a year holds all twelve months; the current and previous year are re-read each run.",
             "", "Technologies (ERU's own categories)",
             "Nuclear_MW = JE; Coal_MW = Steam_all_fuels_MW = PE (steam plant: coal, lignite AND biomass steam units, not split by fuel); Gas_MW = PPE combined cycle + PSE gas / combustion turbines and engines (on any fuel);",
             "Hydro_MW = VE; PumpedStorage_MW = PVE; Wind_MW = VTE; Solar_MW = FVE. No oil, bioenergy, other or battery split at source (blank, not zero). Total_MW = sum of the available fuels, excluding pumped storage.",
             "ERU excludes licence-exempt plants above 100 kW not yet reported to the market operator (463.9 MW at 30 Jun 2026, 453.3 MW of it PV): solar is slightly understated.",
             "", "Method",
             "Reports are whole files: a report is read again only when its Last-Modified / size changes (sheet Releases). Months a report has not reached yet are blank in its sheet and are skipped."]
    write(path, sheets, notes, ["Source", "Technologies (ERU's own categories)", "Method"])
    print(f"Czechia: {len(monthly)} months {monthly.index.min().date()} - {monthly.index.max().date()}")
    return monthly


# ------------------------------------------------------------------------------------------------ Switzerland
CH_URL = "https://data.geo.admin.ch/ch.bfe.elektrizitaetsproduktionsanlagen/csv/2056/ch.bfe.elektrizitaetsproduktionsanlagen.zip"
CH_SUB = {"subcat_1": "Hydro", "subcat_2": "Solar", "subcat_3": "Wind", "subcat_4": "Bioenergy", "subcat_5": "Other",
          "subcat_6": "Nuclear", "subcat_7": "Oil", "subcat_8": "Gas", "subcat_9": "Coal", "subcat_10": "Other"}


def ch_rebuild(zbytes):
    z = zipfile.ZipFile(io.BytesIO(zbytes))
    pl = pd.read_csv(z.open("ElectricityProductionPlant.csv"))
    pl["start"] = pd.to_datetime(pl["BeginningOfOperation"], errors="coerce")
    pl = pl.dropna(subset=["start"])
    pl["kW"] = pd.to_numeric(pl["TotalPower"], errors="coerce").fillna(0.0)
    pl["kW0"] = pd.to_numeric(pl["InitialPower"], errors="coerce").fillna(pl["kW"])
    pl["fuel"] = pl["SubCategory"].map(CH_SUB).fillna("Other")
    pumped = pl["PlantCategory"].eq("plantcat_6")
    pl.loc[pumped, "fuel"] = "PumpedStorage"
    first = pd.Timestamp("2021-01-01")
    last_start = pl["start"].max()
    last_row = pd.Timestamp(last_start.year, last_start.month, 1) + pd.offsets.MonthBegin(1)
    dates = pd.date_range(first, last_row, freq="MS")
    rec = {}
    for d in dates:
        s = pl[pl["start"] < d]
        by = s.groupby("fuel")["kW"].sum() / 1000.0
        rec[d] = {f"{f}_MW": float(by.get(f, 0.0)) for f in FUELS + ["PumpedStorage"]}
        rec[d]["Solar_initial_basis_MW"] = float(s.loc[s["fuel"] == "Solar", "kW0"].sum() / 1000.0)
    df = pd.DataFrame.from_dict(rec, orient="index")
    df.index.name = "date"
    df["Hydro_incl_pumped_MW"] = df["Hydro_MW"] + df["PumpedStorage_MW"]
    return df, last_start, len(pl)


def run_ch(out_dir):
    path = os.path.join(out_dir, "switzerland_power_capacity_monthly.xlsx")
    old = load_monthly(path)
    head = SES.head(CH_URL, timeout=60, allow_redirects=True)
    detail = f"Last-Modified {head.headers.get('Last-Modified')}; ETag {head.headers.get('ETag')}; size {head.headers.get('Content-Length')}"
    prev = stored_release(path, "BFE plant register zip")
    snaps = read_sheet(path, "Snapshots", index_col=0)
    if not old.empty and prev == detail:
        print("  Swiss register unchanged since the last run; stored months kept")
        return old
    content = fetch(CH_URL).content
    sha = hashlib.sha256(content).hexdigest()
    df, last_start, nplants = ch_rebuild(content)
    monthly = standardise(df, STD)
    monthly = pd.concat([monthly, df[["Solar_initial_basis_MW", "Hydro_incl_pumped_MW"]]], axis=1)
    monthly["provisional"] = False
    monthly.iloc[-3:, monthly.columns.get_loc("provisional")] = True
    monthly["PumpedStorage_MW"] = df["PumpedStorage_MW"].round(1)
    monthly["Storage_MW"] = np.nan
    snap_row = {"asof_utc": now_s(), "register_last_modified": head.headers.get("Last-Modified"), "latest_start_date": str(last_start.date()),
                "plants": nplants}
    snap_row.update({f"{c}": monthly.iloc[-1][c] for c in ["Hydro_MW", "PumpedStorage_MW", "Wind_MW", "Solar_MW", "Nuclear_MW", "Gas_MW", "Bioenergy_MW", "Other_MW"]})
    snap_row["row_date"] = str(monthly.index[-1].date())
    snap = pd.DataFrame([snap_row]).set_index("asof_utc")
    if snaps is not None and not snaps.empty:
        snap = pd.concat([snaps, snap[~snap.index.isin(snaps.index)]])
    chk = entsoe_check(out_dir, "switzerland", monthly, {"Hydro_MW": "Hydro_MW", "PumpedStorage_MW": "PumpedStorage_MW", "Nuclear_MW": "Nuclear_MW"})
    notes = ["Switzerland - monthly installed generation capacity by fuel, MW, rebuilt from the BFE / Pronovo plant register.",
             "",
             "Source", "BFE 'Elektrizitaetsproduktionsanlagen' (Pronovo register; data.geo.admin.ch, ch.bfe.elektrizitaetsproduktionsanlagen), CSV zip updated monthly: " + CH_URL,
             "Row dated 1 Mar = capacity at the end of February: the sum of the plants in the CURRENT register whose BeginningOfOperation is before that date (same date convention as the ENTSO-E annual files).",
             "", "Limits",
             "Decommissioned plant is not in the register, so earlier months are understated by anything retired since. The register has one (current) power per plant: later expansions are back-dated to the first start (Solar_initial_basis_MW = power at first start, ~6.5% below the current total).",
             "Plants register with a lag: the last three rows are PROVISIONAL (column provisional) and rise when re-read. PV not registered with Pronovo is missing. ENTSO-E A68 has no Swiss wind or solar.",
             "Hydro_MW excludes pumped-storage plants (plant category 'Pumpspeicherkraftwerk', PumpedStorage_MW); Hydro_incl_pumped_MW is their sum. Other_MW = waste incineration (KVA) and geothermal. Total_MW excludes pumped storage.",
             "", "Method",
             "Each time the zip changes (Last-Modified / ETag, sheet Releases) it is downloaded and the whole series is rebuilt; 'Snapshots' keeps what each run's register said for its latest row (real-time view, since a rebuild revises history)."]
    sheets = {"Monthly": monthly}
    if len(chk):
        sheets["Check_ENTSOE"] = chk
    sheets["Snapshots"] = snap
    sheets["Releases"] = releases_frame([("BFE plant register zip", now_s(), detail, sha)])
    write(path, sheets, notes, ["Source", "Limits", "Method"])
    print(f"Switzerland: {len(monthly)} months {monthly.index.min().date()} - {monthly.index.max().date()}, {nplants} plants")
    return monthly


# ------------------------------------------------------------------------------------------------ Poland
PL_INDEX = "https://www.ure.gov.pl/pl/oze/potencjal-krajowy-oze"
PL_PAGE_FALLBACK = "https://www.ure.gov.pl/pl/oze/potencjal-krajowy-oze/8108,Instalacje-odnawialnych-zrodel-energii-stan-na-30-czerwca-2026-r.html"
PL_MONTHS = {"stycznia": 1, "lutego": 2, "marca": 3, "kwietnia": 4, "maja": 5, "czerwca": 6, "lipca": 7, "sierpnia": 8,
             "wrzesnia": 9, "września": 9, "pazdziernika": 10, "października": 10, "listopada": 11, "grudnia": 12}
PL_CODE = {"PV": "Solar_MW", "PVA": "Solar_MW", "WI": "Wind_MW", "WIL": "Wind_MW", "WO": "Hydro_MW", "BG": "Biogas_MW", "BM": "Biomass_MW",
           "ITPO": "Other_MW", "H2": "Other_MW", "WS": "Cofiring_WS_MW"}


def pl_parse(content):
    xl = pd.ExcelFile(io.BytesIO(content))
    d = xl.parse(xl.sheet_names[0], header=None)
    hdr = None
    for i in range(min(12, len(d))):
        row = [str(v) for v in d.iloc[i].tolist()]
        if any("rodzaj" in v.lower() for v in row):
            hdr = i
            break
    if hdr is None:
        raise RuntimeError("header row not found")
    row = [str(v) for v in d.iloc[hdr].tolist()]
    ck = next(k for k, v in enumerate(row) if "rodzaj" in v.lower())
    ci = next(k for k, v in enumerate(row) if "moc" in v.lower())
    dd = d.iloc[hdr + 1:]
    mw = pd.to_numeric(dd[ci].astype(str).str.replace(",", ".").str.replace(r"[^0-9.]", "", regex=True), errors="coerce")
    code = dd[ck].astype(str).str.strip().str.upper()
    sums = mw.groupby(code).sum()
    out, unknown = {}, {}
    for c, v in sums.items():
        if c in PL_CODE:
            out[PL_CODE[c]] = out.get(PL_CODE[c], 0.0) + float(v)
        elif c not in ("NAN", "") and v:
            unknown[c] = round(float(v), 1)
    return out, unknown


def run_pl(out_dir):
    path = os.path.join(out_dir, "poland_power_capacity_quarterly.xlsx")
    old = load_monthly(path, "Quarterly")
    rel_old = read_sheet(path, "Releases", index_col=0)
    r = fetch(PL_INDEX)
    m = re.search(r'href="(/pl/oze/potencjal-krajowy-oze/8108,[^"]+)"', r.text)
    page = "https://www.ure.gov.pl" + html.unescape(m.group(1)) if m else PL_PAGE_FALLBACK
    r = fetch(page)
    rels = re.findall(r'href="(/download/9/\d+/[^"]+)">\s*<strong>\s*([^<]+?)\s*</strong>', r.text)
    print(f"  Poland: {len(rels)} releases listed")
    rows, rel_rows, unknown_all = {}, [], {}
    done = set(old.index) if not old.empty else set()
    for k, (href, lab) in enumerate(rels):
        mm = re.search(r"(\d{1,2})\s+([a-zżźćńółęąś]+)\s+(\d{4})", lab.lower())
        if not mm or mm.group(2) not in PL_MONTHS:
            continue
        mo, yr = PL_MONTHS[mm.group(2)], int(mm.group(3))
        if (yr, mo) < (2019, 12):
            continue
        d = pd.Timestamp(yr, mo, 1)
        latest = k < 2
        if d in done and not latest:
            continue
        url = "https://www.ure.gov.pl" + href
        try:
            x = fetch(url)
            if x.status_code != 200:
                raise RuntimeError(f"HTTP {x.status_code}")
            vals, unk = pl_parse(x.content)
            rows[d] = vals
            rel_rows.append((lab, now_s(), f"{url}; size {len(x.content)}", hashlib.sha256(x.content).hexdigest()))
            if unk:
                unknown_all[lab] = unk
            print(f"  Poland {lab}: {', '.join(f'{a} {b:.0f}' for a, b in vals.items())}")
        except Exception as e:  # noqa: BLE001
            print(f"  Poland {lab}: {type(e).__name__}: {e}")
            rel_rows.append((lab, now_s(), f"failed: {type(e).__name__}: {e}", ""))
    new = pd.DataFrame.from_dict(rows, orient="index")
    if len(new):
        # date = first of the month AFTER the quarter-end month (stock at 00:00), like the monthly files
        new.index = [i + pd.offsets.MonthBegin(1) for i in new.index]
    monthly = merge_rows(old, new)
    if monthly.empty:
        raise RuntimeError("no Polish data")
    for c in ["Biogas_MW", "Biomass_MW", "Cofiring_WS_MW"]:
        if c not in monthly:
            monthly[c] = np.nan
    monthly["Bioenergy_MW"] = monthly["Biogas_MW"].fillna(0) + monthly["Biomass_MW"].fillna(0)
    for c in ["Hydro_MW", "Wind_MW", "Solar_MW", "Other_MW"]:
        if c not in monthly:
            monthly[c] = np.nan
    monthly["Total_MW"] = monthly[["Hydro_MW", "Wind_MW", "Solar_MW", "Bioenergy_MW", "Other_MW"]].sum(axis=1, min_count=1)
    cols = ["Hydro_MW", "Wind_MW", "Solar_MW", "Bioenergy_MW", "Other_MW", "Total_MW", "Biogas_MW", "Biomass_MW", "Cofiring_WS_MW"]
    monthly = monthly[cols].round(1).sort_index()
    monthly.index.name = "date"
    chk = entsoe_check(out_dir, "poland", monthly, {"Hydro_MW": "Hydro_MW", "Wind_MW": "Wind_MW", "Solar_MW": "Solar_MW", "Bioenergy_MW": "Bioenergy_MW"})
    if rel_old is not None:
        keep = [(i, r_["checked_utc"], r_["detail"], r_["sha256"]) for i, r_ in rel_old.iterrows() if i not in {x_[0] for x_ in rel_rows}]
        rel_rows = keep + rel_rows
    notes = ["Poland - QUARTERLY installed renewable capacity by technology, MW, URE register (PARTIAL: not a national total).",
             "",
             "Source", "URE 'Instalacje odnawialnych zrodel energii - stan na <date>' xlsx releases (" + PL_INDEX + "): the plants in URE's concession / register scope with installed electrical power and technology code.",
             "One row per release; the row dated 1 Jan holds the stock at 31 December, 1 Apr = 31 March, 1 Jul = 30 June, 1 Oct = 30 September (same convention as ENTSO-E). No month between releases is shown or interpolated.",
             "", "Limits",
             "Prosumer micro-installations are NOT in the register, so Solar_MW is about half of ENTSO-E's figure; wind, hydro and biogas are close to ENTSO-E. Cofiring_WS_MW (biomass co-firing in coal units) counts the same plant again and is NOT in Total_MW.",
             "Codes: PV/PVA solar, WI/WIL wind, WO hydro, BG biogas, BM biomass, ITPO waste thermal treatment and H2 (Other), WS co-firing. Thermal plant has no monthly or quarterly source (PSE publishes none).",
             "Releases before Dec 2019 are not read. Older layouts differ (columns); the parser finds the technology and power columns by header.",
             "", "Method",
             "Releases already stored are not downloaded again except the two newest (revisions); sheet Releases lists every file and its hash."]
    sheets = {"Quarterly": monthly}
    if len(chk):
        sheets["Check_ENTSOE"] = chk
    sheets["Releases"] = releases_frame(rel_rows)
    write(path, sheets, notes, ["Source", "Limits", "Method"])
    print(f"Poland: {len(monthly)} releases {monthly.index.min().date()} - {monthly.index.max().date()}; unknown codes {unknown_all}")
    return monthly


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out-dir", default=OUT_DEFAULT)
    ap.add_argument("--only", default="hu,cz,ch,pl")
    a = ap.parse_args()
    os.makedirs(a.out_dir, exist_ok=True)
    failed = []
    for key, fn in (("hu", run_hu), ("cz", run_cz), ("ch", run_ch), ("pl", run_pl)):
        if key not in a.only.split(","):
            continue
        try:
            print(f"== {key}")
            fn(a.out_dir)
        except Exception as e:  # noqa: BLE001
            import traceback
            traceback.print_exc()
            failed.append(key)
    if failed:
        print("FAILED:", failed)
        sys.exit(1)


if __name__ == "__main__":
    main()
