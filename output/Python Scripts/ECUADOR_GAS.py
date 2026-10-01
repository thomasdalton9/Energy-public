"""
Ecuador natural gas: production and use by sector, monthly from 2021,
from EP Petroecuador's statistical reports (public PDFs):
  https://www.eppetroecuador.ec/?p=3721  ("Cifras Institucionales")
    - Informes Estadisticos Mensuales: current year to date (one PDF)
    - Informes Estadisticos Anuales: one PDF per past year (Jan-Dec)

Ecuador's only commercial gas is the offshore Amistad field (Gulf of
Guayaquil). Petroecuador's dispatch table splits its sales into:
  GAS NATURAL (MMBTU)          - pipeline gas to the Termogas Machala
                                 power plant  -> Power
  GAS NATURAL LICUADO (MMBTU)  - LNG from the Bajo Alto plant, trucked
                                 to industry (mostly Cuenca ceramics) -> Industrial
plus Amistad production in barrels of oil equivalent per day.

The PDF text layer splits numbers with stray spaces ("5 93.016"), so
values are read by x-position: each number fragment is assigned to the
nearest month column of the table header, then the fragments are joined.

Units: MMBtu per day (monthly dispatch / days in month); also approx.
million cubic feet per day at 1,030 Btu/ft3. Production BOE/d is
converted at 5.8 MMBtu/BOE.

IMPORTS -> "Total demand" sheet. Ecuador has no LNG terminal and EP
Petroecuador imports no natural gas (its imported derivatives are naphtha,
diesel and LPG - BCE "Reporte del Sector Petrolero"; its 2023 gas import
tender was declared void; the 2024-26 Karpowership barges burn fuel oil).
The only gas imports are small private LNG cargoes in ISO containers for
industry (Sycar: from Panama / AES Colon from Jan 2022, overland from Peru /
Limagas from Nov 2024), recorded by customs under HS 2711.11. They come
from UN Comtrade (Ecuador's own customs returns, reporter 218), public
preview API, one month per call:
  https://comtradeapi.un.org/public/v1/preview/C/M/HS?reporterCode=218&period=YYYYMM&cmdCode=2711,271111,271121&flowCode=M
A month with a 2711 (petroleum gases, mostly LPG) row but no 271111 /
271121 row is a published month with zero gas imports; a month with no
rows at all is not (yet) published monthly. Years without monthly data
get Comtrade's annual figure in the annual table and notes only - it is
not spread over months. Found via ECUADOR_GAS_IMPORTS_DISCOVERY(2).py.
Not reachable from the editing sandbox; runs in GitHub Actions.
"""
import argparse
import io
import os
import re
import sys
import time
from urllib.parse import urljoin

import pandas as pd
import pdfplumber
import requests

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))  # repo root, for xlsx_notes
import xlsx_notes

PAGE = "https://www.eppetroecuador.ec/?p=3721"
H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0 Safari/537.36"}
T = (10, 180)
DATA_START = 2021
MONTHS = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto", "septiembre", "octubre",
          "noviembre", "diciembre"]
BTU_PER_FT3 = 1030.0
MMBTU_PER_BOE = 5.8
MMBTU_PER_MCM = 35.3147e6 * BTU_PER_FT3 / 1e6  # million cubic metres -> MMBtu (36,374)
MMBTU_PER_T_LNG = 52.0  # GIIGNL rule of thumb, gross calorific value
COMTRADE = "https://comtradeapi.un.org/public/v1/preview/C/{freq}/HS"
GAS_HS = ("271111", "271121")
CENACE_XLSX = "output/Data and Chart Outputs/ecuador_power_generation_daily.xlsx"


def out(*a):
    print(*a, flush=True)


def list_reports():
    """(label, url) for the current monthly report and each annual report
    from DATA_START on, from the 'Informes Estadisticos' blocks."""
    r = requests.get(PAGE, headers=H, timeout=T)
    r.raise_for_status()
    html = r.text
    i = html.find("Mensuales")
    j = html.find("Exploraci", i)
    block = html[i:j] if i > 0 and j > i else html
    reports = []
    for href, txt in re.findall(r'<a[^>]+href="([^"]+)"[^>]*>(.*?)</a>', block, re.S | re.I):
        t = re.sub(r"<[^>]+>|\s+", " ", txt).replace("&#8211;", "-").strip()
        years = [int(y) for y in re.findall(r"(20\d\d)", t)]
        if not years or max(years) < DATA_START:
            continue
        reports.append((t, urljoin(PAGE, href.replace("&amp;", "&"))))
    return reports


def es_number(s):
    s = s.replace(" ", "")
    if s in ("", "-", "--"):
        return 0.0
    if "," in s:
        s = s.replace(".", "").replace(",", ".")
    else:
        s = s.replace(".", "")
    try:
        return float(s)
    except ValueError:
        return None


def month_columns(words):
    """[(month number, x centre)] from the table header line."""
    by_top = {}
    for w in words:
        t = w["text"].lower().strip(".")
        if t in MONTHS:
            by_top.setdefault(round(w["top"]), []).append((MONTHS.index(t) + 1, (w["x0"] + w["x1"]) / 2))
    if not by_top:
        return []
    best = max(by_top.values(), key=len)
    return sorted(best, key=lambda mc: mc[1])


def parse_report(content):
    disp, prod = {}, {}
    with pdfplumber.open(io.BytesIO(content)) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ""
            if "GAS" not in text.upper():
                continue
            ym = re.search(r"A[ÑN]O\s+(20\d\d)", text, re.I) or re.search(r"(20\d\d)", text)
            if not ym:
                continue
            year = int(ym.group(1))
            words = page.extract_words(keep_blank_chars=False, x_tolerance=1.5)
            cols = month_columns(words)
            if len(cols) < 1:
                continue
            per_day = re.search(r"barriles\s*/\s*d[ií]a", text, re.I)
            volumes = re.search(r"Cifras en barriles(?!\s*/)", text, re.I) and not re.search(r"US\s*\$", text)
            if volumes and not disp and re.search(r"GAS NATURAL\s*\(MMBTU\)", text):
                g, l = None, None
                for ln_word in [w for w in words if w["text"].upper() == "GAS"]:
                    rest = " ".join(w["text"] for w in words if abs(w["top"] - ln_word["top"]) < 3
                                    and w["x0"] >= ln_word["x0"])[:40].upper()
                    if rest.startswith("GAS NATURAL LICUADO") and l is None:
                        l = _row_at(words, ln_word, cols, year)
                    elif rest.startswith("GAS NATURAL (MMBTU)") and g is None:
                        g = _row_at(words, ln_word, cols, year)
                if g:
                    disp["Power_MMBtu_month"] = g
                if l:
                    disp["Industrial_LNG_MMBtu_month"] = l
            if per_day and not prod and re.search(r"GAS CAMPO AMISTAD", text, re.I):
                for w in words:
                    if w["text"].upper() == "GAS":
                        rest = " ".join(x["text"] for x in words if abs(x["top"] - w["top"]) < 3
                                        and x["x0"] >= w["x0"])[:30].upper()
                        if rest.startswith("GAS CAMPO AMISTAD"):
                            prod["Amistad_production_BOE_per_day"] = _row_at(words, w, cols, year)
                            break
    return disp, prod


def _row_at(words, first_word, cols, year):
    """Row values to the right of a row label that starts at first_word."""
    top = first_word["top"]
    label_end = max(w["x1"] for w in words if abs(w["top"] - top) < 3 and not re.fullmatch(r"[\d.,\-]+", w["text"])
                    and w["x0"] >= first_word["x0"] and w["x0"] < cols[0][1] - 30)
    line = [w for w in words if abs(w["top"] - top) < 3 and w["x0"] > label_end]
    xs = [x for _, x in cols]
    spacing = (xs[-1] - xs[0]) / max(len(xs) - 1, 1) if len(xs) > 1 else 40
    frag = {m: [] for m, _ in cols}
    for w in sorted(line, key=lambda w: w["x0"]):
        if not re.fullmatch(r"[\d.,\-]+", w["text"]):
            continue
        xc = (w["x0"] + w["x1"]) / 2
        m, x = min(cols, key=lambda mc: abs(mc[1] - xc))
        if abs(x - xc) <= spacing * 0.55:
            frag[m].append(w["text"])
    vals = {}
    for m, parts in frag.items():
        if parts:
            v = es_number("".join(parts))
            if v is not None:
                vals[pd.Timestamp(year=year, month=m, day=1)] = v
    return vals


# ------------------------------------------------------------------ imports (UN Comtrade)
IMPORT_COLS = ["Month", "Comtrade_status", "LNG_271111_tonnes", "LNG_271111_USD", "NG_gaseous_271121_tonnes",
               "NG_gaseous_271121_USD", "Partners"]


def comtrade(freq, period):
    """Ecuador's imports of HS 2711 / 271111 / 271121 for one period: a list
    of rows ([] = nothing published), or None if the request failed."""
    params = {"reporterCode": 218, "period": period, "cmdCode": "2711," + ",".join(GAS_HS), "flowCode": "M",
              "includeDesc": "true"}
    for attempt in range(5):
        try:
            r = requests.get(COMTRADE.format(freq=freq), params=params, headers=H, timeout=(10, 60))
        except requests.RequestException as e:
            out(f"  comtrade {period}: {type(e).__name__}")
            time.sleep(5)
            continue
        if r.status_code == 429:
            time.sleep(6 * (attempt + 1))
            continue
        time.sleep(1.2)
        if r.status_code != 200:
            out(f"  comtrade {period}: HTTP {r.status_code} {r.text[:120]}")
            return None
        return r.json().get("data") or []
    return None


def _plain(x):
    """Totals rows only (no mode-of-transport / customs-procedure / second-partner breakdowns)."""
    return x.get("motCode") in (0, None) and x.get("customsCode") in ("C00", None) and x.get("partner2Code") in (0, None)


def summarise(rows):
    """One record (tonnes, USD, partners) from a period's Comtrade rows."""
    rec = {"Reported": any(x.get("cmdCode") == "2711" and x.get("partnerCode") == 0 and _plain(x) for x in rows)}
    partners = {}
    for code, col in (("271111", "LNG_271111"), ("271121", "NG_gaseous_271121")):
        w = [x for x in rows if x.get("cmdCode") == code and x.get("partnerCode") == 0 and _plain(x)]
        by_partner = 0.0
        for x in rows:
            if x.get("cmdCode") == code and x.get("partnerCode") not in (0, None) and _plain(x):
                name = x.get("partnerDesc") or str(x.get("partnerCode"))
                partners[name] = partners.get(name, 0) + (x.get("netWgt") or 0) / 1000
                by_partner += (x.get("netWgt") or 0) / 1000
        # the World row sometimes lacks a weight (e.g. 2025) while the partner rows carry it
        world = (w[0].get("netWgt") or 0) / 1000 if w else 0.0
        rec[f"{col}_tonnes"] = round(max(world, by_partner), 3)
        rec[f"{col}_USD"] = round(w[0].get("primaryValue") or 0, 2) if w else 0.0
        if w:
            rec["Reported"] = True
    rec["Partners"] = "; ".join(f"{k} {v:,.1f} t" for k, v in sorted(partners.items(), key=lambda kv: -kv[1])
                                if v >= 0.05)
    return rec


def pull_imports(have_m, have_a, last_month):
    """Monthly Comtrade records from DATA_START to last_month, and annual
    records for each full year. Incremental: months already published are
    kept (the latest three are re-read for revisions); only unpublished
    months are asked for again. Annual records are re-read until all 12
    months of that year are published monthly."""
    done = {}
    for _, row in have_m.iterrows():
        if row.get("Comtrade_status") == "monthly":
            done[str(row["Month"])] = {c: row.get(c) for c in IMPORT_COLS}
    recheck = set(sorted(done)[-3:])
    recs, failed = {}, 0
    for p in pd.period_range(f"{DATA_START}-01", last_month, freq="M").astype(str):
        if p in done and p not in recheck:
            recs[p] = done[p]
            continue
        rows = comtrade("M", p.replace("-", ""))
        if rows is None:
            failed += 1
            recs[p] = done.get(p, {"Month": p, "Comtrade_status": "request failed"})
            continue
        rec = summarise(rows)
        rec["Comtrade_status"] = "monthly" if rec.pop("Reported") else "not published monthly"
        rec["Month"] = p
        recs[p] = rec
    m = pd.DataFrame([recs[p] for p in sorted(recs)]).reindex(columns=IMPORT_COLS)
    pub = m["Comtrade_status"] == "monthly"
    out(f"Comtrade monthly: {int(pub.sum())} of {len(m)} months published; {failed} failed requests")
    tonnes = m[["LNG_271111_tonnes", "NG_gaseous_271121_tonnes"]].astype(float).fillna(0).sum(axis=1)
    days = pd.to_datetime(m["Month"]).dt.days_in_month
    m["Imports_MMBtu_per_day"] = (tonnes * MMBTU_PER_T_LNG / days).round(0).where(pub)

    arecs = {int(r["Year"]): r.to_dict() for _, r in have_a.iterrows()} if len(have_a) else {}
    for y in range(DATA_START, int(last_month[:4])):
        mon = m[m["Month"].str[:4] == str(y)]
        n_pub = int((mon["Comtrade_status"] == "monthly").sum())
        if arecs.get(y, {}).get("Months_published_monthly") == 12 and n_pub == 12:
            continue
        rows = comtrade("A", str(y))
        if rows is None:
            continue
        rec = summarise(rows)
        t = rec["LNG_271111_tonnes"] + rec["NG_gaseous_271121_tonnes"]
        ndays = 366 if pd.Timestamp(year=y, month=12, day=31).dayofyear == 366 else 365
        arecs[y] = {"Year": y, "Comtrade_status": "annual" if rec["Reported"] else "not published",
                    "Months_published_monthly": n_pub,
                    "LNG_271111_tonnes": rec["LNG_271111_tonnes"], "LNG_271111_USD": rec["LNG_271111_USD"],
                    "NG_gaseous_271121_tonnes": rec["NG_gaseous_271121_tonnes"], "Partners": rec["Partners"],
                    "Imports_MMBtu_per_day_annual_avg": round(t * MMBTU_PER_T_LNG / ndays, 0),
                    "Sum_of_published_months_tonnes": round(tonnes[mon.index][mon["Comtrade_status"] == "monthly"]
                                                            .sum(), 3)}
    a = pd.DataFrame([arecs[y] for y in sorted(arecs)])
    return m, a


def cenace_gas_mwh():
    """Monthly mean gas-fired generation (MWh/d) from the CENACE workbook, for
    months with at least 20 days of data (cross-check of gas burned for power)."""
    try:
        d = pd.read_excel(CENACE_XLSX, sheet_name="Daily")
        d["date"] = pd.to_datetime(d["date"], errors="coerce")
        g = d.dropna(subset=["date"]).set_index("date")["Gas_MWh"].astype(float)
    except Exception:
        return pd.Series(dtype=float)
    n = g.groupby(g.index.strftime("%Y-%m")).agg(["mean", "count"])
    return n["mean"].where(n["count"] >= 20).dropna().round(0)


def total_demand(df, m, a):
    """Monthly total demand = domestic (Petroecuador) + imports (Comtrade), by
    use. Imported LNG is counted as industry: the import licences are for the
    industrial sector and no imported gas reaches the Machala power plant."""
    t = df[["Month", "Power_MMBtu_per_day", "Industrial_LNG_MMBtu_per_day"]].rename(columns={
        "Power_MMBtu_per_day": "Domestic_power_MMBtu_per_day",
        "Industrial_LNG_MMBtu_per_day": "Domestic_industry_MMBtu_per_day"})
    t = t.merge(m[["Month", "Comtrade_status", "Imports_MMBtu_per_day", "Partners"]], on="Month", how="left")
    t = t.rename(columns={"Imports_MMBtu_per_day": "Imports_industry_MMBtu_per_day", "Partners": "Import_origin"})
    known = t["Imports_industry_MMBtu_per_day"].notna()
    t["Imports_power_MMBtu_per_day"] = pd.Series(0.0, index=t.index).where(known)
    annual = {int(r["Year"]): r for _, r in a.iterrows()} if len(a) else {}

    def status(row):
        if row["Comtrade_status"] == "monthly":
            return "customs month published"
        r = annual.get(int(row["Month"][:4]))
        if r is not None and r.get("Comtrade_status") == "annual":
            return (f"monthly not published; {row['Month'][:4]} annual total "
                    f"{r['LNG_271111_tonnes'] + r['NG_gaseous_271121_tonnes']:,.0f} t = avg "
                    f"{r['Imports_MMBtu_per_day_annual_avg']:,.0f} MMBtu/d (not allocated to months)")
        return "not yet published"
    t["Imports_status"] = t.apply(status, axis=1)
    t["Domestic_total_MMBtu_per_day"] = t[["Domestic_power_MMBtu_per_day", "Domestic_industry_MMBtu_per_day"]].sum(
        axis=1, min_count=1)
    t["Total_power_MMBtu_per_day"] = (t["Domestic_power_MMBtu_per_day"] + t["Imports_power_MMBtu_per_day"])
    t["Total_industry_MMBtu_per_day"] = (t["Domestic_industry_MMBtu_per_day"] + t["Imports_industry_MMBtu_per_day"])
    t["Total_demand_MMBtu_per_day"] = t["Total_power_MMBtu_per_day"] + t["Total_industry_MMBtu_per_day"]
    t["Imports_share_pct"] = (100 * t["Imports_industry_MMBtu_per_day"] / t["Total_demand_MMBtu_per_day"]).round(1)
    for c in ("Domestic_total", "Imports_industry", "Total_power", "Total_industry", "Total_demand"):
        t[f"{c}_mcm_per_day"] = (t[f"{c}_MMBtu_per_day"] / MMBTU_PER_MCM).round(4)
    t["CENACE_gas_generation_MWh_per_day"] = t["Month"].map(cenace_gas_mwh())
    t["Implied_heat_rate_MMBtu_per_MWh"] = (t["Domestic_power_MMBtu_per_day"]
                                            / t["CENACE_gas_generation_MWh_per_day"]).round(2)
    return t[["Month", "Domestic_power_MMBtu_per_day", "Domestic_industry_MMBtu_per_day",
              "Domestic_total_MMBtu_per_day", "Imports_power_MMBtu_per_day", "Imports_industry_MMBtu_per_day",
              "Total_power_MMBtu_per_day", "Total_industry_MMBtu_per_day", "Total_demand_MMBtu_per_day",
              "Imports_share_pct", "Domestic_total_mcm_per_day", "Imports_industry_mcm_per_day",
              "Total_power_mcm_per_day", "Total_industry_mcm_per_day", "Total_demand_mcm_per_day", "Imports_status",
              "Import_origin", "CENACE_gas_generation_MWh_per_day", "Implied_heat_rate_MMBtu_per_MWh"]]


def month_ranges(m):
    """'2022-02..2023-12, 2025-01..2025-12' for the months not published monthly."""
    miss = [pd.Period(x, freq="M") for x in m.loc[m["Comtrade_status"] != "monthly", "Month"]]
    runs = []
    for p in miss:
        if runs and p == runs[-1][1] + 1:
            runs[-1][1] = p
        else:
            runs.append([p, p])
    return ", ".join(str(a) if a == b else f"{a}..{b}" for a, b in runs) or "none"


def read_sheet(path, sheet):
    try:
        d = pd.read_excel(path, sheet_name=sheet)
    except Exception:
        return pd.DataFrame()
    return d.drop(columns=[c for c in d.columns if str(c).startswith("Unnamed")])


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", default="output/Data and Chart Outputs/ecuador_gas.xlsx")
    ap.add_argument("--full", action="store_true", help="re-read every annual report, ignoring the archive")
    args = ap.parse_args()

    have = pd.DataFrame()
    if os.path.exists(args.out) and not args.full:
        have = pd.read_excel(args.out, sheet_name="Gas by use")
        have = have.drop(columns=[c for c in have.columns if str(c).startswith("Unnamed")])
        have["Month"] = pd.to_datetime(have["Month"].astype(str))
    complete_years = set()
    if len(have):
        n = have.dropna(subset=["Power_MMBtu_per_day"]).groupby(have["Month"].dt.year)["Month"].nunique()
        complete_years = {int(y) for y, k in n.items() if k == 12}
        out(f"archive: {len(have)} months; complete years {sorted(complete_years)}")

    reports = list_reports()
    out(f"{len(reports)} reports listed: {[t for t, _ in reports]}")
    rows = {}
    for label, url in reports:
        years = [int(y) for y in re.findall(r"(20\d\d)", label)]
        if max(years) in complete_years and not re.search(r"mes|enero\s*-", label, re.I):
            continue
        try:
            r = requests.get(url, headers=H, timeout=T)
            if r.content[:4] != b"%PDF":
                out(f"  {label}: not a PDF")
                continue
            disp, prod = parse_report(r.content)
        except Exception as e:
            out(f"  {label}: ERR {type(e).__name__}: {str(e)[:150]}")
            continue
        n_d = {k: len(v) for k, v in disp.items()}
        n_p = {k: len(v) for k, v in prod.items()}
        out(f"  {label}: dispatch months {n_d}, production months {n_p}")
        for series in (disp, prod):
            for col, vals in series.items():
                for m, v in vals.items():
                    rows.setdefault(m, {})[col] = v

    new = pd.DataFrame.from_dict(rows, orient="index").sort_index()
    if new.empty and have.empty:
        out("nothing parsed")
        sys.exit(1)
    if not new.empty:
        new.index.name = "Month"
        new = new.reset_index()
        days = new["Month"].dt.days_in_month
        for col in ("Power_MMBtu_month", "Industrial_LNG_MMBtu_month"):
            if col not in new.columns:
                new[col] = None
        new["Power_MMBtu_per_day"] = (new["Power_MMBtu_month"] / days).round(0)
        new["Industrial_LNG_MMBtu_per_day"] = (new["Industrial_LNG_MMBtu_month"] / days).round(0)
        new = new.drop(columns=["Power_MMBtu_month", "Industrial_LNG_MMBtu_month"])
    df = pd.concat([have, new], ignore_index=True) if len(have) else new
    df = df.drop_duplicates("Month", keep="last").sort_values("Month").reset_index(drop=True)
    df = df[df["Month"].dt.year >= DATA_START]
    df["Total_sales_MMBtu_per_day"] = df[["Power_MMBtu_per_day", "Industrial_LNG_MMBtu_per_day"]].sum(axis=1, min_count=1)
    df["Total_sales_MMcf_per_day_approx"] = (df["Total_sales_MMBtu_per_day"] / BTU_PER_FT3).round(2)
    if "Amistad_production_BOE_per_day" in df.columns:
        df["Amistad_production_MMBtu_per_day"] = (df["Amistad_production_BOE_per_day"] * MMBTU_PER_BOE).round(0)
    df["Month"] = df["Month"].dt.strftime("%Y-%m")
    out(df.tail(14).to_string(index=False))
    months = pd.period_range(df["Month"].min(), df["Month"].max(), freq="M").astype(str)
    gaps = sorted(set(months) - set(df.dropna(subset=["Power_MMBtu_per_day"])["Month"]))
    out(f"{len(df)} months {df['Month'].min()}..{df['Month'].max()}; missing dispatch: {gaps or 'none'}")

    # imports (UN Comtrade) and the total-demand view
    have_m = pd.DataFrame() if args.full else read_sheet(args.out, "Imports monthly")
    have_a = pd.DataFrame() if args.full else read_sheet(args.out, "Imports annual")
    if len(have_m):
        have_m["Month"] = have_m["Month"].astype(str).str[:7]
    imp_m, imp_a = pull_imports(have_m, have_a, df["Month"].max())
    tot = total_demand(df, imp_m, imp_a)
    out(imp_m[imp_m["LNG_271111_tonnes"].astype(float).fillna(0) > 1].to_string(index=False))
    out(imp_a.to_string(index=False))
    out(tot.tail(8)[["Month", "Domestic_power_MMBtu_per_day", "Domestic_industry_MMBtu_per_day",
                     "Imports_industry_MMBtu_per_day", "Total_demand_MMBtu_per_day", "Imports_status"]].to_string(index=False))
    last_full = tot.dropna(subset=["Total_demand_MMBtu_per_day"]).tail(1)
    def num(r, c):
        v = r.get(c)
        return 0.0 if v is None or pd.isna(v) else float(v)

    annual_lines = [f"  {int(r['Year'])}: {r['Comtrade_status']}, "
                    f"{num(r, 'LNG_271111_tonnes') + num(r, 'NG_gaseous_271121_tonnes'):,.1f} t "
                    f"(avg {num(r, 'Imports_MMBtu_per_day_annual_avg'):,.0f} MMBtu/d); "
                    f"{int(num(r, 'Months_published_monthly'))} months published monthly, summing to "
                    f"{num(r, 'Sum_of_published_months_tonnes'):,.1f} t. "
                    f"Origin: {r['Partners'] if isinstance(r.get('Partners'), str) and r['Partners'] else '-'}"
                    for _, r in imp_a.iterrows()]

    notes = [
        "UNITS",
        "MMBtu per day = monthly dispatch in MMBtu (as published) / days in month. MMcf/d approx at 1,030 Btu per "
        f"cubic foot. Amistad production is published in barrels of oil equivalent per day; MMBtu at "
        f"{MMBTU_PER_BOE} MMBtu/BOE.",
        "",
        "SECTORS",
        "Power_MMBtu_per_day: Petroecuador's 'GAS NATURAL (MMBTU)' dispatches - pipeline gas from the Amistad field "
        "to the Termogas Machala power plant. Industrial_LNG_MMBtu_per_day: 'GAS NATURAL LICUADO (MMBTU)' - LNG "
        "from the Bajo Alto liquefaction plant, trucked to industrial users (mainly ceramics in Cuenca). These two "
        "are effectively all of Ecuador's commercial natural gas.",
        "",
        "COVERAGE",
        f"Monthly from {DATA_START}. Past years from Petroecuador's annual statistical reports, the current year from "
        "the year-to-date monthly report. Runs are incremental: complete years in this archive aren't re-read.",
        "",
        "TOTAL DEMAND (sheet 'Total demand')",
        "Total demand = domestic supply by use (Petroecuador, above) + imports (customs, UN Comtrade). "
        "Power = domestic pipeline gas to Termogas Machala + imports for power; industry = Bajo Alto LNG + "
        "imported LNG. mcm/d = million cubic metres per day at 1,030 Btu per cubic foot "
        f"({MMBTU_PER_MCM:,.0f} MMBtu per million m3).",
        "Imports for power are zero: Ecuador has no LNG import terminal and EP Petroecuador imports no natural gas "
        "(its imported derivatives are naphtha, diesel and LPG - BCE 'Reporte del Sector Petrolero'); its 2023 gas "
        "import tender was declared void and its Sept-2026 request for information (60-100 MMcf/d of LNG for "
        "Machala) has not led to deliveries. The 2024-26 emergency power barges (Karpowership) burn fuel oil, not "
        "gas. LPG is not natural gas and is excluded.",
        "Imports for industry: small private LNG cargoes in ISO containers (Sycar; from Panama/AES Colon from "
        "Jan 2022, trucked from Peru/Limagas from Nov 2024), customs code HS 2711.11. Tonnes converted at "
        f"{MMBTU_PER_T_LNG:.0f} MMBtu per tonne of LNG (gross); HS 2711.21 (gaseous) volumes are negligible "
        "(samples) and use the same factor. Assigned to industry because the import licence (Ministerio de "
        "Energia y Minas, up to 7.3 bcf/yr) is for the industrial sector.",
        "Imports_status says, for each month, whether the customs month is published. Total_demand is left blank "
        f"where it is not. Months not published monthly (this run): {month_ranges(imp_m)}. Recent months lag by "
        "several months and are asked for again on every run. Annual totals for years without full monthly data "
        "are below and in 'Imports annual'; they are NOT spread over months.",
        "Annual imports (UN Comtrade, Ecuador customs):",
        *annual_lines,
        (f"Latest month with both domestic and import data: {last_full['Month'].iloc[0]} - total "
         f"{last_full['Total_demand_MMBtu_per_day'].iloc[0]:,.0f} MMBtu/d, imports "
         f"{last_full['Imports_share_pct'].iloc[0]:.1f}% of it." if len(last_full) else ""),
        "Cross-check: CENACE gas-fired generation (Ecuador power workbook) for months with 20+ days of data; "
        "Implied_heat_rate = domestic power gas / CENACE gas MWh. A plausible heat rate (about 8-11 MMBtu/MWh) "
        "means the Machala plant runs on domestic gas only.",
        "",
        "SOURCE",
        f"EP Petroecuador, Informes Estadisticos (Cifras Institucionales): {PAGE}",
        "UN Comtrade (Ecuador customs returns via Banco Central del Ecuador / SENAE), HS 2711.11 and 2711.21 "
        "imports: https://comtradeplus.un.org/ (public API comtradeapi.un.org/public/v1/preview).",
    ]
    os.makedirs(os.path.dirname(os.path.abspath(args.out)), exist_ok=True)
    xlsx_notes.write_workbook(args.out, {"Gas by use": df, "Total demand": tot, "Imports monthly": imp_m,
                                         "Imports annual": imp_a}, notes,
                              {"UNITS", "SECTORS", "COVERAGE", "TOTAL DEMAND (sheet 'Total demand')", "SOURCE"})
    out(f"Saved {args.out}")


if __name__ == "__main__":
    main()
