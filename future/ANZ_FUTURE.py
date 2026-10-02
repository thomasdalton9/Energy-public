"""
Australia + NZ "future" workbook (one-off, run on demand): project pipeline, closures and official scenarios.

  NEM pipeline   AEMO NEM Generation Information (quarterly workbook, latest found):
                 https://www.aemo.com.au/-/media/files/electricity/nem/planning_and_forecasting/generation_information/
                 <year>/nem-generation-information-<month>-<year>.xlsx
                 sheet "Generator Information": every existing, committed, anticipated and proposed unit with its
                 status, capacity, technology, commissioning date and expected closure year
  NZ scenarios   MBIE Electricity Demand and Generation Scenarios (EDGS 2024): results workbook (generation and
                 capacity by fuel to 2050, by scenario)

Writes australia_nz_future.xlsx. Also writes "NEM capacity rebuilt": capacity by fuel at each month-end from 2021,
rebuilt from the units' commissioning dates and closures - a check on (and possible replacement for) the
snapshot-only history in au_power_capacity.xlsx.

Usage: python3 future/ANZ_FUTURE.py [--out "output/Data and Chart Outputs/australia_nz_future.xlsx"]
"""
import argparse
import io
import os
import re
import sys
from datetime import date

import pandas as pd

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import xlsx_notes  # noqa: E402
from future_common import FUELS, by_year_fuel, col, fuel_from_text, get, read_table  # noqa: E402

GENINFO = ("https://www.aemo.com.au/-/media/files/electricity/nem/planning_and_forecasting/generation_information/"
           "{y}/nem-generation-information-{m}-{y}.xlsx")
MBIE = ("https://www.mbie.govt.nz/building-and-energy/energy-and-natural-resources/energy-statistics-and-modelling/"
        "energy-modelling/electricity-demand-and-generation-scenarios/")
DEFAULT_OUT = os.path.join(ROOT, "output", "Data and Chart Outputs", "australia_nz_future.xlsx")
MONTHS = ["january", "february", "march", "april", "may", "june", "july", "august", "september", "october",
          "november", "december"]


def latest_geninfo():
    today = date.today()
    for y in (today.year, today.year - 1):
        for mi in range(11, -1, -1):
            for m in (MONTHS[mi], MONTHS[mi][:3]):
                url = GENINFO.format(y=y, m=m)
                try:
                    r = get(url)
                except Exception:  # noqa: BLE001
                    continue
                if r.content[:2] == b"PK":
                    return url, r.content
    raise RuntimeError("no NEM Generation Information workbook found")


def nem(url, content):
    d = read_table(content, "Generator Information", [r"Survey ID", r"Site Name"])
    print(f"Generation Information: {len(d)} rows; columns {list(d.columns)}", flush=True)
    status = col(d, r"^Status$|Commitment Status|Unit Status|Project Status")
    print(f"  status values: {d[status].value_counts().to_dict()}", flush=True)
    tech, detail = col(d, r"Technology Type"), col(d, r"Technology Detail", required=False)
    gasfuel = col(d, r"Fuel Type", required=False)
    cap = col(d, r"Nameplate Capacity \(MW\)|Upper Nameplate", r"Unit Capacity \(MW", r"Capacity \(MW")
    cnt = col(d, r"Unit Count", required=False)
    mw = pd.to_numeric(d[cap], errors="coerce")
    if cnt and re.search(r"Unit Capacity", cap):
        mw = mw * pd.to_numeric(d[cnt], errors="coerce").fillna(1)
    d["MW"] = mw
    d["Fuel"] = (d[tech].astype(str) + " " + (d[detail].astype(str) if detail else "") + " " +
                 (d[gasfuel].astype(str) if gasfuel else "")).map(fuel_from_text)
    region = col(d, r"^Region$")
    sheets = {}
    st = d.pivot_table(index=status, columns="Fuel", values="MW", aggfunc="sum").fillna(0)
    st = st.reindex(columns=[f for f in FUELS if f in st.columns]).round(0)
    st.index.name = "Status"
    sheets["NEM pipeline by status"] = st
    start = col(d, r"Full Commercial Use Date|Commercial Use|Commissioning Date|Expected.*(Operation|Commissioning)",
                required=False)
    close = col(d, r"Expected Closure Year|Closure Year", required=False)
    print(f"  date columns: start={start}, close={close}", flush=True)
    stv = d[status].astype(str)
    pipeline = d[~stv.str.contains(r"In Service|Existing|Operating", case=False, regex=True)]
    if start:
        yr = pd.to_datetime(pipeline[start], errors="coerce").dt.year
        sheets["NEM additions by year"] = by_year_fuel(pipeline.assign(_yr=yr), "_yr", "Fuel", "MW",
                                                       start=date.today().year - 1)
    if close:
        existing = d[stv.str.contains(r"In Service|Existing|Operating", case=False, regex=True)]
        sheets["NEM closures by year"] = by_year_fuel(existing, close, "Fuel", "MW", start=date.today().year)
        # capacity history rebuilt from commissioning dates (existing units only), month-ends from 2021
        if start:
            when = pd.to_datetime(existing[start], errors="coerce")
            known = when.notna()
            months = pd.date_range("2021-01-01", pd.Timestamp(date.today()), freq="MS")
            rows = {}
            for m in months:
                live = existing[known & (when <= m + pd.offsets.MonthEnd(0))]
                rows[m] = live.groupby("Fuel")["MW"].sum()
            hist = pd.DataFrame(rows).T.fillna(0)
            hist = hist.reindex(columns=[f for f in FUELS if f in hist.columns]).round(0)
            hist.index.name = "date"
            sheets["NEM capacity rebuilt"] = hist
            print(f"  rebuilt history: {known.sum()} of {len(existing)} existing units have a start date "
                  f"({existing.loc[known, 'MW'].sum() / 1000:.1f} of {existing['MW'].sum() / 1000:.1f} GW)", flush=True)
    keep = [c for c in (col(d, r"^Site Name$"), col(d, r"Site Owner", required=False), region, status, tech,
                        detail, "Fuel", "MW", start, close) if c]
    sheets["NEM units"] = d[keep].reset_index(drop=True)
    by_region = d[~stv.str.contains(r"In Service|Existing|Operating", case=False, regex=True)].pivot_table(
        index=region, columns="Fuel", values="MW", aggfunc="sum").fillna(0).round(0)
    by_region.index.name = "Region"
    sheets["NEM pipeline by region"] = by_region
    return sheets


def nz():
    html = get(MBIE).text
    links = re.findall(r'href="([^"]+results[^"]*\.xlsx)"', html, re.I)
    if not links:
        raise RuntimeError("no EDGS results workbook link")
    url = links[0] if links[0].startswith("http") else "https://www.mbie.govt.nz" + links[0]
    content = get(url).content
    xl = pd.ExcelFile(io.BytesIO(content))
    print(f"EDGS results {url}: sheets {xl.sheet_names}", flush=True)
    sheets = {}
    for sh in xl.sheet_names:
        if not re.search(r"generation|capacity", sh, re.I):
            continue
        d = pd.read_excel(xl, sh)
        print(f"  {sh}: {d.shape} columns {list(d.columns)}", flush=True)
        try:
            yc, vc, sc = col(d, r"TimePeriod|Year"), col(d, r"^Value$"), col(d, r"Scenario")
        except KeyError:
            continue
        cat = col(d, r"Fuel|Technology|Generation type|Plant", r"Variable", required=False)
        ref = d[d[sc].astype(str).str.contains("Reference", case=False)] if d[sc].astype(str).str.contains(
            "Reference", case=False).any() else d
        p = ref.pivot_table(index=yc, columns=cat, values=vc, aggfunc="sum") if cat else ref.groupby(yc)[vc].sum().to_frame()
        p.index = pd.to_datetime(p.index.astype(int).astype(str) + "-01-01")
        p.index.name = "Year"
        sheets[f"NZ {sh}"[:31]] = p.round(2)
    return sheets, url


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", default=DEFAULT_OUT)
    args = ap.parse_args()
    sheets, src = {}, []
    try:
        url, content = latest_geninfo()
        print(f"NEM Generation Information: {url}", flush=True)
        sheets.update(nem(url, content))
        src.append(f"AEMO NEM Generation Information: {url}")
    except Exception as e:  # noqa: BLE001
        print(f"NEM Generation Information failed: {type(e).__name__}: {e}", flush=True)
    try:
        s, url = nz()
        sheets.update(s)
        src.append(f"MBIE Electricity Demand and Generation Scenarios (EDGS 2024) results: {url}")
    except Exception as e:  # noqa: BLE001
        print(f"MBIE EDGS failed: {type(e).__name__}: {e}", flush=True)
    if not sheets:
        raise SystemExit("nothing fetched")
    notes = [
        "UNITS",
        "NEM: MW by fuel. Pipeline by status: every unit in AEMO's Generation Information by its commitment status "
        "(in service, committed, anticipated, proposed, ...). Additions by year: not-yet-operating units by "
        "expected commercial use year. Closures by year: operating units by expected closure year. Capacity "
        "rebuilt: operating units' capacity at each month-end from their commissioning dates (units retired before "
        "this publication are not in it, so earlier months understate closed coal).",
        "NZ: MBIE EDGS 2024 Reference scenario, as published (generation GWh / capacity MW).",
        "",
        "COVERAGE",
        "NEM (QLD, NSW, VIC, SA, TAS). WA (WEM) and AEMO's ESOO/GSOO/ISP data are not included (AEMO's web pages "
        "refuse automated access).",
        "",
        "SOURCE",
        *src,
    ]
    os.makedirs(os.path.dirname(os.path.abspath(args.out)), exist_ok=True)
    xlsx_notes.write_workbook(args.out, sheets, notes, {"UNITS", "COVERAGE", "SOURCE"})
    print(f"Saved {args.out}: {list(sheets)}", flush=True)


if __name__ == "__main__":
    main()
