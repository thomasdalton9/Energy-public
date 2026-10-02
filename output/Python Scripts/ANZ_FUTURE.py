"""
Australia + NZ "future" workbook (one-off, run on demand): project pipeline, closures and official scenarios.

  NEM pipeline   AEMO NEM Generation Information (quarterly workbook, latest found):
                 https://www.aemo.com.au/-/media/files/electricity/nem/planning_and_forecasting/generation_information/
                 <year>/nem-generation-information-<month>-<year>.xlsx
                 sheet "Generator Information": every existing, committed, anticipated and proposed unit with its
                 status, capacity, technology, commissioning date and expected closure year
  NZ scenarios   MBIE Electricity Demand and Generation Scenarios (EDGS 2024): results workbook (generation and
                 capacity by fuel to 2050, by scenario)

Writes australia_nz_future.xlsx, including "NEM capacity outlook": in-service capacity by fuel, plus committed/anticipated projects by their commercial-use
year, minus announced closures by closure year.

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
    agg = col(d, r"Aggregated Nameplate Capacity \(MW AC\)", required=False)
    unit = col(d, r"Unit Capacity \(MW AC\)", required=False)
    cnt = col(d, r"Unit Count", required=False)
    mw = pd.to_numeric(d[agg], errors="coerce") if agg else pd.Series(float("nan"), index=d.index)
    if unit:   # units without an aggregated figure: unit capacity x unit count
        mw = mw.fillna(pd.to_numeric(d[unit], errors="coerce") * pd.to_numeric(d[cnt], errors="coerce").fillna(1)
                       if cnt else pd.to_numeric(d[unit], errors="coerce"))
    site = col(d, r"Max Site Capacity", required=False)
    if site:   # single-unit sites: the site capacity
        mw = mw.fillna(pd.to_numeric(d[site], errors="coerce"))
    d["MW"] = mw
    print(f"  MW: {d['MW'].notna().sum()} of {len(d)} rows have a capacity ({d['MW'].sum() / 1000:,.1f} GW)", flush=True)
    text = lambda c: d[c].fillna("").astype(str) if c else ""  # noqa: E731 - NaN would blank the whole label
    d["Fuel"] = (text(tech) + " " + text(detail) + " " + text(gasfuel)).map(fuel_from_text)
    print(f"  fuels: {d.groupby('Fuel')['MW'].sum().round(0).to_dict()}", flush=True)
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
    live = stv.str.contains(r"In Service", case=False)
    firm = stv.str.contains(r"Committed|Anticipated|In Commissioning", case=False)
    pipeline = d[~live & ~stv.str.contains("Withdraw", case=False)]
    this_year = date.today().year
    if start:
        yr = pd.to_datetime(pipeline[start], errors="coerce").dt.year
        sheets["NEM additions by year"] = by_year_fuel(pipeline.assign(_yr=yr), "_yr", "Fuel", "MW",
                                                       start=this_year - 1, end=2040)
    if close:
        existing = d[live]
        sheets["NEM closures by year"] = by_year_fuel(existing, close, "Fuel", "MW", start=this_year, end=2060)
        # capacity outlook: in service today, plus committed / anticipated / commissioning projects by their
        # commercial-use year, minus announced closures by closure year
        cy = pd.to_numeric(existing[close], errors="coerce")
        sy = pd.to_datetime(d.loc[firm, start], errors="coerce").dt.year if start else None
        rows = {}
        for y in range(this_year, 2041):
            base = existing[~(cy <= y)].groupby("Fuel")["MW"].sum()
            if sy is not None:
                base = base.add(d.loc[firm][sy.fillna(this_year) <= y].groupby("Fuel")["MW"].sum(), fill_value=0)
            rows[pd.Timestamp(f"{y}-01-01")] = base
        out = pd.DataFrame(rows).T.fillna(0)
        out = out.reindex(columns=[f for f in FUELS if f in out.columns]).round(0)
        out.index.name = "Year"
        sheets["NEM capacity outlook"] = out
    keep = [c for c in (col(d, r"^Site Name$"), col(d, r"Site Owner", required=False), region, status, tech,
                        detail, "Fuel", "MW", start, close) if c]
    sheets["NEM units"] = d[keep].reset_index(drop=True)
    by_region = pipeline.pivot_table(
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
        if not re.search(r"generation|build", sh, re.I):
            continue
        d = pd.read_excel(xl, sh)
        print(f"  {sh}: {d.shape} columns {list(d.columns)}\n{d.head(4).to_string()[:900]}", flush=True)
        try:
            yc, vc = col(d, r"TimePeriod|Year"), col(d, r"^Value$|MW|Capacity")
        except KeyError:
            continue
        sc = col(d, r"Scenario", required=False)
        if sc:
            print(f"    scenarios: {d[sc].dropna().unique().tolist()}", flush=True)
            ref = d[sc].astype(str).str.contains("Reference", case=False)
            d = d[ref] if ref.any() else d
        cat = col(d, r"Commodity|Fuel|Technology|Plant type", required=False)
        var = col(d, r"^Variable$", required=False)
        groups = d.groupby(var) if var else [(sh, d)]
        for v, g in groups:
            p = g.pivot_table(index=yc, columns=cat, values=vc, aggfunc="sum") if cat else g.groupby(yc)[vc].sum().to_frame()
            p.index = pd.to_datetime(pd.to_numeric(p.index, errors="coerce").astype("Int64").astype(str) + "-01-01",
                                     errors="coerce")
            p = p[p.index.notna()]
            p = p.loc[:, p.abs().sum() > 0]
            p.index.name = "Year"
            if not p.empty:
                sheets[f"NZ {v}"[:31]] = p.round(2)
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
        "expected commercial use year. Closures by year: in-service units by expected closure year. Capacity "
        "outlook: in-service capacity, plus committed / anticipated / commissioning projects from their commercial "
        "use year, minus announced closures from their closure year (publicly announced projects not included).",
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
