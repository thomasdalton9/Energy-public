"""
Nigeria grid power generation by plant, QUARTERLY, from the Nigerian Electricity Regulatory Commission (NERC).

Why NERC and not the grid operator: the system operator (NISO, formerly NCC/TCN) has no public data feed. niso.ng
does not resolve, nsong.org is a login + CAPTCHA 'Dispatch Tools Portal', and NERC's own weekly 'Harmonized Reports'
(2020) and monthly operational factsheets are picture-only PDFs. The one structured raw series that is public is
NERC's Quarterly Report (PDF with real text), section 2.1.3 'Quarterly generation': average hourly generation
(MWh/h) of each grid-connected power plant for the quarter and the one before. NERC reports NISO's metered
generation of the grid-connected plants (about 28 plants in 2026/Q2: 5 hydro, 2 steam, 19 OCGT, 2 CCGT).
  https://nerc.gov.ng/resource-category/nerc-reports/

Granularity is QUARTERLY (no daily or monthly series is public). Only the plants NERC lists are included: embedded
and captive generation, solar and wind outside the grid, and island plants (e.g. Maiduguri) are NOT in it.
Fuel split: NERC gives hydro vs thermal only. Every thermal plant on the grid is a gas turbine or a gas-fired steam
unit, so 'Thermal' is booked to the Gas column and the Units tab says so (NERC itself does not state the fuel).

Sheets: 'Quarterly' (date = first day of the quarter; Hydro_GWh, Gas_GWh, Total_GWh, average GW, hours),
'Plants' (average MWh/h per plant and quarter), 'Reports' (every PDF read, quarter found, checks).
Each report carries its own quarter AND the quarter before; the report of a quarter wins, the earlier-quarter
column only fills a quarter no report of its own was read for (basis column).

Incremental: the committed workbook is the history store; a PDF already read is not downloaded again except the
newest report (NERC revises the prior-quarter column).

Usage: python3 NIGERIA_NERC_POWER.py [--out PATH] [--diag]   (--diag re-reads every report)
"""

import argparse
import calendar
import datetime as dt
import os
import re
import sys

import pandas as pd
import requests
import urllib3

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))  # repo root, for xlsx_notes

import xlsx_notes  # noqa: E402

urllib3.disable_warnings()

OUT = "output/Data and Chart Outputs/nigeria_power_generation_quarterly.xlsx"
CATEGORY = "https://nerc.gov.ng/resource-category/nerc-reports/"
H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36"}
HYDRO = {"Kainji", "Jebba", "Shiroro", "Zungeru", "Dadin-Kowa"}
ORD = {"FIRST": 1, "SECOND": 2, "THIRD": 3, "FOURTH": 4}
PLANT = re.compile(r"^[A-Z][A-Za-z0-9\-\. ]*?_\d+$")


def get(url, retries=3, **kw):
    for i in range(retries):
        try:
            r = requests.get(url, headers=H, timeout=(10, 120), verify=False, **kw)
            if r.status_code == 200:
                return r
            print(f"  HTTP {r.status_code} {url}", flush=True)
            if r.status_code in (404, 403):
                return None
        except requests.RequestException as e:
            print(f"  {type(e).__name__} {url}: {str(e)[:100]}", flush=True)
    return None


def list_reports():
    """Quarterly report PDFs on NERC's 'NERC Reports' category (newest first)."""
    found, seen = [], set()
    for pg in range(1, 15):
        r = get(CATEGORY + (f"page/{pg}/" if pg > 1 else ""))
        if r is None:
            break
        n0 = len(found)
        for u in re.findall(r'href=["\']([^"\']+\.pdf)["\']', r.text, re.I):
            u = u.replace(" ", "%20")
            name = u.rsplit("/", 1)[-1].lower()
            if u in seen:
                continue
            seen.add(u)
            if re.search(r"annual", name):
                continue
            if re.search(r"q[1-4]|quarter", name):
                found.append(u)
        if len(found) == n0 and pg > 3 and not re.search(r"\.pdf", r.text):
            break
    print(f"{len(found)} quarterly report PDFs listed", flush=True)
    return found


def num(tok):
    t = tok.strip()
    if t in ("-", "–", "−", "—"):
        return "dash"
    if re.search(r"[A-Za-z]", t) or len(t) > 14:
        return None
    c = re.sub(r"[^\d.,\-]", "", t)
    if re.fullmatch(r"-?\d[\d,]*\.\d+", c):
        return float(c.replace(",", ""))
    return None


def quarter_of(doc):
    head = " ".join(doc[i].get_text() for i in range(min(2, len(doc))))
    m = re.search(r"(FIRST|SECOND|THIRD|FOURTH)\s+QUARTER\s+(\d{4})", head, re.I)
    if m:
        return int(m.group(2)), ORD[m.group(1).upper()]
    m = re.search(r"(\d{4})\s*/\s*Q([1-4])", head)
    return (int(m.group(1)), int(m.group(2))) if m else None


def canon(name):
    """Plant name: footnote digit glued to the unit number removed (Alaoji_17 -> Alaoji_1), spacing/case normalised."""
    name = re.sub(r"_([12])\d$", r"_\1", name.strip())
    base, unit = name.rsplit("_", 1)
    base = " ".join(base.replace("Dadin Kowa", "Dadin-Kowa").split())
    return f"{base.title() if base.islower() else base}_{unit}".replace("Ibom power", "Ibom Power")


def parse_lines(lines):
    recs, total, k = {}, None, 0
    while k < len(lines):
        ln = lines[k]
        if PLANT.match(ln) or ln == "Total":
            vals, j = [], k + 1
            while j < len(lines):
                toks = lines[j].split() if lines[j] else []
                nv = [num(t) for t in toks]
                if toks and all(v is not None for v in nv):
                    vals.extend(nv)
                elif re.fullmatch(r"\d{1,2}", lines[j] or "x") and not vals:
                    pass                      # footnote marker between name and values
                else:
                    break
                j += 1
                if len(vals) >= 4:
                    break
            if len(vals) >= 2:
                prev = 0.0 if vals[0] == "dash" else vals[0]
                cur = 0.0 if vals[1] == "dash" else vals[1]
                if ln == "Total":
                    total = (prev, cur)
                else:
                    recs.setdefault(canon(ln), (prev, cur))
            k = j
        else:
            k += 1
    return recs, total


def parse_plants(doc):
    """(rows {plant: (prev, cur)}, total (prev, cur), page) from the 'Average Hourly Generation' table, which may run
    over two or three pages (the page holding the Total row is tried alone, then with the pages before it)."""
    for i, pg in enumerate(doc):
        t = pg.get_text()
        if "Average Hourly Generation" not in t or not re.search(r"(?m)^Total\s*$", t):
            continue
        for back in (0, 1, 2):
            if i - back < 0:
                break
            lines = []
            for n in range(i - back, i + 1):
                lines += [ln.strip() for ln in doc[n].get_text().split("\n")]
            recs, total = parse_lines(lines)
            if len(recs) >= 8 and total and total[1]:
                cur_sum = sum(v[1] for v in recs.values())
                if abs(cur_sum - total[1]) / total[1] < 0.005:
                    return recs, total, i + 1
        recs, total = parse_lines([ln.strip() for ln in t.split("\n")])
        if len(recs) >= 8 and total:
            return recs, total, i + 1
    return None


def keyfacts(doc):
    """Total quarterly generation (GWh) and hydro share (%) from the report's text, for reports without the plant table."""
    txt = re.sub(r"\s+", " ", " ".join(p.get_text() for p in doc[: min(len(doc), 40)]))
    tot = share = None
    pats = [r"total generation of ([\d,]+\.\d+) ?GWh",
            r"([\d,]+\.\d+) ?GWh,? Total quarterly (?:energy )?generation",
            r"Total (?:quarterly )?(?:energy |electricity )?generat\w+[^.\d]{0,80}?(?:was|of|stood at) ([\d,]+\.\d+) ?GWh",
            r"total generation was ([\d,]+\.\d+) ?GWh"]
    for pt in pats:
        m = re.search(pt, txt, re.I)
        if m:
            tot = float(m.group(1).replace(",", ""))
            break
    if tot is None:
        m = re.search(r"total electric energy generated (?:in [\d/Q]+ )?was ([\d,]{7,}) ?MWh", txt, re.I)
        if m:
            tot = float(m.group(1).replace(",", "")) / 1000
    for pt in (r"([\d.]+)% Share of total quarterly generation from Hydro ?power Plants",
               r"([\d.]+)% Share of Hydro ?Power plants in the energy mix",
               r"contribution of hydro ?power plants to the energy mix in [\d/Q]+ was ([\d.]+)%"):
        m = re.search(pt, txt, re.I)
        if m:
            share = float(m.group(1))
            break
    return tot, share


def qstart(y, q):
    return pd.Timestamp(y, 3 * q - 2, 1)


def hours(ts):
    end = ts + pd.DateOffset(months=3)
    return int((end - ts).days * 24)


def prev_q(y, q):
    return (y, q - 1) if q > 1 else (y - 1, 4)


def read_report(url, diag=False):
    import fitz
    r = get(url)
    if r is None or not r.content.startswith(b"%PDF"):
        return None, "download failed"
    doc = fitz.open(stream=r.content, filetype="pdf")
    q = quarter_of(doc)
    if q is None:
        return None, "quarter not found"
    if q[0] < 2019:
        return {"quarter": q}, "skipped: pre-2019 layout, quarter not reliably identified"
    parsed = parse_plants(doc)
    if parsed is None:
        tot, share = keyfacts(doc)
        if tot and 5000 < tot < 14000:
            note = f"key facts text: total {tot:,.2f} GWh, hydro share {share if share is not None else 'n/a'}%"
            return {"quarter": q, "kf": (tot, share), "note": note}, "ok (key facts)"
        return {"quarter": q}, "no plant table"
    recs, total, page = parsed
    cur_sum = sum(v[1] for v in recs.values())
    prev_sum = sum(v[0] for v in recs.values())
    chk = max(abs(cur_sum - total[1]) / total[1] if total[1] else 1, abs(prev_sum - total[0]) / total[0] if total[0] else 1)
    txt = " ".join(p.get_text() for p in doc)
    m = re.search(r"total generation\s+of\s+([\d,]+\.\d+)\s*GWh", txt, re.I)
    stated = float(m.group(1).replace(",", "")) if m else None
    h = hours(qstart(*q))
    gwh_calc = total[1] * h / 1000
    note = f"page {page}; {len(recs)} plants; plant sum vs Total {chk:.3%}"
    if stated:
        note += f"; stated {stated:,.2f} GWh vs calc {gwh_calc:,.2f}"
    status = "ok" if chk < 0.005 and (stated is None or abs(stated - gwh_calc) / stated < 0.005) else "check failed"
    return {"quarter": q, "recs": recs, "total": total, "note": note}, status


def load(path, sheet):
    try:
        return pd.read_excel(path, sheet_name=sheet)
    except (FileNotFoundError, ValueError, KeyError, OSError):
        return pd.DataFrame()


NOTES = [
    "UNITS",
    "Quarterly: one row per calendar quarter (date = first day of the quarter). Hydro_GWh / Gas_GWh / Total_GWh = NERC's "
    "average hourly generation (MWh/h) x hours in the quarter / 1000; Hydro_GW / Gas_GW / Total_GW are the average GW. "
    "Gas = all thermal plants on the grid (gas turbines and gas-fired steam units; NERC reports 'thermal' only). "
    "basis = 'report' (NERC's report for that quarter) or 'prior-quarter column' (read from the next quarter's report "
    "because no report of its own was read).",
    "Plants: average MWh/h per grid-connected plant and quarter, as NERC prints it (plant names as in the report; "
    "Alaoji_1 shows 0 when NERC prints a dash = not available for dispatch).",
    "Reports: every NERC report PDF read, the quarter found, the checks (plants add to NERC's Total within 0.5% and "
    "total GWh matches the report's stated 'total generation') and status.",
    "",
    "COVERAGE",
    "Quarterly only. Grid-connected plants NERC reports on (about 28 in 2026/Q2); embedded/captive generation, "
    "off-grid solar and wind and island plants are not included. Quarters without a readable table are gaps, never "
    "filled. The system operator (NISO) publishes nothing public: niso.ng does not resolve, nsong.org is a login "
    "portal, NERC's weekly Harmonized Reports and monthly factsheets are picture-only PDFs.",
    "",
    "SOURCE",
    "Nigerian Electricity Regulatory Commission (NERC), Quarterly Reports, section 2.1.3 'Quarterly generation' "
    f"(average hourly generation by plant): {CATEGORY}",
]


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--out", default=OUT)
    ap.add_argument("--diag", action="store_true", help="re-read every report")
    args = ap.parse_args()

    reports = load(args.out, "Reports")
    done = set(reports.loc[reports["status"].astype(str).str.startswith(("ok", "skipped")), "url"]) if not reports.empty else set()
    plants = load(args.out, "Plants")
    if not plants.empty:
        plants["date"] = pd.to_datetime(plants["date"])
        plants = plants.set_index("date")
    basis, kf = {}, {}
    qs = load(args.out, "Quarterly")
    if not qs.empty:
        qs["date"] = pd.to_datetime(qs["date"])
        basis = dict(zip(qs["date"], qs["basis"]))
        for r in qs.itertuples():
            if str(r.basis).startswith("key facts"):
                kf[r.date] = (r.Total_GWh, (100 * r.Hydro_GWh / r.Total_GWh) if r.Total_GWh and pd.notna(r.Hydro_GWh) else None)

    urls = list_reports()
    if not urls:
        print("No report list from NERC.", flush=True)
        sys.exit(1 if plants.empty else 0)
    newest = urls[0]
    rows = {r["url"]: r for r in reports.to_dict("records")} if not reports.empty else {}
    plants_d = {d: plants.loc[d].dropna().to_dict() for d in plants.index} if not plants.empty else {}

    for u in urls:
        if u in done and u != newest and not args.diag:
            continue
        print(f"Reading {u}", flush=True)
        res, status = read_report(u)
        rec = {"url": u, "file": u.rsplit("/", 1)[-1], "quarter": "", "status": status, "note": ""}
        if res:
            y, q = res["quarter"]
            rec["quarter"] = f"{y}/Q{q}"
            rec["note"] = res.get("note", "")
        print(f"  -> {rec['quarter']} {status} {rec['note']}", flush=True)
        rows[u] = rec
        if not status.startswith("ok"):
            continue
        y, q = res["quarter"]
        cur_d = qstart(y, q)
        if "kf" in res:
            if cur_d not in plants_d:
                kf[cur_d] = res["kf"]
                basis[cur_d] = "key facts (report text)"
            continue
        py, pq = prev_q(y, q)
        prv_d = qstart(py, pq)
        plants_d[cur_d] = {k: v[1] for k, v in res["recs"].items()}
        basis[cur_d] = "report"
        kf.pop(cur_d, None)
        if prv_d not in plants_d or basis.get(prv_d) != "report":
            plants_d[prv_d] = {k: v[0] for k, v in res["recs"].items()}
            basis[prv_d] = "prior-quarter column"
            kf.pop(prv_d, None)

    if not plants_d and not kf:
        print("Nothing parsed.", flush=True)
        sys.exit(1)
    pl = pd.DataFrame.from_dict(plants_d, orient="index").sort_index() if plants_d else pd.DataFrame()
    pl.index.name = "date"
    pl = pl.dropna(axis=1, how="all")
    hyd = [c for c in pl.columns if c.rsplit("_", 1)[0] in HYDRO]
    th = [c for c in pl.columns if c not in hyd]
    print("Plants classed hydro:", hyd, "\nThermal:", th, flush=True)
    recs = {}
    for d in pl.index:
        h = hours(d)
        hy, tt = pl.loc[d, hyd].fillna(0).sum(), pl.loc[d, th].fillna(0).sum()
        recs[d] = {"Hydro_GWh": hy * h / 1000, "Gas_GWh": tt * h / 1000, "basis": basis.get(d, "report")}
    for d, (tot, share) in kf.items():
        if d in recs:
            continue
        recs[d] = {"Hydro_GWh": tot * share / 100 if share is not None else None,
                   "Gas_GWh": tot * (1 - share / 100) if share is not None else None,
                   "Total_GWh": tot, "basis": basis.get(d, "key facts (report text)")}
    q = pd.DataFrame.from_dict(recs, orient="index").sort_index()
    q.index.name = "date"
    q["Total_GWh"] = q["Total_GWh"].fillna(q["Hydro_GWh"] + q["Gas_GWh"]) if "Total_GWh" in q else q["Hydro_GWh"] + q["Gas_GWh"]
    q["hours"] = [hours(d) for d in q.index]
    for c in ("Hydro", "Gas", "Total"):
        q[f"{c}_GW"] = q[f"{c}_GWh"] / q["hours"]
    q = q[["Hydro_GWh", "Gas_GWh", "Total_GWh", "Hydro_GW", "Gas_GW", "Total_GW", "hours", "basis"]]
    q = q.round({"Hydro_GWh": 2, "Gas_GWh": 2, "Total_GWh": 2, "Hydro_GW": 4, "Gas_GW": 4, "Total_GW": 4})
    rp = pd.DataFrame(list(rows.values())).sort_values("quarter", ascending=False).reset_index(drop=True)
    xlsx_notes.write_workbook(args.out, {"Quarterly": q, "Plants": pl.round(2), "Reports": rp},
                              NOTES, {ln for ln in NOTES if ln and ln.isupper()})
    print(f"Saved {args.out}: {len(q)} quarters {q.index.min():%Y-%m} to {q.index.max():%Y-%m}", flush=True)
    print(q.to_string(), flush=True)


if __name__ == "__main__":
    main()
