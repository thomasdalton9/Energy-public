"""
India installed generating capacity by fuel, monthly, from CEA's (Central Electricity Authority) monthly
"All India Installed Capacity (Utilities)" report, as published on CEA's National Power Portal (NPP):
  https://npp.gov.in/public-reports/cea/monthly/installcap/YYYY/MON/capacity1-YYYY-MM.xls   (MON = JAN..DEC)
listed at https://npp.gov.in/publishedReports (the same figures as cea.nic.in/installed-capacity-report,
IC_allocation_as_on_*.xlsx, which keeps only the latest month online; NPP keeps every month from 2019 and
answers GitHub's runners). Found via discovery_archive/asia/CAPPRICE_DISCOVERY1-2.py.

Each file: capacity (MW) by region and sector (state / private / central) and mode - Thermal (Coal, Lignite,
Gas, Diesel), Nuclear, Hydro (large hydro, >25 MW) and RES (MNRE renewables) - plus a 'Break up of RES' row:
Small hydro, Wind, Bio-power (biomass/bagasse cogeneration, waste to energy) and Solar. The all-India total
row and the RES break-up are used:
  Coal = coal + lignite; Gas; Oil = diesel; Nuclear; Hydro = large hydro + small hydro; Wind; Solar;
  Bioenergy = biomass / bagasse cogeneration + waste to energy; Other = any RES not in the break-up (normally 0)

Writes output/Data and Chart Outputs/india_power_capacity.xlsx (standard capacity layout,
south_america/power_capacity_std.py):
  Monthly  date (1st of the month the report is 'as on' the end of), <Fuel>_MW, Total_MW, plus detail columns
           Large_hydro_MW, Small_hydro_MW, Lignite_MW, Waste_to_energy_MW, Grand_total_reported_MW (CEA's own
           grand total, a check on the sum)
  Files    month, file URL and its Last-Modified date

Incremental: each month's file is fetched once (months not yet saved, from 2019-01), plus the latest
REVISE_MONTHS again (CEA revises the latest report). NPP posts a month's report about 2-3 weeks after the
month ends. Runs on the 1st and 15th.

    python3 asia/INDIA_CEA_CAPACITY.py [--out PATH]
"""
import argparse
import io
import os
import re
import sys
import time
from datetime import date

import pandas as pd
import requests

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)
sys.path.insert(0, os.path.join(ROOT, "south_america"))
import power_capacity_std as cap_std  # noqa: E402

NPP = "https://npp.gov.in/public-reports/cea/monthly/installcap"
H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0 Safari/537.36"}
T = (20, 120)
START = pd.Timestamp("2019-01-01")
REVISE_MONTHS = 2
OUT = os.path.join(ROOT, "output", "Data and Chart Outputs", "india_power_capacity.xlsx")
MON = ["JAN", "FEB", "MAR", "APR", "MAY", "JUN", "JUL", "AUG", "SEP", "OCT", "NOV", "DEC"]
# header text -> key; regex, first match wins (checked on the column's joined header text, lower case)
MODE_KEYS = [("coal", r"\bcoal\b"), ("lignite", r"lignite"), ("gas", r"\bgas\b"), ("diesel", r"diesel"),
             ("nuclear", r"nuclear"), ("hydro", r"hydro"), ("res", r"\bres\b|res\s*\*|mnre")]
RE_KEYS = [("small_hydro", r"small\s*hydro"), ("wind", r"wind"), ("waste", r"waste"),
           ("bio", r"bm\s*power|bio|bagasse|cogen|congen"), ("solar", r"solar"), ("total", r"total")]
EXTRA = ["Large_hydro_MW", "Small_hydro_MW", "Lignite_MW", "Waste_to_energy_MW", "Grand_total_reported_MW"]


def out(*a):
    print(*a, flush=True)


def url(m):
    return f"{NPP}/{m.year}/{MON[m.month - 1]}/capacity1-{m.year}-{m.month:02d}.xls"


def fetch(m):
    for i in range(3):
        try:
            r = requests.get(url(m), headers=H, timeout=T)
            if r.status_code == 404 or (r.status_code == 200 and r.content[:4] != b"\xd0\xcf\x11\xe0"
                                        and b"<html" in r.content[:500].lower()):
                return None, None
            r.raise_for_status()
            return r.content, r.headers.get("last-modified")
        except requests.RequestException as e:
            if i == 2:
                out(f"  {m:%Y-%m}: {e}")
                return None, None
            time.sleep(5 * (i + 1))


def num(v):
    x = pd.to_numeric(str(v).replace(",", "").strip(), errors="coerce") if not pd.isna(v) else float("nan")
    return float(x)


def labels(df, rows):
    """Joined header text per column over the given rows (lower case)."""
    return {j: " ".join(str(df.iat[i, j]) for i in rows if not pd.isna(df.iat[i, j])).lower()
            for j in range(df.shape[1])}


def value_at(df, i, j, taken=()):
    """Value in row i at column j, or the next non-empty cell to the right (merged header cells) that is not
    another mapped column."""
    for k in range(j, min(j + 3, df.shape[1])):
        if k != j and k in taken:
            break
        x = num(df.iat[i, k])
        if pd.notna(x):
            return x
    return float("nan")


def parse(content):
    df = pd.read_excel(io.BytesIO(content), header=None)
    col1 = df.astype(str).apply(lambda c: c.str.strip().str.lower())
    find = lambda pat: [i for i in range(len(df)) if col1.iloc[i].str.contains(pat, regex=True).any()]  # noqa: E731
    r_all = find(r"^total of all")
    r_head = find(r"^region$")
    r_first = find(r"state sector")
    if not r_all or not r_head or not r_first:
        raise ValueError("layout: all-India total / header row not found")
    r_all, r_head, r_first = r_all[0], r_head[0], r_first[0]
    lab = labels(df, range(r_head, r_first))
    cols = {}
    for key, pat in MODE_KEYS:
        for j, t in lab.items():
            if j not in cols.values() and re.search(pat, t) and "total" not in t.replace("grand total", ""):
                cols[key] = j
                break
    g = [j for j, t in lab.items() if "grand total" in t]
    taken = set(cols.values()) | set(g)
    v = {k: value_at(df, r_all, j, taken) for k, j in cols.items()}
    v["grand"] = value_at(df, r_all, g[0], taken) if g else float("nan")
    # RES break-up block
    rb = [i for i in find(r"small hydro") if i > r_all]
    if rb:
        rlab = labels(df, [rb[0], rb[0] + 1])
        rv = next((i for i in range(rb[0] + 1, min(rb[0] + 6, len(df)))
                   if sum(pd.notna(num(x)) and num(x) > 0 for x in df.iloc[i]) >= 3), None)
        if rv is not None:
            rtaken = {j for j, t in rlab.items() if t.strip()}
            for key, pat in RE_KEYS:
                for j, t in rlab.items():
                    if re.search(pat, t) and f"re_{key}" not in v:
                        if key == "bio" and "waste" in t:
                            continue
                        v[f"re_{key}"] = value_at(df, rv, j, rtaken)
                        break
    return v


def to_row(v):
    z = lambda k: 0.0 if pd.isna(v.get(k, float("nan"))) else v[k]  # noqa: E731
    res = z("res")
    parts = {"Small_hydro": z("re_small_hydro"), "Wind": z("re_wind"), "Solar": z("re_solar"),
             "Bio": z("re_bio"), "Waste": z("re_waste")}
    other = max(0.0, res - sum(parts.values())) if res else 0.0
    fuels = {"Coal": z("coal") + z("lignite"), "Gas": z("gas"), "Oil": z("diesel"), "Nuclear": z("nuclear"),
             "Hydro": z("hydro") + parts["Small_hydro"], "Wind": parts["Wind"], "Solar": parts["Solar"],
             "Bioenergy": parts["Bio"] + parts["Waste"], "Other": other if other > 1 else 0.0}
    # Lignite_MW blank where the report has no lignite column (before Dec 2021 lignite sits inside coal)
    lig = v.get("lignite", float("nan"))
    extra = {"Large_hydro_MW": z("hydro"), "Small_hydro_MW": parts["Small_hydro"],
             "Lignite_MW": lig if pd.notna(lig) and lig > 0 else float("nan"),
             "Waste_to_energy_MW": parts["Waste"], "Grand_total_reported_MW": v.get("grand")}
    return fuels, extra


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", default=OUT)
    a = ap.parse_args()
    monthly = cap_std.load_monthly(a.out)
    if "Lignite_MW" in monthly:   # saved before the fix: 0 = no lignite column in that report, not zero capacity
        monthly["Lignite_MW"] = monthly["Lignite_MW"].mask(monthly["Lignite_MW"] <= 0)
    files = cap_std.load_sheet(a.out, "Files")
    last = pd.Timestamp(date.today().replace(day=1)) - pd.DateOffset(months=1)
    months = pd.date_range(START, last, freq="MS")
    have = set(monthly.index) if not monthly.empty else set()
    recent = set(sorted(have)[-REVISE_MONTHS:]) if have else set()
    todo = [m for m in months if m not in have or m in recent or m > max(have, default=START)]
    out(f"India capacity: {len(have)} months saved; checking {len(todo)}")
    rows, extras, stamps = {}, {}, []
    for m in todo:
        content, lm = fetch(m)
        if content is None:
            out(f"  {m:%Y-%m}: not published")
            continue
        try:
            v = parse(content)
        except Exception as e:  # noqa: BLE001
            out(f"  {m:%Y-%m}: parse failed ({type(e).__name__}: {e})")
            continue
        fuels, extra = to_row(v)
        tot = sum(fuels.values())
        grand = extra["Grand_total_reported_MW"]
        flag = "" if pd.isna(grand) or abs(tot - grand) < 0.005 * grand else f"  !! sum {tot:.0f} vs grand {grand:.0f}"
        if tot < 250000 or tot > 1500000:
            out(f"  {m:%Y-%m}: implausible total {tot:.0f} MW - skipped ({v})")
            continue
        rows[m], extras[m] = fuels, extra
        stamps.append({"month": m, "url": url(m), "last_modified": lm})
        out(f"  {m:%Y-%m}: total {tot / 1000:.1f} GW (coal {fuels['Coal'] / 1000:.1f}, solar {fuels['Solar'] / 1000:.1f}, "
            f"wind {fuels['Wind'] / 1000:.1f}){flag}")
        time.sleep(0.3)
    if rows:
        new = cap_std.standard(pd.DataFrame.from_dict(rows, orient="index"))
        new = new.join(pd.DataFrame.from_dict(extras, orient="index").round(1))
        monthly = pd.concat([monthly[~monthly.index.isin(new.index)], new]).sort_index() if not monthly.empty else new
        st = pd.DataFrame(stamps)
        files = pd.concat([files[~files["month"].isin(st["month"])], st]) if not files.empty and "month" in files else st
        files = files.sort_values("month")
    if monthly.empty:
        raise SystemExit("No India capacity data")
    monthly.index.name = "date"
    monthly = monthly[cap_std.COLUMNS + [c for c in EXTRA if c in monthly]]
    notes = [
        "UNITS",
        "Monthly: installed capacity, MW, all India (utilities), as on the last day of the month (row dated the 1st). "
        "Coal = coal + lignite; Gas; Oil = diesel; Nuclear; Hydro = large hydro (CEA 'Hydro', >25 MW) + small hydro "
        "(MNRE, <=25 MW); Wind; Solar (ground-mounted, rooftop and off-grid solar as MNRE reports it); Bioenergy = "
        "biomass/bagasse cogeneration + waste to energy; Other = any RES not covered by the break-up. Total_MW = sum of "
        "fuels; Grand_total_reported_MW = CEA's own grand total (a check). Detail columns split hydro and show lignite "
        "and waste to energy. Lignite_MW is blank before Dec 2021: earlier reports have no lignite column (lignite "
        "is inside coal there, so Coal_MW is consistent throughout).",
        "Mar 2025: CEA's report drops gas by 5.06 GW (25,188 -> 20,132 MW), coal by 2.8 GW and nuclear by 0.1 GW "
        "while solar / wind / hydro grow, so the total falls 2.6 GW. This is CEA's own revision (retired / "
        "de-rated units taken out), not a parse error: every month's fuel sum equals CEA's reported grand total.",
        "Files: the NPP file read for each month and its Last-Modified date.",
        "",
        "COVERAGE",
        f"Monthly from {monthly.index.min():%b %Y} to {monthly.index.max():%b %Y}. Utilities only (captive plants "
        "are not included). Renewable capacity is MNRE's figure as carried in CEA's report. A month's report appears "
        "on NPP about 2-3 weeks after the month ends.",
        "",
        "SOURCE",
        "CEA (Central Electricity Authority, Ministry of Power), All India Installed Capacity (Utilities) monthly report, "
        "via the National Power Portal: https://npp.gov.in/publishedReports "
        "(files https://npp.gov.in/public-reports/cea/monthly/installcap/YYYY/MON/capacity1-YYYY-MM.xls); also published "
        "at https://cea.nic.in/installed-capacity-report/?lang=en.",
    ]
    cap_std.write(a.out, monthly, {"Files": files.set_index("month") if not files.empty and "month" in files else files},
                  notes, {"UNITS", "COVERAGE", "SOURCE"})


if __name__ == "__main__":
    main()
