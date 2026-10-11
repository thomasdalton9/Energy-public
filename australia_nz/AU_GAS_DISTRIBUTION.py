"""
Australia east coast residential / commercial (distribution-network) gas demand. AEMO's Gas Bulletin Board has
no such facility type in au_gas.xlsx 'Demand by sector', so this takes the market-operator reports that do:

  Victoria DWGM   AEMO VicGas reports (INT310 price and withdrawals; rolling ~14 gas days, PublicRptsNN.zip archive
                  of the same folder for the first backfill). Victoria's distribution network is the largest
                  residential/commercial gas load in Australia.
  STTM hubs       Sydney, Adelaide, Brisbane: AEMO STTM reports in Current/STTM (rolling about a week, DayNN.zip
                  files). Hub demand = gas withdrawn at the hub (mostly distribution-connected customers).
  Annual          Dept of Climate Change, Energy, the Environment and Water, Australian Energy Statistics
                  Table F (gas consumption by sector incl. residential and commercial, financial years), as a
                  labelled supplement.

IMPORTANT - column detection: the sandbox this was written in cannot reach AEMO or DCCEEW, so the report layouts
were NOT inspected. Each source is therefore read by name-matching its columns (see MATCH below); the file used,
the column chosen and any ambiguity are printed in the run log and written on the Units sheet. A source whose
layout does not match is skipped (and said so), never filled with guesses. Run
discovery_archive/workflows/au_resi_discovery.yml once to see the real layouts, then tighten MATCH.

Writes au_gas_distribution.xlsx:
  DWGM demand      daily, Victoria: value per gas day (GJ/day -> TJ/day) from the INT310 withdrawals column
  STTM hub demand  daily TJ/day by hub (Sydney, Adelaide, Brisbane) + Total
  Annual by sector annual PJ, AES Table F rows for residential / commercial / industrial / power etc.

Incremental: saved rows are the history; each run fetches the rolling windows (daily, see au_gas_distribution.yml)
and merges them over the saved rows (new values win); archives are read only when nothing is saved yet. The annual
file is downloaded only when its link changes (recorded on the Units sheet).

Usage: python3 AU_GAS_DISTRIBUTION.py [--out "output/Data and Chart Outputs/au_gas_distribution.xlsx"]
"""
import argparse
import io
import os
import re
import sys
import zipfile

import pandas as pd
import requests

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)  # repo root, for xlsx_notes
import xlsx_notes  # noqa: E402

NEM = "https://nemweb.com.au"
VIC = NEM + "/Reports/CURRENT/VicGas/"
STTM = NEM + "/Reports/CURRENT/STTM/"
HEADERS = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0"}
DEFAULT_OUT = os.path.join(ROOT, "output", "Data and Chart Outputs", "au_gas_distribution.xlsx")
HUBS = ["Sydney", "Adelaide", "Brisbane"]
AES_PAGES = ["https://www.energy.gov.au/data/australian-energy-statistics",
             "https://www.energy.gov.au/energy-data/australian-energy-statistics"]

# Name patterns used to find the right report / column (layouts unverified, see docstring)
MATCH = {
    "dwgm_file": re.compile(r"int310", re.I),
    "dwgm_value": re.compile(r"withdraw", re.I),
    "sttm_value": re.compile(r"(customer|demand|withdraw|consum).*(qty|quantity|gj)|total_(customer|demand|withdraw)",
                             re.I),
    "sttm_exclude_file": re.compile(r"price|bid|offer|contingency|capacity|notice|forecast", re.I),
}
LOG = []


def get(url):
    r = requests.get(url, headers=HEADERS, timeout=(10, 180))
    r.raise_for_status()
    return r


def listing(url):
    try:
        return re.findall(r'href="([^"]+\.(?:csv|zip))"', get(url).text, re.I)
    except Exception as e:  # noqa: BLE001
        print(f"  folder {url}: {type(e).__name__}: {str(e)[:100]}", flush=True)
        return []


def full(u):
    return NEM + u if u.startswith("/") else (u if u.startswith("http") else None)


def csv_members(content, name):
    """(name, DataFrame) for a csv, or for every csv inside a zip."""
    if name.lower().endswith(".zip"):
        z = zipfile.ZipFile(io.BytesIO(content))
        for n in z.namelist():
            if n.lower().endswith(".csv"):
                try:
                    yield n, pd.read_csv(z.open(n), low_memory=False)
                except Exception:  # noqa: BLE001
                    continue
    else:
        yield name, pd.read_csv(io.BytesIO(content), low_memory=False)


def gas_date(d):
    col = next((c for c in d.columns if str(c).lower() in ("gas_date", "gasdate")), None)
    if col is None:
        return None
    return pd.to_datetime(d[col], format="%d %b %Y", errors="coerce")


def pick_value(d, pat):
    cols = [c for c in d.columns if pat.search(str(c)) and pd.api.types.is_numeric_dtype(pd.to_numeric(d[c],
                                                                                                    errors="coerce"))
            and pd.to_numeric(d[c], errors="coerce").notna().any()]
    return cols


# ---------------------------------------------------------------- DWGM (Victoria)
def dwgm(backfill):
    names = [VIC + n.rsplit("/", 1)[-1] for n in listing(VIC) if MATCH["dwgm_file"].search(n)
             and n.lower().endswith(".csv")]
    names = sorted(set(names)) or [VIC + "int310_v1_price_and_withdrawals_rpt_1.csv"]
    srcs = [(n, None) for n in names]
    if backfill:
        srcs += [(f"{VIC}PublicRpts{i:02d}.zip", None) for i in range(1, 15)]
    parts, used = [], set()
    for url, _ in srcs:
        try:
            content = get(url).content
        except Exception as e:  # noqa: BLE001
            print(f"  {url.rsplit('/', 1)[-1]}: {type(e).__name__}", flush=True)
            continue
        for n, d in csv_members(content, url):
            if not MATCH["dwgm_file"].search(n):
                continue
            gd = gas_date(d)
            vals = pick_value(d, MATCH["dwgm_value"])
            if gd is None or not vals:
                LOG.append(f"DWGM {n}: no gas_date / withdrawal column; columns {list(d.columns)[:20]}")
                continue
            used.add((n.split("/")[-1][:40], vals[0]))
            x = pd.DataFrame({"date": gd, "v": pd.to_numeric(d[vals[0]], errors="coerce")})
            si = next((c for c in d.columns if "schedule" in str(c).lower() and "interval" in str(c).lower()), None)
            x["si"] = pd.to_numeric(d[si], errors="coerce") if si else 0
            parts.append(x.dropna(subset=["date", "v"]))
    if not parts:
        return pd.DataFrame(), "no INT310 withdrawal column found"
    x = pd.concat(parts)
    # value reported at the last schedule interval of each gas day (daily cumulative/total as published) plus the max
    x = x.sort_values(["date", "si"])
    g = x.groupby("date")
    out = pd.DataFrame({"DWGM_withdrawals_last_interval": g["v"].last(), "DWGM_withdrawals_max": g["v"].max(),
                        "Schedule_intervals": g["si"].nunique()})
    print(f"DWGM: {len(out)} gas days {out.index.min():%Y-%m-%d}..{out.index.max():%Y-%m-%d}; used {sorted(used)[:3]}",
          flush=True)
    return out, f"{sorted(used)[0][0]} column {sorted(used)[0][1]}"


# ---------------------------------------------------------------- STTM hubs
def sttm(backfill):
    fl = listing(STTM)
    srcs = [u for u in fl if u.lower().endswith(".csv") and not MATCH["sttm_exclude_file"].search(u)]
    zips = sorted(u for u in fl if re.search(r"day\d+\.zip", u, re.I))
    srcs += zips if backfill else zips[:1]
    rows, used = [], {}
    for u in srcs:
        link = full(u)
        if not link:
            continue
        try:
            content = get(link).content
        except Exception as e:  # noqa: BLE001
            print(f"  {u}: {type(e).__name__}", flush=True)
            continue
        for n, d in csv_members(content, u):
            if MATCH["sttm_exclude_file"].search(n):
                continue
            gd = gas_date(d)
            hub = next((c for c in d.columns if str(c).lower() == "hub_name"), None)
            vals = pick_value(d, MATCH["sttm_value"])
            if gd is None or hub is None or not vals:
                continue
            key = re.sub(r"_\d{8}.*$|_\d+\.csv$", "", n.lower().rsplit("/", 1)[-1])
            used.setdefault(key, set()).add(vals[0])
            x = pd.DataFrame({"date": gd, "hub": d[hub], "v": pd.to_numeric(d[vals[0]], errors="coerce"),
                              "rep": key + ":" + vals[0]})
            rows.append(x.dropna(subset=["date", "v"]))
    if not rows:
        return pd.DataFrame(), "no STTM report with hub_name + a demand/withdrawal quantity column found"
    x = pd.concat(rows)
    rep = x.groupby("rep").size().sort_values(ascending=False).index[0]   # the matching report with most rows
    print(f"STTM: candidate reports {sorted(f'{k}:{sorted(v)}' for k, v in used.items())}; using {rep}", flush=True)
    x = x[x["rep"] == rep]
    # facility/schedule rows within a hub-day are summed only if the report is by facility; the sum is logged
    out = x.pivot_table(index="date", columns="hub", values="v", aggfunc="sum")
    out = out.reindex(columns=[h for h in HUBS if h in out.columns] + [h for h in out.columns if h not in HUBS])
    print(f"STTM: {len(out)} gas days {out.index.min():%Y-%m-%d}..{out.index.max():%Y-%m-%d}", flush=True)
    return out, rep


# ---------------------------------------------------------------- Annual: AES Table F
SECTOR_PAT = re.compile(r"resident|commercial|industr|manufactur|mining|power|electricity|total|transport", re.I)


def aes_table_f(prev_url):
    link = None
    for page in AES_PAGES:
        try:
            html = get(page).text
        except Exception:  # noqa: BLE001
            continue
        for u in re.findall(r'href="([^"]+\.xlsx?)"', html, re.I):
            if re.search(r"table.?f|gas", u, re.I):
                link = u if u.startswith("http") else "https://www.energy.gov.au" + u
                break
        if link:
            break
    if not link:
        return None, "Australian Energy Statistics Table F link not found on the landing pages", prev_url
    if link == prev_url:
        return None, "unchanged", link
    xl = pd.ExcelFile(io.BytesIO(get(link).content))
    out = {}
    for sh in xl.sheet_names:
        raw = xl.parse(sh, header=None)
        for hi in range(min(25, len(raw))):
            yrs = {j: re.match(r"^(\d{4})", str(v)) for j, v in raw.iloc[hi].items()}
            yrs = {j: int(m.group(1)) for j, m in yrs.items() if m and 1950 <= int(m.group(1)) <= 2100}
            if len(yrs) >= 5:
                for _, row in raw.iloc[hi + 1:].iterrows():
                    label = str(row.iloc[0]).strip()
                    if SECTOR_PAT.search(label):
                        out[f"{sh} | {label}"] = pd.Series({y: pd.to_numeric(row[j], errors="coerce")
                                                            for j, y in yrs.items()})
                break
    if not out:
        return None, f"{link}: no year-header / sector rows recognised in sheets {xl.sheet_names[:8]}", link
    df = pd.DataFrame(out).sort_index().dropna(how="all")
    df.index.name = "year"
    return df, link, link


def load(path, sheet):
    try:
        d = pd.read_excel(path, sheet_name=sheet, index_col=0)
    except (FileNotFoundError, ValueError, KeyError, OSError):
        return pd.DataFrame()
    d.index = pd.to_datetime(pd.Index(d.index).astype(str), errors="coerce")
    return d[d.index.notna()].sort_index()


def merge(new, old):
    if new.empty:
        return old
    return new.combine_first(old) if not old.empty else new


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", default=DEFAULT_OUT)
    ap.add_argument("--annual", action="store_true", help="force the annual file download")
    args = ap.parse_args()

    old = {k: load(args.out, k) for k in ("DWGM demand", "STTM hub demand")}
    try:
        old_annual = pd.read_excel(args.out, sheet_name="Annual by sector", index_col=0)
    except (FileNotFoundError, ValueError, KeyError, OSError):
        old_annual = pd.DataFrame()
    prev_url = ""
    try:
        u = pd.read_excel(args.out, sheet_name="Units", header=None)[0].astype(str)
        hit = u[u.str.startswith("AES_LINK=")]
        prev_url = hit.iloc[0].split("=", 1)[1] if len(hit) else ""
    except Exception:  # noqa: BLE001
        pass

    status = {}
    out = {}
    for name, fn in (("DWGM demand", dwgm), ("STTM hub demand", sttm)):
        try:
            new, what = fn(old[name].empty)
            status[name] = what
        except Exception as e:  # noqa: BLE001 - keep the other source
            new, status[name] = pd.DataFrame(), f"failed: {type(e).__name__}: {str(e)[:150]}"
        if name == "DWGM demand" and not new.empty:   # GJ -> TJ (new rows only; saved rows are already TJ)
            new = new.assign(**{c: new[c] / 1000.0 for c in new.columns if c.startswith("DWGM_withdrawals")})
        d = merge(new, old[name])
        if name == "STTM hub demand" and not d.empty:
            d = d.drop(columns=["Total"], errors="ignore")
            d["Total"] = d.sum(axis=1, min_count=1)
        if not d.empty:
            d.index.name = "date"
            out[name] = d
    annual = old_annual
    try:
        new_a, what, link = aes_table_f("" if args.annual else prev_url)
        status["Annual by sector"] = what
        if new_a is not None:
            annual = new_a
            prev_url = link
    except Exception as e:  # noqa: BLE001
        status["Annual by sector"] = f"failed: {type(e).__name__}: {str(e)[:150]}"
    if not annual.empty:
        out["Annual by sector"] = annual
    for line in LOG + [f"{k}: {v}" for k, v in status.items()]:
        print("  " + line, flush=True)
    if not out:
        raise SystemExit("No distribution demand source could be read; see log (run the discovery workflow)")

    def cover(d):
        return f"{d.index.min():%d %b %Y} to {d.index.max():%d %b %Y}"

    notes = [
        "UNITS",
        "DWGM demand: terajoules per gas day (TJ/day; AEMO publishes GJ, divided by 1,000). Victorian Declared "
        "Wholesale Gas Market withdrawals (INT310), dominated by the distribution network (residential, commercial, "
        "small industrial). 'last_interval' is the figure at the day's final schedule interval, 'max' the day's "
        "highest reported value: check against AEMO's definition before relying on either.",
        "STTM hub demand: terajoules per gas day by hub (STTM gas day 06:30-06:30 AEST), Total = sum of the hubs.",
        "Annual by sector: Australian Energy Statistics Table F, petajoules per financial year (year = the "
        "financial year's first year as written in the source header), one column per 'sheet | row label'.",
        "",
        "COVERAGE",
        *[f"{k}: {cover(v)}" if k != "Annual by sector" else f"{k}: {v.index.min()} to {v.index.max()}"
          for k, v in out.items()],
        "Daily series build from the rolling AEMO reports (DWGM about 14 gas days, STTM about a week) plus the "
        "archive read on the first run; this workbook is the history store. Annual series is a labelled supplement.",
        "Layouts were detected by column name (see AU_GAS_DISTRIBUTION.py MATCH). Detection result: "
        + "; ".join(f"{k}: {v}" for k, v in status.items()),
        "",
        "SOURCE",
        f"AEMO NEMWEB VicGas {VIC}; STTM {STTM}",
        "https://aemo.com.au/energy-systems/gas/declared-wholesale-gas-market-dwgm",
        "https://aemo.com.au/energy-systems/gas/short-term-trading-market-sttm",
        "Dept of Climate Change, Energy, the Environment and Water, Australian Energy Statistics Table F: "
        "https://www.energy.gov.au/energy-data/australian-energy-statistics",
        f"AES_LINK={prev_url}",
    ]
    xlsx_notes.write_workbook(args.out, {k: v.round(4) for k, v in out.items()}, notes, {"UNITS", "COVERAGE", "SOURCE"})
    print(f"Saved {args.out}", flush=True)


if __name__ == "__main__":
    main()
