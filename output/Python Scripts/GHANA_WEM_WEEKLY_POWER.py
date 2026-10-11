"""
Ghana daily power generation by plant, from the Energy Commission's weekly Wholesale Electricity Market (WEM)
statistics PDFs.

GRIDCo (the grid operator) and VRA publish no dispatch data table. The Energy Commission (energycom.gov.gh,
the electricity regulator) publishes every week a PDF 'Weekly Wholesale Electricity Market (WEM) Statistics' made
from the market operator's dispatch data. Its table 'Power Plant generation (GWh)' gives GWh per DAY for each
plant: Akosombo, Kpong, Bui (hydro), Bui Solar, VRA Kaleo (solar), and the thermal plants (SEAP/Sunon Asogli, TAPCO,
TICO, TT1PP, TT2PP, CENIT, Amandi, Karpowership, AMERI, KTPP, Cenpower, AKSA, Bridge Power, Genser, ...) and Imports,
with a Total column and Total row (used to check every table).
  Listing:  https://www.energycom.gov.gh/index.php/planning/weekly-wholesale-electricity-market-wem-statistics
  Year pages: .../category/<id>-<year>?limit=100 ; PDFs are '...?download=<id>:<slug>'.

Incremental: the Files sheet of the committed workbook records every PDF (download id) already read; only new ids
are downloaded, plus the latest two (the Commission re-issues the newest weeks) and any that failed before. A day that
appears in two PDFs takes the newer PDF.

Fuel: the PDFs name plants, not fuels, so plants are grouped Hydro / Solar / Thermal (gas, light crude, heavy fuel -
not split) / Import. 'Daily' = GWh per day by group; 'Plants_GWh' = GWh per day per plant as published.

Usage: python3 GHANA_WEM_WEEKLY_POWER.py [--out PATH] [--from-year 2022] [--budget-min 40] [--dump-failed DIR]
"""
print("STARTING", flush=True)

import argparse
import os
import re
import sys
import time

import pandas as pd
import requests

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.dirname(HERE))  # repo root, for xlsx_notes
sys.path.insert(0, HERE)

import ghana_wem_parse as wem  # noqa: E402
import xlsx_notes  # noqa: E402

OUT = "output/Data and Chart Outputs/ghana_wem_weekly_generation_daily.xlsx"
EC = "https://www.energycom.gov.gh"
LANDING = EC + "/index.php/planning/weekly-wholesale-electricity-market-wem-statistics"
H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36"}
HYDRO = {"AKOSOMBO", "KPONG", "BUI"}
SOLAR = {"BUI SOLAR", "VRA KALEO", "NAVRONGO", "LAWRA", "BXC SOLAR", "MEINERGY"}
ALIAS = {"SAPP": "SEAP", "AKSA ANWOMASO": "AKSA Anwomaso"}   # 'SAPP' (2022 typo for SEAP): same slot, totals check
KEEP_LATEST = 2


def canon(name):
    """Plant name as published, whitespace-normalised, with fixed spellings for the hydro / solar / import rows."""
    n = " ".join(name.split())
    u = n.upper()
    if u in ALIAS:
        return ALIAS[u]
    if u in HYDRO:
        return u
    if u in SOLAR:
        return {"BUI SOLAR": "BUI Solar", "VRA KALEO": "VRA Kaleo"}.get(u, n.title())
    if u == "IMPORT":
        return "Import"
    return n


def group(plant):
    u = plant.upper()
    if u in HYDRO:
        return "Hydro"
    if u in SOLAR:
        return "Solar"
    if u == "IMPORT":
        return "Import"
    return "Thermal"


def repair_dates(df):
    """The header days of a table are consecutive; a mistyped year or day in one header (seen: 2021, 2027, 2028 in a
    2025 file) is replaced by the day implied by the other headers. Returns (df, number repaired); ValueError when the
    headers do not agree on any anchor."""
    cols = list(df.columns)
    offs = pd.Series([d - pd.Timedelta(days=i) for i, d in enumerate(cols)])
    anchor = offs.mode().iloc[0]
    if (offs == anchor).sum() < max(4, len(cols) // 2 + 1):
        raise ValueError("header days not consecutive")
    fixed = [anchor + pd.Timedelta(days=i) for i in range(len(cols))]
    n = sum(1 for a, b in zip(cols, fixed) if a != b)
    out = df.copy()
    out.columns = pd.DatetimeIndex(fixed)
    return out, n


def in_window(first, last, year):
    """A weekly file belongs to its year's page, give or take the week across New Year."""
    return (first >= pd.Timestamp(year - 1, 12, 20)) and (last <= pd.Timestamp(year + 1, 1, 10))


MONTH_NAMES = ["january", "february", "march", "april", "may", "june", "july", "august", "september", "october", "november", "december"]


def fix_year(df, name, year):
    """Some weeks print a wrong year in every header day (e.g. February 2022 headed '1-Feb-21'). When the file name names
    the month (and the days are not already inside the page's year window) the year of the headers is replaced by the
    year in the file name, or the page's year. Returns (df, True) when changed."""
    if in_window(df.columns.min(), df.columns.max(), year):
        return df, False
    low = name.lower()
    mon = [i + 1 for i, m in enumerate(MONTH_NAMES) if m in low]
    if len(mon) != 1:
        raise ValueError(f"days {df.columns.min():%Y-%m-%d}..{df.columns.max():%Y-%m-%d} outside the {year} page")
    yr = re.search(r"(20\d\d)", name)
    target = int(yr.group(1)) if yr else year
    fixed = []
    for d in df.columns:
        if d.month != mon[0]:
            raise ValueError(f"header month {d:%b} does not match file name month")
        fixed.append(pd.Timestamp(target, d.month, d.day))
    out = df.copy()
    out.columns = pd.DatetimeIndex(fixed)
    if not in_window(out.columns.min(), out.columns.max(), year):
        raise ValueError("days still outside the page's year after correction")
    return out, True


def categories():
    """[(year, category path)] from the landing page."""
    r = requests.get(LANDING, headers=H, timeout=(10, 60))
    r.raise_for_status()
    found = {}
    for m in re.finditer(r"weekly-wholesale-electricity-market-wem-statistics/category/(\d+)-(\d{4})", r.text):
        found[int(m.group(2))] = f"{m.group(1)}-{m.group(2)}"
    return sorted(found.items())


def list_year(cat):
    url = f"{LANDING}/category/{cat}?limit=100"
    r = requests.get(url, headers=H, timeout=(10, 60))
    r.raise_for_status()
    seen, out = set(), []
    for m in re.finditer(r'href="([^"]*\?download=(\d+):[^"]*)"', r.text):
        i = int(m.group(2))
        if i in seen:
            continue
        seen.add(i)
        out.append((i, m.group(1).replace("&amp;", "&").split("&start=")[0]))
    return out


def read_pdf(content):
    import fitz
    doc = fitz.open(stream=content, filetype="pdf")
    last = None
    for page in doc:
        words = [[w[0], w[1], w[2], w[3], w[4]] for w in page.get_text("words")]
        try:
            return wem.parse_words(words)
        except ValueError as e:
            last = e
    raise ValueError(str(last) if last else "empty pdf")


def load(path, sheet, dates=True):
    try:
        df = pd.read_excel(path, sheet_name=sheet, index_col=0)
    except (FileNotFoundError, ValueError, KeyError, OSError):
        return pd.DataFrame()
    if dates:
        df.index = pd.to_datetime(df.index, errors="coerce")
        df = df[df.index.notna()]
    return df


NOTES = [
    "UNITS",
    "Plants_GWh: GWh generated per DAY by each plant exactly as the weekly table 'Power Plant generation (GWh)' prints it "
    "(names as published; 'SAPP' in 2022 is read as SEAP). 'Import' is electricity imported into Ghana. Blank = the plant is "
    "not in that week's table (plants and names change over time).",
    "Daily: GWh per day by group - Hydro (Akosombo, Kpong, Bui), Solar (Bui Solar, VRA Kaleo), Thermal (all other plants: gas, "
    "light crude and heavy fuel oil; the PDFs do not give fuels), Generation (= Hydro + Solar + Thermal, excludes Import), Import, "
    "Total (= Generation + Import = the table's Total).",
    "Files: every PDF read (download id, year, file name, first and last day, status). A table whose plant rows do not add up to "
    "its own Total row (more than 0.1 GWh + 0.5% on any day) is NOT used and is listed as 'check failed'.",
    "",
    "COVERAGE",
    "Daily from the first week read (2022 by default) to the latest week the Commission has posted (the Commission posts the "
    "weeks in batches, often a month or more late; the 2026 list ended at June week 4 on 11-Oct-2026). Weeks it never posted are "
    "gaps, not filled. Plants outside the dispatch table (rooftop and other embedded solar, mines' own plants) are not in it.",
    "",
    "SOURCE",
    f"Energy Commission of Ghana, Weekly Wholesale Electricity Market (WEM) Statistics (PDF reports): {LANDING}",
    "Pulled on the 1st and 15th by GitHub Actions (ghana_wem_weekly.yml).",
]
TITLES = ["UNITS", "COVERAGE", "SOURCE"]


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--out", default=OUT)
    ap.add_argument("--from-year", type=int, default=2022)
    ap.add_argument("--budget-min", type=float, default=40)
    ap.add_argument("--dump-failed", default="", help="directory to save PDFs that fail to parse (for debugging)")
    args = ap.parse_args()
    t0 = time.time()

    files = load(args.out, "Files", dates=False)
    plants = load(args.out, "Plants_GWh")
    done = set()
    if not files.empty:
        okf = files[files["status"] == "ok"]
        good = [i for i, r in okf.iterrows()
                if pd.notna(r["first_day"]) and in_window(pd.Timestamp(r["first_day"]), pd.Timestamp(r["last_day"]), int(r["year"]))]
        done = set(good)   # files saved with days outside their year (mistyped headers) are read again
    if not plants.empty:   # drop days from the same mistyped headers saved by an earlier run
        plants = plants[(plants.index >= pd.Timestamp(args.from_year - 1, 12, 20)) & (plants.index <= pd.Timestamp.now().normalize())]
    cats = [(y, c) for y, c in categories() if y >= args.from_year]
    print("Year pages:", cats, flush=True)
    todo, order = [], []
    for y, c in cats:
        for i, u in list_year(c):
            order.append(i)
            todo.append((i, y, u))
    newest = set(sorted(order)[-KEEP_LATEST:])
    todo = [t for t in todo if t[0] not in done or t[0] in newest]
    print(f"{len(order)} PDFs listed, {len(todo)} to read", flush=True)

    rec = {} if files.empty else files.to_dict("index")
    rows = {}   # date -> {plant: GWh} from this run, id order so later PDFs win
    new_cols = set()
    for n, (i, y, u) in enumerate(sorted(todo)):
        if time.time() - t0 > args.budget_min * 60:
            print("Time budget reached; the next run continues.", flush=True)
            break
        url = EC + u if u.startswith("/") else u
        r, name = None, ""
        try:
            r = requests.get(url, headers=H, timeout=(10, 60))
            r.raise_for_status()
            name = re.findall(r'filename="?([^";]+)', r.headers.get("content-disposition", ""))
            name = name[0] if name else ""
            if r.content[:4] != b"%PDF":
                raise ValueError("not a PDF")
            df, tot, warns = read_pdf(r.content)
            df.index = [canon(x) for x in df.index]
            df, n_fix = repair_dates(df)
            tot = pd.Series(tot.values, index=df.columns) if (tot is not None and len(tot) == df.shape[1]) else tot
            df, yfix = fix_year(df, name, y)
            tot = pd.Series(tot.values, index=df.columns) if (tot is not None and len(tot) == df.shape[1]) else tot
            if df.index.duplicated().any():
                raise ValueError("duplicate plant rows")
            s = df.sum()
            if tot is not None:
                tot = tot[~tot.index.duplicated()]
                diff = (s.reindex(tot.index) - tot).abs()
                bad = bool((diff > 0.1 + 0.005 * tot.abs()).any() or diff.isna().any())
            else:
                bad, diff = False, pd.Series([0.0])
            if bad:
                rec[i] = {"year": y, "file": name, "first_day": df.columns.min(), "last_day": df.columns.max(),
                          "status": "check failed", "note": f"max diff vs Total row {diff.max():.2f} GWh"}
                print(f"  {i} {name}: check failed", flush=True)
                continue
            for d in df.columns:
                rows.setdefault(d, {}).update({p: df.at[p, d] for p in df.index})
            rec[i] = {"year": y, "file": name, "first_day": df.columns.min(), "last_day": df.columns.max(),
                      "status": "ok", "note": "; ".join(([f"{n_fix} header day(s) repaired"] if n_fix else []) + (["header year corrected from the file name"] if yfix else []) + warns)[:200]}
            print(f"  [{time.time()-t0:.0f}s] {i} {name}: {df.shape[0]} plants x {df.shape[1]} days {df.columns.min():%Y-%m-%d}..{df.columns.max():%Y-%m-%d}", flush=True)
        except Exception as e:  # noqa: BLE001
            rec[i] = {"year": y, "file": name, "first_day": pd.NaT, "last_day": pd.NaT, "status": "failed",
                      "note": f"{type(e).__name__}: {str(e)[:150]}"}
            print(f"  {i}: FAILED {type(e).__name__}: {str(e)[:150]}", flush=True)
            if args.dump_failed:
                os.makedirs(args.dump_failed, exist_ok=True)
                try:
                    open(os.path.join(args.dump_failed, f"{i}.pdf"), "wb").write(r.content)  # noqa
                except Exception:  # noqa: BLE001
                    pass

    if not rows and plants.empty:
        print("Nothing parsed.", flush=True)
        sys.exit(1)
    new = pd.DataFrame.from_dict(rows, orient="index").sort_index() if rows else pd.DataFrame()
    if not new.empty:
        new.index = pd.DatetimeIndex(new.index)
        old = plants[~plants.index.isin(new.index)] if not plants.empty else plants
        # a re-read day replaces the whole saved day (plants can be renamed)
        plants = pd.concat([old, new]).sort_index()
    plants.index.name = "date"
    plants = plants.astype(float)
    cols = {g: [c for c in plants.columns if group(c) == g] for g in ("Hydro", "Solar", "Thermal", "Import")}
    daily = pd.DataFrame({f"{g}_GWh": plants[c].sum(axis=1, min_count=1) if c else float("nan") for g, c in cols.items()})
    daily["Generation_GWh"] = daily[["Hydro_GWh", "Solar_GWh", "Thermal_GWh"]].sum(axis=1, min_count=1)
    daily["Total_GWh"] = daily[["Generation_GWh", "Import_GWh"]].sum(axis=1, min_count=1)
    daily = daily.round(3)
    daily.index.name = "date"
    files_out = pd.DataFrame.from_dict(rec, orient="index").sort_index()
    files_out.index.name = "download_id"
    os.makedirs(os.path.dirname(args.out) or ".", exist_ok=True)
    xlsx_notes.write_workbook(args.out, {"Daily": daily, "Plants_GWh": plants.round(3), "Files": files_out}, NOTES, TITLES)
    print(f"Saved {len(daily)} days {daily.index.min():%Y-%m-%d}..{daily.index.max():%Y-%m-%d}; files ok {int((files_out.status == 'ok').sum())}, "
          f"failed {int((files_out.status == 'failed').sum())}, check failed {int((files_out.status == 'check failed').sum())}", flush=True)
    m = daily.resample("MS").agg(["sum", "count"])
    print((daily[["Hydro_GWh", "Thermal_GWh", "Solar_GWh", "Import_GWh", "Total_GWh"]].resample("MS").sum() / 1).tail(12).round(0).to_string(), flush=True)


if __name__ == "__main__":
    main()
