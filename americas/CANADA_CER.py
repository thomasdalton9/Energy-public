"""
Canada Energy Regulator (CER) open data (no key), two workbooks:

  canada_cer_gas.xlsx    "Gas trade monthly": natural gas exports (and imports where the file has them) by region,
                         million m3 per day, from the CER's monthly export/import CSVs.
  canada_cer_power.xlsx  "Electricity trade monthly": electricity exports and imports by province, GWh per month
                         (CER electricity-exports-and-imports-monthly.csv);
                         "Generation annual": generation by province and fuel, GWh per YEAR, from Canada's Energy
                         Future 2026 (historical years of the 'Current Measures' scenario, then its projection);
                         "Fossil shares": each province's coal / gas / oil share of fossil generation by year,
                         used by NORTH_AMERICA_MASTER to split StatCan's monthly fossil generation beyond the last
                         year of StatCan's own annual fuel table.

The CER files are whole-file downloads with no history query: each is downloaded only when its Last-Modified header is
newer than the one recorded on the Units sheet.

Usage: python3 CANADA_CER.py [--gas-out ...] [--power-out ...]
"""
import argparse
import io
import os
import re
import sys

import pandas as pd
import requests

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)  # repo root, for xlsx_notes
import xlsx_notes  # noqa: E402

DEFAULT_DIR = os.path.join(ROOT, "output", "Data and Chart Outputs")
BASE = "https://www.cer-rec.gc.ca/open/"
HEADERS = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) "
                         "Chrome/124.0 Safari/537.36"}
GAS_FILES = ["imports-exports/natural-gas-exports-monthly.csv", "imports-exports/natural-gas-exports-and-imports-monthly.csv"]
ELEC_FILE = "imports-exports/electricity-exports-and-imports-monthly.csv"
EF_FILE = "energy/energyfutures2026/electricity-generation-2026.csv"
EF_SCENARIO = "Current Measures"
FOSSIL = {"Coal & Coke": "Coal", "Natural Gas": "Gas", "Oil": "Oil"}


def fetch(path):
    """(text, last-modified) of one CER CSV; HTML error pages (the CER answers 200 for unknown files) are rejected."""
    r = requests.get(BASE + path, headers=HEADERS, timeout=(10, 300))
    r.raise_for_status()
    if "html" in r.headers.get("content-type", "") or r.content[:1] == b"<":
        raise ValueError(f"{path}: server returned an HTML page, not a CSV")
    try:
        text = r.content.decode("utf-8-sig")
    except UnicodeDecodeError:   # the electricity file is Windows-1252 (Qu\xe9bec)
        text = r.content.decode("cp1252", "replace")
    return text, r.headers.get("last-modified", "")


def last_modified(path):
    try:
        r = requests.head(BASE + path, headers=HEADERS, timeout=(10, 60), allow_redirects=True)
        return r.headers.get("last-modified", "")
    except requests.RequestException:
        return ""


def read_state(path):
    state = {}
    try:
        for line in pd.read_excel(path, sheet_name="Units")["Notes"].astype(str):
            m = re.match(r"lm\[(.+?)\]=(.*)", line)
            if m:
                state[m.group(1)] = m.group(2)
    except Exception:  # noqa: BLE001
        pass
    return state


def load_sheet(path, sheet):
    try:
        d = pd.read_excel(path, sheet_name=sheet, index_col=0)
    except (FileNotFoundError, ValueError, KeyError, OSError):
        return pd.DataFrame()
    return d


def pick(cols, *pats):
    for p in pats:
        for c in cols:
            if re.search(p, str(c), re.I):
                return c
    return None


def gas_trade(text):
    d = pd.read_csv(io.StringIO(text))
    print("  gas file columns:", list(d.columns), "rows", len(d), flush=True)
    per = pick(d.columns, r"^period$", r"date", r"month")
    reg = pick(d.columns, r"^region", r"province", r"export region", r"^port")
    vol = pick(d.columns, r"10\^?3 ?m3|1000 m3|thousand", r"volume")
    flow = pick(d.columns, r"^flow$", r"^activity$", r"^trade")
    if not (per and vol):
        raise ValueError(f"cannot find period/volume columns in {list(d.columns)}")
    print(f"  using period={per!r} region={reg!r} volume={vol!r} flow={flow!r}", flush=True)
    d["month"] = pd.to_datetime(d[per], errors="coerce", dayfirst=False)
    if d["month"].isna().mean() > 0.5 and "Year" in d and "Month" in d:
        d["month"] = pd.to_datetime(d["Year"].astype(str) + " " + d["Month"].astype(str), errors="coerce")
    d["v"] = pd.to_numeric(d[vol], errors="coerce")
    d = d.dropna(subset=["month", "v"])
    d["month"] = d["month"].dt.to_period("M").dt.to_timestamp()
    if flow:
        # "Imports" and "Imports (See Disclaimer)" are the same flow under two labels in different years
        d["flow"] = d[flow].astype(str).str.replace(r"\s*\(.*?\)", "", regex=True).str.strip().str.title()
    else:
        d["flow"] = "Exports"
    key = d[reg].astype(str) if reg else "Total"
    d["col"] = d["flow"] + " - " + key if reg else d["flow"]
    w = d.pivot_table(index="month", columns="col", values="v", aggfunc="sum")
    # thousand m3 per month -> million m3 per day
    out = w.div(w.index.days_in_month, axis=0) / 1000.0
    for fl in d["flow"].unique():
        cols = [c for c in out.columns if c.startswith(fl + " -")]
        if cols:
            out[f"{fl} - Total"] = out[cols].sum(axis=1, min_count=1)
    out = out.round(3)
    out.columns = [f"{c} (mcm/d)" for c in out.columns]
    out.index.name = "Month"
    return out


def elec_trade(text):
    d = pd.read_csv(io.StringIO(text))
    print("  electricity file columns:", list(d.columns), "rows", len(d), flush=True)
    for c in ("Activity", "Source", "Destination"):
        if c in d:
            print(f"   {c}: {sorted(d[c].dropna().unique())[:40]}", flush=True)
    e = pick(d.columns, r"energy")
    d["month"] = pd.to_datetime(d["Period"], errors="coerce")
    d["gwh"] = pd.to_numeric(d[e], errors="coerce") / 1000.0
    d = d.dropna(subset=["month", "gwh"])
    d["month"] = d["month"].dt.to_period("M").dt.to_timestamp()
    ex = d[(d["Activity"].str.lower() == "exports") & (d["Destination"].str.lower() == "total")]
    im = d[(d["Activity"].str.lower() == "imports") & (d["Source"].str.lower() == "total")]
    if im.empty:   # imports listed by province instead: use the destination column as the province
        im = d[(d["Activity"].str.lower() == "imports") & (d["Destination"].str.lower() != "total")]
        im = im.assign(Source=im["Destination"])
    out = pd.DataFrame()
    for lab, x, col in (("Exports", ex, "Source"), ("Imports", im, "Destination" if (im.get("Source") == "Total").any() else "Source")):
        w = x.pivot_table(index="month", columns="Source", values="gwh", aggfunc="sum") if lab == "Exports" else \
            x.pivot_table(index="month", columns="Destination" if "Destination" in x and x["Destination"].nunique() > 1 else "Source",
                          values="gwh", aggfunc="sum")
        w.columns = [f"{lab} - {c}_GWh" for c in w.columns]
        out = out.join(w, how="outer") if not out.empty else w
    out["Exports - Canada_GWh"] = out[[c for c in out.columns if c.startswith("Exports -")]].sum(axis=1, min_count=1)
    out["Imports - Canada_GWh"] = out[[c for c in out.columns if c.startswith("Imports -") and "Canada" not in c]].sum(axis=1, min_count=1)
    out.index.name = "Month"
    return out.round(1)


def generation_annual(text):
    d = pd.read_csv(io.StringIO(text))
    d = d[d["Scenario"] == EF_SCENARIO]
    d = d[d["Region"] != "Canada"]
    w = d.pivot_table(index="Year", columns=["Region", "Variable"], values="Value", aggfunc="sum")
    w.columns = [f"{r}|{v}" for r, v in w.columns]
    shares = {}
    for region in sorted(d["Region"].unique()):
        fos = pd.DataFrame({FOSSIL[v]: w.get(f"{region}|{v}") for v in FOSSIL}).fillna(0)
        tot = fos.sum(axis=1).replace(0, float("nan"))
        for g in FOSSIL.values():
            shares[f"{region}|{g}"] = (fos[g] / tot)
    sh = pd.DataFrame(shares).round(4)
    w.index.name = sh.index.name = "Year"
    return w.round(1), sh


def save(path, sheets, notes, state):
    os.makedirs(os.path.dirname(os.path.abspath(path)), exist_ok=True)
    lines = notes + [""] + [f"lm[{k}]={v}" for k, v in state.items()]
    xlsx_notes.write_workbook(path, sheets, lines, {"UNITS", "COVERAGE", "SOURCE", "STATE"})
    print(f"Saved {path}: " + ", ".join(f"{k} {len(v)} rows" for k, v in sheets.items()), flush=True)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--gas-out", default=os.path.join(DEFAULT_DIR, "canada_cer_gas.xlsx"))
    ap.add_argument("--power-out", default=os.path.join(DEFAULT_DIR, "canada_cer_power.xlsx"))
    ap.add_argument("--force", action="store_true")
    args = ap.parse_args()
    ok = 0

    print("CER natural gas trade", flush=True)
    state, saved = read_state(args.gas_out), load_sheet(args.gas_out, "Gas trade monthly")
    try:
        got = None
        for f in GAS_FILES:
            lm = last_modified(f)
            if not args.force and not saved.empty and lm and state.get(f) == lm:
                print(f"  {f}: unchanged ({lm})", flush=True)
                got = "unchanged"
                break
            try:
                text, lm = fetch(f)
                out = gas_trade(text)
                out.index = pd.to_datetime(out.index)
                if not saved.empty:
                    saved.index = pd.to_datetime(saved.index)
                    out = out.combine_first(saved[[c for c in saved.columns if c in out.columns]])   # retired columns drop out
                out = out[out.index >= "2015-01-01"]
                out.index = out.index.strftime("%Y-%m")
                out.index.name = "Month"
                notes = [
                    "UNITS",
                    "Million cubic metres per day (mcm/d): the CER's monthly volume (10^3 m3) / days in month / 1000. "
                    "One column per flow (Exports / Imports) and CER region, plus a Total.",
                    "",
                    "COVERAGE",
                    f"Canada, monthly, {out.index.min()} to {out.index.max()}. Source file: {BASE}{f} (Last-Modified {lm}).",
                    "",
                    "SOURCE",
                    "Canada Energy Regulator, Natural Gas Exports / Exports and Imports open data: "
                    "https://www.cer-rec.gc.ca/en/data-analysis/energy-commodities/natural-gas/index.html",
                ]
                save(args.gas_out, {"Gas trade monthly": out}, notes, {**state, f: lm})
                got = f
                ok += 1
                break
            except Exception as e:  # noqa: BLE001
                print(f"  {f} FAILED: {type(e).__name__}: {e}", flush=True)
        if got == "unchanged":
            ok += 1
    except Exception as e:  # noqa: BLE001
        print(f"GAS FAILED: {type(e).__name__}: {e}", flush=True)

    print("CER electricity trade and generation", flush=True)
    state = read_state(args.power_out)
    sheets, new_state, changed = {}, dict(state), False
    saved_trade = load_sheet(args.power_out, "Electricity trade monthly")
    saved_gen, saved_sh = load_sheet(args.power_out, "Generation annual"), load_sheet(args.power_out, "Fossil shares")
    sheets.update({"Electricity trade monthly": saved_trade, "Generation annual": saved_gen, "Fossil shares": saved_sh})
    for f, kind in ((ELEC_FILE, "trade"), (EF_FILE, "ef")):
        try:
            lm = last_modified(f)
            have = not (saved_trade.empty if kind == "trade" else saved_gen.empty)
            if not args.force and have and lm and state.get(f) == lm:
                print(f"  {f}: unchanged ({lm})", flush=True)
                continue
            text, lm = fetch(f)
            if kind == "trade":
                out = elec_trade(text)
                out = out[out.index >= "2015-01-01"]
                out.index = out.index.strftime("%Y-%m")
                out.index.name = "Month"
                sheets["Electricity trade monthly"] = out
            else:
                sheets["Generation annual"], sheets["Fossil shares"] = generation_annual(text)
            new_state[f] = lm
            changed = True
        except Exception as e:  # noqa: BLE001
            print(f"  {f} FAILED: {type(e).__name__}: {e}", flush=True)
    sheets = {k: v for k, v in sheets.items() if not v.empty}
    if changed and sheets:
        notes = [
            "UNITS",
            "Electricity trade monthly: GWh per month (CER energy in MW.h / 1000), exports and imports by province "
            "(exports: the exporting province, all destinations; imports: the importing province where the file "
            "gives it). Generation annual: GWh per YEAR by '<province>|<fuel>' from Canada's Energy Future 2026, "
            f"scenario '{EF_SCENARIO}' (years up to the latest historical year are the CER's compiled history, later "
            "years its projection). Fossil shares: coal / gas / oil as a share of the province's coal + gas + oil "
            "generation in that year.",
            "",
            "COVERAGE",
            "Canada and provinces; trade monthly from 2015, generation annual 2005-2050 (projection after the "
            "latest historical year).",
            "",
            "SOURCE",
            f"Canada Energy Regulator open data: {BASE}{ELEC_FILE} ; {BASE}{EF_FILE} "
            "(https://www.cer-rec.gc.ca/en/data-analysis/canada-energy-future/index.html)",
        ]
        save(args.power_out, sheets, notes, new_state)
        ok += 1
    elif sheets:
        ok += 1
    if not ok:
        raise SystemExit("No CER file could be updated")


if __name__ == "__main__":
    main()
