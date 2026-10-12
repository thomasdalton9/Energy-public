"""
Official ANNUAL installed generating capacity for Ghana and Morocco (input to middle_east/capacity/EAST_MED_AFRICA_CAPACITY_COMPLETE.py).

  Ghana    Energy Commission 'National Energy Statistics' booklet, Table 3.1 'Grid Installed and Dependable Capacity (MW)':
           hydro, thermal, other renewables (solar and waste) at year end, 2000 - last year, excluding off-grid and distributed
           generation (the same booklet as GHANA_ENERGY_COMMISSION.py; the list page names the newest edition).
           -> output/Data and Chart Outputs/ghana_capacity_annual.xlsx
  Morocco  HCP 'Annuaire statistique du Maroc' energy workbook, table 5-4 'Evolution de la puissance installee ...' (ONEE data):
           installed MW by source (hydro, coal steam, fuel steam, gas turbines, diesel, combined cycle, gasoil, wind, solar), three
           years per yearbook; the newest yearbook wins per year. Yearbook files come from the same HCP list page and archive
           formats as africa/MOROCCO_GAS.py (which this reuses for the download and archive extraction).
           -> output/Data and Chart Outputs/morocco_capacity_annual.xlsx

Both workbooks: sheet 'Monthly' (date = 1 January of the year whose END the stock is, <Fuel>_MW in the standard capacity layout,
Total_MW; annual rows as the other annual capacity workbooks), a detail sheet with the source's own categories, and 'Release'.
Ghana thermal (gas, light crude oil, dual-fuel plants) is shown as Gas and 'Other renewables' (solar and waste, 0.1 MW biogas) as Solar,
as the Energy Commission does not split them; Morocco gas turbines (mostly gasoil-fired) are shown as Other, not Gas.
Incremental: a source is downloaded again only when the Ghana booklet id or a Moroccan yearbook file changes (see 'Release').

Usage: python3 AFRICA_CAPACITY_ANNUAL.py [--data-dir DIR] [--force]
"""
import argparse
import hashlib
import io
import os
import re
import sys
import tempfile
import time

import pandas as pd
import requests
import urllib3

urllib3.disable_warnings()
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import xlsx_notes  # noqa: E402

DATA_DIR = os.path.join(ROOT, "output", "Data and Chart Outputs")
H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36"}
GH_LIST = "https://www.energycom.gov.gh/index.php/planning/energy-statistics"
GH_BASE = "https://www.energycom.gov.gh"
FIRST_YEAR = 2015     # nothing older is needed (the complete build starts in 2021)


def log(m):
    print(m, flush=True)


def get(url, tries=4, **kw):
    last = None
    for k in range(tries):
        try:
            r = requests.get(url, headers=H, timeout=(15, 240), verify=False, **kw)
            if r.status_code < 400:
                return r
            last = f"HTTP {r.status_code}"
        except requests.RequestException as e:
            last = type(e).__name__
        time.sleep(4 * (k + 1))
    raise RuntimeError(f"{last} for {url}")


def read_release(path):
    try:
        return pd.read_excel(path, sheet_name="Release")
    except Exception:  # noqa: BLE001
        return None


# ------------------------------------------------------------------ Ghana
def ghana_parse(pages):
    """Table 3.1 rows 'YEAR hydro thermal other|- total | dependable ...' -> {year: (hydro, thermal, other, total)}."""
    rows = {}
    for text in pages:
        if not re.search(r"Table 3\.1:\s*Grid Installed and Dependable Capacity", text or "") or "....." in (text or "")[:400]:
            continue
        for line in text.split("\n"):
            m = re.match(r"^((?:19|20)\d\d)\*?\s+(.*)$", line.strip())
            if not m:
                continue
            toks = m.group(2).split()
            if len(toks) < 8 or not all(re.fullmatch(r"[\d,]+(?:\.\d+)?|-", t) for t in toks[:8 if len(toks) == 8 else 9]):
                continue
            v = [0.0 if t == "-" else float(t.replace(",", "")) for t in toks]
            if len(v) not in (8, 9, 10):
                continue
            # installed block = first four numbers (hydro, thermal, other, total)
            h, t, o, tot = v[0], v[1], v[2], v[3]
            if abs(h + t + o - tot) > 1.5:
                continue
            rows[int(m.group(1))] = (h, t, o, tot)
        if rows:
            break
    return rows


def ghana(data_dir, force):
    import pdfplumber
    out = os.path.join(data_dir, "ghana_capacity_annual.xlsx")
    r = get(GH_LIST)
    found = sorted({(int(y), int(i)) for i, y in re.findall(r"energy-statistics\?download=(\d+):(\d{4})-energy-statistics", r.text)}, reverse=True)
    if not found:
        log("Ghana: no booklet link found")
        return False
    ed, did = found[0]
    rel = read_release(out)
    if rel is not None and not force and int(rel.loc[0, "Download_id"]) == did:
        log(f"Ghana: booklet {ed} (id {did}) already read")
        return False
    url = f"{GH_BASE}/index.php/planning/energy-statistics?download={did}:{ed}-energy-statistics"
    g = get(url)
    with pdfplumber.open(io.BytesIO(g.content)) as pdf:
        pages = [p.extract_text() or "" for p in pdf.pages]
    rows = ghana_parse(pages)
    rows = {y: v for y, v in rows.items() if y >= FIRST_YEAR}
    if not rows:
        log("Ghana: Table 3.1 not found / not parsed")
        return False
    d = pd.DataFrame(rows, index=["Hydro_MW", "Thermal_MW", "Other_renewables_MW", "Total_MW"]).T.sort_index()
    d.insert(0, "Year", d.index)
    d.index = pd.to_datetime([f"{y}-01-01" for y in d.index])
    d.index.name = "date"
    monthly = pd.DataFrame({"Year": d["Year"], "Hydro_MW": d["Hydro_MW"], "Gas_MW": d["Thermal_MW"], "Solar_MW": d["Other_renewables_MW"],
                            "Total_MW": d["Total_MW"]})
    release = pd.DataFrame([{"Edition": ed, "Download_id": did, "URL": url, "Bytes": len(g.content),
                             "SHA256": hashlib.sha256(g.content).hexdigest()[:16], "Years": f"{d['Year'].min()}-{d['Year'].max()}"}])
    notes = ["Ghana installed generating capacity, year end, MW (annual; rows are dated 1 January of the year whose END the stock is)", "",
             "Source", f"Energy Commission of Ghana, National Energy Statistics {ed} (data to {d['Year'].max()}), Table 3.1 'Grid Installed and Dependable "
             f"Capacity (MW)': {url}",
             "Grid-connected capacity only (excludes off-grid and distributed generation). Gas_MW = the Commission's 'Thermal' (gas, light crude oil, "
             "dual-fuel plants; not split by fuel). Solar_MW = 'Other Renewables' (solar and waste: 132.5 MW solar + 0.1 MW biogas in 2024). Hydro = "
             "Akosombo, Kpong, Bui, Tsatsadu. 'Thermal_MW' and 'Other_renewables_MW' on the Table 3.1 sheet keep the source names.",
             "Dependable capacity (also in the table) is not used."]
    xlsx_notes.write_workbook(out, {"Monthly": monthly, "Table 3.1": d.drop(columns="Year"), "Release": release}, notes,
                              ["Source"])
    log(f"Ghana: {len(d)} years {d['Year'].min()}-{d['Year'].max()} saved")
    return True


# ------------------------------------------------------------------ Morocco
MA_LABELS = [   # (regex on the label, item)
    (r"^hydraulique", "Hydro"), (r"charbon", "Coal_steam"), (r"fuel vapeur", "Fuel_steam"), (r"turbines? [aà] gaz", "Gas_turbines"),
    (r"^diesel", "Diesel"), (r"cycle combin", "Combined_cycle"), (r"^gasoil", "Gasoil"), (r"^eolien\s*$", "Wind"), (r"^solaire\s*$", "Solar"),
]


def morocco_sheet(df):
    """ONEE table 5-4 frame -> {year: {item: MW}} (the three years of the yearbook), first block only (not the autoproducers)."""
    hdr, years = None, None
    for i in range(min(len(df), 15)):
        vals = [v for v in df.iloc[i].values if isinstance(v, (int, float)) and not pd.isna(v) and 1990 < v < 2100]
        if len(vals) >= 2:
            hdr, years = i, [int(v) for v in vals]
            break
    if hdr is None:
        return {}
    out = {y: {} for y in years}
    for i in range(hdr + 1, len(df)):
        row = df.iloc[i].values
        label = next((str(v).strip() for v in row if isinstance(v, str) and v.strip()), "")
        low = label.lower()
        if "autoproducteur" in low:
            break
        nums = [v for v in row if isinstance(v, (int, float)) and not pd.isna(v)]
        if re.match(r"^puissances? install", low) and len(nums) >= len(years):
            for y, v in zip(years, nums):
                out[y]["Total_reported"] = float(v)
            continue
        if len(nums) < len(years):
            continue
        for rx, item in MA_LABELS:
            if re.search(rx, low):
                for y, v in zip(years, nums):
                    out[y][item] = float(v)
                break
    return {y: v for y, v in out.items() if "Total_reported" in v and "Hydro" in v}


def morocco(data_dir, force):
    import MOROCCO_GAS as mg
    out = os.path.join(data_dir, "morocco_capacity_annual.xlsx")
    rel = read_release(out)
    r = mg.get(mg.HCP_LIST, timeout=(20, 90))
    items = {}
    for m in re.finditer(r'href="(/file/(\d+)/)"', r.text):
        ctx = re.sub(r"<[^>]+>|\s+", " ", r.text[m.end():m.end() + 600])
        t = re.search(r"Annuaire[^|]{0,80}?ann[ée]e\s*(\d{4})", ctx)
        if t and int(t.group(1)) >= 2019:
            items[m.group(2)] = int(t.group(1))
    log(f"Morocco: yearbooks {sorted(items.values())}")
    stored = {}
    if rel is not None and not force:
        stored = {str(a): b for a, b in zip(rel["File_id"], rel["Rows"])}
    old = None
    if rel is not None and not force:
        try:
            old = pd.read_excel(out, sheet_name="ONEE detail")
        except Exception:  # noqa: BLE001
            old = None
    recs, relrows, changed = [], [], False
    for fid, yb in sorted(items.items(), key=lambda kv: kv[1]):
        if fid in stored and str(stored[fid]) not in ("0", "nan") and old is not None:
            relrows.append({"File_id": fid, "Yearbook": yb, "Rows": stored[fid]})
            recs += old[old["Yearbook"] == yb].to_dict("records")
            continue
        try:
            rr = mg.get(f"https://www.hcp.ma/file/{fid}/", timeout=(20, 300))
            n = 0
            with tempfile.TemporaryDirectory() as td:
                for b in mg._extract(rr.content, td):
                    x = pd.read_excel(b, sheet_name=None, header=None)
                    for sname, df in x.items():
                        if "puissance" not in " ".join(str(v) for v in df.head(12).astype(str).values.ravel()).lower():
                            continue
                        parsed = morocco_sheet(df)
                        for y, v in parsed.items():
                            for k, val in v.items():
                                recs.append({"Yearbook": yb, "Year": y, "Item": k, "MW": val}); n += 1
                        if parsed:
                            break
            relrows.append({"File_id": fid, "Yearbook": yb, "Rows": n})
            changed = True
            log(f"  yearbook {yb}: {n} values")
        except Exception as e:  # noqa: BLE001
            log(f"  yearbook {yb}: {type(e).__name__}: {str(e)[:120]}")
            relrows.append({"File_id": fid, "Yearbook": yb, "Rows": 0})
    if not recs:
        return False
    if not changed and old is not None:
        return False
    det = pd.DataFrame(recs)
    # newest yearbook wins per year and item
    best = det.sort_values("Yearbook").drop_duplicates(["Year", "Item"], keep="last")
    w = best.pivot(index="Year", columns="Item", values="MW").sort_index()
    w = w[w.index >= FIRST_YEAR]
    parts = ["Coal_steam", "Fuel_steam", "Gas_turbines", "Diesel", "Combined_cycle", "Gasoil", "Hydro", "Wind", "Solar"]
    for c in parts:
        if c not in w:
            w[c] = float("nan")
    chk = w[parts].sum(axis=1, min_count=1)
    w["Check_sum_minus_reported"] = (chk - w["Total_reported"]).round(1)
    monthly = pd.DataFrame({"Year": w.index, "Hydro_MW": w["Hydro"].values, "Coal_MW": w["Coal_steam"].values,
                            "Oil_MW": (w["Fuel_steam"] + w["Diesel"] + w["Gasoil"]).values, "Gas_MW": w["Combined_cycle"].values,
                            "Other_MW": w["Gas_turbines"].values, "Wind_MW": w["Wind"].values, "Solar_MW": w["Solar"].values,
                            "Total_MW": w["Total_reported"].values},
                           index=pd.to_datetime([f"{y}-01-01" for y in w.index]))
    monthly.index.name = "date"
    wide = w.copy()
    wide.index = pd.to_datetime([f"{y}-01-01" for y in wide.index])
    wide.index.name = "date"
    notes = ["Morocco installed generating capacity (ONEE), year end, MW (annual; rows dated 1 January of the year whose END the stock is)", "",
             "Source", "Haut-Commissariat au Plan, Annuaire statistique du Maroc, energy workbook table 5-4 'Evolution de la puissance installee, de la "
             "production d'electricite' (Office National de l'Electricite et de l'Eau Potable): https://www.hcp.ma/downloads/Annuaires-statistiques-du-Maroc-Format-Excel_t22392.html",
             "Hydro_MW = hydraulique; Coal_MW = charbon vapeur (Jorf Lasfar, Safi); Oil_MW = fuel vapeur + diesel + gasoil; Gas_MW = cycle combine "
             "(Tahaddart, Ain Beni Mathar); Other_MW = turbines a gaz (mostly gasoil-fired, not split by fuel); Wind_MW = eolien (ONEE, third parties, IPP, "
             "law 13-09); Solar_MW = solaire (ONEE, MASEN, law 13-09). Total_MW is the table's own total (ONEE system; main autoproducers are listed "
             "separately in the table and not included). The 'ONEE detail' sheet keeps every yearbook's own values; 'Check_sum_minus_reported' on 'Years' "
             "is the sum of the sources less the reported total."]
    sheets = {"Monthly": monthly, "Years": wide, "ONEE detail": det.set_index("Yearbook"), "Release": pd.DataFrame(relrows)}
    xlsx_notes.write_workbook(out, sheets, notes, ["Source"])
    log(f"Morocco: years {w.index.min()}-{w.index.max()} saved")
    return True


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--data-dir", default=DATA_DIR)
    ap.add_argument("--force", action="store_true")
    ap.add_argument("--only", choices=["ghana", "morocco"])
    args = ap.parse_args()
    ok = 0
    for name, fn in (("ghana", ghana), ("morocco", morocco)):
        if args.only and args.only != name:
            continue
        try:
            ok += bool(fn(args.data_dir, args.force))
        except Exception as e:  # noqa: BLE001
            log(f"{name}: FAILED {type(e).__name__}: {e}")
    gh = os.environ.get("GITHUB_OUTPUT")
    if gh:
        with open(gh, "a") as f:
            f.write(f"updated={1 if ok else 0}\n")


if __name__ == "__main__":
    main()
