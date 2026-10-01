"""
Peru generation capacity (potencia efectiva) by technology and energy
resource from COES (Comite de Operacion Economica del Sistema
Interconectado Nacional) - free, public, no key.

Source: COES annual statistics ('Estadisticas Anuales'), chapter 2 'Estado
de la infraestructura del SEIN' Excel workbook for each year:
  https://www.coes.org.pe/Portal/publicaciones/estadisticas/
  -> Publicaciones/Estadisticas Anuales/<year>/Excel/Capitulo 02_*.xlsx
     (downloaded through /Portal/browser/download?url=...; the file names
     differ by year, so the script lists each year's Excel folder through
     the portal's file browser, POST /Portal/browser/vistadatos)
The workbook lists every generating unit in the SEIN at the end of the year
with its company, TIPO DE GENERACION, TECNOLOGIA, TIPO DE RECURSO
ENERGETICO and POTENCIA EFECTIVA (MW); this script sums the units by
energy resource. COES' monthly bulletins (Publicaciones/Boletines) carry
no capacity tables, so the finest granularity is annual: one row per year,
dated 1 January, holding the effective capacity at 31 December of that year.
Found via discovery_archive/south_america/CAPACITY_PBUE_DISCOVERY.py .. 9.

Incremental: years already in the workbook are kept; new years (COES
publishes a year's statistics during the following year) and the latest
year (in case COES replaces the file) are downloaded.

Usage: python3 PERU_COES_CAPACITY.py [--out PATH] [--test]
"""

print("STARTING", flush=True)

import argparse
import html
import io
import os
import re
import sys
import unicodedata
from datetime import date, datetime
from urllib.parse import quote

import pandas as pd
import requests

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import power_capacity_common as pcc  # noqa: E402

PORTAL = "https://www.coes.org.pe/Portal/"
BASE = "Publicaciones/Estadisticas Anuales/"
HEADERS = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) "
                         "Chrome/124.0 Safari/537.36"}
TIMEOUT = (15, 180)
DEFAULT_OUT = os.path.join("output", "Data and Chart Outputs", "peru_power_capacity.xlsx")
GEN_TYPES = {"HIDROELECTRICA", "TERMOELECTRICA", "SOLAR", "EOLICA", "EOLICO", "GEOTERMICA", "BIOMASA"}

S = requests.Session()
S.headers.update(HEADERS)


def key(s):
    s = unicodedata.normalize("NFKD", str(s)).encode("ascii", "ignore").decode().upper()
    return " ".join(s.split())


def fuel_of(resource, gen_type):
    """Standard fuel for a COES unit. Gas is tested before water: 'Gas Natural de Aguaytia' contains 'AGUA'."""
    r, g = key(resource), key(gen_type)
    if "GAS NATURAL" in r or r == "GAS":
        return "Gas"
    if any(t in r for t in ("DIESEL", "RESIDUAL", "R500", "R6", "D2", "PETROLEO", "NAFTA", "REFINERIA", "RFG",
                            "FUEL")):
        return "Oil"
    if "CARBON" in r:
        return "Coal"
    if "BAGAZO" in r or "BIOMASA" in r or "BIOGAS" in r:
        return "Bioenergy"
    if g == "HIDROELECTRICA" or r in ("AGUA", "HIDRO", "HIDRAULICA"):
        return "Hydro"
    if g == "SOLAR" or r == "SOLAR":
        return "Solar"
    if g.startswith("EOLIC") or r.startswith("EOLIC") or r == "VIENTO":
        return "Wind"
    if "GEOTERM" in r or g == "GEOTERMICA":
        return "Other"
    return None


def browse(path):
    """(path, 'D'|'F') items of a COES portal folder."""
    r = S.post(PORTAL + "browser/vistadatos", data={"baseDirectory": path, "url": path, "indicador": "",
                                                    "initialLink": "", "orderFolder": ""}, timeout=TIMEOUT)
    r.raise_for_status()
    out = []
    for m in re.finditer(r"openBlob\('([^']+)',\s*'(\w)'", r.text):
        it = (html.unescape(m.group(1)), m.group(2))
        if it not in out:
            out.append(it)
    return out


def chapter2_url(year):
    files = [p for p, k in browse(f"{BASE}{year}/Excel/") if k == "F"]
    hit = [p for p in files if re.search(r"Cap[ií]tulo 0?2[_ ]", p) and p.lower().endswith((".xlsx", ".xls"))]
    if not hit:
        return None, files
    return PORTAL + "browser/download?url=" + quote(hit[0]), files


def unit_table(content):
    """The per-unit list: (sheet, DataFrame with company, gen_type, plant, unit, technology, resource, mw).
    Every sheet / header row with 'POTENCIA EFECTIVA' and 'TIPO DE GENERACION' columns is read as one contiguous
    block of unit rows (a block ends after three rows without a generation type, which separates it from the
    commissioning / retirement table and any other table further down); the largest block is the unit list."""
    xl = pd.ExcelFile(io.BytesIO(content))
    best = None
    for sheet in xl.sheet_names:
        raw = pd.read_excel(xl, sheet_name=sheet, header=None)
        for i in range(min(30, len(raw))):
            cells = {key(v): j for j, v in enumerate(raw.iloc[i]) if isinstance(v, str)}
            mw = [j for k, j in cells.items() if k.startswith("POTENCIA EFECTIVA")]
            gen = [j for k, j in cells.items() if k.startswith("TIPO DE GENERACION")]
            if not (mw and gen):
                continue
            res = [j for k, j in cells.items() if "RECURSO" in k or "COMBUSTIBLE" in k or k == "FUENTE"]
            col = lambda names: next((j for k, j in cells.items() if any(k.startswith(n) for n in names)), None)  # noqa: E731
            cols = {"company": col(["EMPRESA"]), "gen_type": gen[0], "plant": col(["CENTRAL"]),
                    "unit": col(["UNIDAD", "GRUPO"]), "technology": col(["TECNOLOGIA"]),
                    "resource": res[0] if res else None, "mw": mw[0]}
            rows, misses = [], 0
            for r in range(i + 1, len(raw)):
                g = raw.iat[r, cols["gen_type"]]
                v = raw.iat[r, cols["mw"]]
                ok = (key(g) in GEN_TYPES and not isinstance(v, (datetime, date, pd.Timestamp))
                      and pd.notna(pd.to_numeric(v, errors="coerce")))
                if not ok:
                    misses += 1
                    if rows and misses >= 3:
                        break
                    continue
                misses = 0
                rows.append({c: (raw.iat[r, j] if j is not None else None) for c, j in cols.items()})
            if rows and (best is None or len(rows) > len(best[1])):
                best = (sheet, rows)
    if best is None:
        raise ValueError(f"no unit list with 'POTENCIA EFECTIVA' and 'TIPO DE GENERACION' in {xl.sheet_names}")
    d = pd.DataFrame(best[1])
    d["mw"] = pd.to_numeric(d["mw"], errors="coerce")
    return best[0], d


def resource_summary(content):
    """COES' own summary table 'Potencia efectiva por tipo de recurso energetico' (Cuadro 2.5 / C5): (sheet,
    {resource label: MW}, published total) - or None when the workbook has no such table. The rows must add up
    to the table's TOTAL row."""
    xl = pd.ExcelFile(io.BytesIO(content))
    for sheet in xl.sheet_names:
        raw = pd.read_excel(xl, sheet_name=sheet, header=None)
        for i in range(min(15, len(raw))):
            cells = {key(v): j for j, v in enumerate(raw.iloc[i]) if isinstance(v, str)}
            lab = next((j for k, j in cells.items() if k.startswith("TIPO DE RECURSO")), None)
            mw = next((j for k, j in cells.items() if k.startswith("POTENCIA EFECTIVA")), None)
            if lab is None or mw is None:
                continue
            vals, total = {}, None
            for r in range(i + 1, min(i + 40, len(raw))):
                name, v = raw.iat[r, lab], pd.to_numeric(raw.iat[r, mw], errors="coerce")
                if not isinstance(name, str) or pd.isna(v):
                    if vals and not isinstance(name, str):
                        break
                    continue
                if key(name).startswith("TOTAL"):
                    total = float(v)
                    break
                vals[name.strip()] = float(v)
            if vals and total is not None and abs(sum(vals.values()) - total) <= 1:
                return sheet, vals, total
    return None


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", default=DEFAULT_OUT)
    ap.add_argument("--test", action="store_true")
    args = ap.parse_args()

    years = sorted(int(re.search(r"(\d{4})/?$", p).group(1)) for p, k in browse(BASE)
                   if k == "D" and re.search(r"(\d{4})/?$", p))
    years = [y for y in years if y >= pcc.START.year]
    old = pcc.load_monthly(args.out)
    old_units = pcc.load_sheet(args.out, "COES units")
    old_files = pcc.load_sheet(args.out, "COES files")
    have = set(old.index.year) if not old.empty else set()
    todo = [y for y in years if y not in have or y == max(years)]
    print(f"COES annual statistics years {years}; archive {sorted(have)}; fetching {todo}", flush=True)

    rows, units, sources, basis, labels = {}, [], {}, {}, {}
    for y in todo:
        url, files = chapter2_url(y)
        if url is None:
            print(f"  {y}: no chapter-2 workbook yet (files: {files[:5]})", flush=True)
            continue
        r = S.get(url, timeout=TIMEOUT)
        r.raise_for_status()
        if len(r.content) < 1000:
            print(f"  {y}: {url} returned {len(r.content)} bytes - skipped", flush=True)
            continue
        summary = resource_summary(r.content)
        try:
            sheet, d = unit_table(r.content)
        except ValueError as e:
            sheet, d = None, None
            print(f"  {y}: {e}", flush=True)
        if d is not None:
            d["fuel"] = [fuel_of(res, g) for res, g in zip(d["resource"], d["gen_type"])]
            unknown = d[d["fuel"].isna()]
            if not unknown.empty:
                print(f"  {y}: unmapped resources -> Other: {sorted(set(unknown['resource'].map(str)))}", flush=True)
                d["fuel"] = d["fuel"].fillna("Other")
            print(f"  {y}: unit list sheet {sheet!r}: {len(d)} units, {d['mw'].sum():,.1f} MW", flush=True)
        if summary:
            print(f"  {y}: summary sheet {summary[0]!r}: total {summary[2]:,.1f} MW: {summary[1]}", flush=True)
        # The unit list is used when it matches COES' own summary total (or there is no summary); otherwise the
        # summary by energy resource (in 2021-2022 the unit list has merged cells and company subtotals).
        if d is not None and (summary is None or abs(d["mw"].sum() - summary[2]) <= 0.01 * summary[2]):
            d.insert(0, "year", y)
            units.append(d)
            by = d.groupby("fuel")["mw"].sum()
            basis[y] = f"unit list (sheet {sheet!r}, {len(d)} units)"
            for res, f in zip(d["resource"].map(str), d["fuel"]):
                if res != "nan":
                    labels.setdefault(y, {}).setdefault(f, set()).add(res)
            if args.test:
                print(d.groupby(["gen_type", "resource", "fuel"], dropna=False)["mw"].agg(["count", "sum"])
                      .to_string(), flush=True)
        elif summary:
            by = {}
            for name, v in summary[1].items():
                f = fuel_of(name, "") or "Other"
                labels.setdefault(y, {}).setdefault(f, set()).add(name)
                by[f] = by.get(f, 0.0) + v
            by = pd.Series(by)
            basis[y] = f"summary by energy resource (sheet {summary[0]!r})"
        else:
            print(f"  {y}: neither a unit list nor a summary table - skipped", flush=True)
            continue
        rows[pd.Timestamp(year=y, month=1, day=1)] = {f"{f}_MW": v for f, v in by.items()}
        sources[y] = url
        print(f"  {y}: using {basis[y]}: {by.sum():,.1f} MW: {', '.join(f'{k} {v:,.1f}' for k, v in by.items())}",
              flush=True)

    new = pcc.standardise(pd.DataFrame.from_dict(rows, orient="index")) if rows else pd.DataFrame()
    if not old.empty:
        old = pcc.standardise(old)
        monthly = pd.concat([old[~old.index.isin(new.index)], new]).sort_index() if not new.empty else old
    else:
        monthly = new
    if monthly.empty:
        sys.exit("No data fetched and no archive - nothing to write")
    unit_df = pd.concat(units, ignore_index=True) if units else pd.DataFrame()
    if not old_units.empty:
        old_units = old_units.reset_index()
        fetched = [t.year for t in rows]
        unit_df = pd.concat([old_units[~old_units["year"].isin(fetched)], unit_df], ignore_index=True)
    if not unit_df.empty:
        unit_df = unit_df.sort_values(["year", "fuel", "mw"], ascending=[True, True, False]).set_index("year")
    files = pd.DataFrame([{"year": y, "url": sources[y], "basis": basis[y],
                           "resources": " | ".join(f"{f}: {', '.join(sorted(v))}" for f, v in labels[y].items())}
                          for y in sorted(sources)])
    if not old_files.empty:
        old_files = old_files.reset_index()
        files = pd.concat([old_files[~old_files["year"].isin(list(sources))], files], ignore_index=True)
    files = files.sort_values("year").set_index("year")
    all_labels = {}
    for txt in files["resources"].dropna():
        for part in str(txt).split(" | "):
            f, _, names = part.partition(": ")
            all_labels.setdefault(f, set()).update(n for n in names.split(", ") if n)
    print(monthly.to_string(), flush=True)

    val = pcc.validation_lines(monthly, "Peru")
    print("\n".join(val), flush=True)
    notes = pcc.unit_notes("year") + [
        "Effective capacity (potencia efectiva) of the units COES lists in the SEIN at 31 December; each row is dated "
        "1 January of that year. COES' monthly bulletins carry no capacity table, so there is no monthly series.",
        "",
        "COVERAGE",
        f"{monthly.index.min():%Y} to {monthly.index.max():%Y} ({len(monthly)} years), from 2021. Peru's national "
        "grid (SEIN) only; isolated systems and self-generators outside the COES are not included. A year appears "
        "once COES publishes its annual statistics (during the following year).",
        "",
        "SOURCE",
        "COES - Estadisticas Anuales, chapter 2 'Estado de la infraestructura del SEIN' (Excel), unit list: "
        "https://www.coes.org.pe/Portal/publicaciones/estadisticas/ ; files: "
        + "; ".join(f"{y}: {r.url} - {r.basis}" for y, r in files.iterrows()) + ".",
        "The per-unit list is summed by energy resource when its total matches COES' own summary table "
        "'Potencia efectiva por tipo de recurso energetico'; otherwise (2021-2022 workbooks, whose unit list has "
        "merged cells and company subtotals) that summary table is used directly.",
        "Script: south_america/PERU_COES_CAPACITY.py (scheduled by .github/workflows/peru_power_capacity.yml).",
        "",
        "MAPPING (COES 'TIPO DE RECURSO ENERGETICO', or the generation type for hydro / solar / wind)",
    ] + [f"{f}_MW = {', '.join(sorted(all_labels[f]))}" for f in pcc.FUELS if f in all_labels] + [
        "Diesel, residual fuel oil and refinery gas / naphtha go in Oil_MW; bagasse and biogas in Bioenergy_MW. No "
        "nuclear; Other_MW = any resource not matched above.",
        "",
        "VALIDATION",
    ] + val
    extra = {"COES files": files}
    if not unit_df.empty:
        extra["COES units"] = unit_df
    pcc.write(args.out, monthly, notes, extra)
    print(f"Saved {args.out}", flush=True)


if __name__ == "__main__":
    main()
