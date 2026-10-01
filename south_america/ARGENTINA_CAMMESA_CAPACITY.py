"""
Argentina installed generation capacity by technology, monthly from January
2021, from CAMMESA's "Potencia Instalada" statistics workbook (Sintesis
Mensual > Estadisticas), published at
  https://cammesaweb.cammesa.com/download/potencia-instalada/
  (file: https://microfe.cammesa.com/static-content/CammesaWeb/download-manager-files/
         Sintesis%20Mensual/Estadisticas/Potencia%20Instalada.xlsx)
No key needed. CAMMESA updates it monthly (Sep-2026 upload: data to Aug-2026).

The workbook holds, in sheet 'RESUMEN por MAQUINA':
  - CAMMESA's year-end installed MW by machine type, 2002 to last December,
    plus the current year to date (the latest month);
  - a machine-by-machine list for the latest month, with the enabling date
    (FECHA HABILITACION) of machines enabled in recent years;
and in sheet 'Potencia dada de baja' each machine retired since 2021 with its
month of retirement.

Monthly series: each December (and the latest month) equals CAMMESA's own
published figure. Within each year, months move with the dated changes:
machines enabled that year (from the machine list) are added from their
enabling month and retired machines removed from their retirement month. Any
part of a year's change CAMMESA does not date (re-ratings, machines enabled
and later retired, reclassifications such as the 2024 step down in large
hydro) is applied in December (or in the latest month for the current year)
and listed in sheet 'Undated changes'.

The CAMMESA nemo API (api.cammesa.com/pub-svc) has no public capacity
publication, and the Secretaria de Energia open-data copy
(datos.energia.gob.ar, potencia-instalada.csv) stops in Feb 2020. Found via
discovery_archive/south_america/POWER_CAPACITY_PROBE.py, PROBE3.py, PROBE4.py.

Usage: python3 ARGENTINA_CAMMESA_CAPACITY.py [--out <xlsx>]
"""
import argparse
import datetime as dt
import io
import numbers
import os
import re
import sys

import pandas as pd
import requests

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import power_capacity_std as std  # noqa: E402

OUT = "output/Data and Chart Outputs/argentina_power_capacity.xlsx"
PAGE = "https://cammesaweb.cammesa.com/download/potencia-instalada/"
XLSX = ("https://microfe.cammesa.com/static-content/CammesaWeb/download-manager-files/Sintesis%20Mensual/"
        "Estadisticas/Potencia%20Instalada.xlsx")
HEADERS = {"User-Agent": "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124 Safari/537.36"}
REFRESH_MONTHS = 2
COAL_PLANTS = re.compile(r"TURBIO", re.I)   # Rio Turbio, Argentina's coal-fired steam plant (Santa Cruz)

# CAMMESA annual-table machine type ('TIPO') -> standard fuel. Thermal capacity is published by machine type, not
# fuel: combined cycles, gas turbines and steam turbines burn natural gas (most can switch to gasoil/fuel oil), the
# diesel engines burn gasoil/fuel oil. Coal (Rio Turbio's steam units) is split out of 'Turbovapor' by plant.
TYPE_TO_FUEL = {
    "Hidráulica > 50 MW": "Hydro", "Hidráulica < 50 MW": "Hydro",
    "Ciclos Combinados": "Gas", "Turbina a gas": "Gas", "Turbovapor": "Gas",
    "Nuclear": "Nuclear", "Motor Diesel": "Oil",
    "Eólica": "Wind", "Solar": "Solar", "Biogas": "Bioenergy", "Biomasa": "Bioenergy",
}
# machine-list TECNOLOGIA -> annual-table TIPO
TECH_TO_TYPE = {
    "Ciclos Combinados": "Ciclos Combinados", "Turbina a gas": "Turbina a gas", "Turbovapor": "Turbovapor",
    "Nuclear": "Nuclear", "Motor Diesel": "Motor Diesel", "Eólica": "Eólica", "Solar": "Solar",
    "Biogas": "Biogas", "Biomasa": "Biomasa", "Hidráulica renovable": "Hidráulica < 50 MW",
    "Hidráulica": "Hidráulica > 50 MW",
}


def download():
    for url in (XLSX, None):
        if url is None:   # fall back to the download-manager link on the CAMMESA page
            page = requests.get(PAGE, headers=HEADERS, timeout=(20, 90)).text
            links = re.findall(r"https://cammesaweb\.cammesa\.com/download/potencia-instalada/\?wpdmdl=\d+[^\"'<\s]*",
                                page)
            if not links:
                break
            url = links[0].replace("&#038;", "&")
        for attempt in range(3):
            try:
                r = requests.get(url, headers=HEADERS, timeout=(20, 300))
                if r.ok and r.content[:2] == b"PK":
                    print(f"Downloaded {r.url}: {len(r.content):,} B (last-modified {r.headers.get('last-modified')})",
                          flush=True)
                    return r.content, r.url, r.headers.get("last-modified")
                print(f"  {url}: HTTP {r.status_code} {r.content[:20]!r}", flush=True)
            except requests.RequestException as e:
                print(f"  {url}: attempt {attempt + 1} {type(e).__name__}", flush=True)
    raise RuntimeError("Could not download CAMMESA Potencia Instalada.xlsx")


def num(v):
    return pd.to_numeric(pd.Series([v]).astype(str).str.replace(",", ".", regex=False), errors="coerce").iloc[0]


def parse_annual(sheet):
    """CAMMESA year-end MW by machine type: DataFrame index=year, columns=TIPO."""
    first = sheet.iloc[:, 0].astype(str).str.strip()
    hdr = first.index[first.eq("TIPO")][0]
    years = [int(float(y)) if pd.notna(y) and str(y).strip() not in ("", "nan") else None for y in sheet.iloc[hdr, 1:]]
    out = {}
    for i in range(hdr + 1, len(sheet)):
        label = str(sheet.iloc[i, 0]).strip()
        if label.upper().startswith("POTENCIA") or label in ("", "nan"):
            break
        out[label] = {y: num(v) for y, v in zip(years, sheet.iloc[i, 1:]) if y}
    return pd.DataFrame(out)


def parse_machines(sheet):
    """Machine list for the latest month (columns AÑO, MES, MAQUINA, ..., POTENCIA INSTALADA, FECHA HABILITACION)."""
    first = sheet.iloc[:, 0].astype(str).str.strip()
    hdr = first.index[first.eq("AÑO") & sheet.iloc[:, 1].astype(str).str.strip().eq("MES")][0]
    names = [str(c).strip() for c in sheet.iloc[hdr]]
    cols = {}
    for key, pat in [("mes", r"^MES$"), ("maquina", r"^MAQUINA$"), ("central", r"^CENTRAL$"),
                     ("agente_desc", r"AGENTE DESC"), ("tipo_maquina", r"TIPO MAQUINA"), ("tecnologia", r"TECNOLOG"),
                     ("mw", r"POTENCIA INSTALADA"), ("habilitacion", r"HABILITACI")]:
        idx = [i for i, n in enumerate(names) if re.search(pat, n, re.I)]
        cols[key] = idx[0] if idx else None
    rows = []
    for i in range(hdr + 1, len(sheet)):
        r = sheet.iloc[i]
        if pd.isna(r.iloc[cols["maquina"]]):
            continue
        rows.append({k: (r.iloc[c] if c is not None else None) for k, c in cols.items()})
    m = pd.DataFrame(rows)
    m["mw"] = pd.to_numeric(m["mw"], errors="coerce")
    m["mes"] = pd.to_datetime(m["mes"], errors="coerce")
    m["habilitacion"] = pd.to_datetime(m["habilitacion"], errors="coerce")
    m["tecnologia"] = m["tecnologia"].astype(str).str.strip()
    return m[m["mw"].notna()]


def parse_retirements(sheet):
    """Machines retired since 2021: maquina, central, tecnologia, MW, month. Rows whose columns are shifted are read
    by value type (the MW is the last plain number, the month the last date)."""
    first = sheet.iloc[:, 0].astype(str).str.strip()
    hdr = first.index[first.eq("AÑO")][0]
    rows = []
    for i in range(hdr + 1, len(sheet)):
        r = list(sheet.iloc[i])
        if pd.isna(r[1]):
            continue
        dates = [v for v in r if isinstance(v, dt.datetime) and not pd.isna(v)]
        nums = [float(v) for v in r[2:] if isinstance(v, numbers.Number) and not isinstance(v, bool)
                and not pd.isna(v)]
        texts = [str(v).strip() for v in r if isinstance(v, str)]
        tech = next((t for t in texts if t in TECH_TO_TYPE), None)
        if tech is None:
            tech = next((t for t in texts if t in ("TG", "DI", "TV", "CC")), None)
            tech = {"TG": "Turbina a gas", "DI": "Motor Diesel", "TV": "Turbovapor", "CC": "Ciclos Combinados"}.get(tech)
        rows.append({"maquina": r[1], "central": r[2], "agente_desc": r[4], "tecnologia": tech,
                     "mw": nums[-1] if nums else None,
                     "month": pd.Timestamp(dates[-1]).to_period("M").to_timestamp() if dates else pd.NaT})
    return pd.DataFrame(rows)


def type_of(row):
    t = TECH_TO_TYPE.get(row["tecnologia"])
    if t == "Hidráulica > 50 MW" and str(row.get("tipo_maquina", "")).upper() in ("HR", "HI<50"):
        return "Hidráulica < 50 MW"
    return t or row["tecnologia"]


def fuel_of(tipo, central):
    if tipo == "Turbovapor" and central is not None and COAL_PLANTS.search(str(central)):
        return "Coal"
    return TYPE_TO_FUEL.get(tipo, "Other")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", default=OUT)
    args = ap.parse_args()

    content, url, modified = download()
    xl = pd.ExcelFile(io.BytesIO(content))
    resumen = xl.parse("RESUMEN por MAQUINA", header=None)
    annual = parse_annual(resumen)
    machines = parse_machines(resumen)
    baja_sheet = next((s for s in xl.sheet_names if "baja" in s.lower()), None)
    retired = parse_retirements(xl.parse(baja_sheet, header=None)) if baja_sheet else pd.DataFrame()
    latest = machines["mes"].max().to_period("M").to_timestamp()
    cur_year = latest.year
    print(f"Annual table {annual.index.min()}-{annual.index.max()}; machine list for {latest:%Y-%m}: {len(machines)} "
          f"machines, {machines['mw'].sum():,.1f} MW (table {cur_year}: {annual.loc[cur_year].sum():,.1f} MW); "
          f"{machines['habilitacion'].notna().sum()} with an enabling date; {len(retired)} retirements", flush=True)
    unknown = sorted(set(machines["tecnologia"]) - set(TECH_TO_TYPE))
    if unknown:
        print(f"  machine technologies not in the annual table mapping (-> Other): {unknown}", flush=True)
    missing_types = sorted(set(annual.columns) - set(TYPE_TO_FUEL))
    if missing_types:
        print(f"  annual-table types counted as Other: {missing_types}", flush=True)

    machines["tipo"] = machines.apply(type_of, axis=1)
    machines["fuel"] = [fuel_of(t, f"{c} {d}")
                        for t, c, d in zip(machines["tipo"], machines["central"], machines["agente_desc"])]
    coal = machines[machines["fuel"] == "Coal"]
    if not retired.empty:
        retired["tipo"] = retired["tecnologia"].map(TECH_TO_TYPE).fillna("Other")
        retired["fuel"] = [fuel_of(t, f"{c} {d}")
                           for t, c, d in zip(retired["tipo"], retired["central"], retired["agente_desc"])]

    # year-end (and latest month) anchors by fuel, from CAMMESA's table; coal taken out of Turbovapor by plant
    anchors = pd.DataFrame({f: annual[[t for t in annual.columns if TYPE_TO_FUEL.get(t, "Other") == f]].sum(axis=1)
                            for f in std.FUELS})
    # coal units in service at each year end (by enabling date; undated units count from the start)
    anchors["Coal"] = [coal.loc[coal["habilitacion"].isna() | (coal["habilitacion"].dt.year <= y), "mw"].sum()
                       for y in anchors.index]
    anchors["Gas"] = anchors["Gas"] - anchors["Coal"]
    months = pd.date_range(std.START, latest, freq="MS")

    adds = machines[machines["habilitacion"].notna()].assign(month=lambda d: d["habilitacion"].dt.to_period("M")
                                                             .dt.to_timestamp())
    add_m = adds.groupby(["month", "fuel"])["mw"].sum().unstack(fill_value=0.0)
    ret_m = (retired.dropna(subset=["mw", "month"]).groupby(["month", "fuel"])["mw"].sum().unstack(fill_value=0.0)
             if not retired.empty else pd.DataFrame())
    dated = pd.DataFrame(0.0, index=months, columns=std.FUELS)
    dated = dated.add(add_m.reindex(index=months, columns=std.FUELS, fill_value=0.0), fill_value=0.0)
    if not ret_m.empty:
        dated = dated.sub(ret_m.reindex(index=months, columns=std.FUELS, fill_value=0.0), fill_value=0.0)

    rows, undated = {}, []
    for year in range(std.START.year, cur_year + 1):
        if year - 1 not in anchors.index:
            continue
        level = anchors.loc[year - 1, std.FUELS].astype(float).copy()
        ym = [m for m in months if m.year == year]
        for m in ym:
            level = level + dated.loc[m, std.FUELS]
            rows[m] = level.copy()
        if ym and year in anchors.index:
            end = ym[-1]
            resid = anchors.loc[year, std.FUELS].astype(float) - rows[end]
            rows[end] = rows[end] + resid
            undated.append(resid.rename(end))
    by_fuel = pd.DataFrame(rows).T.sort_index()
    by_fuel = by_fuel.clip(lower=0.0)

    saved = std.load_monthly(args.out)
    keep = saved.index[saved.index < (saved.index.max() - pd.DateOffset(months=REFRESH_MONTHS - 1))] \
        if not saved.empty else pd.DatetimeIndex([])
    monthly = std.standard(by_fuel)
    if len(keep):
        monthly = pd.concat([saved.loc[keep, std.COLUMNS], monthly.drop(index=keep, errors="ignore")]).sort_index()

    undated_df = pd.DataFrame(undated).round(1)
    undated_df.index.name = "applied_in"
    dated_out = dated.loc[:, dated.ne(0).any()].round(1)
    dated_out.index.name = "date"
    annual_out = annual.copy()
    annual_out.index.name = "year (Dec; latest year = latest month)"
    annual_out["TOTAL"] = annual_out.sum(axis=1)
    mach_out = machines[["maquina", "central", "agente_desc", "tipo_maquina", "tecnologia", "tipo", "fuel", "mw",
                         "habilitacion"]].rename(columns={"mw": "MW"}).set_index("maquina")
    ret_out = retired.set_index("maquina") if not retired.empty else pd.DataFrame()
    check = std.ember_check("Argentina", monthly)
    if not check.empty:
        check = check.set_index("year")

    notes = [
        "UNITS",
        "MW installed at the end of each month ('date' is the 1st of the month), as counted by CAMMESA for the "
        "wholesale electricity market (MEM, national grid SADI).",
        "",
        "COVERAGE",
        f"Monthly, {monthly.index.min():%b %Y} to {monthly.index.max():%b %Y}. Every December, and the latest month, "
        "equals CAMMESA's published installed capacity; months in between move with the machines CAMMESA lists as "
        "enabled (by FECHA HABILITACION) or retired (by MES BAJA) in that month. Changes CAMMESA does not date are "
        "applied in December / the latest month (sheet 'Undated changes'; e.g. CAMMESA's large-hydro total steps "
        "down from 10,834 MW in 2023 to 9,639 MW in 2024 with no machine retirement listed).",
        "",
        "SOURCE",
        f"CAMMESA, Sintesis Mensual > Estadisticas > 'Potencia Instalada' ({PAGE}); file {XLSX} (sheets 'RESUMEN por "
        "MAQUINA' and 'Potencia dada de baja'). Updated weekly by GitHub Actions (argentina_power_capacity.yml).",
        "",
        "MAPPING (CAMMESA machine type -> column)",
        "Hydro_MW = Hidraulica > 50 MW (including the Rio Grande and Los Reyunos pumped-storage plants) + Hidraulica "
        "< 50 MW. Nuclear_MW = Nuclear (Atucha I/II, Embalse). Wind_MW = Eolica. Solar_MW = Solar. Bioenergy_MW = "
        "Biogas + Biomasa.",
        "CAMMESA publishes thermal capacity by machine type, not fuel: Gas_MW = Ciclos Combinados + Turbina a gas + "
        "Turbovapor (natural-gas plants; most can also burn gasoil or fuel oil); Oil_MW = Motor Diesel (gasoil/fuel "
        "oil engines). Coal_MW would take the Rio Turbio coal plant out of Turbovapor by name, but Rio Turbio is not in "
        "CAMMESA's machine list, so Coal_MW is 0; steam units able to burn coal (e.g. San Nicolas) stay in Gas_MW "
        "because CAMMESA does not split steam turbines by fuel.",
        "Other_MW = any machine type outside this list (none in the current file). Argentina has no geothermal; "
        "batteries are not in the CAMMESA table. Total_MW = sum of the fuel columns.",
        "",
        "SHEETS",
        "Monthly: standard table. CAMMESA annual by type: CAMMESA's year-end MW by machine type (latest year = latest "
        "month). Dated changes: MW enabled minus MW retired in each month, by fuel. Undated changes: the part of each "
        "year's change applied in December. Machines (latest): CAMMESA's machine list for the latest month. "
        "Retirements: machines retired since 2021. Ember check: December values vs Ember's yearly capacity.",
    ]
    std.write(args.out, monthly, {"CAMMESA annual by type": annual_out, "Dated changes": dated_out,
                                  "Undated changes": undated_df, "Machines (latest)": mach_out,
                                  "Retirements": ret_out, "Ember check": check},
              notes, {"UNITS", "COVERAGE", "SOURCE", "MAPPING (CAMMESA machine type -> column)", "SHEETS"})


if __name__ == "__main__":
    main()
