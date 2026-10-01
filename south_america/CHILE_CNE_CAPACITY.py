"""
Chile installed generation capacity by technology, monthly from January
2021, from CNE's monthly "Reporte Capacidad Instalada Generacion" workbook
(https://www.cne.cl/wp-content/uploads/YYYY/MM/Capacidad_Instalada_Generacion.xlsx,
linked from https://www.cne.cl/estadisticas/electricidad/). No key needed.

The workbook is a snapshot of every generating plant in operation in the
national grid (SEN, one row per plant) and the medium-size systems of Los
Lagos, Aysen, Magallanes and Easter Island (one row per unit): sheet 'base'
with the plant's technology (tipo_de_energia), start of service
(fecha_puesta_servicio_central) and net capacity (potencia_neta_mw). CNE
replaces it every month (the September 2026 upload holds August 2026), and
earlier uploads are not kept on cne.cl, so:
  - months before the first run are rebuilt from today's snapshot, counting
    each plant from the month it entered service. Plants retired since 2021
    (notably the coal units closed under the decarbonisation plan) are not in
    the snapshot, so those months understate capacity that has since closed;
  - from the first run on, each saved month is kept as it was computed and
    only the latest two months are recomputed from each new upload, so
    retirements show up from then on. Sheet 'Snapshots' keeps each upload's
    own totals by technology.

The newest upload is found through cne.cl's public WordPress media search,
as in CHILE_POWER_DAILY.py (whose download helper is reused). CEN's own site
and API are not used (browser challenge / key). Found via
discovery_archive/south_america/POWER_CAPACITY_PROBE.py and PROBE2.py.

Usage: python3 CHILE_CNE_CAPACITY.py [--out <xlsx>]
"""
import argparse
import io
import os
import re
import sys

import pandas as pd
import requests

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import CHILE_POWER_DAILY as cne  # noqa: E402  (cne.cl media search + download helper)
import power_capacity_std as std  # noqa: E402

OUT = "output/Data and Chart Outputs/chile_power_capacity.xlsx"
SEARCH = "Capacidad_Instalada_Generacion"
REFRESH_MONTHS = 2
MONTHS_ES = {m: i + 1 for i, m in enumerate(["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto",
                                             "septiembre", "octubre", "noviembre", "diciembre"])}

# CNE tipo_de_energia -> standard fuel
MAPPING = {
    "Hidráulica Embalse": "Hydro", "Hidráulica Pasada": "Hydro", "Mini Hidráulica Pasada": "Hydro",
    "Gas Natural": "Gas", "Propano": "Gas",
    "Eólica": "Wind",
    "Solar": "Solar", "Solar Fotovoltaica": "Solar", "Solar-CSP": "Solar",
    "Carbón": "Coal", "Carbón - Petcoke": "Coal", "Petcoke": "Coal",
    "Petróleo Diesel": "Oil", "Fuel Oil Nro. 6": "Oil",
    "Biomasa": "Bioenergy", "Biogas": "Bioenergy", "Biomasa-Petróleo N°6": "Bioenergy",
    "Geotérmica": "Other", "Cogeneración": "Other",
}


def latest_upload():
    r = requests.get(cne.MEDIA_API, params={"search": SEARCH, "per_page": 50, "_fields": "date,modified,source_url"},
                     headers=cne.HEADERS, timeout=(10, 60))
    r.raise_for_status()
    items = [i for i in r.json() if re.search(r"/Capacidad_Instalada_Generacion(-\d+)?\.xlsx$", i.get("source_url", ""),
                                              re.I)]
    if items:
        best = max(items, key=lambda i: (i.get("date") or "", i.get("modified") or ""))
        return best["source_url"], best.get("modified") or best.get("date")
    r = requests.get(cne.STATS_PAGE, headers=cne.HEADERS, timeout=(10, 60))
    m = re.findall(r"https://www\.cne\.cl/wp-content/uploads/\d{4}/\d{2}/Capacidad_Instalada_Generacion(?:-\d+)?\.xlsx",
                   r.text)
    if not m:
        raise RuntimeError("No Capacidad_Instalada_Generacion.xlsx found on cne.cl")
    return sorted(m)[-1], ""


def report_month(xl, base):
    """Month the snapshot describes: PORTADA note 'capacidad instalada de agosto de 2026', else fecha_act - 1 month."""
    try:
        cover = xl.parse("PORTADA", header=None).astype(str).values.ravel()
        for cell in cover:
            m = re.search(r"capacidad instalada de (\w+) de (\d{4})", cell, re.I)
            if m and m.group(1).lower() in MONTHS_ES:
                return pd.Timestamp(int(m.group(2)), MONTHS_ES[m.group(1).lower()], 1)
    except Exception:  # noqa: BLE001
        pass
    act = pd.to_datetime(base.get("fecha_act"), errors="coerce").max()
    return (act - pd.DateOffset(months=1)).to_period("M").to_timestamp()


def parse(content):
    xl = pd.ExcelFile(io.BytesIO(content))
    raw = xl.parse("base", header=None)
    hdr = raw.index[raw.iloc[:, 0].astype(str).str.strip().str.lower().eq("sistema")][0]
    base = raw.iloc[hdr + 1:].copy()
    base.columns = [str(c).strip() for c in raw.iloc[hdr]]
    base = base[base["sistema"].notna()]
    base["MW"] = pd.to_numeric(base["potencia_neta_mw"].astype(str).str.replace(",", ".", regex=False),
                               errors="coerce").fillna(0.0)
    base["tipo_de_energia"] = base["tipo_de_energia"].astype(str).str.strip()
    base["fuel"] = base["tipo_de_energia"].map(MAPPING).fillna("Other")
    base["start"] = pd.to_datetime(base["fecha_puesta_servicio_central"], errors="coerce")
    base["month"] = base["start"].dt.to_period("M").dt.to_timestamp()
    return base, report_month(xl, base)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", default=OUT)
    args = ap.parse_args()

    url, modified = latest_upload()
    print(f"Newest CNE capacity upload: {url} (modified {modified})", flush=True)
    allrows, snap_month = parse(cne.download(url))
    # plants still in test operation ('En Pruebas') are listed but not counted, matching CNE's own SEN total
    in_op = allrows["estado"].astype(str).str.strip().str.lower().str.startswith("en operaci")
    base = allrows[in_op].copy()
    if (~in_op).any():
        print(f"  not counted (estado != En Operacion): {allrows.loc[~in_op, 'MW'].sum():,.1f} MW in "
              f"{(~in_op).sum()} rows", flush=True)
    other = sorted(base.loc[~base["tipo_de_energia"].isin(MAPPING), "tipo_de_energia"].unique())
    if other:
        print(f"  unmapped technologies counted as Other: {other}", flush=True)
    print(f"Snapshot for {snap_month:%Y-%m}: {len(base):,} rows, {base['MW'].sum():,.1f} MW net "
          f"(SEN {base.loc[base['sistema'] == 'SEN', 'MW'].sum():,.1f} MW); "
          f"{base['start'].isna().sum()} rows without a start date", flush=True)

    months = pd.date_range(std.START, snap_month, freq="MS")
    b = base.copy()
    b["month"] = b["month"].fillna(months[0]).clip(lower=months[0])   # undated or pre-2021 -> in from the start
    add = b.groupby(["month", "fuel"])["MW"].sum().unstack(fill_value=0.0)
    by_fuel = add.reindex(add.index.union(months), fill_value=0.0).sort_index().cumsum().reindex(months)
    add_t = b.groupby(["month", "tipo_de_energia"])["MW"].sum().unstack(fill_value=0.0)
    by_tech_new = add_t.reindex(add_t.index.union(months), fill_value=0.0).sort_index().cumsum().reindex(months).round(1)

    saved = std.load_monthly(args.out)
    keep = saved.index[saved.index < (saved.index.max() - pd.DateOffset(months=REFRESH_MONTHS - 1))] \
        if not saved.empty else pd.DatetimeIndex([])
    monthly = std.standard(by_fuel)
    old_tech = std.load_sheet(args.out, "By CNE technology", index_col=0)
    by_tech = by_tech_new
    if len(keep):
        monthly = pd.concat([saved.loc[keep, std.COLUMNS], monthly.drop(index=keep, errors="ignore")]).sort_index()
        if not old_tech.empty:
            old_tech.index = pd.to_datetime(old_tech.index, errors="coerce")
            by_tech = pd.concat([old_tech[old_tech.index.isin(keep)], by_tech_new.drop(index=keep, errors="ignore")]
                                ).sort_index().fillna(0.0)
        print(f"kept {len(keep)} saved month(s) to {keep.max():%Y-%m}", flush=True)
    by_tech.index.name = "date"

    # each upload's own totals (one row per snapshot month), kept across runs
    snap = base.groupby("tipo_de_energia")["MW"].sum().round(2).to_frame().T
    snap.index = pd.DatetimeIndex([snap_month], name="snapshot_month")
    snap.insert(0, "source_url", url)
    snap.insert(1, "Total_MW", round(base["MW"].sum(), 2))
    snaps = std.load_sheet(args.out, "Snapshots", index_col=0)
    if not snaps.empty:
        snaps.index = pd.to_datetime(snaps.index, errors="coerce")
        snaps = snaps.drop(index=[snap_month], errors="ignore")
    snaps = pd.concat([snaps, snap]).sort_index()
    snaps.index.name = "snapshot_month"

    by_system = base.pivot_table(index="tipo_de_energia", columns="sistema", values="MW", aggfunc="sum",
                                 fill_value=0.0).round(2)
    by_system["fuel"] = by_system.index.map(lambda t: MAPPING.get(t, "Other"))
    plants = allrows[["sistema", "subsistema", "central", "estado", "start", "tipo_de_energia", "fuel", "MW",
                   "medio_generacion", "region_nombre"]].rename(columns={"MW": "potencia_neta_MW"}) \
        .sort_values(["fuel", "potencia_neta_MW"], ascending=[True, False]).set_index("central")
    check = std.ember_check("Chile", monthly)
    if not check.empty:
        check = check.set_index("year")

    notes = [
        "UNITS",
        "MW of net installed capacity (potencia neta) in service at the end of each month; 'date' is the 1st of the "
        "month.",
        "",
        "COVERAGE",
        f"Monthly, {monthly.index.min():%b %Y} to {monthly.index.max():%b %Y}: the national grid (SEN) plus the "
        "medium-size systems of Los Lagos, Aysen, Magallanes and Easter Island, including small distributed "
        "generators (PMGD); plants still in test operation (En Pruebas) are not counted. The latest month is the "
        "CNE upload's report month (the September upload holds August).",
        "RETIREMENTS: CNE keeps only the current snapshot on cne.cl and it has no retirement dates, so months filled "
        "in on the first run (Jan 2021 to Aug 2026) count only plants still in service in Aug 2026, each from the month "
        "it entered service. Capacity retired in between (mainly coal units closed in 2021-2025, plus some diesel) is "
        "missing from those months, so Coal_MW reads flat while actual coal capacity fell. From the first run on, "
        "saved months are kept and only the latest two are recomputed from each new upload, so later retirements "
        "show. Battery storage is not in CNE's capacity report.",
        "",
        "SOURCE",
        "CNE (Comision Nacional de Energia), Estadisticas > Electricidad > 'Capacidad Instalada de Generacion', "
        "https://www.cne.cl/estadisticas/electricidad/ ; file e.g. "
        "https://www.cne.cl/wp-content/uploads/2026/09/Capacidad_Instalada_Generacion.xlsx (sheet 'base'). Updated "
        "weekly by GitHub Actions (chile_power_capacity.yml).",
        "",
        "MAPPING (CNE tipo_de_energia -> column)",
        "Hydro_MW = Hidraulica Embalse + Hidraulica Pasada + Mini Hidraulica Pasada (Chile has no pumped storage in "
        "operation). Gas_MW = Gas Natural + Propano. Wind_MW = Eolica. Solar_MW = Solar / Solar Fotovoltaica (PV) + "
        "Solar-CSP (Cerro Dominador). Coal_MW = Carbon + Carbon - Petcoke + Petcoke.",
        "Oil_MW = Petroleo Diesel + Fuel Oil Nro. 6. Bioenergy_MW = Biomasa + Biogas + Biomasa-Petroleo N6 (biomass "
        "boilers with fuel-oil backup). Other_MW = Geotermica (Cerro Pabellon) + Cogeneracion. Nuclear_MW = 0.",
        "Total_MW = sum of the fuel columns.",
        "",
        "SHEETS",
        "Monthly: standard table. By CNE technology: the same by CNE label. Snapshots: each CNE upload's own "
        "totals by technology (one row per report month, kept across runs). By system (latest): latest MW by "
        "technology and system. Plants (latest): one row per plant/unit in the latest upload. Ember check: December "
        "values vs Ember's yearly capacity.",
    ]
    std.write(args.out, monthly, {"By CNE technology": by_tech, "Snapshots": snaps, "By system (latest)": by_system,
                                  "Plants (latest)": plants, "Ember check": check},
              notes, {"UNITS", "COVERAGE", "SOURCE", "MAPPING (CNE tipo_de_energia -> column)", "SHEETS"})


if __name__ == "__main__":
    main()
