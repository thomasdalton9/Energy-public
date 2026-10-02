"""
Brazil installed generation capacity by technology, monthly from January
2021, rebuilt from ANEEL's open-data registers (dadosabertos.aneel.gov.br,
no key). Reuses BRAZIL_ANEEL_CAPACITY.py's SIGA download and number parsing.

  1. SIGA (siga-empreendimentos-geracao.csv): every centralised plant with
     its phase, technology (SigTipoGeracao), fuel, start of commercial
     operation (DatEntradaOperacao) and verified capacity
     (MdaPotenciaFiscalizadaKw). Plants in operation are counted from the
     month of DatEntradaOperacao.
  2. Mini and micro distributed generation, MMGD ("Relacao de
     empreendimentos de Mini e Micro Geracao Distribuida",
     empreendimento-geracao-distribuida.parquet): ~4.7 million rooftop and
     small systems (99% solar) with MdaPotenciaInstaladaKW and the
     registration date DthAtualizaCadastralEmpreend, counted from that month.

Neither register carries a retirement date: SIGA drops plants once they
are decommissioned, so the back-filled history (2021 to the first run)
counts only plants still registered today - capacity retired since 2021
is missing from the earlier months. From the first run on, each saved
month is kept as it was computed (only the latest two months are
recomputed), so retirements after that are captured.

Found via discovery_archive/south_america/POWER_CAPACITY_PROBE.py and
POWER_CAPACITY_PROBE2.py; ANEEL's own year-end totals by plant type
(empreendimento-operacao-historico.csv) are read for a check sheet.

Usage: python3 BRAZIL_ANEEL_CAPACITY_MONTHLY.py [--out <xlsx>] [--no-gd]
"""
import argparse
import io
import os
import sys

import pandas as pd
import requests

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import BRAZIL_ANEEL_CAPACITY as siga  # noqa: E402
import power_capacity_std as std  # noqa: E402

OUT = "output/Data and Chart Outputs/brazil_power_capacity.xlsx"
GD_URL = ("https://dadosabertos.aneel.gov.br/dataset/5e0fafd2-21b9-4d5b-b622-40438d40aba2/resource/"
          "cd29f6eb-e08d-4db7-b6fb-ed6e3b682d27/download/empreendimento-geracao-distribuida.parquet")
HIST_URL = ("https://dadosabertos.aneel.gov.br/dataset/306a6fdb-beb9-4296-bf18-77fa0e076ef1/resource/"
            "e61fd029-5e78-43be-bed4-873b7b11f04c/download/empreendimento-operacao-historico.csv")
REFRESH_MONTHS = 2
BACKFILL_TO = pd.Timestamp("2026-10-01")   # months before the first run: rebuilt from today's register every run

TYPE_TO_FUEL = {"UHE": "Hydro", "PCH": "Hydro", "CGH": "Hydro", "EOL": "Wind", "UFV": "Solar", "UTN": "Nuclear"}
# UTE (thermal) plants by DscFonteCombustivel (SIGA) / DscFonteGeracao (MMGD)
THERMAL_TO_FUEL = {
    "Gás natural": "Gas", "Gás Natural": "Gas",
    "Petróleo": "Oil",
    "Carvão mineral": "Coal",
    "Agroindustriais": "Bioenergy", "Floresta": "Bioenergy", "Resíduos sólidos urbanos": "Bioenergy",
    "Resíduos animais": "Bioenergy", "Biocombustíveis líquidos": "Bioenergy",
    "Outros Fósseis": "Other",
}
GD_THERMAL_GAS = {"Gás Natural"}


def siga_fuel(tipo, fonte):
    if tipo in TYPE_TO_FUEL:
        return TYPE_TO_FUEL[tipo]
    return THERMAL_TO_FUEL.get(fonte, "Other")


def gd_fuel(tipo, fonte):
    if tipo in TYPE_TO_FUEL:
        return TYPE_TO_FUEL[tipo]
    if tipo == "UTE":
        return "Gas" if fonte in GD_THERMAL_GAS else "Bioenergy"
    return "Other"


def month_floor(s):
    return pd.to_datetime(s, errors="coerce").dt.to_period("M").dt.to_timestamp()


def cumulative(frame, months, key):
    """Plants (month_start, key, MW) -> capacity in service at the end of each month, columns by key."""
    f = frame.copy()
    f["month"] = f["month"].fillna(pd.Timestamp("1900-01-01")).clip(lower=pd.Timestamp("1900-01-01"))
    f.loc[f["month"] < months[0], "month"] = months[0]   # everything before 2021 is in the first month
    add = f.groupby(["month", key])["MW"].sum().unstack(fill_value=0.0)
    add = add.reindex(add.index.union(months), fill_value=0.0).sort_index().cumsum()
    return add.reindex(months).fillna(0.0)


def load_siga():
    raw = siga.fetch_raw()
    asof = pd.to_datetime(raw["DatGeracaoConjuntoDados"], errors="coerce").max()
    op = raw[raw["DscFaseUsina"] == "Operação"].copy()
    op["MW"] = siga.to_float_br(op["MdaPotenciaFiscalizadaKw"]) / 1000.0
    op["fuel"] = [siga_fuel(t, f) for t, f in zip(op["SigTipoGeracao"], op["DscFonteCombustivel"])]
    op["month"] = month_floor(op["DatEntradaOperacao"])
    op["type"] = op["SigTipoGeracao"].where(op["SigTipoGeracao"] != "UTE", "UTE " + op["DscFonteCombustivel"].astype(str))
    print(f"SIGA as of {asof:%Y-%m-%d}: {len(raw):,} rows, {len(op):,} in operation, {op['MW'].sum() / 1000:,.2f} GW",
          flush=True)
    return op, asof


def load_gd():
    r = requests.get(GD_URL, headers=siga.HEADERS, timeout=(15, 900))
    r.raise_for_status()
    gd = pd.read_parquet(io.BytesIO(r.content), columns=["SigTipoGeracao", "DscFonteGeracao", "MdaPotenciaInstaladaKW",
                                                         "DthAtualizaCadastralEmpreend", "DatGeracaoConjuntoDados"])
    asof = pd.to_datetime(gd["DatGeracaoConjuntoDados"], errors="coerce").max()
    gd["MW"] = pd.to_numeric(gd["MdaPotenciaInstaladaKW"], errors="coerce").fillna(0.0) / 1000.0
    gd["fuel"] = [gd_fuel(t, f) for t, f in zip(gd["SigTipoGeracao"], gd["DscFonteGeracao"])]
    gd["month"] = month_floor(gd["DthAtualizaCadastralEmpreend"])
    gd["type"] = "MMGD " + gd["SigTipoGeracao"].fillna("unknown").astype(str)
    print(f"MMGD as of {asof:%Y-%m-%d}: {len(gd):,} systems, {gd['MW'].sum() / 1000:,.2f} GW", flush=True)
    return gd[["MW", "fuel", "month", "type"]], asof


def aneel_history():
    """ANEEL's own year-end installed capacity by plant type (centralised plants), MW."""
    try:
        r = requests.get(HIST_URL, headers=siga.HEADERS, timeout=(15, 120))
        r.raise_for_status()
        h = pd.read_csv(io.BytesIO(r.content), sep=";", encoding="utf-8", dtype=str)
    except Exception as e:  # noqa: BLE001
        print(f"  ANEEL history check skipped: {e}", flush=True)
        return pd.DataFrame()
    h["MW"] = siga.to_float_br(h["MdaPotenciaInstaladaKW"]) / 1000.0
    h["date"] = pd.to_datetime(h["AnoReferencia"] + "-" + h["MesReferencia"].str.zfill(2) + "-01", errors="coerce")
    return h.pivot_table(index="date", columns="SigTipoGeracao", values="MW", aggfunc="sum")


def thermal_retirement_gap(hist, plants, months):
    """MW of thermal (UTE) capacity in ANEEL's published totals but missing from the rebuild from today's SIGA
    (plants retired since), at each month: ANEEL - rebuild at each published date, interpolated between dates and
    tapered to 0 at the latest month (the register is complete for today)."""
    if hist.empty or "UTE" not in hist:
        return pd.Series(0.0, index=months)
    ours = cumulative(plants[(plants["source"] == "SIGA") & plants["type"].str.startswith("UTE")].assign(t="UTE"),
                      months, "t")["UTE"]
    pts = {}
    for d in hist.index[hist.index >= std.START]:
        m = d.to_period("M").to_timestamp()
        if m in ours.index and pd.notna(hist.at[d, "UTE"]):
            pts[m] = float(hist.at[d, "UTE"]) - float(ours[m])
    pts[months[-1]] = 0.0
    gap = pd.Series(pts).sort_index().reindex(months).interpolate().bfill().fillna(0.0)
    return gap.clip(lower=0.0)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", default=OUT)
    ap.add_argument("--no-gd", action="store_true", help="skip the 106 MB MMGD download (test runs)")
    args = ap.parse_args()

    op, asof = load_siga()
    last_month = (asof - pd.Timedelta(days=1)).to_period("M").to_timestamp()
    months = pd.date_range(std.START, last_month, freq="MS")
    frames = [op[["MW", "fuel", "month", "type"]].assign(source="SIGA")]
    if not args.no_gd:
        gd, _ = load_gd()
        frames.append(gd.assign(source="MMGD"))
    plants = pd.concat(frames, ignore_index=True)

    by_fuel_new = cumulative(plants, months, "fuel")
    by_type_new = cumulative(plants, months, "type").round(1)
    plants["fuel_source"] = plants["fuel"] + " (" + plants["source"] + ")"
    by_src_new = cumulative(plants, months, "fuel_source").round(1)

    # incremental: saved months stay as computed then (they reflect the register at that time); the latest
    # REFRESH_MONTHS saved months and any new months come from this run's registers
    saved = std.load_monthly(args.out)
    keep = saved.index[saved.index < (saved.index.max() - pd.DateOffset(months=REFRESH_MONTHS - 1))] \
        if not saved.empty else pd.DatetimeIndex([])
    # back-filled months: thermal plants retired since 2021 are missing from today's register; add the gap to
    # ANEEL's published year-end UTE totals (interpolated) under Oil - the retirements are mostly oil/diesel units
    hist = aneel_history()
    gap = thermal_retirement_gap(hist, plants, months)
    back = gap.index < BACKFILL_TO
    by_fuel_new.loc[back, "Oil"] = by_fuel_new.loc[back, "Oil"] + gap[back] if "Oil" in by_fuel_new else gap[back]
    print(f"  retired thermal added to back-filled months (Oil): {gap.iloc[0]:,.0f} MW in {months[0]:%b %Y}, "
          f"{gap[back].iloc[-1]:,.0f} MW in {gap[back].index[-1]:%b %Y}", flush=True)
    keep = keep[keep >= BACKFILL_TO]   # back-filled months are always recomputed (with the adjustment)
    monthly = std.standard(by_fuel_new)
    if len(keep):
        monthly = pd.concat([saved.loc[keep, std.COLUMNS], monthly.drop(index=keep, errors="ignore")]).sort_index()
        print(f"kept {len(keep)} saved month(s) {keep.min():%Y-%m}..{keep.max():%Y-%m}; the rest recomputed", flush=True)

    def merge_saved(new, sheet):
        old = std.load_sheet(args.out, sheet, index_col=0)
        if old.empty or not len(keep):
            return new
        old.index = pd.to_datetime(old.index, errors="coerce")
        old = old[old.index.isin(keep)]
        return pd.concat([old, new.drop(index=keep, errors="ignore")]).sort_index().fillna(0.0)

    by_src = merge_saved(by_src_new, "By source")
    by_type = merge_saved(by_type_new, "By ANEEL type")
    by_src.index.name = by_type.index.name = "date"

    latest = (plants.groupby(["source", "type", "fuel"])["MW"].agg(["count", "sum"])
              .rename(columns={"count": "plants", "sum": "capacity_MW"}).round(1).reset_index().set_index("type"))
    check_hist = pd.DataFrame()
    if not hist.empty:
        ours = cumulative(plants[plants["source"] == "SIGA"].assign(
            t=lambda d: d["type"].str.split(" ").str[0]), months, "t")
        rows = []
        for d in hist.index[hist.index >= std.START]:
            m = d.to_period("M").to_timestamp()
            if m not in ours.index:
                continue
            for t in hist.columns:
                a, b = float(hist.loc[d, t]), float(ours.loc[m].get(t, 0.0))
                if pd.isna(a):
                    continue
                rows.append({"date": m, "aneel_type": t, "aneel_published_MW": round(a, 1),
                             "rebuilt_from_SIGA_MW": round(b, 1),
                             "diff_pct": round(100 * (b - a) / a, 1) if a else None})
        check_hist = pd.DataFrame(rows).set_index("date") if rows else pd.DataFrame()
    check = std.ember_check("Brazil", monthly)
    if not check.empty:
        check = check.set_index("year")

    notes = [
        "UNITS",
        "MW installed at the end of each month ('date' is the 1st of the month). Centralised plants: ANEEL's verified "
        "capacity (MdaPotenciaFiscalizadaKw). Distributed generation: installed capacity declared to ANEEL "
        "(MdaPotenciaInstaladaKW; for solar this is generally the inverter/module rating reported by the distributor).",
        "",
        "COVERAGE",
        f"Monthly, {monthly.index.min():%b %Y} to {monthly.index.max():%b %Y}. Centralised plants from SIGA (SIN and "
        "isolated systems) plus mini and micro distributed generation (MMGD). The latest month is the register as of "
        f"{asof:%Y-%m-%d} (month to date).",
        "RETIREMENTS: neither register has a decommissioning date. Retired plants are removed from SIGA, so a rebuild "
        "from today's register misses capacity retired since 2021 (mainly old oil/diesel thermal units). Back-filled "
        "months (Jan 2021 to Sep 2026) are therefore recomputed on every run and topped up to ANEEL's own published "
        "thermal (UTE) totals: the gap at each published date (2.2 GW end-2021, 1.5 GW end-2022, ...) is "
        "interpolated between dates, tapered to 0 at the latest month and added to Oil_MW. From Oct 2026 each saved "
        "month is kept as computed (only the latest two are recomputed), so later retirements show up directly.",
        "Each SIGA plant is counted in full from its start of commercial operation (DatEntradaOperacao, the first "
        "unit); ANEEL's unit-by-unit release list (unidades-geradoras-liberadas-operacao-comercial) gives similar "
        "yearly additions (2021-2025: 7.6/8.3/10.3/10.8/7.5 GW vs 7.5/8.4/10.3/10.2/7.3 GW here). "
        "MMGD systems are counted from DthAtualizaCadastralEmpreend (date of registration/last register update), "
        "the only date in the file.",
        "",
        "SOURCE",
        "ANEEL open data (https://dadosabertos.aneel.gov.br): SIGA - Sistema de Informacoes de Geracao, "
        f"{siga.URL} ; Relacao de empreendimentos de Mini e Micro Geracao Distribuida, {GD_URL} ; ANEEL year-end "
        f"totals by plant type for the check sheet, {HIST_URL}. Updated weekly by GitHub Actions "
        "(brazil_power_capacity.yml).",
        "",
        "MAPPING (ANEEL -> column)",
        "Hydro_MW = UHE (large hydro, including the pumped-storage capable plants, which ANEEL registers as UHE) + "
        "PCH (small) + CGH (mini) + MMGD hydro. Wind_MW = EOL + MMGD wind. Solar_MW = UFV (centralised PV) + MMGD solar "
        "(distributed PV, ~54 GW). Nuclear_MW = UTN (Angra 1 and 2).",
        "Thermal plants (UTE) by DscFonteCombustivel: Gas_MW = 'Gas natural' (natural gas, including process and "
        "refinery gas listed under it); Oil_MW = 'Petroleo' (diesel, fuel oil, other oil products and refinery gas); "
        "Coal_MW = 'Carvao mineral' (coal, including blast-furnace and coal-derived process gas); Bioenergy_MW = "
        "'Agroindustriais' (sugarcane bagasse etc.), 'Floresta' (wood, black liquor, charcoal), 'Residuos solidos "
        "urbanos', 'Residuos animais', 'Biocombustiveis liquidos' + MMGD biogas/biomass. MMGD natural-gas units -> Gas.",
        "Other_MW = 'Outros Fosseis' (process heat) and anything unclassified. Brazil has no geothermal; ANEEL has no "
        "battery storage in these registers.",
        "Total_MW = sum of the fuel columns.",
        "",
        "SHEETS",
        "Monthly: standard table. By source: each fuel split into centralised (SIGA) and distributed (MMGD). By ANEEL "
        "type: MW by ANEEL plant type (UTE split by fuel). Latest by type: plant count and MW in the latest registers. "
        "ANEEL history check: ANEEL's published year-end MW by plant type vs the rebuild from today's SIGA (the gap is "
        "mostly retired plants, see RETIREMENTS). Ember check: December values vs Ember's yearly capacity.",
    ]
    std.write(args.out, monthly, {"By source": by_src, "By ANEEL type": by_type, "Latest by type": latest,
                                  "ANEEL history check": check_hist, "Ember check": check},
              notes, {"UNITS", "COVERAGE", "SOURCE", "MAPPING (ANEEL -> column)", "SHEETS"})


if __name__ == "__main__":
    main()
