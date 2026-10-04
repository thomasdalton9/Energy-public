"""
Indonesia installed RENEWABLE generating capacity by type, monthly, from the Ministry of Energy and Mineral
Resources (ESDM), Directorate General of New & Renewable Energy and Energy Conservation (Ditjen EBTKE).

Source: the "Data Angka" box on https://ebtke.esdm.go.id/ is fed by the open JSON endpoint
https://ebtke.esdm.go.id/api/api/konten/data-angka (no key). Its dataAngkaKapasitasPembangkit record holds
the latest month's national renewable capacity (MW) by plant type: PLTA (hydro), PLTM (mini hydro), PLTMH
(micro hydro), PLTP (geothermal), PLTS (solar), PLTS Atap (rooftop solar), PLTB (wind), PLTBm (biomass),
PLTBg (biogas), PLTSa (waste-to-energy), PLTBn (biofuel), PLT Hybrid, and the total. Only the latest month
is public (the history routes need a login), so the monthly series is built up by this pull: each run reads
the saved workbook (the history store) and adds or replaces the row for the month the record is dated.
Found via discovery_archive/asia/INDONESIA_DISCOVERY7.py / 8.py (rounds 1-8 found no public sub-annual
official generation or demand data for Indonesia - see those scripts).

Writes output/Data and Chart Outputs/indonesia_renewable_capacity.xlsx:
  Monthly  standard capacity layout (south_america/power_capacity_std.py): date (1st of the data month),
           Hydro_MW (PLTA+PLTM+PLTMH), Solar_MW (PLTS+PLTS Atap), Wind_MW (PLTB), Bioenergy_MW
           (PLTBm+PLTBg+PLTSa+PLTBn), Other_MW (PLTP geothermal + PLT Hybrid), fossil columns 0 (NOT covered),
           Total_MW = renewable total.
  Detail   the same months by EBTKE plant type (MW), plus EBTKE's own total and when EBTKE entered the record.
Runs on the 1st and 15th (EBTKE updates the figure every month or two, with a ~2 month lag).

    python3 asia/INDONESIA_EBTKE_CAPACITY.py [--out PATH]
"""
import argparse
import os
import sys
import time

import pandas as pd
import requests

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)
sys.path.insert(0, os.path.join(ROOT, "south_america"))
import power_capacity_std as cap_std  # noqa: E402

URL = "https://ebtke.esdm.go.id/api/api/konten/data-angka"
H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0 Safari/537.36",
     "Accept": "application/json", "Referer": "https://ebtke.esdm.go.id/"}
T = (15, 90)
OUT = os.path.join(ROOT, "output", "Data and Chart Outputs", "indonesia_renewable_capacity.xlsx")
# EBTKE field -> (detail column, standard fuel)
TYPES = [("plta", "PLTA_hydro_MW", "Hydro"), ("pltm", "PLTM_minihydro_MW", "Hydro"),
         ("pltmh", "PLTMH_microhydro_MW", "Hydro"), ("pltp", "PLTP_geothermal_MW", "Other"),
         ("plts", "PLTS_solar_MW", "Solar"), ("plts_atap", "PLTS_Atap_rooftop_solar_MW", "Solar"),
         ("pltb", "PLTB_wind_MW", "Wind"), ("pltbm", "PLTBm_biomass_MW", "Bioenergy"),
         ("pltbg", "PLTBg_biogas_MW", "Bioenergy"), ("pltsa", "PLTSa_waste_MW", "Bioenergy"),
         ("pltbn", "PLTBn_biofuel_MW", "Bioenergy"), ("plt_hybrid", "PLT_Hybrid_MW", "Other")]
MONTHS = {m: i + 1 for i, m in enumerate(
    ["january", "february", "march", "april", "may", "june", "july", "august", "september", "october",
     "november", "december"])}
MONTHS.update({m: i + 1 for i, m in enumerate(
    ["januari", "februari", "maret", "april", "mei", "juni", "juli", "agustus", "september", "oktober",
     "november", "desember"])})


def out(*a):
    print(*a, flush=True)


def fetch():
    for i in range(4):
        try:
            r = requests.get(URL, headers=H, timeout=T)
            r.raise_for_status()
            return r.json()["data"]["dataAngkaKapasitasPembangkit"]
        except (requests.RequestException, ValueError, KeyError, TypeError) as e:
            if i == 3:
                raise
            out(f"  retry: {e!r}")
            time.sleep(5 * (i + 1))


def to_rows(rec):
    month = MONTHS.get(str(rec.get("bulan", "")).strip().lower())
    if not month:
        raise SystemExit(f"Unrecognised month in EBTKE record: {rec.get('bulan')!r}")
    d = pd.Timestamp(int(rec["tahun"]), month, 1)
    det = {col: pd.to_numeric(rec.get(k), errors="coerce") for k, col, _ in TYPES}
    det["EBTKE_total_MW"] = pd.to_numeric(rec.get("total"), errors="coerce")
    det["entered"] = str(rec.get("updated_at") or rec.get("created_at") or "")[:10]
    detail = pd.DataFrame([det], index=pd.DatetimeIndex([d], name="date"))
    by = {}
    for _, col, fuel in TYPES:
        by[fuel] = by.get(fuel, 0.0) + (0.0 if pd.isna(det[col]) else float(det[col]))
    return detail, cap_std.standard(pd.DataFrame([by], index=[d]))


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", default=OUT)
    a = ap.parse_args()
    rec = fetch()
    out(f"EBTKE record: {rec}")
    if str(rec.get("satuan", "MW")).strip().upper() != "MW":
        raise SystemExit(f"Unexpected unit {rec.get('satuan')!r}")
    detail_new, row = to_rows(rec)
    total = float(row["Total_MW"].iloc[0])
    ebtke_total = float(detail_new["EBTKE_total_MW"].iloc[0])
    if not 5000 < total < 200000:
        raise SystemExit(f"Implausible renewable total {total} MW")
    if abs(total - ebtke_total) > 1.0:
        out(f"  note: sum of types {total:.2f} MW vs EBTKE total {ebtke_total:.2f} MW")

    monthly = cap_std.load_monthly(a.out)
    monthly = pd.concat([monthly[~monthly.index.isin(row.index)], row]).sort_index() if not monthly.empty else row
    monthly.index.name = "date"
    detail = cap_std.load_sheet(a.out, "Detail", index_col=0)
    if not detail.empty:
        detail.index = pd.to_datetime(detail.index, errors="coerce")
        detail = detail[detail.index.notna() & ~detail.index.isin(detail_new.index)]
        detail = pd.concat([detail, detail_new]).sort_index()
    else:
        detail = detail_new
    detail.index.name = "date"

    notes = [
        "UNITS",
        "Monthly: installed renewable generating capacity, MW, standard fuel columns (one row per data month, dated "
        "the 1st). Hydro = PLTA + PLTM (mini) + PLTMH (micro); Solar = PLTS + PLTS Atap (rooftop); Wind = PLTB; "
        "Bioenergy = PLTBm (biomass) + PLTBg (biogas) + PLTSa (waste) + PLTBn (biofuel); Other = PLTP (geothermal) "
        "+ PLT Hybrid. Total_MW = renewable total.",
        "Gas_MW, Coal_MW, Oil_MW, Nuclear_MW are 0 here because EBTKE reports renewables only. This is NOT the "
        "whole fleet: fossil capacity (roughly 75 GW) is not in this workbook.",
        "Detail: the same months by EBTKE plant type (MW), EBTKE's own total, and the date EBTKE entered the record.",
        "",
        "COVERAGE",
        "National (PLN, IPP and captive/off-grid renewable plants as counted by Ditjen EBTKE). EBTKE publishes only "
        "the latest month, so the history starts with this pull's first run "
        f"({monthly.index.min():%b %Y}) and grows as EBTKE posts new months (about every 1-2 months, ~2 months' "
        "lag). A month EBTKE posts and replaces between two runs (1st and 15th) can be missed.",
        "",
        "SOURCE",
        "Ministry of Energy and Mineral Resources (ESDM), Ditjen EBTKE, 'Data Angka - Kapasitas Pembangkit EBT': "
        "https://ebtke.esdm.go.id/ (JSON https://ebtke.esdm.go.id/api/api/konten/data-angka).",
    ]
    cap_std.write(a.out, monthly, {"Detail": detail}, notes, {"UNITS", "COVERAGE", "SOURCE"})


if __name__ == "__main__":
    main()
