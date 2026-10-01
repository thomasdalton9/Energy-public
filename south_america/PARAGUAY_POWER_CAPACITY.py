"""
Paraguay installed generation capacity (Paraguay's share), annual, in the standard capacity layout
(sheet "Monthly": date, Hydro_MW ... Other_MW, Total_MW; one row per year dated 1 January).

  Hydro = 50% of Itaipu (20 x 700 MW = 14,000 MW; itaipu.gov.br, read live) + 50% of Yacyreta
          (20 groups, 3,200 MW; eby.gov.py 'Datos tecnicos', read live) + ANDE's own hydro (Acaray, 210 MW),
          taken as Ember's Paraguay hydro capacity less the two binational halves (ANDE's statistics site is
          behind a bot captcha, so Ember - which compiles from ANDE - is the fallback for ANDE's plants).
  Oil, Bioenergy, Solar, other = Ember yearly capacity for Paraguay (fallback, labelled).

Both binational plants have been complete since 2011 (Yacyreta's 20th group) and 2007 (Itaipu's 20th
unit); Ember carries the latest year forward until it publishes the next.

Usage: python3 PARAGUAY_POWER_CAPACITY.py [--out PATH]
"""
import argparse
import io
import os
import re
import sys

import pandas as pd
import requests

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import power_capacity_std as std  # noqa: E402

OUT = "output/Data and Chart Outputs/paraguay_power_capacity.xlsx"
UA = {"User-Agent": "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36"}
ITAIPU_MW, YACYRETA_MW = 14000.0, 3200.0   # used only if the live pages cannot be read


def itaipu_mw(s):
    try:
        t = s.get("https://www.itaipu.gov.br/energia/geracao", timeout=60).text
        m = re.search(r'Capacidade instalada</div>.{0,400}?data-to-value="([\d.,]+)"', t, re.S)
        if m:
            return float(m.group(1).replace(",", ".")) * 1000, "itaipu.gov.br (live)"
    except Exception as e:  # noqa: BLE001
        print(f"  Itaipu page: {e}", flush=True)
    return ITAIPU_MW, "Itaipu Binacional, 20 x 700 MW (constant)"


def yacyreta_mw(s):
    try:
        t = s.get("https://www.eby.gov.py/datos-tecnicos/", timeout=60).text
        m = re.search(r"POTENCIA M[AÁ]XIMA INSTALADA,?\s*20 GRUPOS\s*([\d.]+)\s*MW", re.sub(r"<[^>]+>", " ", t), re.I)
        if m:
            return float(m.group(1).replace(".", "")), "eby.gov.py Datos tecnicos (live)"
    except Exception as e:  # noqa: BLE001
        print(f"  EBY page: {e}", flush=True)
    return YACYRETA_MW, "EBY, 20 groups (constant)"


def ember_capacity():
    r = requests.get(std.EMBER_URL, timeout=(15, 300), headers={"User-Agent": "gas-demand-scripts/1.0"})
    r.raise_for_status()
    e = pd.read_csv(io.BytesIO(r.content), low_memory=False)
    e = e[(e["Area"] == "Paraguay") & (e["Category"] == "Capacity") & (e["Subcategory"] == "Fuel") & (e["Unit"] == "GW")]
    return e.pivot_table(index="Year", columns="Variable", values="Value", aggfunc="sum") * 1000.0   # MW


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", default=OUT)
    args = ap.parse_args()
    s = requests.Session()
    s.headers.update(UA)
    it, it_src = itaipu_mw(s)
    yc, yc_src = yacyreta_mw(s)
    print(f"Itaipu {it:,.0f} MW ({it_src}); Yacyreta {yc:,.0f} MW ({yc_src})", flush=True)
    em = ember_capacity()
    years = [y for y in em.index if y >= 2021]
    rows, detail = {}, []
    for y in years:
        e = em.loc[y].fillna(0.0)
        binational = 0.5 * it + 0.5 * yc
        ande_hydro = max(float(e.get("Hydro", 0.0)) - binational, 0.0)
        by = {"Hydro": binational + ande_hydro, "Oil": e.get("Other Fossil", 0.0), "Gas": e.get("Gas", 0.0),
              "Coal": e.get("Coal", 0.0), "Solar": e.get("Solar", 0.0), "Wind": e.get("Wind", 0.0),
              "Bioenergy": e.get("Bioenergy", 0.0), "Nuclear": e.get("Nuclear", 0.0),
              "Other": e.get("Other Renewables", 0.0)}
        rows[pd.Timestamp(y, 1, 1)] = by
        detail.append({"year": y, "Itaipu_total_MW": it, "Itaipu_Paraguay_50pct_MW": 0.5 * it,
                       "Yacyreta_total_MW": yc, "Yacyreta_Paraguay_50pct_MW": 0.5 * yc,
                       "ANDE_hydro_Acaray_MW (Ember hydro - binational halves)": round(ande_hydro, 1),
                       "Ember_hydro_MW": float(e.get("Hydro", 0.0))})
    monthly = std.standard(pd.DataFrame.from_dict(rows, orient="index"))
    plants = pd.DataFrame(detail).set_index("year")
    check = std.ember_check("Paraguay", monthly)
    notes = [
        "UNITS",
        "MW installed, Paraguay's share; one row per YEAR dated 1 January (Ember's yearly capacity years). Paraguay owns "
        "50% of Itaipu (with Brazil) and 50% of Yacyreta (with Argentina); only those halves are counted.",
        "",
        "SOURCES",
        f"Itaipu: {it:,.0f} MW, {it_src} - https://www.itaipu.gov.br/energia/geracao",
        f"Yacyreta: {yc:,.0f} MW, {yc_src} - https://www.eby.gov.py/datos-tecnicos/",
        "ANDE's own plants (Acaray hydro 210 MW, small thermal and solar): Ember yearly capacity data "
        "(https://ember-energy.org/data/yearly-electricity-data/), used because ande.gov.py is behind a bot captcha. "
        "ANDE hydro = Ember's Paraguay hydro capacity less the two binational halves. Oil = Ember 'Other Fossil'.",
        "",
        "SHEETS",
        "Monthly: standard layout. Plants: the binational capacities and the ANDE residual. Ember check: this "
        "workbook vs Ember's totals (should match: Ember also counts Paraguay's halves).",
    ]
    std.write(args.out, monthly, {"Plants": plants, "Ember check": check}, notes, {"UNITS", "SOURCES", "SHEETS"})


if __name__ == "__main__":
    main()
