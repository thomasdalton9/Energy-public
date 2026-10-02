"""
Spain daily generation by technology from Red Eléctrica de España (REE, the Spanish TSO) REData API
(apidatos.ree.es, no key): "generacion/estructura-generacion", time_trunc=day (MWh), national total.
Pumped-storage turbining is left out (it shifts energy). Cogeneration, steam turbine and other
fossil/waste technologies are in Other_MWh; combined cycle and gas turbine are Gas.
"""
import sys

import pandas as pd

import europe_common as ec

URL = "https://apidatos.ree.es/en/datos/generacion/estructura-generacion"
EXPLICIT = {"combined cycle": "Gas_MWh", "gas turbine": "Gas_MWh", "diesel engines": "Oil_MWh", "fuel + gas": "Oil_MWh",
            "steam turbine": "Other_MWh", "cogeneration": "Other_MWh", "non-renewable waste": "Other_MWh",
            "renewable waste": "Bioenergy_MWh", "other renewables": "Other_MWh", "solar thermal": "Solar_MWh",
            "hydro-wind": "Other_MWh"}


def column(title):
    t = title.lower()
    for k, v in EXPLICIT.items():
        if k in t:
            return v
    return ec.bucket(t)


def fetch(start, end):
    frames, titles = [], set()
    for a, b in ec.month_chunks(start, end):
        j = ec.get_json(URL, {"start_date": f"{a.isoformat()}T00:00", "end_date": f"{b.isoformat()}T23:59",
                              "time_trunc": "day"}, timeout=(15, 180))
        for item in j.get("included", []):
            title = item["attributes"]["title"]
            if title.lower().startswith("total") or title.lower() in ("generation", "demand"):
                continue
            titles.add(title)
            col = column(title)
            if col is None:
                continue
            for v in item["attributes"]["values"]:
                frames.append((v["datetime"][:10], col, v["value"]))
        print(f"  {a} to {b}: {len(frames)} values so far", file=sys.stderr, flush=True)
    print("technologies:", sorted(titles), file=sys.stderr)
    if not frames:
        return pd.DataFrame()
    d = pd.DataFrame(frames, columns=["date", "col", "mwh"])
    d["date"] = pd.to_datetime(d["date"]).dt.date
    wide = d.pivot_table(index="date", columns="col", values="mwh", aggfunc="sum")
    return ec.to_daily(wide)


NOTES = ec.STANDARD_NOTES + [
    "Spain: REE REData 'estructura-generacion' (national, MWh per day by technology; the day is the Spanish local "
    "day, not UTC). Combined cycle + gas turbine = Gas; solar PV + solar thermal = Solar; cogeneration, steam "
    "turbine, other renewables, waste = Other (cogeneration is mostly gas). Pumped-storage turbining excluded.",
    "",
    "COVERAGE",
    f"Daily from {ec.DEFAULT_START} to yesterday; last {ec.RELOAD_DAYS} days re-pulled each run.",
    "",
    "SOURCE",
    "Red Eléctrica de España REData: https://www.ree.es/en/datos/generation/generation-structure",
]

if __name__ == "__main__":
    ec.run("spain_power_generation_daily.xlsx", fetch, NOTES, __doc__)
