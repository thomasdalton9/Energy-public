"""
Denmark daily generation by type from Energinet (the Danish TSO) Energi Data Service, dataset
GenerationProdTypeExchange (hourly MWh per price area DK1/DK2, no key). Where a hour has both a Preliminary and a
Final version the Final one is used. SolarPowerSelfCon (self-consumed solar) is not counted.
"""
import sys

import pandas as pd

import europe_common as ec

URL = "https://api.energidataservice.dk/dataset/GenerationProdTypeExchange"
MAP = {"OffshoreWindPower": "Wind_MWh", "OnshoreWindPower": "Wind_MWh", "HydroPower": "Hydro_MWh",
       "SolarPower": "Solar_MWh", "Biomass": "Bioenergy_MWh", "Biogas": "Bioenergy_MWh", "Waste": "Other_MWh",
       "FossilGas": "Gas_MWh", "FossilOil": "Oil_MWh", "FossilHardCoal": "Coal_MWh"}


def fetch(start, end):
    frames = []
    for a, b in ec.month_chunks(start, end):
        j = ec.get_json(URL, {"start": f"{a.isoformat()}T00:00", "end": f"{(b + pd.Timedelta(days=1)).date().isoformat()}T00:00",
                              "limit": 0, "sort": "TimeUTC asc"})
        recs = j.get("records", [])
        print(f"  {a} to {b}: {len(recs)} rows", file=sys.stderr, flush=True)
        if recs:
            frames.append(pd.DataFrame(recs))
    if not frames:
        return pd.DataFrame()
    d = pd.concat(frames)
    d["final"] = (d["Version"] == "Final").astype(int)
    d = d.sort_values(["TimeUTC", "PriceArea", "final"]).drop_duplicates(["TimeUTC", "PriceArea"], keep="last")
    d["date"] = pd.to_datetime(d["TimeUTC"]).dt.date
    out = pd.DataFrame(index=sorted(d["date"].unique()))
    for src, col in MAP.items():
        if src in d:
            s = d.groupby("date")[src].sum(min_count=1)
            out[col] = out.get(col, 0) + s.reindex(out.index).fillna(0)
    hours = d.groupby("date")["TimeUTC"].nunique()
    out = out[hours.reindex(out.index) >= 20]
    return ec.to_daily(out)


NOTES = ec.STANDARD_NOTES + [
    "Denmark: Energinet GenerationProdTypeExchange, hourly MWh per price area (DK1 + DK2). Wind = offshore + onshore; "
    "Bioenergy = biomass + biogas; Other = waste. Self-consumed solar not counted; central/decentral CHP is split by "
    "fuel (biomass, gas, coal, oil, waste) in the source.",
    "",
    "COVERAGE",
    f"Daily from {ec.DEFAULT_START} (or the dataset's first record) to yesterday; last {ec.RELOAD_DAYS} days re-pulled each run.",
    "",
    "SOURCE",
    "Energinet Energi Data Service: https://www.energidataservice.dk/tso-electricity/GenerationProdTypeExchange",
]

if __name__ == "__main__":
    ec.run("denmark_power_generation_daily.xlsx", fetch, NOTES, __doc__)
