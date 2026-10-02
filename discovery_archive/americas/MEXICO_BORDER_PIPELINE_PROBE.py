"""One-off probe: do EIA's state-to-state capacity / pipeline-project workbooks carry US-Mexico border crossings? Log only."""
import io
import pandas as pd
import requests

UA = {"User-Agent": "Mozilla/5.0 (energy-data research)"}
BASE = "https://www.eia.gov/naturalgas/pipelines/"
for f in ["EIA-StatetoStateCapacity_Jan2026.xlsx", "EIA-NaturalGasPipelineProjectsAug2026.xlsx"]:
    try:
        r = requests.get(BASE + f, headers=UA, timeout=90)
        print(f"## {f} -> {r.status_code} {len(r.content)}")
        xl = pd.read_excel(io.BytesIO(r.content), sheet_name=None, header=None)
        for sh, df in xl.items():
            print(f"-- sheet {sh!r} {df.shape}")
            mask = df.astype(str).apply(lambda c: c.str.contains("Mexic|Border|Sasabe|Presidio|Ojinaga|Roma|Rio Grande|Nogales|Douglas|Clint|Agua Prieta", case=False, regex=True)).any(axis=1)
            print(df.head(4).to_string(max_colwidth=40)[:900])
            print(f"   rows mentioning Mexico/border: {int(mask.sum())}")
            print(df[mask].head(40).to_string(max_colwidth=45)[:6000])
    except Exception as e:
        print("##", f, "ERROR", type(e).__name__, str(e)[:150])
