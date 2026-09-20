import pandas as pd

years = range(2010, 2027)  # 2010 up to and including 2026
cols = ["val_gerhidraulica", "val_gertermica", "val_gereolica", "val_gersolar"]

dfs = []

for year in years:
    url = f"https://ons-aws-prod-opendata.s3.amazonaws.com/dataset/balanco_energia_subsistema_ho/BALANCO_ENERGIA_SUBSISTEMA_{year}.parquet"
    print(f"Downloading {year}...")
    df_year = pd.read_parquet(url)
    dfs.append(df_year)

print("Combining all years...")
df = pd.concat(dfs, ignore_index=True)
print(f"Total rows: {len(df)}")

# Force generation columns to numeric
for col in cols:
    df[col] = pd.to_numeric(df[col], errors="coerce")

df["date"] = pd.to_datetime(df["din_instante"]).dt.date

daily = df.groupby("date")[cols].mean()
daily.columns = ["hydro_mw", "thermal_mw", "wind_mw", "solar_mw"]

daily.to_csv("brazil_generation_daily_2010_2026.csv")
print("Saved to brazil_generation_daily_2010_2026.csv")
print(daily.head())
print(daily.tail())