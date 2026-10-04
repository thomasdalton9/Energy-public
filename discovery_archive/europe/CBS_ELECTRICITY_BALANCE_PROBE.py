"""
Probe: CBS StatLine 84575NED 'Elektriciteitsbalans; aanbod en verbruik' (monthly) - columns, units and the last 30 months, to use the
total solar production (incl. rooftop) and total production/consumption in the Dutch power balance. Prints only.
"""
import sys

import pandas as pd
import requests

H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0 Safari/537.36", "Accept": "application/json"}
B = "https://opendata.cbs.nl/ODataApi/odata/84575NED"
props = requests.get(f"{B}/DataProperties", params={"$format": "json"}, headers=H, timeout=90).json().get("value", [])
print("columns:")
for p in props:
    if p.get("Type") in ("Topic", "TopicGroup"):
        print("  ", p.get("Key"), "|", p.get("Type"), "|", p.get("Unit"), "|", p.get("Title"))
rows, url = [], f"{B}/TypedDataSet?$format=json&$top=2000"
while url:
    j = requests.get(url, headers=H, timeout=120).json()
    rows += j.get("value", [])
    url = j.get("odata.nextLink")
df = pd.DataFrame(rows)
print("rows", len(df), "periods", df["Perioden"].iloc[0], "..", df["Perioden"].iloc[-1])
mon = df[df["Perioden"].str.contains("MM")]
pd.set_option("display.width", 250, "display.max_columns", 40)
print(mon.tail(30).to_string())
sys.exit(0)
