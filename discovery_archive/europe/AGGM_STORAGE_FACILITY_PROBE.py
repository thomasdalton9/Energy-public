"""
Probe: AGGM data monitor storage series by facility (Haidach RAG + GSA/SEFE, 7Fields, OMV, Uniper, MG-Ost, SSO SEFE Haidach allocations to THE vs VTP,
Entry/ExitSpeicher*_MGM-Allokationen, Netto*SpeicherOesterreich). Prints monthly sums (TWh, from daily kWh) Oct 2021 - latest, and the daily
stock (ArbeitsgasVolumen) month-end value for the stock series. Prints only.
"""
import re
import sys
from datetime import date, timedelta

import pandas as pd
import requests

UA = "Mozilla/5.0"
B = "https://platform.aggm.at/vis-service/api/"
names = sorted(set(re.findall(r'"name":"([^"]+)"', requests.get(B + "ts/attributes", headers={"User-Agent": UA}, timeout=90).text)))
pat = re.compile(r"haidach|7fields|Speicher.*(Einspeise|Ausspeise)|(Entry|Exit)Speicher|NettoE.*Speicher|Speicher.*Arbeitsgas|SSO|Deutschland.*Speicher", re.I)
sel = [n for n in names if pat.search(n) and not re.search(r"Kap|Unterbrech|Anteil|Nutzbares|NetFlow", n)]
print("series", len(sel), flush=True)
rows = {}
s = date(2021, 10, 1)
end = date.today()
while s < end:
    e = min(s + timedelta(days=180), end)
    for i in range(0, len(sel), 15):
        ch = sel[i:i + 15]
        body = {"rangeType": "individual", "from": f"{s:%Y-%m-%d}T06:00:00", "to": f"{e + timedelta(days=1):%Y-%m-%d}T06:00:00", "granularity": "day", "timeseries": ch}
        try:
            r = requests.post(B + "ts/values", json=body, headers={"User-Agent": UA, "Content-Type": "application/json", "Accept": "application/json"}, timeout=120)
            for cd in r.json()["timeSeriesData"]["chartData"]:
                for p in cd["dataSet"]:
                    if p.get("y") is None:
                        continue
                    t = pd.Timestamp(p["x"], unit="ms", tz="UTC").tz_convert("Europe/Vienna").tz_localize(None).normalize()
                    rows[(t, cd["header"]["name"])] = p["y"] / 1e9        # kWh -> TWh
        except Exception as ex:  # noqa: BLE001
            print("fail", s, i, type(ex).__name__, str(ex)[:80], flush=True)
    s = e + timedelta(days=1)
df = pd.Series(rows).unstack()
df.index.name = "date"
df.to_csv("aggm_storage_daily.csv")
stock = [c for c in df if "Arbeitsgas" in c]
flow = [c for c in df if c not in stock]
pd.set_option("display.width", 250, "display.max_columns", 60, "display.max_rows", 200)
short = {c: re.sub(r"_?(Speicher|Menge|ArbeitsgasVolumen|MGM-Allokationen)", "", c)[:14] for c in df}
m = df[flow].resample("MS").sum(min_count=1).round(2)
chg = df[stock].resample("MS").last().diff().round(2)
pick = [c for c in flow if re.search(r"Austria_|Netto|Haidach|7Fields|SEFE_Speicher(Ein|Aus)|RAG_Speicher(Ein|Aus)|EntrySpeicher(GESAMT)|ExitSpeicher(GESAMT)", c)]
print("MONTHLY FLOWS TWh (rows=month)\n" + m[pick].rename(columns=short).loc["2023-10":].to_string(), flush=True)
pk2 = [c for c in stock if re.search(r"Austria|Haidach|7Fields|SEFE_Speicher|RAG_|RAG-Sum|MGOst|OMV", c)]
print("MONTHLY STOCK CHANGE TWh (rows=month)\n" + chg[pk2].rename(columns=short).loc["2023-10":].to_string(), flush=True)
print("STOCK NAMES", stock)
sys.exit(0)
