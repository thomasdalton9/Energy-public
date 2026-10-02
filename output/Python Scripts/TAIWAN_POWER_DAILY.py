"""
Taiwan power generation by fuel, from Taipower (Taiwan's grid operator and utility) open data - no key, no login.

Source: Taipower open-data file d006010, "net generation by unit" (10-minute MW per unit with its fuel type; one JSON
of about 200 MB covering the latest published window - currently a quarter):
  https://service.taipower.com.tw/data/opendata/apply/file/d006010/001.json
  records.NET_P[]: FUEL_TYPE (Chinese fuel label), UNIT_NAME, DATETIME (Taiwan time), NET_P (MW)
The file is downloaded only when its Last-Modified changes (recorded on the Release sheet); it is streamed, summed to
days (MW x 10/60 = MWh) and merged into the stored workbook, so history accumulates window by window. A day is stored
only with all 144 ten-minute slots.

Output (standard layout): Daily - date, Hydro/Gas/Wind/Solar/Coal/Nuclear/Oil/Bioenergy/Other _MWh, Pumped_net_MWh,
Battery_net_MWh, Total_MWh; Release - source Last-Modified and window; Fuel types - every Taipower fuel label seen and
the column it maps to.

    python3 asia/TAIWAN_POWER_DAILY.py --out "output/Data and Chart Outputs/taiwan_power_generation_daily.xlsx"
"""
import argparse
import os
import sys
import tempfile
from collections import defaultdict

import pandas as pd
import requests

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))  # repo root, for xlsx_notes
import xlsx_notes  # noqa: E402

URL = "https://service.taipower.com.tw/data/opendata/apply/file/d006010/001.json"
H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0 Safari/537.36"}
SLOTS = 144
OUT_FUELS = ["Hydro", "Gas", "Wind", "Solar", "Coal", "Nuclear", "Oil", "Bioenergy", "Other"]
COLS = [f"{f}_MWh" for f in OUT_FUELS] + ["Pumped_net_MWh", "Battery_net_MWh"]
# substring of Taipower's fuel label -> output column (first match wins; order matters)
FUEL_MAP = [("抽蓄", "Pumped_net"), ("儲能", "Battery_net"), ("核", "Nuclear"), ("燃氣", "Gas"), ("燃煤", "Coal"),
            ("燃油", "Oil"), ("柴油", "Oil"), ("重油", "Oil"), ("水力", "Hydro"), ("風力", "Wind"), ("太陽", "Solar"),
            ("生質", "Bioenergy"), ("汽電", "Other"), ("其它", "Other"), ("其他", "Other"), ("地熱", "Other"),
            ("垃圾", "Other")]


def column_for(label):
    for key, col in FUEL_MAP:
        if key in label:
            return col + ("_MWh" if not col.endswith("_net") else "_MWh")
    return "Other_MWh"


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", default="output/Data and Chart Outputs/taiwan_power_generation_daily.xlsx")
    ap.add_argument("--force", action="store_true", help="download even if Last-Modified is unchanged")
    args = ap.parse_args()

    old, release = pd.DataFrame(), {}
    if os.path.exists(args.out):
        try:
            old = pd.read_excel(args.out, sheet_name="Daily")
            old["date"] = pd.to_datetime(old["date"])
            old = old.set_index("date")
            rel = pd.read_excel(args.out, sheet_name="Release")
            release = dict(zip(rel["item"], rel["value"].astype(str)))
        except Exception as e:  # noqa: BLE001
            print(f"could not read stored workbook ({type(e).__name__}); rebuilding")
    head = requests.head(URL, headers=H, timeout=60, allow_redirects=True)
    lm = head.headers.get("Last-Modified", "") or head.headers.get("ETag", "")
    print(f"HEAD {head.status_code}, Last-Modified/ETag: {lm!r}; stored: {release.get('last_modified')!r}")
    if lm and lm == release.get("last_modified") and not old.empty and not args.force:
        print("source unchanged since the last pull - nothing to do")
        return

    tmp = tempfile.NamedTemporaryFile(suffix=".json", delete=False)
    try:
        with requests.get(URL, headers=H, timeout=(15, 600), stream=True) as r:
            r.raise_for_status()
            size = 0
            for chunk in r.iter_content(1 << 20):
                tmp.write(chunk)
                size += len(chunk)
        tmp.close()
        print(f"downloaded {size / 1e6:.0f} MB")
        import ijson  # noqa: PLC0415
        mwh = defaultdict(float)               # (date, column) -> MWh
        slots = defaultdict(set)               # date -> {timestamps}
        labels = defaultdict(int)
        first = last = None
        with open(tmp.name, "rb") as f:
            for rec in ijson.items(f, "records.NET_P.item", use_float=True):
                label = str(rec.get("FUEL_TYPE", ""))
                ts = str(rec.get("DATETIME", ""))
                try:
                    p = float(rec.get("NET_P"))
                except (TypeError, ValueError):
                    continue
                day = ts[:10]
                labels[label] += 1
                mwh[(day, column_for(label))] += p * 10 / 60
                slots[day].add(ts)
                first = ts if first is None or ts < first else first
                last = ts if last is None or ts > last else last
    finally:
        os.remove(tmp.name)
    print(f"window {first} to {last}; fuel labels seen: " + "; ".join(f"{k}={v}" for k, v in sorted(labels.items())))
    complete = {d for d, s in slots.items() if len(s) >= SLOTS}
    rows = defaultdict(dict)
    for (d, c), v in mwh.items():
        if d in complete:
            rows[d][c] = v
    new = pd.DataFrame.from_dict(rows, orient="index").reindex(columns=COLS)
    new.index = pd.to_datetime(new.index)
    new["Total_MWh"] = new[[f"{f}_MWh" for f in OUT_FUELS]].sum(axis=1, min_count=1)
    daily = pd.concat([old[~old.index.isin(new.index)] if len(old) else old, new]).sort_index().round(1)
    daily.index = daily.index.strftime("%Y-%m-%d")
    daily.index.name = "date"
    mapping = pd.DataFrame({"Taipower fuel label": list(labels), "rows": list(labels.values()),
                            "maps to": [column_for(k) for k in labels]}).set_index("Taipower fuel label")
    rel = pd.DataFrame({"item": ["last_modified", "window_first", "window_last"],
                        "value": [lm, first, last]}).set_index("item")
    lines = ["UNITS", "MWh per day = sum of 10-minute net MW x 10/60 over every unit, by fuel. Pumped_net_MWh and Battery_net_MWh",
             "are signed (negative = pumping / charging) and are not in Total_MWh. Other = cogeneration (汽電共生), other",
             "renewables and anything unmapped (see the 'Fuel types' sheet).", "",
             "SOURCE", "Taipower open data d006010 (net generation by unit, 10-minute), https://service.taipower.com.tw/",
             "", "COVERAGE", "From March 2026 (earliest window Taipower currently publishes); history accumulates in this workbook",
             "as new windows appear. Only days with all 144 slots are kept."]
    os.makedirs(os.path.dirname(os.path.abspath(args.out)), exist_ok=True)
    xlsx_notes.write_workbook(args.out, {"Daily": daily, "Release": rel, "Fuel types": mapping}, lines,
                              {"UNITS", "SOURCE", "COVERAGE"})
    print(f"Saved {args.out}: {len(daily)} days {daily.index.min()} to {daily.index.max()}")


if __name__ == "__main__":
    main()
