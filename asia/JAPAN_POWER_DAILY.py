"""
Japan power generation by fuel, from the 10 regional transmission operators (TSOs) - no key, no login.

Each TSO (Hokkaido, Tohoku, Tokyo/TEPCO PG, Chubu, Hokuriku, Kansai, Chugoku, Shikoku, Kyushu, Okinawa) publishes
its own area supply-demand actuals as one CSV per month ("eria_jukyu_YYYYMM_NN.csv", NN = area number 01-10;
Shift-JIS or UTF-8, two header rows, 30-minute rows, values in MW averaged over the interval): area demand, nuclear,
thermal split LNG / coal / oil / other, hydro, geothermal, biomass, solar, wind, pumped storage, battery, interconnector.
The fuel split exists from roughly FY2023-24; earlier files only have total thermal, so this pull starts 2024-04-01.

Output (workbook, standard layout used by add_charts.power_daily and the JKT master):
  Daily    - date, Hydro/Gas/Wind/Solar/Coal/Nuclear/Oil/Bioenergy/Other _MWh, Demand_MWh, Pumped_net_MWh,
             Battery_net_MWh, Total_MWh (sum of the 10 areas; only days on which all 10 areas have data)
  By area  - date, area, same columns, one row per area and day (the store the Daily sheet is rebuilt from)
  Units    - units and notes

Incremental: By area is read back and only months from a week before the last stored day are fetched (a short
revision window). A day is stored only when it has all of its intervals (48, or 24 for hourly files).

    python3 asia/JAPAN_POWER_DAILY.py --out "output/Data and Chart Outputs/japan_power_generation_daily.xlsx"
"""
import argparse
import io
import os
import sys
import unicodedata
import zipfile
from datetime import date, timedelta

import pandas as pd
import requests

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))  # repo root, for xlsx_notes
import xlsx_notes  # noqa: E402

H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0 Safari/537.36"}
T = (15, 90)
DATA_START = date(2024, 4, 1)
REVISION_DAYS = 7

# area, name, monthly CSV url, optional per-day url (Tohoku's rolling file), optional (zip url, member), shift minutes
YM = "{ym}"
ZONES = [
    ("01", "Hokkaido", "https://www.hepco.co.jp/network/con_service/public_document/supply_demand_results/csv/eria_jukyu_{ym}_01.csv"),
    ("02", "Tohoku", "https://setsuden.nw.tohoku-epco.co.jp/common/demand/eria_jukyu_{ym}_02.csv"),
    ("03", "Tokyo", "https://www.tepco.co.jp/forecast/html/images/eria_jukyu_{ym}_03.csv"),
    ("04", "Chubu", "https://powergrid.chuden.co.jp/denki_yoho_content_data/eria_jukyu_{ym}_04.csv"),
    ("05", "Hokuriku", "https://www.rikuden.co.jp/nw/denki-yoho/csv/eria_jukyu_{ym}_05.csv"),
    ("06", "Kansai", "https://www.kansai-td.co.jp/interchange/denkiyoho/area-performance/eria_jukyu_{ym}_06.csv"),
    ("07", "Chugoku", "https://www.energia.co.jp/nw/jukyuu/sys/eria_jukyu_{ym}_07.csv"),
    ("08", "Shikoku", "https://www.yonden.co.jp/nw/supply_demand/csv/eria_jukyu_{ym}_08.csv"),
    ("09", "Kyushu", "https://www.kyuden.co.jp/td_area_jukyu/csv/eria_jukyu_{ym}_09.csv"),
    ("10", "Okinawa", "https://www.okiden.co.jp/business-support/service/supply-and-demand/csv/eria_jukyu_{ym}_10.csv"),
]
TOHOKU_DAILY = "https://setsuden.nw.tohoku-epco.co.jp/common/demand/realtime_jukyu/realtime_jukyu_{d}_02.csv"
CHUBU_ZIP = "https://powergrid.chuden.co.jp/denki_yoho_content_data/eria_jukyu_{fy}.zip"   # fiscal-year archive
SHIFT_MIN = {"Kyushu": -30}   # Kyushu labels intervals by their end, the others by their start

# source column (after NFKC normalisation) -> output column; thermal "other", geothermal and "other" go to Other
FUEL_COLS = {"原子力": "Nuclear", "火力(LNG)": "Gas", "火力(石炭)": "Coal", "火力(石油)": "Oil", "火力(その他)": "Other",
             "水力": "Hydro", "地熱": "Other", "バイオマス": "Bioenergy", "太陽光発電実績": "Solar",
             "風力発電実績": "Wind", "その他": "Other"}
SIGNED = {"揚水": "Pumped_net", "蓄電池": "Battery_net"}   # kept signed (negative = charging / pumping)
OUT_FUELS = ["Hydro", "Gas", "Wind", "Solar", "Coal", "Nuclear", "Oil", "Bioenergy", "Other"]
VALUE_COLS = [f"{f}_MWh" for f in OUT_FUELS] + ["Demand_MWh", "Pumped_net_MWh", "Battery_net_MWh"]
DAILY_SHEET, AREA_SHEET = "Daily", "By area"


def norm(s):
    return unicodedata.normalize("NFKC", str(s)).strip()


def decode(raw):
    for enc in ("utf-8-sig", "cp932"):
        try:
            return raw.decode(enc)
        except UnicodeDecodeError:
            continue
    return raw.decode("utf-8", errors="replace")


def parse_csv(raw, shift_min=0):
    """One TSO monthly (or daily) CSV -> frame indexed by interval start with MW columns, or None."""
    lines = decode(raw).splitlines()
    head = next((i for i, ln in enumerate(lines[:8]) if "DATE" in norm(ln).upper() and "TIME" in norm(ln).upper()), None)
    if head is None:
        return None
    df = pd.read_csv(io.StringIO("\n".join(lines[head:])), dtype=str)
    df.columns = [norm(c) for c in df.columns]
    if "DATE" not in df.columns or "TIME" not in df.columns:
        return None
    ts = pd.to_datetime(df["DATE"].map(norm) + " " + df["TIME"].map(norm), errors="coerce", format="mixed")
    out = pd.DataFrame(index=ts)
    for col in df.columns:
        if col in FUEL_COLS or col in SIGNED:
            out[col] = pd.to_numeric(df[col].map(lambda v: norm(v).replace(",", "")), errors="coerce").values
        elif "需要" in col and "Demand" not in out:
            out["Demand"] = pd.to_numeric(df[col].map(lambda v: norm(v).replace(",", "")), errors="coerce").values
    out = out[out.index.notna()]
    out.index = out.index + pd.Timedelta(minutes=shift_min)
    return out if len(out) else None


def to_daily(mw):
    """30-minute (or hourly) MW frame -> daily MWh in the output columns; incomplete days dropped."""
    if mw is None or mw.empty:
        return pd.DataFrame(columns=VALUE_COLS)
    step_h = mw.index.to_series().diff().dropna().median() / pd.Timedelta(hours=1)
    if not 0.2 < step_h < 1.1:
        return pd.DataFrame(columns=VALUE_COLS)
    mw = mw[~mw.index.duplicated(keep="last")].sort_index()
    day = mw.index.normalize()
    expected = round(24 / step_h)
    e = pd.DataFrame(index=mw.index)
    for f in OUT_FUELS:
        srcs = [c for c, t in FUEL_COLS.items() if t == f and c in mw]
        e[f"{f}_MWh"] = mw[srcs].clip(lower=0).sum(axis=1, min_count=1) * step_h if srcs else float("nan")
    e["Demand_MWh"] = mw["Demand"] * step_h if "Demand" in mw else float("nan")
    for c, t in SIGNED.items():
        e[f"{t}_MWh"] = mw[c] * step_h if c in mw else float("nan")
    d = e.groupby(day).sum(min_count=1)
    d = d[mw.groupby(day).size().reindex(d.index) >= expected]
    d.index.name = "date"
    return d[VALUE_COLS].round(1)


def months(start, end):
    d = start.replace(day=1)
    while d <= end:
        yield d
        d = (d.replace(day=28) + timedelta(days=4)).replace(day=1)


def get(session, url):
    try:
        r = session.get(url, headers=H, timeout=T)
    except requests.RequestException as e:
        return None, f"{type(e).__name__}"
    return (r.content, "200") if r.status_code == 200 and len(r.content) > 200 else (None, str(r.status_code))


def fetch_month(session, zone, month):
    area, name, url = zone
    shift = SHIFT_MIN.get(name, 0)
    raw, status = get(session, url.format(ym=f"{month:%Y%m}"))
    if raw is not None:
        parsed = parse_csv(raw, shift)
        if parsed is None:
            print(f"  {name} {month:%Y-%m}: http 200 but not parseable; first lines: {decode(raw)[:300]!r}")
        return parsed, status
    if name == "Chubu":
        fy = month.year if month.month >= 4 else month.year - 1
        z, zs = get(session, CHUBU_ZIP.format(fy=fy))
        if z is not None:
            try:
                with zipfile.ZipFile(io.BytesIO(z)) as zf:
                    member = f"eria_jukyu_{month:%Y%m}_04.csv"
                    if member in zf.namelist():
                        return parse_csv(zf.read(member), shift), "zip"
            except zipfile.BadZipFile:
                pass
    if name == "Tohoku":
        frames, d = [], month
        while d.month == month.month and d <= date.today():
            raw, _ = get(session, TOHOKU_DAILY.format(d=f"{d:%Y%m%d}"))
            if raw is not None:
                f = parse_csv(raw, shift)
                if f is not None:
                    frames.append(f)
            d += timedelta(days=1)
        if frames:
            return pd.concat(frames), "daily files"
    return None, status


def load_store(path):
    if not os.path.exists(path):
        return pd.DataFrame()
    try:
        d = pd.read_excel(path, sheet_name=AREA_SHEET)
        d["date"] = pd.to_datetime(d["date"])
        return d
    except Exception as e:  # noqa: BLE001
        print(f"could not read stored {AREA_SHEET} sheet ({type(e).__name__}); rebuilding from {DATA_START}")
        return pd.DataFrame()


def national(store):
    ok = store.groupby("date")["area"].nunique()
    full = ok[ok >= len(ZONES)].index
    d = store[store["date"].isin(full)].groupby("date")[VALUE_COLS].sum(min_count=1)
    d["Total_MWh"] = d[[f"{f}_MWh" for f in OUT_FUELS]].sum(axis=1, min_count=1)
    return d.round(1)


NOTES = [
    "UNITS",
    "MWh per day, summed from 30-minute (or hourly) MW-average rows. Fuel columns exclude pumped storage, batteries and",
    "interconnector flows; Pumped_net_MWh and Battery_net_MWh are signed (negative = pumping / charging). Demand_MWh is",
    "the TSO's area demand. Other = thermal 'other' + geothermal + 'other' as labelled by each TSO.",
    "",
    "SOURCE",
    "The 10 regional TSOs' own area supply-demand CSVs (eria_jukyu_YYYYMM_NN.csv): Hokkaido Electric Power Network, Tohoku",
    "Electric Power Network, TEPCO Power Grid, Chubu Electric Power Grid, Hokuriku Electric Power T&D, Kansai Transmission",
    "& Distribution, Energia Communications / Chugoku, Shikoku Electric Power T&D, Kyushu Electric Power T&D, Okinawa",
    "Electric Power. Solar includes rooftop (as reported by each TSO).",
    "",
    "COVERAGE",
    f"From {DATA_START:%Y-%m-%d} (earlier files only report total thermal). 'Daily' counts only days on which all 10 areas",
    "have a complete set of intervals; 'By area' keeps every complete area-day.",
]


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", default="output/Data and Chart Outputs/japan_power_generation_daily.xlsx")
    ap.add_argument("--from-date", default=DATA_START.isoformat())
    args = ap.parse_args()
    start = date.fromisoformat(args.from_date)

    store = load_store(args.out)
    session = requests.Session()
    new, report = [], []
    for zone in ZONES:
        # incremental per area: a week before that area's last stored day (a new area starts at DATA_START)
        zstart = start
        if not store.empty and (store["area"] == zone[1]).any():
            zlast = store.loc[store["area"] == zone[1], "date"].max().date()
            zstart = max(start, zlast - timedelta(days=REVISION_DAYS))
            store = store[~((store["area"] == zone[1]) & (store["date"] >= pd.Timestamp(zstart)))]
        got, last_status = 0, ""
        for m in months(zstart, date.today()):
            mw, last_status = fetch_month(session, zone, m)
            d = to_daily(mw)
            if mw is not None and not len(d):
                step = mw.index.to_series().diff().dropna().median()
                print(f"  {zone[1]} {m:%Y-%m}: parsed {len(mw)} rows ({mw.index.min()} to {mw.index.max()}, step {step}) "
                      f"but no complete day; columns {list(mw.columns)[:6]}")
            if len(d):
                d = d[d.index >= pd.Timestamp(zstart)]
            if len(d):
                d = d.reset_index()
                d.insert(1, "area", zone[1])
                new.append(d)
                got += len(d)
        report.append(f"{zone[1]}: from {zstart}, {got} area-days fetched (last http {last_status})")
        print(report[-1])
    if new:
        store = pd.concat([store] + new, ignore_index=True)
    if store.empty:
        sys.exit("no data fetched from any TSO")
    store = store.drop_duplicates(["date", "area"], keep="last").sort_values(["date", "area"])
    daily = national(store)
    areas = store.copy()
    areas["date"] = areas["date"].dt.strftime("%Y-%m-%d")
    daily.index = daily.index.strftime("%Y-%m-%d")
    os.makedirs(os.path.dirname(os.path.abspath(args.out)), exist_ok=True)
    xlsx_notes.write_workbook(args.out, {DAILY_SHEET: daily, AREA_SHEET: areas.set_index("date")},
                              NOTES + ["", "LAST RUN"] + report, {"UNITS", "SOURCE", "COVERAGE", "LAST RUN"})
    print(f"Saved {args.out}: {len(daily)} national days, {len(areas)} area-days")


if __name__ == "__main__":
    main()
