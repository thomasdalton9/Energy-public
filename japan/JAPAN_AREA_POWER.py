"""
Pull Japan's power generation by source from the general transmission and distribution operators' own "area supply and
demand results" files (エリア需給実績データ) and keep it as a growing archive
(output/Data and Chart Outputs/japan_power_generation_daily.xlsx).

Source: the nine operators of the OCCTO-linked grid publish one CSV per month and area in the same format (30-minute
average MW: area demand, nuclear, thermal split LNG / coal / oil / other, hydro, geothermal, biomass, solar and wind
actuals with solar and wind curtailment, pumped storage, batteries, inter-area lines, other): Hokkaido EPCO NW, Tohoku EPCO NW,
TEPCO PG, Chubu EPCO PG, Hokuriku EPT&D, KEPCO T&D, Energia (Chugoku), YONDEN T&D, Kyushu EPT&D. URL patterns are in AREAS.
Okinawa (Okinawa Electric Power, an island grid of about 1 GW) is not reachable from GitHub Actions and is left out. All
files reachable from Actions (probe 10 Oct 2026, discovery_archive/japan/).

Sheet 'Areas': one row per (date, area), MWh per day = sum of the day's 30-minute MW averages x 0.5 h, per source.
A blank column in an operator's file (a source the area has none of) stays blank. Days with fewer than 48 half-hours
(the file's first or last partial day, a missing month) are not rows.
Sheet 'Daily': the nine areas summed, only for days when every area has a row. Hydro, Gas (LNG), Wind, Solar, Coal,
Nuclear, Oil, Bioenergy (biomass), Geothermal, Other (other thermal + 'other'), pumped storage and battery net output,
area demand, solar and wind curtailment (MWh the operators report as output-controlled), Total_MWh = the sources
above (pumped storage, batteries and inter-area flows are not generation and are left out). Inter-area lines net to
about zero in the sum.
Sheet 'Coverage': per area and month: rows read, or the HTTP status when the file is not there.

Incremental: the committed workbook is the history store; months already stored (all days present) are not
fetched again, except the last two months of each area (operators revise). Months a file is not published for
(HTTP 404) are recorded on 'Coverage' and tried again only when they are within the last 3 months.

    python3 japan/JAPAN_AREA_POWER.py --out "output/Data and Chart Outputs/japan_power_generation_daily.xlsx"
"""
import argparse
import io
import os
import sys
import time
import unicodedata

import pandas as pd
import requests

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))  # repo root, for xlsx_notes
import xlsx_notes  # noqa: E402

UA = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36"}
START = pd.Timestamp("2024-04-01")   # the operators' monthly eria_jukyu files start Oct 2023 - Apr 2024 (older data is in other formats)
DEFAULT_OUT = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "output",
                           "Data and Chart Outputs", "japan_power_generation_daily.xlsx")

# area -> URL template ({ym} = YYYYMM)
AREAS = {
    "Hokkaido": "https://www.hepco.co.jp/network/con_service/public_document/supply_demand_results/csv/eria_jukyu_{ym}_01.csv",
    "Tohoku": "https://setsuden.nw.tohoku-epco.co.jp/common/demand/eria_jukyu_{ym}_02.csv",
    "Tokyo": "https://www.tepco.co.jp/forecast/html/images/eria_jukyu_{ym}_03.csv",
    "Chubu": "listing:chubu",
    "Hokuriku": "https://www.rikuden.co.jp/nw/denki-yoho/csv/eria_jukyu_{ym}_05.csv",
    "Kansai": "listing:kansai",
    "Chugoku": "https://www.energia.co.jp/nw/jukyuu/sys/eria_jukyu_{ym}_07.csv",
    "Shikoku": "https://www.yonden.co.jp/nw/supply_demand/csv/eria_jukyu_{ym}_08.csv",
    "Kyushu": "https://www.kyuden.co.jp/td_area_jukyu/csv/eria_jukyu_{ym}_09.csv",
}
# source column (NFKC-normalised Japanese header) -> workbook column
COLS = {
    "エリア需要": "Demand_MWh", "原子力": "Nuclear_MWh", "火力(LNG)": "Gas_MWh", "火力(石炭)": "Coal_MWh",
    "火力(石油)": "Oil_MWh", "火力(その他)": "Other thermal_MWh", "水力": "Hydro_MWh", "地熱": "Geothermal_MWh",
    "バイオマス": "Bioenergy_MWh", "太陽光発電実績": "Solar_MWh", "太陽光出力制御量": "Solar_Curtailed_MWh",
    "風力発電実績": "Wind_MWh", "風力出力制御量": "Wind_Curtailed_MWh", "揚水": "Pumped_MWh", "蓄電池": "Battery_MWh",
    "連系線": "Interconnection_MWh", "その他": "Other misc_MWh",
}
GEN = ["Hydro_MWh", "Gas_MWh", "Wind_MWh", "Solar_MWh", "Coal_MWh", "Nuclear_MWh", "Oil_MWh", "Bioenergy_MWh",
       "Geothermal_MWh", "Other_MWh"]
LAST_MONTHS_REFETCH = 2
RETRY_404_MONTHS = 3


def listing_urls(kind, sess):
    """{YYYYMM: (url, zip member or None)} for the two operators whose files are listed by a script-built page.
    Chubu: getFilesInfo.php (JSON) lists one zip per year (eria_jukyu_YYYY.zip, holding the monthly eria_jukyu_YYYYMM_04.csv)
    and the two most recent months as plain CSVs; its monthly *_keito.zip packages hold only total demand and generation, not
    by source. Kansai: /interchange/denkiyoho/area-performance/filelist.json lists eria_jukyu_YYYYMM_06.csv."""
    import re
    found = {}
    if kind == "chubu":
        base = "https://powergrid.chuden.co.jp"
        r = sess.get(base + "/denkiyoho/resource/php/getFilesInfo.php", headers=UA, timeout=(10, 60))
        r.raise_for_status()
        entries = r.json()
        for e in entries:   # one zip per calendar year holding the monthly eria_jukyu_YYYYMM_04.csv files
            m = re.match(r"^eria_jukyu_(\d{4})\.zip$", str(e.get("filename") or ""))
            if m and e.get("path"):
                for mm in range(1, 13):
                    found[f"{m.group(1)}{mm:02d}"] = (base + e["path"], f"eria_jukyu_{m.group(1)}{mm:02d}_04")
        for e in entries:   # the two most recent months as plain CSVs
            m = re.match(r"^eria_jukyu_(\d{6})_04\.csv$", str(e.get("filename") or ""))
            if m and e.get("path"):
                found[m.group(1)] = (base + str(e["path"]), None)
    else:
        base = "https://www.kansai-td.co.jp/interchange/denkiyoho/area-performance/"
        r = sess.get(base + "filelist.json", headers=UA, timeout=(10, 60))
        r.raise_for_status()
        for m in re.finditer(r"eria_jukyu_(\d{6})_06\.csv", r.text):
            found[m.group(1)] = (base + m.group(0), None)
    return found


def unzip_text(content, member=None):
    import zipfile
    z = zipfile.ZipFile(io.BytesIO(content))
    names = [n for n in z.namelist() if n.lower().endswith(".csv")]
    pick = [n for n in names if member and member in n] if member else ([n for n in names if "eria_jukyu" in n.lower()] or names)
    if not pick:
        raise ValueError("no CSV in zip: " + ", ".join(z.namelist()[:6]))
    b = z.read(pick[0])
    for enc in ("cp932", "utf-8-sig"):
        try:
            return b.decode(enc)
        except UnicodeDecodeError:
            pass
    raise ValueError("undecodable")


def parse(text):
    """One operator CSV -> daily MWh per source (complete days only)."""
    lines = text.splitlines()
    hdr = next(i for i, l in enumerate(lines[:6]) if "DATE" in l and "TIME" in l)
    d = pd.read_csv(io.StringIO("\n".join(lines[hdr:])), dtype=str)
    d = d.dropna(subset=["DATE", "TIME"])
    d.columns = [unicodedata.normalize("NFKC", str(c)).strip() for c in d.columns]
    d["t"] = pd.to_datetime(d["DATE"].str.strip().str.replace("/", "-"), errors="coerce", format="mixed") \
        + pd.to_timedelta(d["TIME"].str.strip().apply(lambda s: s + ":00" if s.count(":") == 1 else s), errors="coerce")
    d = d.dropna(subset=["t"]).sort_values("t")
    if len(d) and d["t"].iloc[0].strftime("%H:%M") == "00:30":   # interval-end labels: 00:30 is the first half-hour
        d["t"] = d["t"] - pd.Timedelta(minutes=30)
    d["date"] = d["t"].dt.normalize()
    out = {}
    for src, col in COLS.items():
        if src in d.columns:
            out[col] = pd.to_numeric(d[src].str.replace(",", ""), errors="coerce")
    v = pd.DataFrame(out, index=d.index)
    v["date"] = d["date"]
    n = v.groupby("date").size()
    day = v.groupby("date").sum(min_count=1) * 0.5   # MW average over 30 minutes -> MWh
    return day.loc[n[n == 48].index]


_ZIPS = {}


def fetch(url, sess, member=None):
    last = None
    for i in range(3):
        try:
            if url.lower().endswith(".zip") and url in _ZIPS:
                return unzip_text(_ZIPS[url], member), 200
            r = sess.get(url, headers=UA, timeout=(10, 120))
            if r.status_code == 404:
                return None, 404
            r.raise_for_status()
            if url.lower().endswith(".zip"):
                _ZIPS[url] = r.content
                return unzip_text(r.content, member), 200
            for enc in ("cp932", "utf-8-sig"):
                try:
                    return r.content.decode(enc), 200
                except UnicodeDecodeError:
                    pass
            raise ValueError("undecodable")
        except ValueError as e:   # member missing from the zip: not worth retrying
            return None, str(e)[:60]
        except Exception as e:  # noqa: BLE001
            last = e
            time.sleep(3 * (i + 1))
    return None, f"{type(last).__name__}"


def main():
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--out", default=DEFAULT_OUT)
    ap.add_argument("--start", default=None)
    args = ap.parse_args()
    start = pd.Timestamp(args.start) if args.start else START
    old_areas = old_cov = None
    if os.path.exists(args.out):
        try:
            old_areas = pd.read_excel(args.out, sheet_name="Areas", parse_dates=["date"])
            old_cov = pd.read_excel(args.out, sheet_name="Coverage")
        except Exception as e:  # noqa: BLE001
            print(f"  stored workbook unreadable ({type(e).__name__}); starting again", file=sys.stderr)
            old_areas = old_cov = None
    today = pd.Timestamp.today().normalize()
    months = pd.date_range(start.to_period("M").to_timestamp(), today.to_period("M").to_timestamp(), freq="MS")
    have = {}   # (area, ym) -> days stored
    if old_areas is not None:
        c = old_areas.assign(ym=old_areas["date"].dt.strftime("%Y%m")).groupby(["area", "ym"]).size()
        have = c.to_dict()
    missing_log = {}
    if old_cov is not None:
        missing_log = {(r.area, str(r.month)): r.status for r in old_cov.itertuples() if str(r.status) != "ok"}
    sess = requests.Session()
    new_rows, cov = [], {}
    n_req = 0
    for area, tpl in AREAS.items():
        listed = None
        if tpl.startswith("listing:"):
            try:
                listed = listing_urls(tpl.split(":")[1], sess)
                print(f"  {area}: {len(listed)} monthly files listed", (min(listed), max(listed)) if listed else "")
            except Exception as e:  # noqa: BLE001
                print(f"  {area}: file list failed ({type(e).__name__}: {e})", file=sys.stderr)
                continue
        for k, m in enumerate(months):
            ym = m.strftime("%Y%m")
            recent = (today.to_period("M") - m.to_period("M")).n < LAST_MONTHS_REFETCH
            if not recent and have.get((area, ym), 0) >= m.days_in_month - 1:
                continue
            if (area, ym) in missing_log and (today.to_period("M") - m.to_period("M")).n >= RETRY_404_MONTHS:
                cov[(area, ym)] = missing_log[(area, ym)]
                continue
            if listed is not None and ym not in listed:
                cov[(area, ym)] = "not listed"
                continue
            if listed is not None:
                text, status = fetch(listed[ym][0], sess, listed[ym][1])
            else:
                text, status = fetch(tpl.format(ym=ym), sess)
            n_req += 1
            if text is None:
                cov[(area, ym)] = str(status)
                continue
            try:
                day = parse(text)
            except Exception as e:  # noqa: BLE001
                cov[(area, ym)] = f"parse {type(e).__name__}: {e}"[:80]
                print(f"  {area} {ym}: parse failed {type(e).__name__}: {e}", file=sys.stderr)
                continue
            day = day[(day.index >= m) & (day.index < m + pd.offsets.MonthBegin(1))]
            day.insert(0, "area", area)
            new_rows.append(day.reset_index())
            cov[(area, ym)] = "ok"
            time.sleep(0.2)
    print(f"  {n_req} files requested")
    new = pd.concat(new_rows) if new_rows else pd.DataFrame()
    if old_areas is not None:
        if len(new):
            key = new[["area", "date"]].drop_duplicates()
            old_keep = old_areas.merge(key, on=["area", "date"], how="left", indicator=True)
            old_keep = old_keep[old_keep["_merge"] == "left_only"].drop(columns="_merge")
            areas = pd.concat([old_keep, new])
        else:
            areas = old_areas
    else:
        areas = new
    if areas.empty:
        raise SystemExit("no data fetched")
    areas = areas.sort_values(["date", "area"]).reset_index(drop=True)
    # coverage sheet: previous + this run's
    cov_rows = {}
    if old_cov is not None:
        for r in old_cov.itertuples():
            cov_rows[(r.area, str(r.month))] = r.status
    cov_rows.update(cov)
    for (area, ym), n in areas.assign(ym=areas["date"].dt.strftime("%Y%m")).groupby(["area", "ym"]).size().items():
        cov_rows[(area, ym)] = "ok"
    coverage = pd.DataFrame([(a, m, s) for (a, m), s in sorted(cov_rows.items())], columns=["area", "month", "status"])
    # national daily
    need = list(AREAS)
    cnt = areas.groupby("date")["area"].nunique()
    full = cnt[cnt == len(need)].index
    a = areas[areas["date"].isin(full)].drop(columns="area").groupby("date").sum(min_count=1)
    daily = pd.DataFrame(index=a.index)
    other = a.get("Other thermal_MWh", 0).fillna(0) + a.get("Other misc_MWh", 0).fillna(0)
    for col in ("Hydro_MWh", "Gas_MWh", "Wind_MWh", "Solar_MWh", "Coal_MWh", "Nuclear_MWh", "Oil_MWh", "Bioenergy_MWh",
                "Geothermal_MWh"):
        daily[col] = a.get(col)
    daily["Other_MWh"] = other
    daily["Total_MWh"] = daily[GEN].sum(axis=1, min_count=1)
    for col in ("Pumped_MWh", "Battery_MWh", "Demand_MWh", "Solar_Curtailed_MWh", "Wind_Curtailed_MWh"):
        daily[col] = a.get(col)
    daily.index.name = "date"
    daily = daily.round(1)
    areas_out = areas.set_index("date").round(1)
    os.makedirs(os.path.dirname(os.path.abspath(args.out)), exist_ok=True)
    xlsx_notes.write_workbook(args.out, {"Daily": daily, "Areas": areas_out, "Coverage": coverage.set_index("area")},
                              NOTES, TITLES)
    last = daily.index.max()
    avg = daily["Total_MWh"].tail(30).mean() / 24 / 1000
    dem = daily["Demand_MWh"].tail(30).mean() / 24 / 1000
    print(f"Saved Daily {len(daily)} days ({daily.index.min():%Y-%m-%d}..{last:%Y-%m-%d}); last-30-day mean generation "
          f"{avg:.1f} GW vs demand {dem:.1f} GW; areas rows {len(areas)} -> {args.out}")
    bad = coverage[coverage["status"] != "ok"]
    if len(bad):
        print("  months not available:", bad.groupby("area").size().to_dict())


NOTES = [
    "UNITS",
    "MWh per day = sum of the 30-minute average MW values x 0.5 h. Daily sheet: Japan's nine OCCTO-linked areas summed "
    "(Okinawa not included), days with all nine areas only. Gas_MWh is the operators' 'thermal (LNG)' line; Other_MWh is "
    "'thermal (other)' plus 'other'; Pumped_MWh and Battery_MWh are net output (negative = charging) and, with inter-area "
    "flows, are not part of Total_MWh. Solar_Curtailed_MWh / Wind_Curtailed_MWh are the output-control volumes the "
    "operators report.",
    "",
    "SOURCE",
    "Each general transmission and distribution operator's 'area supply and demand results' monthly CSV (Hokkaido EPCO "
    "Network, Tohoku EPCO Network, TEPCO Power Grid, Chubu EPCO Power Grid, Hokuriku Electric Power Transmission & "
    "Distribution, Kansai Transmission and Distribution, Energia Chugoku, Shikoku T&D, Kyushu T&D). URL patterns in "
    "japan/JAPAN_AREA_POWER.py (AREAS).",
    "",
    "UPDATES",
    "Incremental: months already stored are not fetched again except the last two; operators may revise earlier values.",
]
TITLES = {"UNITS", "SOURCE", "UPDATES"}

if __name__ == "__main__":
    main()
