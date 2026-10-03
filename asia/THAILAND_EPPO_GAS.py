"""
Thailand natural gas from EPPO (Energy Policy and Planning Office, Ministry of Energy) energy statistics,
natural gas tables: https://www.eppo.go.th/data-energy-statistic/energy-statistic/gas-energy-stat/
Found via discovery_archive/asia/BD_TH_DISCOVERY*.py.

  Table 3.1-1M  Production and Import of Natural Gas, MMSCFD, monthly from 1986: each Gulf of Thailand field
                (Erawan, Bongkot, Bongkot Tai, Pailin, Arthit, Tantawan, Phu Horm, Sirikit, Lanta, Nam Phong,
                Jasmin, Yoong Thong, JDA, others), domestic total; imports by Myanmar pipeline (Yadana, Yetagun,
                Zawtika) and LNG; grand total. T03_01_01-1.xls = history, T03_01_01.xls = recent years (overrides).
  Table 3.2-2M  Consumption of Natural Gas by Sector, MMSCFD, monthly from 1986: electricity (EGAT, IPP, SPP),
                industry, GSP (gas separation plants), NGV. T03_02_02-1.xls / T03_02_02.xls.

Writes output/Data and Chart Outputs/thailand_gas.xlsx:
  Supply     monthly MMSCFD by field / import source, plus Domestic_total, Myanmar_pipeline, LNG, Imports_total, Total
  Demand     monthly MMSCFD: Power_EGAT, Power_IPP, Power_SPP, Power_total, Industry, GSP, NGV, Total

EPPO only re-publishes whole files (small): they are downloaded every run, merged over the saved sheets (new rows win), and the files' Last-Modified dates are
recorded on the Units sheet. Runs on the 1st and 15th.

    python3 asia/THAILAND_EPPO_GAS.py
"""
import argparse
import io
import os
import re
import sys

import pandas as pd
import requests

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)
import xlsx_notes  # noqa: E402

PAGES = ["https://www.eppo.go.th/data-energy-statistic/energy-statistic/gas-energy-stat/"
         "%E0%B8%81%E0%B9%8A%E0%B8%B2%E0%B8%8B%E0%B8%98%E0%B8%A3%E0%B8%A3%E0%B8%A1%E0%B8%8A%E0%B8%B2%E0%B8%95%E0%B8%B4-2/",
         "https://www.eppo.go.th/epposite/info/stat/natural-gas"]
FALLBACK = "https://www.eppo.go.th/wp-content/uploads/2026/04/{key}.xls"
H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0 Safari/537.36"}
T = (20, 120)
OUT = os.path.join(ROOT, "output", "Data and Chart Outputs", "thailand_gas.xlsx")
MONTHS = {m: i for i, m in enumerate(["JAN", "FEB", "MAR", "APR", "MAY", "JUN", "JUL", "AUG", "SEP", "OCT", "NOV",
                                      "DEC"], start=1)}
MYANMAR = ("Yadana", "Yetagun", "Zawtika")


def out(*a):
    print(*a, flush=True)


def table_links():
    links = {}
    for page in PAGES:
        try:
            r = requests.get(page, headers=H, timeout=T)
            if r.status_code != 200:
                continue
        except requests.RequestException as e:
            out(f"  {page}: {e}")
            continue
        for u, key in re.findall(r'href="([^"]+/(T03_0[12]_0[12](?:-1)?)\.xls)"', r.text):
            links.setdefault(key, u)
    for key in ("T03_01_01-1", "T03_01_01", "T03_02_02-1", "T03_02_02"):
        links.setdefault(key, FALLBACK.format(key=key))
    return links


def clean(name):
    name = re.sub(r"\s+", " ", str(name)).strip()
    fixes = {"Yetakun": "Yetagun", "Tan tawan": "Tantawan", "Phu horm": "Phu_Horm", "Oths.": "Other_fields",
             "Total Import": "Imports_total", "Grand Total": "Total", "Nam phong": "Nam_Phong",
             "Yoong thong": "Yoong_Thong", "Bongkot Tai": "Bongkot_Tai"}
    return fixes.get(name, name.replace(" ", "_"))


def parse(content, sub_key):
    """EPPO monthly table -> DataFrame (month start index). The sub-header row is the one holding sub_key
    ('Erawan' / 'EGAT'); the group row above it names the columns with no sub-header ('Grand Total', 'Industry')
    and the group 'Total' columns ('Domestic Production' + 'Total')."""
    df = pd.read_excel(io.BytesIO(content), header=None)
    h = next(i for i in range(len(df)) if any(str(v).strip() == sub_key for v in df.iloc[i]))
    group, sub = df.iloc[h - 1].ffill(), df.iloc[h]
    names = {}
    for j in range(1, df.shape[1]):
        s, g = sub.iloc[j], group.iloc[j]
        s = None if pd.isna(s) else str(s).strip()
        g = None if pd.isna(g) else str(g).strip()
        if s and s.lower() != "total":
            names[j] = clean(s)
        elif s:   # a group's total
            names[j] = {"domestic production": "Domestic_total", "electricity": "Power_total"}.get(
                (g or "").lower(), clean(f"{g} total"))
        elif g and str(df.iloc[h - 1, j]).strip() == g:   # its own group header, no sub-header
            names[j] = clean(g)
    rows, year = {}, None
    for i in range(len(df)):
        c0 = str(df.iat[i, 0]).strip()
        m = re.match(r"^(\d{4})(?:\.0)?\b", c0)
        if m and 1980 < int(m.group(1)) < 2100:
            year = int(m.group(1))
        mon = MONTHS.get(c0[:3].upper())
        if year and mon:
            rows.setdefault(pd.Timestamp(year, mon, 1), {k: pd.to_numeric(df.iat[i, j], errors="coerce")
                                                         for j, k in names.items()})   # first block wins
    return pd.DataFrame.from_dict(rows, orient="index").sort_index()


def fetch(links, keys, sub_key):
    frames, stamps = [], []
    for key in keys:   # history first, then the recent-years file overrides it
        try:
            r = requests.get(links[key], headers=H, timeout=T)
            r.raise_for_status()
            f = parse(r.content, sub_key)
        except Exception as e:  # noqa: BLE001
            out(f"  {key}: {type(e).__name__}: {e}")
            continue
        out(f"  {key}: {len(f)} months {f.index.min():%Y-%m}..{f.index.max():%Y-%m}" if len(f) else f"  {key}: empty")
        frames.append(f)
        stamps.append(f"{links[key].rsplit('/', 1)[-1]} (Last-Modified {r.headers.get('last-modified')})")
    if not frames:
        return pd.DataFrame(), stamps
    d = pd.concat(frames)
    return d[~d.index.duplicated(keep="last")].sort_index(), stamps


def keep_saved(path, sheet, new):
    """Merge a freshly parsed table over the saved sheet (new rows win), so a failed or truncated download never
    drops history or a whole sheet."""
    try:
        old = pd.read_excel(path, sheet_name=sheet, index_col=0)
        old.index = pd.to_datetime(old.index)
    except (FileNotFoundError, ValueError):
        return new
    if new is None or new.empty:
        return old
    return pd.concat([old[~old.index.isin(new.index)], new]).sort_index()


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", default=OUT)
    args = ap.parse_args()
    links = table_links()
    out(f"links: {links}")
    supply, s1 = fetch(links, ("T03_01_01-1", "T03_01_01"), "Erawan")
    demand, s2 = fetch(links, ("T03_02_02-1", "T03_02_02"), "EGAT")
    if supply.empty and demand.empty:
        raise SystemExit("No EPPO gas tables")
    if not supply.empty:
        supply["Myanmar_pipeline"] = supply[[c for c in MYANMAR if c in supply]].sum(axis=1, min_count=1)
        supply = supply[supply.get("Total", supply.sum(axis=1)) > 0].round(1)
        supply.index.name = "date"
        out(supply[[c for c in ("Domestic_total", "Myanmar_pipeline", "LNG", "Total") if c in supply]].tail(4).to_string())
    if not demand.empty:
        demand = demand.rename(columns={"EGAT": "Power_EGAT", "IPP": "Power_IPP", "SPP": "Power_SPP"})
        demand = demand[demand.get("Total", demand.sum(axis=1)) > 0].round(1)
        demand.index.name = "date"
        out(demand.tail(4).to_string())
    notes = [
        "UNITS",
        "MMSCFD = million standard cubic feet per day, monthly average (EPPO's unit; heat value 1,000 BTU/SCF). "
        "1 MMSCFD = 0.0283 million m3 per day (the charts show mcm/d).",
        "Supply: domestic production by field (Gulf of Thailand; JDA = Malaysia-Thailand Joint Development Area) and "
        "Domestic_total; imports: Yadana, Yetagun, Zawtika (Myanmar pipeline; Myanmar_pipeline = their sum), LNG, "
        "Imports_total; Total = grand total supply.",
        "Demand: Power_EGAT / Power_IPP / Power_SPP and Power_total (gas for electricity), Industry, GSP (gas "
        "separation plants), NGV (vehicles), Total.",
        "",
        "COVERAGE",
        (f"Supply monthly {supply.index.min():%Y-%m}..{supply.index.max():%Y-%m}; " if not supply.empty else "") +
        (f"demand monthly {demand.index.min():%Y-%m}..{demand.index.max():%Y-%m}. " if not demand.empty else "") +
        "EPPO updates about two months after the month.",
        "Files read this run: " + "; ".join(s1 + s2),
        "",
        "SOURCE",
        "EPPO (Energy Policy and Planning Office, Ministry of Energy, Thailand), natural gas statistics, tables 3.1-1 "
        "and 3.2-2 (data from PTT): https://www.eppo.go.th/data-energy-statistic/energy-statistic/gas-energy-stat/",
    ]
    supply, demand = keep_saved(args.out, "Supply", supply), keep_saved(args.out, "Demand", demand)
    for v in (supply, demand):
        v.index.name = "date"
    sheets = {k: v for k, v in (("Supply", supply), ("Demand", demand)) if not v.empty}
    xlsx_notes.write_workbook(args.out, sheets, notes, {"UNITS", "COVERAGE", "SOURCE"})
    out(f"Saved {args.out}")


if __name__ == "__main__":
    main()
