"""
Peru hydro reservoirs: useful volume (volumen util, hm3) of the SEIN's
seasonal reservoirs and lagoons by basin - Lake Junin, the Mantaro-basin
lagoons, the Rimac system, the Chili system (Arequipa), Aricota,
Sibinacocha and the rest - and the national total in hm3 and as % of
total useful capacity, weekly, from COES. No key needed.

Source: COES 'Informe Semanal de Evaluacion de la Operacion del SEIN'
(https://www.coes.org.pe/Portal/PostOperacion/Informes/EvaluacionSemanal),
one Excel workbook per COES operating week (Saturday-Friday), listed by the
portal's file browser under
    Post Operacion/Informes/Evaluacion Semanal/<year>/SEMANAL N° <w> (dd.mm.yyyy - dd.mm.yyyy)/
and downloaded through /Portal/browser/download?url=... . Since 2024 it has
  - 'EVOLUCION DE VOLUMENES DE LOS EMBALSES Y LAGUNAS' (5.2; 6.2 in early
    2024): the useful volume at the end of every week of this year and the
    three before, by basin (12 series: Junin, Mantaro sub-basins, Rimac,
    Chili, Locumba/Aricota, Vilcanota/Sibinacocha, Paucartambo, Santa,
    Pativilca/Viconga, San Gaban, Canete/Paucarcocha, Huallaga/Chaglla);
  - 'VOLUMEN UTIL DE LOS EMBALSES Y LAGUNAS' (5.1): every reservoir /
    lagoon row with its volume at the start and end of the week and its
    useful capacity.
Before 2024 the reports used another layout (company blocks, not all
reservoirs), so the series starts with the earliest year the 2024
reports look back to: 2021. Found via
discovery_archive/south_america/HYDRO_PE_EC_UY_DISCOVERY.py (rounds 3-4).
COES publishes no daily table covering all lagoons: its daily IDCOS / IEOD
hydrology annexes carry only a handful of seasonal reservoirs (Junin,
Sibinacocha, Aricota, Viconga) and the Mantaro lagoons only as discharges.
So this series is WEEKLY (end of each COES week, a Friday).

National total = sum of the 12 basin series. National % = total / the
total useful capacity in the newest report's 5.1 table (1,970 hm3 in
Sep-2026), applied to every week - COES' capacity figures have changed
between reports (e.g. Junin 376 -> 315 hm3 in 2024), so one fixed
denominator keeps the % comparable across years.

A week's value comes from the newest report that covers it (each report
restates four years). Each run lists the year folders and reads the newest
report of every year from 2024 that has not been read yet ('Reports read');
the newest report also refreshes the latest-week detail sheet.

Usage: python3 PERU_HYDRO_RESERVOIRS.py [--out PATH] [--start-year 2021]
"""

print("STARTING", flush=True)

import argparse
import html
import io
import os
import re
import sys
import time
import unicodedata
from datetime import date, timedelta
from urllib.parse import quote

import pandas as pd
import requests

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))  # repo root, for xlsx_notes

import xlsx_notes  # noqa: E402

PORTAL = "https://www.coes.org.pe/Portal/"
BASE = "Post Operación/Informes/Evaluacion Semanal/"
PAGE = PORTAL + "PostOperacion/Informes/EvaluacionSemanal"
HEADERS = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) "
                         "Chrome/124.0 Safari/537.36"}
TIMEOUT = (15, 240)
OUT = os.path.join("output", "Data and Chart Outputs", "peru_hydro_reservoirs.xlsx")
START_YEAR = 2021
NEW_LAYOUT_FROM = 2024   # first year whose weekly reports carry the 5.1 / 5.2 volume tables

# group -> keywords matched (accent-free, upper case) against COES' series / row labels. First match wins.
BASINS = [
    ("Junin", ["JUNIN"]),
    ("Rimac", ["RIMAC", "SHEQUE", "QUISHA", "SACSA", "ANTACOTO", "MARCAPOMACOCHA", "YURACMAYO"]),
    ("Mantaro_lagoons", ["MANTARO", "ELECTROPERU", "LAGUNAS STATKRAFT", "POMACOCHA", "VICHECOCHA", "CARHUACOCHA",
                         "CHILLICOCHA", "CHILICOCHA", "HUICHICOCHA", "COYLLORCOCHA", "YURAJCOCHA"]),
    ("Chili", ["CHILI", "EL PANE", "BAMPUTANE", "ESPANOLES", "CHALHUANCA", "PILLONES", "FRAYLE", "AGUADA BLANCA"]),
    ("Aricota", ["ARICOTA", "LOCUMBA"]),
    ("Sibinacocha", ["SIBINACOCHA", "VILCANOTA"]),
    ("Paucartambo", ["PAUCARTAMBO", "JAICO", "PACCHAPATA", "MACHAY", "MATACOCHA", "HUANGUSH"]),
    ("Santa", ["SANTA", "AGUASCOCHA", "RAJUCOLTA", "CULLICOCHA"]),
    ("Viconga", ["VICONGA", "PATIVILCA"]),
    ("San_Gaban", ["SAN GABAN", "AJOYAJOTA", "PARINAJOTA", "ISOCOCHA"]),
    ("Canete", ["PAUCARCOCHA", "CANETE"]),
    ("Chaglla", ["CHAGLLA", "HUALLAGA"]),
]
GROUPS = [g for g, _ in BASINS]
GROUP_LABEL = {"Junin": "Lake Junin (Chinchaycocha)",
               "Mantaro_lagoons": "Mantaro sub-basin lagoons (Electroperu, Statkraft, Carhuacocha, Chillicocha, "
                                  "Huichicocha, Vichecocha ...)",
               "Rimac": "Rimac system (Sheque, Sacsa, Antacoto/Marcapomacocha, Yuracmayo ...)",
               "Chili": "Chili system, Arequipa (El Frayle, Pillones, El Pane, Bamputane, Aguada Blanca, Chalhuanca, "
                        "Los Espanoles)",
               "Aricota": "Aricota (Locumba)", "Sibinacocha": "Sibinacocha (Vilcanota)",
               "Paucartambo": "Paucartambo (Jaico/Pacchapata/Altos Machay, Matacocha/Huangush)",
               "Santa": "Santa (Aguascocha, Rajucolta, Cullicocha)", "Viconga": "Viconga (Pativilca)",
               "San_Gaban": "San Gaban lagoons (Ajoyajota, Parinajota, Isococha ...)",
               "Canete": "Paucarcocha (Canete)", "Chaglla": "Chaglla (Huallaga)"}

S = requests.Session()
S.headers.update(HEADERS)


def key(s):
    s = unicodedata.normalize("NFKD", str(s)).encode("ascii", "ignore").decode().upper()
    return " ".join(s.split())


def group_of(label):
    k = key(label)
    for g, words in BASINS:
        if any(w in k for w in words):
            return g
    return None


def browse(path):
    for attempt in range(4):
        try:
            r = S.post(PORTAL + "browser/vistadatos", data={"baseDirectory": path, "url": path, "indicador": "",
                                                            "initialLink": "", "orderFolder": ""}, timeout=TIMEOUT)
            r.raise_for_status()
            break
        except requests.RequestException as e:
            if attempt == 3:
                raise
            print(f"    retry browse {path}: {type(e).__name__}", flush=True)
            time.sleep(5 * (attempt + 1))
    out = []
    for m in re.finditer(r"openBlob\('([^']+)',\s*'(\w)'", r.text):
        it = (html.unescape(m.group(1)), m.group(2))
        if it not in out:
            out.append(it)
    return out


def download(path):
    for attempt in range(4):
        try:
            r = S.get(PORTAL + "browser/download?url=" + quote(path), timeout=TIMEOUT)
            r.raise_for_status()
            return r.content
        except requests.RequestException as e:
            if attempt == 3:
                raise
            print(f"    retry download: {type(e).__name__}", flush=True)
            time.sleep(5 * (attempt + 1))


WEEK = re.compile(r"SEMANAL\s*N\D{0,3}(\d+)\s*\((\d\d)\.(\d\d)\.(\d{4})\s*-\s*(\d\d)\.(\d\d)\.(\d{4})\)", re.I)


def list_weeks(year):
    """[(folder, week number, start date, end date)] for one year, from the folder names."""
    out = []
    for p, k in browse(f"{BASE}{year}/"):
        m = WEEK.search(p)
        if k == "D" and m:
            g = list(map(int, m.groups()))
            out.append((p, g[0], date(g[3], g[2], g[1]), date(g[6], g[5], g[4])))
    return sorted(out, key=lambda w: w[1])


def week_end(calendar, year, week):
    """End date (Friday) of COES week `week` of `year`, from the folder names; weeks without a folder are
    counted on from the nearest listed week of that year."""
    weeks = calendar.get(year) or {}
    if week in weeks:
        return weeks[week]
    if not weeks:
        return None
    k = min(weeks, key=lambda w: abs(w - week))
    return weeks[k] + timedelta(days=7 * (week - k))


def sheet_with(xl, title):
    for sh in xl.sheet_names:
        df = pd.read_excel(xl, sheet_name=sh, header=None, nrows=15)
        if any(title in key(v) for v in df.values.ravel() if isinstance(v, str)):
            if key(sh) not in ("INDICE", "PORTADA"):
                return sh
    return None


def parse_evolution(xl, calendar):
    """'Evolucion de volumenes' sheet -> long frame: date, group, volume_hm3 (end-of-week, four years)."""
    sh = sheet_with(xl, "EVOLUCION DE VOLUMENES DE LOS EMBALSES")
    if sh is None:
        raise RuntimeError("no 'EVOLUCION DE VOLUMENES DE LOS EMBALSES Y LAGUNAS' sheet")
    df = pd.read_excel(xl, sheet_name=sh, header=None)
    rows = []
    seen = set()
    for i in range(min(30, len(df))):
        for j in range(df.shape[1]):
            v = df.iat[i, j]
            if not (isinstance(v, str) and key(v).startswith("VOLUMEN UTIL")):
                continue
            g = group_of(v)
            if g is None:
                print(f"    WARNING: series not matched to a group: {v!r}", flush=True)
                continue
            if g in seen:
                raise RuntimeError(f"two series map to {g}: {v!r}")
            seen.add(g)
            # the next row holds the years over the columns to the right of the week column j
            years = {}
            for c in range(j + 1, min(j + 6, df.shape[1])):
                y = pd.to_numeric(df.iat[i + 1, c], errors="coerce")
                if pd.notna(y) and 2000 < y < 2100:
                    years[c] = int(y)
            for r in range(i + 2, len(df)):
                wk = pd.to_numeric(df.iat[r, j], errors="coerce")
                if pd.isna(wk):
                    if r > i + 3:
                        break
                    continue
                for c, y in years.items():
                    vol = pd.to_numeric(df.iat[r, c], errors="coerce")
                    d = week_end(calendar, y, int(wk))
                    if pd.notna(vol) and d is not None:
                        rows.append({"date": pd.Timestamp(d), "group": g, "volume_hm3": float(vol), "year": y,
                                     "week": int(wk)})
    if len(seen) < 10:
        raise RuntimeError(f"only {len(seen)} basin series found in sheet {sh}")
    return pd.DataFrame(rows), sorted(seen)


def parse_table(xl):
    """'Volumen util de los embalses y lagunas' (5.1) -> per-reservoir rows of this week: reservoir, group,
    volume at start / end of week, capacity."""
    sh = sheet_with(xl, "UTIL DE LOS EMBALSES Y LAGUNAS")
    if sh is None:
        raise RuntimeError("no 'VOLUMEN UTIL DE LOS EMBALSES Y LAGUNAS' sheet")
    df = pd.read_excel(xl, sheet_name=sh, header=None, nrows=80)
    for i in range(len(df)):
        cells = {j: key(v) for j, v in df.iloc[i].items() if isinstance(v, str)}
        vol_cols = {}
        for j, v in cells.items():
            m = re.search(r"VOLUMEN UTIL\s*(\d{1,2})/(\d{1,2})/(\d{4})", v)
            if m:
                vol_cols[j] = date(int(m.group(3)), int(m.group(2)), int(m.group(1)))
        if len(vol_cols) >= 2:
            cap_cols = sorted(j for j, v in cells.items() if v.startswith("CAPACIDAD"))
            break
    else:
        raise RuntimeError(f"sheet {sh}: no 'VOLUMEN UTIL dd/mm/yyyy' header row")
    label_col = None
    for r in range(max(0, i - 3), i + 1):
        for j, v in df.iloc[r].items():
            if isinstance(v, str) and re.fullmatch(r"(LAGUNAS? ?/ ?EMBALSES?|EMBALSES? ?/ ?LAGUNAS?)", key(v)):
                label_col = j
    if label_col is None:
        label_col = min(vol_cols) - 1
    newest = max(vol_cols.values()).year
    own = sorted(j for j, d in vol_cols.items() if d.year == newest)   # start and end of this week
    cap_j = next((c for c in cap_cols if c > max(own)), None)
    rows = []
    for r in range(i + 1, len(df)):
        label = df.iat[r, label_col]
        if isinstance(label, str) and key(label).startswith(("CUADRO", "FUENTE", "NOTA")):
            break
        if not isinstance(label, str) or not label.strip():
            continue
        name = re.sub(r"^Volumen [UÚ]til de(l)? (Embalse|Laguna)s? ", "", label.strip(), flags=re.I).strip()
        row = {"reservoir": name, "group": group_of(name)}
        for n, j in zip(("start", "end"), own):
            row[f"volume_hm3_{vol_cols[j]:%Y-%m-%d}"] = pd.to_numeric(df.iat[r, j], errors="coerce")
        row["capacity_hm3"] = pd.to_numeric(df.iat[r, cap_j], errors="coerce") if cap_j is not None else None
        rows.append(row)
    if not rows:
        raise RuntimeError(f"sheet {sh}: header found but no reservoir rows")
    return pd.DataFrame(rows)


def load(path, sheet, **kw):
    try:
        return pd.read_excel(path, sheet_name=sheet, **kw)
    except (FileNotFoundError, ValueError, KeyError, OSError):
        return pd.DataFrame()


def build_daily(long, capacity):
    """Long (date, group, volume, report_end) -> one row per week-end date. Newest report wins."""
    long = long.sort_values(["date", "group", "report_end"]).drop_duplicates(["date", "group"], keep="last")
    wide = long.pivot(index="date", columns="group", values="volume_hm3").reindex(columns=GROUPS)
    out = pd.DataFrame(index=wide.index)
    complete = wide.notna().sum(axis=1) >= len(GROUPS) - 1   # tolerate one basin missing in a week
    out["Total_hm3"] = wide.sum(axis=1, min_count=1).where(complete).round(2)
    out["Total_pct"] = (100 * out["Total_hm3"] / capacity).round(2)
    out["Capacity_hm3"] = round(capacity, 2)
    out["Basins_reported"] = wide.notna().sum(axis=1)
    for g in GROUPS:
        out[f"{g}_hm3"] = wide[g].round(3)
    out.index.name = "date"
    return out.sort_index()


def notes(daily, reports, latest):
    last = daily.dropna(subset=["Total_hm3"]).iloc[-1]
    return [
        "UNITS",
        "hm3 = million m3 of USEFUL volume (volumen util: water above each reservoir's minimum operating level), "
        "at the END of each COES operating week (Saturday-Friday; the date is the Friday).",
        "Total_hm3 = sum of COES' 12 basin series (weeks with more than one basin missing are left blank). "
        f"Total_pct = Total_hm3 / Capacity_hm3, where Capacity_hm3 = {last['Capacity_hm3']:,.1f} hm3 is the sum of "
        "the useful capacities in the newest report's reservoir table (sheet 'Latest week by reservoir'), used for "
        "every week. COES' capacity figures have changed between reports (e.g. Junin 376 -> 315 hm3 in 2024; the "
        "San Gaban lagoons 376 -> 69 hm3), so a single denominator keeps the % comparable across years; a basin can "
        "exceed its current capacity in years when COES allowed more (Junin in early 2025).",
        "",
        "COVERAGE",
        f"{daily.index.min():%d-%b-%Y} to {daily.index.max():%d-%b-%Y}, weekly. Latest ({last.name:%d-%b-%Y}): "
        f"{last['Total_hm3']:,.0f} hm3 = {last['Total_pct']:.1f}% of {last['Capacity_hm3']:,.0f} hm3.",
        "COES publishes no daily table covering all reservoirs and lagoons (its daily IDCOS/IEOD annexes carry only "
        "a few, e.g. Junin, Sibinacocha, Aricota), so the series is weekly. History starts in 2021: the 2024 reports "
        "restate 2021-2024; reports before 2024 use another layout (company blocks that miss several reservoirs).",
        "",
        "BASINS (COES' own 12 series in 'Evolucion de volumenes de los embalses y lagunas'; COES' basin series and "
        "the sum of its reservoir rows in table 5.1 can differ by a few hm3, e.g. Rimac)",
        *[f"{g}_hm3: {GROUP_LABEL[g]}" for g in GROUPS],
        "",
        "SOURCE",
        f"COES, Informe Semanal de Evaluacion de la Operacion del SEIN ({PAGE}): sheet 'Evolucion de volumenes de los "
        "embalses y lagunas' (5.2; 6.2 in early 2024) for the weekly basin series, sheet 'Volumen util de los "
        "embalses y lagunas' (5.1) for the per-reservoir table and capacities. One Excel per week, via the portal's "
        "file browser (Post Operacion/Informes/Evaluacion Semanal/<year>/SEMANAL N° <w> (...)/).",
        f"Latest report read: {latest}. Reports read: {len(reports)} (sheet 'Reports read').",
        "Script: south_america/PERU_HYDRO_RESERVOIRS.py (scheduled by .github/workflows/peru_hydro_reservoirs.yml).",
        "",
        "METHOD",
        "Each report restates its year and the three before; a week's value comes from the newest report that "
        "covers it. Each run reads only reports not read before: the newest one of each year from 2024.",
    ]


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--out", default=OUT)
    ap.add_argument("--start-year", type=int, default=START_YEAR)
    args = ap.parse_args()
    os.makedirs(os.path.dirname(os.path.abspath(args.out)), exist_ok=True)

    long = load(args.out, "Weekly by basin (long)")
    if not long.empty:
        long["date"] = pd.to_datetime(long["date"])
        long["report_end"] = pd.to_datetime(long["report_end"])
    reports = load(args.out, "Reports read")
    latest_table = load(args.out, "Latest week by reservoir")
    done = set(reports["report"]) if not reports.empty else set()

    today = date.today()
    calendar, newest = {}, {}
    for y in range(args.start_year, today.year + 1):
        try:
            ws = list_weeks(y)
        except Exception as e:  # noqa: BLE001
            print(f"  {y}: listing FAILED {type(e).__name__}: {e}", flush=True)
            continue
        calendar[y] = {w: end for _, w, _, end in ws}
        if ws and y >= NEW_LAYOUT_FROM:
            newest[y] = ws[-1]
        print(f"  {y}: {len(ws)} weekly folders" + (f", newest {ws[-1][1]} ending {ws[-1][3]}" if ws else ""),
              flush=True)
    todo = [w for y, w in sorted(newest.items()) if w[0] not in done]
    print(f"{len(done)} reports already read; to read: {[f'{w[3]}' for w in todo]}", flush=True)
    new_rep, failed = [], 0
    for folder, wk, a, b in todo:
        try:
            files = [p for p, k in browse(folder) if k == "F" and p.lower().endswith((".xlsx", ".xlsm"))]
            if not files:
                raise RuntimeError("no Excel file in the folder")
            f = sorted(files, key=lambda p: ("SEMANAL" not in key(p), p))[0]
            xl = pd.ExcelFile(io.BytesIO(download(f)))
            evo, groups = parse_evolution(xl, calendar)
            evo["report_end"] = pd.Timestamp(b)
            long = pd.concat([long, evo], ignore_index=True) if not long.empty else evo
            if b >= max(w[3] for w in newest.values()):
                latest_table = parse_table(xl)
                latest_table.insert(0, "report_week_end", pd.Timestamp(b))
            new_rep.append({"report": folder, "file": f.split("/")[-1], "week": wk, "start": a, "end": b,
                            "basins": len(groups), "weeks_read": len(evo),
                            "first": evo["date"].min().date(), "last": evo["date"].max().date()})
            print(f"  report ending {b}: {len(groups)} basins, {len(evo)} basin-weeks "
                  f"{evo['date'].min():%d-%b-%Y}..{evo['date'].max():%d-%b-%Y}", flush=True)
        except Exception as e:  # noqa: BLE001 - one bad report shouldn't stop the run; it is retried next run
            failed += 1
            print(f"  {folder}: FAILED {type(e).__name__}: {str(e)[:200]}", flush=True)
    if new_rep:
        reports = pd.concat([reports, pd.DataFrame(new_rep)], ignore_index=True) if not reports.empty \
            else pd.DataFrame(new_rep)
    if long.empty or latest_table.empty:
        sys.exit("No data read and no archive - nothing to write")
    long = long[long["date"].dt.year >= args.start_year]
    capacity = float(pd.to_numeric(latest_table["capacity_hm3"], errors="coerce").sum())
    daily = build_daily(long, capacity)
    latest = str(latest_table["report_week_end"].iloc[0])[:10]
    nl = notes(daily, reports, latest)
    keep = long.sort_values(["date", "group", "report_end"]).drop_duplicates(["date", "group"], keep="last")
    xlsx_notes.write_workbook(args.out, {"Daily": daily, "Latest week by reservoir": latest_table.set_index("reservoir"),
                                         "Weekly by basin (long)": keep.set_index("date"),
                                         "Reports read": reports.set_index("report")},
                              nl, {ln for ln in nl if ln and ln.split(" (")[0].isupper()})
    print(f"Saved {args.out}: {len(daily)} weeks {daily.index.min():%d-%b-%Y} .. {daily.index.max():%d-%b-%Y}; "
          f"capacity {capacity:,.1f} hm3; failed {failed}", flush=True)
    print(daily[["Total_hm3", "Total_pct", "Basins_reported", "Junin_hm3", "Mantaro_lagoons_hm3",
                 "Rimac_hm3", "Chili_hm3"]].tail(8).to_string(), flush=True)
    unmatched = latest_table[latest_table["group"].isna()]
    if len(unmatched):
        print("WARNING: reservoir rows not matched to a basin:", unmatched["reservoir"].tolist(), flush=True)
    if failed and failed >= len(todo):
        sys.exit(f"all {failed} reports failed")


if __name__ == "__main__":
    main()
