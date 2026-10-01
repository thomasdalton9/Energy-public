"""
Peru hydro reservoirs: useful volume (volumen util, hm3) of the SEIN's
seasonal reservoirs and lagoons - Lake Junin, the Mantaro-basin lagoons
(Electroperu, Statkraft and the other Mantaro sub-basin systems), the
Rimac system (Sheque, Antacoto/Marcapomacocha, Yuracmayo ...), the Chili
system (El Frayle, Pillones, El Pane ...), Aricota, Sibinacocha and the rest
- and the national total in hm3 and as % of total useful capacity, from
COES. No key needed.

Source: COES 'Informe Semanal de Evaluacion de la Operacion'
(https://www.coes.org.pe/Portal/PostOperacion/Informes/EvaluacionSemanal),
one Excel workbook per operating week (Saturday-Friday), listed by the
portal's file browser under
    Post Operacion/Informes/Evaluacion Semanal/<year>/SEMANAL N° <w> (dd.mm.yyyy - dd.mm.yyyy)/
and downloaded through /Portal/browser/download?url=... . Its section
5.1 'VOLUMEN UTIL DE LOS EMBALSES Y LAGUNAS (Millones de m3)' lists every
reservoir / lagoon group with its useful volume at the start (Saturday)
and end (Friday) of the week, % full, and its useful capacity (hm3) - for
the week and for the same week a year earlier.
Found via discovery_archive/south_america/HYDRO_PE_EC_UY_DISCOVERY.py (round 3).
COES publishes no daily table covering all lagoons: its daily IDCOS / IEOD
hydrology annexes carry only a handful of seasonal reservoirs (Junin,
Sibinacocha, Aricota, Viconga) and the Mantaro lagoons only as discharges.
So this series has two readings a week (start and end of each COES week).

National total = sum of all reservoirs in that week's table; national % =
total / sum of their useful capacities (COES' own capacities, 1,970 hm3 in
2026). Basin groups are this script's grouping of COES' rows (BASINS);
group % = group volume / group capacity.

Each report also carries the same week a year earlier, so the reports
from 2021 on give history back to early 2020. A date's value comes from
the report of its own week; the year-earlier columns only fill dates no
report of their own covers.

Incremental: the workbook is the archive ('Reports read' sheet); each run
lists the year folders and reads only weekly reports not yet read (the
backfill runs newest-first within --budget-min minutes).

Usage: python3 PERU_HYDRO_RESERVOIRS.py [--out PATH] [--start-year 2021] [--budget-min N] [--max-reports N]
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
from datetime import date
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

# group -> keywords matched (accent-free, upper case) against COES' row label. First match wins.
BASINS = [
    ("Junin", ["JUNIN"]),
    ("Mantaro_lagoons", ["ELECTROPERU", "LAGUNAS STATKRAFT", "POMACOCHA", "VICHECOCHA", "CARHUACOCHA", "CHILLICOCHA",
                         "CHILICOCHA", "HUICHICOCHA", "COYLLORCOCHA", "YURAJCOCHA"]),
    ("Rimac", ["SHEQUE", "QUISHA", "SACSA", "ANTACOTO", "MARCAPOMACOCHA", "YURACMAYO"]),
    ("Chili", ["EL PANE", "BAMPUTANE", "ESPANOLES", "CHALHUANCA", "PILLONES", "FRAYLE", "AGUADA BLANCA"]),
    ("Aricota", ["ARICOTA"]),
    ("Sibinacocha", ["SIBINACOCHA"]),
    ("Paucartambo", ["JAICO", "PACCHAPATA", "MACHAY", "MATACOCHA", "HUANGUSH"]),
    ("Santa", ["AGUASCOCHA", "RAJUCOLTA", "CULLICOCHA"]),
    ("Viconga", ["VICONGA"]),
    ("San_Gaban", ["AJOYAJOTA", "PARINAJOTA", "ISOCOCHA"]),
    ("Canete", ["PAUCARCOCHA"]),
    ("Chaglla", ["CHAGLLA"]),
]
GROUP_LABEL = {"Junin": "Lake Junin (Chinchaycocha)", "Mantaro_lagoons": "Mantaro-basin lagoons (Electroperu, Statkraft, "
               "Carhuacocha, Chillicocha, Huichicocha ...)", "Rimac": "Rimac system (Sheque, Sacsa, Antacoto/"
               "Marcapomacocha, Yuracmayo)", "Chili": "Chili system, Arequipa (El Frayle, Pillones, El Pane, Bamputane, "
               "Aguada Blanca, Chalhuanca, Los Espanoles)", "Aricota": "Aricota (Locumba)",
               "Sibinacocha": "Sibinacocha (Vilcanota)", "Paucartambo": "Paucartambo (Jaico/Pacchapata/Altos Machay, "
               "Matacocha/Huangush)", "Santa": "Santa (Aguascocha, Rajucolta, Cullicocha)", "Viconga": "Viconga (Pativilca)",
               "San_Gaban": "San Gaban (Ajoyajota, Parinajota, Isococha ...)", "Canete": "Paucarcocha (Canete)",
               "Chaglla": "Chaglla (Huallaga)", "Other": "Rows not matched to a group"}

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
    return "Other"


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
    """[(folder, week number, start date, end date)] for one year."""
    out = []
    for p, k in browse(f"{BASE}{year}/"):
        m = WEEK.search(p)
        if k == "D" and m:
            g = list(map(int, m.groups()))
            out.append((p, g[0], date(g[3], g[2], g[1]), date(g[6], g[5], g[4])))
    return out


def parse_volumes(content):
    """Section 5.1 of a weekly report -> long frame: date, reservoir, volume_hm3, capacity_hm3, own_week (bool)."""
    xl = pd.ExcelFile(io.BytesIO(content))
    for sh in xl.sheet_names:
        df = pd.read_excel(xl, sheet_name=sh, header=None, nrows=80)
        text = " ".join(key(v) for v in df.values.ravel() if isinstance(v, str))
        if "UTIL DE LOS EMBALSES Y LAGUNAS" not in text:
            continue
        # header row: the one with 'VOLUMEN UTIL dd/mm/yyyy' cells
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
            continue   # e.g. the table of contents names the section too
        # the label column: the one with 'LAGUNA / EMBALSE' above, else the first text column left of the data
        label_col = None
        for r in range(max(0, i - 3), i + 1):
            for j, v in df.iloc[r].items():
                if isinstance(v, str) and re.fullmatch(r"(LAGUNAS? ?/ ?EMBALSES?|EMBALSES? ?/ ?LAGUNAS?)", key(v)):
                    label_col = j
        if label_col is None:
            label_col = min(vol_cols) - 1
        newest = max(vol_cols.values()).year
        rows = []
        for r in range(i + 1, len(df)):
            label = df.iat[r, label_col]
            if isinstance(label, str) and key(label).startswith(("CUADRO", "FUENTE", "NOTA")):
                break
            if not isinstance(label, str) or not label.strip():
                continue
            name = re.sub(r"^Volumen [UÚ]til de(l)? (Embalse|Laguna)s? ", "", label.strip(), flags=re.I).strip()
            for j, d in vol_cols.items():
                cap_j = next((c for c in cap_cols if c > j), None)
                vol = pd.to_numeric(df.iat[r, j], errors="coerce")
                cap = pd.to_numeric(df.iat[r, cap_j], errors="coerce") if cap_j is not None else float("nan")
                if pd.notna(vol):
                    rows.append({"date": pd.Timestamp(d), "reservoir": name, "volume_hm3": float(vol),
                                 "capacity_hm3": float(cap) if pd.notna(cap) else None, "own_week": d.year == newest})
        if not rows:
            raise RuntimeError(f"sheet {sh}: header found but no reservoir rows")
        return pd.DataFrame(rows)
    raise RuntimeError("no '5.1 VOLUMEN UTIL DE LOS EMBALSES Y LAGUNAS' sheet")


def load(path, sheet):
    try:
        return pd.read_excel(path, sheet_name=sheet)
    except (FileNotFoundError, ValueError, KeyError, OSError):
        return pd.DataFrame()


def build_daily(raw):
    """Long per-reservoir rows -> one row per date: national total / %, then each group's hm3 and %."""
    raw = raw.copy()
    raw["date"] = pd.to_datetime(raw["date"])
    # a date's own report wins over a year-later report's look-back columns
    raw = raw.sort_values(["date", "reservoir", "own_week", "report"]).drop_duplicates(["date", "reservoir"], keep="last")
    raw["group"] = raw["reservoir"].map(group_of)
    vol = raw.pivot_table(index="date", columns="group", values="volume_hm3", aggfunc="sum")
    cap = raw.pivot_table(index="date", columns="group", values="capacity_hm3", aggfunc="sum")
    out = pd.DataFrame(index=vol.index)
    out["Total_hm3"] = vol.sum(axis=1).round(2)
    out["Total_capacity_hm3"] = cap.sum(axis=1).round(2)
    out["Total_pct"] = (100 * out["Total_hm3"] / out["Total_capacity_hm3"]).round(2)
    out["Reservoirs_reported"] = raw.groupby("date")["reservoir"].nunique()
    for g in [g for g, _ in BASINS] + ["Other"]:
        if g in vol:
            out[f"{g}_hm3"] = vol[g].round(3)
            out[f"{g}_pct"] = (100 * vol[g] / cap[g]).round(2) if g in cap else float("nan")
    out.index.name = "date"
    return out.sort_index()


def notes(daily, reports):
    last = daily.dropna(subset=["Total_hm3"]).iloc[-1]
    groups = [g for g, _ in BASINS] + ["Other"]
    return [
        "UNITS",
        "hm3 = million m3 of USEFUL volume (volumen util: water above each reservoir's minimum operating level). "
        "*_pct = % of that reservoir group's useful capacity as COES states it in the same table.",
        "Total_hm3 = sum of every reservoir/lagoon row in COES' table 5.1 for that date; Total_capacity_hm3 = sum of "
        "their capacities; Total_pct = Total_hm3 / Total_capacity_hm3 (the national % of useful storage). "
        "Reservoirs_reported = rows in COES' table that date.",
        "",
        "COVERAGE",
        f"{daily.index.min():%d-%b-%Y} to {daily.index.max():%d-%b-%Y}: two readings per week - the start (Saturday) "
        "and end (Friday) of each COES operating week. COES publishes no daily table covering all lagoons. "
        f"Latest ({last.name:%d-%b-%Y}): {last['Total_hm3']:,.0f} hm3 = {last['Total_pct']:.1f}% of "
        f"{last['Total_capacity_hm3']:,.0f} hm3.",
        f"Weekly reports read: {len(reports)} (sheet 'Reports read'). Reports from 2021 on also give the same weeks "
        "of the year before, so history starts in early 2020.",
        "",
        "GROUPS (this script's grouping of COES' rows by basin; full detail per row in sheet 'COES table 5.1')",
        *[f"{g}: {GROUP_LABEL[g]}" for g in groups],
        "",
        "SOURCE",
        f"COES, Informe Semanal de Evaluacion de la Operacion del SEIN ({PAGE}), section 5.1 'Volumen util de los "
        "embalses y lagunas (millones de m3)', one Excel per week, via the portal's file browser "
        "(Post Operacion/Informes/Evaluacion Semanal/<year>/SEMANAL N° <w> (...)/Informe_Semanal_SEM<w>_<year>.xlsx).",
        "Script: south_america/PERU_HYDRO_RESERVOIRS.py (scheduled by .github/workflows/peru_hydro_reservoirs.yml).",
        "",
        "METHOD",
        "A date's values come from the weekly report covering that week; the year-earlier columns of later reports "
        "only fill dates that have no report of their own. Incremental: reports already read are skipped.",
    ]


def save(path, raw, reports):
    daily = build_daily(raw)
    table = raw.sort_values(["date", "reservoir", "own_week", "report"]).drop_duplicates(["date", "reservoir"],
                                                                                         keep="last")
    table = table.assign(group=table["reservoir"].map(group_of)).sort_values(["date", "group", "reservoir"])
    nl = notes(daily, reports)
    xlsx_notes.write_workbook(path, {"Daily": daily, "COES table 5.1": table.set_index("date"),
                                     "Reports read": reports.set_index("report")},
                              nl, {ln for ln in nl if ln and ln.split(" (")[0].isupper()})
    return daily


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--out", default=OUT)
    ap.add_argument("--start-year", type=int, default=START_YEAR)
    ap.add_argument("--budget-min", type=float, default=40)
    ap.add_argument("--max-reports", type=int, default=0, help="test: read at most N new reports (spread over years)")
    args = ap.parse_args()
    t0 = time.time()
    os.makedirs(os.path.dirname(os.path.abspath(args.out)), exist_ok=True)

    raw = load(args.out, "COES table 5.1")
    reports = load(args.out, "Reports read")
    if not raw.empty:
        raw["date"] = pd.to_datetime(raw["date"])
    done = set(reports["report"]) if not reports.empty else set()
    weeks = []
    for y in range(args.start_year, date.today().year + 1):
        try:
            ws = list_weeks(y)
        except Exception as e:  # noqa: BLE001
            print(f"  {y}: listing FAILED {type(e).__name__}: {e}", flush=True)
            continue
        print(f"  {y}: {len(ws)} weekly folders", flush=True)
        weeks += ws
    todo = sorted([w for w in weeks if w[0] not in done], key=lambda w: w[3], reverse=True)
    if args.max_reports:
        step = max(1, len(todo) // args.max_reports)
        todo = todo[::step][:args.max_reports]
    print(f"{len(done)} reports already read; {len(todo)} to read", flush=True)
    new_raw, new_rep, failed = [], [], 0
    for n, (folder, wk, a, b) in enumerate(todo, 1):
        if time.time() - t0 > args.budget_min * 60:
            print(f"time budget reached; {len(todo) - n + 1} reports left for later runs", flush=True)
            break
        try:
            files = [p for p, k in browse(folder) if k == "F" and p.lower().endswith((".xlsx", ".xlsm", ".xls"))]
            if not files:
                raise RuntimeError("no Excel file in the folder")
            f = sorted(files, key=lambda p: ("SEMANAL" not in key(p), p))[0]
            df = parse_volumes(download(f))
            df["report"] = folder
            new_raw.append(df)
            own = df[df["own_week"]]
            new_rep.append({"report": folder, "file": f.split("/")[-1], "week": wk, "start": a, "end": b,
                            "rows": own["reservoir"].nunique(), "total_hm3_end": round(own[own["date"] == own["date"].max()]["volume_hm3"].sum(), 2)})
            print(f"  [{n}/{len(todo)}] {b}: {own['reservoir'].nunique()} reservoirs, end-of-week total "
                  f"{new_rep[-1]['total_hm3_end']:,.1f} hm3", flush=True)
        except Exception as e:  # noqa: BLE001 - one bad report shouldn't stop the run; it is retried next run
            failed += 1
            print(f"  [{n}/{len(todo)}] {folder}: FAILED {type(e).__name__}: {str(e)[:200]}", flush=True)
        if new_raw and n % 25 == 0:   # checkpoint
            raw = pd.concat([raw] + new_raw, ignore_index=True)
            reports = pd.concat([reports, pd.DataFrame(new_rep)], ignore_index=True)
            new_raw, new_rep = [], []
            save(args.out, raw, reports)
    if new_raw:
        raw = pd.concat([raw] + new_raw, ignore_index=True)
        reports = pd.concat([reports, pd.DataFrame(new_rep)], ignore_index=True)
    if raw.empty:
        sys.exit("No data read and no archive - nothing to write")
    daily = save(args.out, raw, reports)
    print(f"Saved {args.out}: {len(daily)} dates {daily.index.min():%d-%b-%Y} .. {daily.index.max():%d-%b-%Y}; "
          f"{len(reports)} reports; failed {failed}", flush=True)
    print(daily[["Total_hm3", "Total_capacity_hm3", "Total_pct", "Reservoirs_reported", "Junin_hm3",
                 "Mantaro_lagoons_hm3", "Rimac_hm3"]].tail(8).to_string(), flush=True)
    if "Other_hm3" in daily:
        print("WARNING: rows not matched to a group:",
              sorted(set(raw.loc[raw["reservoir"].map(group_of) == "Other", "reservoir"])), flush=True)
    if failed > max(3, len(todo) // 5):
        sys.exit(f"{failed} reports failed")


if __name__ == "__main__":
    main()
