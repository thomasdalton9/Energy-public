"""
Pull Mexico's real (historical) electricity demand by load zone
("Zona de Carga") from CENACE's "Demanda Real del Sistema - Por
Retiros" report (EstimacionDemandaReal.aspx) and keep a national daily
archive plus the full hourly-by-zone detail.

MECHANISM: CENACE's date pickers are Telerik RadDatePickers. The page is
loaded in Playwright, the report's three pickers (Generales, FI = fecha
inicial, FF = fecha final) are set through Telerik's client API
(window.$find(id).set_selectedDate) - twice, because setting one fires
an async postback that resets the others a moment later - and the live
form is then POSTed with the "Descargar en archivo .zip" button's
name/value (no .click(); see MEXICO_CENACE_HISTORIA_PLAYWRIGHT_PROBE.py).

An FI..FF RANGE returns every operating day in the range in one ZIP
(a whole year is ~34 MB and ~20 s), one CSV per day per settlement:
    "Demanda Real Retiro_<N> Dia Operacion YYYY-MM-DD v<published>.csv"
N is the settlement (LIQUIDACION) number: 0 is published ~14 days after
the operating day, then 1, 2, 3 over the following ~7 months (2017 days
also have 4, L5, L6). The highest N available is used. The pickers'
maxDate (read from the page) is the latest day CENACE has published,
about 15 days behind today - nothing later can be requested.
Found via discovery_archive/americas/MEXICO_CENACE_DATE_DIAGNOSTIC.py and
MEXICO_CENACE_FORMAT_DIAGNOSTIC.py (Oct 2026).

WHY THIS WAS REWRITTEN (Oct 2026): the previous version requested one
day at a time and trusted that the file it got back was for the day it
asked for. It wasn't always: the picker postbacks race, and 19 of its
327 saved days were exact copies of a neighbouring day's file (e.g.
2016-02-14/15/16 identical). It also always chose "the highest numeric
revision", which from March 2017 is a file in a different layout
("Estimacion de Demanda por Retiros", later fully quoted) that its
parser didn't recognise - so every day from 2017-02-27 on was logged as
"no data", while a checkpoint advanced past each "no data" day anyway.
The backfill therefore "finished" at 2026-09-29 holding only 2016-01-27
to 2017-02-26. Now:
  - every row's date comes from the CSV's own filename and is
    cross-checked against its "Dia de Operacion: dd/mm/yyyy" line, and
    any day outside the requested range is rejected;
  - one parser handles all three layouts (plain, quoted, broken-quoted);
  - there is no "attempted" checkpoint: the archive resumes from its own
    last saved day, and a run that can't get a valid ZIP fails loudly.

COVERAGE: from 2016-01-29, the first day with the national
interconnected system (SIN) in the report. CENACE's market data starts
2016-01-27, but 27-28 Jan 2016 cover only Baja California (BCA, ~33,000
MWh/day) and are not national totals, so days without SIN are skipped.
Baja California Sur (BCS) is in the report from 2016-03-23 on and is
included whenever CENACE reports it.

ARCHIVES:
  - mexico_demanda_nacional_daily.xlsx: one row per day - Total_MWh
    (all systems) plus per-system totals and the settlement used.
  - mexico_demanda_por_zona_hourly/YYYY-MM.csv.gz: the hourly-by-zone
    detail, one small file per month (a single multi-year file would
    approach GitHub's 100 MB limit, and rewriting it for every revision
    would bloat the repo).

INCREMENTAL: each run fetches from the day after the latest saved day,
plus a trailing REVISION_WINDOW_DAYS so later settlements (1-3) replace
earlier ones; days already at an equal or higher settlement are left
alone. --max-days bounds how much new history one run takes on.
Not reachable from the editing sandbox (CENACE blocks it) - run via
.github/workflows/mexico_demanda_real.yml.
"""

import argparse
import gzip
import io
import os
import re
import sys
import zipfile
from datetime import date, timedelta

import pandas as pd
from playwright.sync_api import sync_playwright

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))  # repo root, for xlsx_notes

import xlsx_notes

URL = "https://www.cenace.gob.mx/Paginas/SIM/Reportes/EstimacionDemandaReal.aspx"

PICKER_GEN = "ctl00_ContentPlaceHolder1_RadDatePickerVisualizarGeneralesPorRetiros"
PICKER_FI = "ctl00_ContentPlaceHolder1_RadDatePickerFIVisualizarPorRetiros"
PICKER_FF = "ctl00_ContentPlaceHolder1_RadDatePickerFFVisualizarPorRetiros"
DOWNLOAD_BUTTON_NAME = "ctl00$ContentPlaceHolder1$DescargarArchivosCsv_PorRetiros"
DOWNLOAD_BUTTON_VALUE = "Descargar en archivo .zip"

DATA_START = date(2016, 1, 29)  # first day the report includes the SIN (see docstring)
DEFAULT_MAX_DAYS = 5000         # enough for the whole history in one run
CHUNK_DAYS = 183                # one POST per half year (~17 MB ZIP)
REVISION_WINDOW_DAYS = 240      # settlements 1-3 arrive up to ~7 months after the day

REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT_DIR = os.path.join(REPO_ROOT, "output", "Data and Chart Outputs")
DEFAULT_XLSX_OUT = os.path.join(OUT_DIR, "mexico_demanda_nacional_daily.xlsx")
DEFAULT_HOURLY_DIR = os.path.join(OUT_DIR, "mexico_demanda_por_zona_hourly")

FILE_RE = re.compile(r"Retiro_L?(\d+) Dia Operacion (\d{4}-\d{2}-\d{2})")
DIA_RE = re.compile(r"Dia de Operacion:\s*(\d{2})/(\d{2})/(\d{4})")
# Data rows in all three layouts CENACE has used:
#   BCA,ENSENADA,1,115.804,              (2016-2017, trailing comma)
#   "BCA","ENSENADA","1","117.79755"     (quoted)
ROW_RE = re.compile(r'^\s*"?([A-Z]{2,4})"?\s*,\s*"?([^",]+?)"?\s*,\s*"?(\d{1,2})"?\s*,\s*"?'
                    r'(-?\d+(?:\.\d*)?(?:[eE][-+]?\d+)?)"?\s*,?\s*$')
SYSTEMS = ["SIN", "BCA", "BCS"]
DATA_COLUMNS = ["Total_MWh"] + [f"{s}_MWh" for s in SYSTEMS] + ["Liquidacion"]


def log(msg):
    print(msg, file=sys.stderr, flush=True)


# ---------------------------------------------------------------- CENACE

def set_picker(page, picker_id, d):
    page.evaluate(
        """([id, y, m, d]) => { const p = window.$find(id);
            if (p) p.set_selectedDate(new Date(y, m, d)); }""",
        [picker_id, d.year, d.month - 1, d.day],
    )


def read_pickers(page):
    return page.evaluate(
        """(ids) => ids.map(id => { const p = window.$find(id); const s = p && p.get_selectedDate();
            return s ? [s.getFullYear(), s.getMonth() + 1, s.getDate()] : null; })""",
        [PICKER_GEN, PICKER_FI, PICKER_FF],
    )


def load_page(page):
    page.goto(URL, timeout=90000, wait_until="load")
    page.wait_for_selector(f"#{PICKER_GEN}", timeout=30000, state="attached")


def latest_published_day(page):
    """The pickers' maxDate - the last operating day CENACE has published."""
    ymd = page.evaluate(
        """(id) => { const p = window.$find(id); const m = p && p.get_maxDate();
            return m ? [m.getFullYear(), m.getMonth() + 1, m.getDate()] : null; }""",
        PICKER_FI,
    )
    if not ymd:
        raise RuntimeError("Could not read the date pickers' maxDate - page layout changed?")
    return date(*ymd)


def download_range(page, context, first, last, attempts=3):
    """ZIP bytes for operating days first..last (inclusive). The pickers
    are read back before posting; a mismatch reloads the page and retries."""
    want = [[first.year, first.month, first.day], [first.year, first.month, first.day],
            [last.year, last.month, last.day]]
    for attempt in range(1, attempts + 1):
        load_page(page)
        for _ in range(2):  # second pass undoes the postback reset of FI/FF
            set_picker(page, PICKER_GEN, first)
            set_picker(page, PICKER_FI, first)
            set_picker(page, PICKER_FF, last)
            page.wait_for_timeout(2500)
        got = read_pickers(page)
        if got != want:
            log(f"  pickers read back {got}, wanted {want} (attempt {attempt}) - retrying")
            continue
        fields = page.evaluate(
            """() => { const data = {}; for (const el of document.forms[0].elements) {
                if (el.name) data[el.name] = el.value; } return data; }"""
        )
        fields[DOWNLOAD_BUTTON_NAME] = DOWNLOAD_BUTTON_VALUE
        for noise in [k for k in fields if k.endswith((".x", ".y"))]:
            fields.pop(noise)
        resp = context.request.post(URL, form=fields, timeout=300000)
        body = resp.body()
        if body[:2] == b"PK":
            return body
        log(f"  {first}..{last}: HTTP {resp.status}, {resp.headers.get('content-type')}, "
            f"{len(body):,} bytes - not a ZIP (attempt {attempt})")
    raise RuntimeError(f"CENACE did not return a ZIP for {first}..{last} after {attempts} attempts")


def parse_csv(text):
    """Rows (sistema, zona_carga, hora, energia_mwh) from any CENACE layout."""
    rows = []
    for line in text.splitlines():
        m = ROW_RE.match(line)
        if m:
            rows.append((m.group(1), m.group(2).strip(), int(m.group(3)), float(m.group(4))))
    return rows


def parse_zip(body, first, last):
    """{day: (liquidacion, DataFrame)} using the highest settlement per day
    that has data rows (a few late-settlement files hold none - the next
    settlement down is used then). Each file's date comes from its own
    name and must match the 'Dia de Operacion' line inside it and lie
    within first..last."""
    files = {}
    with zipfile.ZipFile(io.BytesIO(body)) as zf:
        for name in zf.namelist():
            m = FILE_RE.search(name)
            if not m:
                log(f"  skipping unrecognised file {name!r}")
                continue
            liq, day = int(m.group(1)), date.fromisoformat(m.group(2))
            if not first <= day <= last:
                raise RuntimeError(f"ZIP for {first}..{last} contains {name!r} - outside the requested range")
            files.setdefault(day, []).append((liq, name))

        out, unparsed = {}, []
        for day, candidates in sorted(files.items()):
            for liq, name in sorted(candidates, reverse=True):
                text = zf.read(name).decode("latin-1", errors="replace")
                dm = DIA_RE.search(text)
                if dm:
                    inner = date(int(dm.group(3)), int(dm.group(2)), int(dm.group(1)))  # dd/mm/yyyy
                    if inner != day:
                        raise RuntimeError(f"{name!r} says Dia de Operacion {inner}, filename says {day}")
                rows = parse_csv(text)
                if rows:
                    out[day] = (liq, pd.DataFrame(rows, columns=["sistema", "zona_carga", "hora", "energia_mwh"]))
                    break
                lines = text.splitlines()
                log(f"  {name!r}: no data rows ({len(text):,} chars) - trying an earlier settlement. "
                    f"Head: {lines[:3]!r} ... tail: {lines[-2:]!r}")
            else:
                unparsed.append(day)
        if unparsed:
            log(f"  {len(unparsed)} day(s) with no parseable file at all: {[d.isoformat() for d in unparsed[:10]]}")
        if files and len(unparsed) > 0.2 * len(files):
            raise RuntimeError(f"{len(unparsed)} of {len(files)} days in {first}..{last} unparseable - layout changed?")
    return out


# --------------------------------------------------------------- archive

def load_national(path):
    try:
        df = pd.read_excel(path, sheet_name="Data", index_col=0)
    except (FileNotFoundError, ValueError):
        return pd.DataFrame(columns=DATA_COLUMNS)
    if "Liquidacion" not in df.columns:
        # Written by the pre-Oct-2026 version: dates not validated and
        # some days mislabelled - discard and rebuild.
        log(f"{path} is from the old unvalidated pull - rebuilding it from scratch.")
        return pd.DataFrame(columns=DATA_COLUMNS)
    df.index = pd.to_datetime(df.index).date
    df.index.name = "date"
    return df


def write_hourly_month(hourly_dir, month, day_frames):
    """Merge {day: DataFrame} into hourly_dir/YYYY-MM.csv.gz (replacing those days)."""
    os.makedirs(hourly_dir, exist_ok=True)
    path = os.path.join(hourly_dir, f"{month}.csv.gz")
    new = pd.concat([df.assign(date=d.isoformat()) for d, df in sorted(day_frames.items())])
    new = new[["date", "sistema", "zona_carga", "hora", "energia_mwh"]]
    if os.path.exists(path):
        old = pd.read_csv(path, dtype={"date": str})
        old = old[~old["date"].isin(new["date"].unique())]
        new = pd.concat([old, new]).sort_values(["date", "sistema", "zona_carga", "hora"], kind="stable")
    tmp = path + f".tmp{os.getpid()}"
    # mtime=0 keeps the gzip bytes identical when the data is unchanged.
    with open(tmp, "wb") as raw, gzip.GzipFile(fileobj=raw, mode="wb", mtime=0) as gz:
        gz.write(new.to_csv(index=False).encode("utf-8"))
    os.replace(tmp, path)


NOTES_LINES = [
    "UNITS",
    "Total_MWh is Mexico's real electricity demand for the day (energy withdrawn, 'demanda real por "
    "retiros'): the sum of every hourly value for every load zone (Zona de Carga) in every system - "
    "a daily energy total in MWh, not an average MW. SIN_MWh / BCA_MWh / BCS_MWh split it by system "
    "(Sistema Interconectado Nacional, Baja California, Baja California Sur).",
    "",
    "SETTLEMENT (Liquidacion)",
    "CENACE publishes each day several times: settlement 0 about 14 days after the operating day, then "
    "1, 2 and 3 over the next ~7 months (2017 days also have 4-6). The highest settlement available is "
    "used and replaced when a later one appears, so the last ~7 months can still be revised. The most "
    "recent day is about 15 days before the pull date - CENACE publishes nothing more recent.",
    "",
    "COVERAGE",
    f"From {DATA_START.isoformat()}, the first day the report includes the national interconnected "
    "system (SIN). 27-28 January 2016 cover only Baja California and are excluded. BCS is in the "
    "report from 2016-03-23 (BCS_MWh blank before that, so early totals exclude BCS's ~5,000 MWh/day).",
    "",
    "DETAIL FILES",
    "Hourly demand by load zone (date, sistema, zona_carga, hora, energia_mwh) is in the "
    "mexico_demanda_por_zona_hourly folder next to this workbook, one gzipped CSV per month (too many "
    "rows for an Excel sheet).",
    "",
    "SOURCE",
    f"CENACE (Centro Nacional de Control de Energia), 'Demanda Real del Sistema - Por Retiros': {URL} "
    "- an ASP.NET report page automated with Playwright, downloading date ranges as ZIPs of daily CSVs. "
    "Each day's date is taken from its file name and checked against the file's own 'Dia de Operacion'.",
]
NOTES_SECTION_TITLES = {"UNITS", "SETTLEMENT (Liquidacion)", "COVERAGE", "DETAIL FILES", "SOURCE"}


def save_national(path, national):
    national = national.sort_index()
    national.index.name = "date"
    xlsx_notes.write_workbook(path, {"Data": national[DATA_COLUMNS]}, NOTES_LINES, NOTES_SECTION_TITLES)


def main():
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--start-date", type=date.fromisoformat, default=DATA_START)
    parser.add_argument("--max-days", type=int, default=DEFAULT_MAX_DAYS,
                        help="max days of NEW history (beyond the archive's end) to fetch this run")
    parser.add_argument("--xlsx-out", default=DEFAULT_XLSX_OUT)
    parser.add_argument("--hourly-dir", default=DEFAULT_HOURLY_DIR)
    args = parser.parse_args()

    national = load_national(args.xlsx_out)
    national = national[national.index >= args.start_date] if not national.empty else national

    with sync_playwright() as p:
        browser = p.chromium.launch()
        context = browser.new_context()
        page = context.new_page()
        load_page(page)
        latest = latest_published_day(page)

        # New history continues from the day after the archive's last day;
        # the trailing revision window is re-fetched so later settlements
        # replace earlier ones.
        new_from = args.start_date if national.empty else max(national.index) + timedelta(days=1)
        fetch_from = new_from if national.empty else \
            max(args.start_date, min(new_from, latest - timedelta(days=REVISION_WINDOW_DAYS)))
        fetch_to = min(latest, new_from + timedelta(days=args.max_days - 1))
        log(f"CENACE latest published day: {latest}. Fetching {fetch_from} to {fetch_to}.")
        if fetch_from > fetch_to:
            log("Nothing to fetch.")
            browser.close()
            return

        updated_days = 0
        chunk_start = fetch_from
        while chunk_start <= fetch_to:
            chunk_end = min(fetch_to, chunk_start + timedelta(days=CHUNK_DAYS - 1))
            body = download_range(page, context, chunk_start, chunk_end)
            days = parse_zip(body, chunk_start, chunk_end)
            changed = {}
            for day, (liq, df) in days.items():
                systems = set(df["sistema"])
                if "SIN" not in systems:
                    log(f"  {day}: no SIN rows (systems {sorted(systems)}) - not a national total, skipped")
                    continue
                if day in national.index and pd.notna(national.at[day, "Liquidacion"]) \
                        and int(national.at[day, "Liquidacion"]) >= liq:
                    continue
                by_sys = df.groupby("sistema")["energia_mwh"].sum()
                row = {"Total_MWh": round(df["energia_mwh"].sum(), 1),
                       **{f"{s}_MWh": (round(by_sys[s], 1) if s in by_sys else None) for s in SYSTEMS},
                       "Liquidacion": liq}
                national.loc[day] = pd.Series(row)
                changed[day] = df
            missing = [chunk_start + timedelta(days=i) for i in range((chunk_end - chunk_start).days + 1)
                       if chunk_start + timedelta(days=i) not in days]
            if missing:
                log(f"  {len(missing)} day(s) absent from CENACE's ZIP: "
                    f"{', '.join(d.isoformat() for d in missing[:10])}{' ...' if len(missing) > 10 else ''}")
            if changed:
                by_month = {}
                for day, df in changed.items():
                    by_month.setdefault(day.strftime("%Y-%m"), {})[day] = df
                for month, frames in sorted(by_month.items()):
                    write_hourly_month(args.hourly_dir, month, frames)
                national.index = pd.Index(national.index, name="date")
                save_national(args.xlsx_out, national)
                updated_days += len(changed)
            log(f"  {chunk_start}..{chunk_end}: {len(days)} day(s) in ZIP, {len(changed)} new/revised")
            chunk_start = chunk_end + timedelta(days=1)

        browser.close()

    if updated_days:
        print(f"Saved {updated_days} new/revised day(s). Archive: {len(national)} days, "
              f"{min(national.index)} to {max(national.index)} -> {args.xlsx_out}")
    else:
        print("No new or revised days this run.")


if __name__ == "__main__":
    main()
