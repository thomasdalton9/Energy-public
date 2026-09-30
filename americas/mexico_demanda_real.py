"""
Pull Mexico's real (historical) electricity demand by load zone
("Zona de Carga") from CENACE's EstimacionDemandaReal.aspx report and
maintain a growing archive, backfilling day by day from history.

BREAKTHROUGH: after a long series of probes (see MEXICO_CENACE_
HISTORIA_PLAYWRIGHT_PROBE.py's docstring for the full debugging
history), the whole download mechanism is now proven end-to-end and,
critically, proven to work for MANY different historical days reused
across a SINGLE browser session (MEXICO_CENACE_MULTIDAY_TEST.py: three
non-consecutive days, ~1965-2026 span, all succeeded in ~18 seconds
total) - so one script run can backfill a real chunk of history
without relaunching the browser per day.

Mechanism: CENACE's date pickers are Telerik RadDatePickers - setting
one via window.$find(id).set_selectedDate(...) triggers a postback that
resets the OTHER pickers on the form back to their server defaults a
moment later, so every day's dates are set in two passes (see
set_all_dates). The live form's fields (including whatever hidden/
ClientState fields Telerik's own JS updated) are then read straight
from the DOM and POSTed via Playwright's own request context - not by
calling .click() on the download button, which reliably times out
(some actionability check never passes) - WebForms only needs the
button's name/value pair present in the POST body to know "which
button was clicked", not an actual click event.

Each successful day returns a ZIP with 1-2 CSVs (CENACE sometimes
republishes a day with a "_2"/"_3" revision suffix - always using the
LAST one, i.e. the latest revision, matching what the site itself would
show by default). The CSV's data rows are "Sistema, Zona de Carga,
Hora, Energia (MWh)" - hourly demand per load zone. This is NOT
reachable from this repo's normal editing sandbox (CENACE blocks the
sandbox's egress proxy) - run via GitHub Actions, see
.github/workflows/mexico_demanda_real.yml.

ARCHIVE DESIGN: hourly-by-zone for a full multi-year backfill would be
millions of rows - too large for a single Excel sheet (1,048,576 row
limit) and slow to open. So this keeps TWO archives:
  - a small .xlsx with one row per day, national total demand (MWh,
    summed across every zone/system) - the headline series, human-
    browsable in Excel.
  - a .csv.gz with the full hourly-by-zone detail, appended
    incrementally (never rewritten in full) - complete granularity for
    anyone who wants to dig into a specific zone or hour.

BACKFILL DESIGN: --max-days bounds how many new days one run processes
(default 90, about 3-9 minutes at ~2-6s/day), so a multi-year backfill
runs across several scheduled invocations rather than one very long
job. Each run reads the national .xlsx to find the latest day already
saved and continues forward from there (or from --start-date, default
2016-01-01, CENACE's own data start per this script's discovery
scripts, if the archive is empty) - stopping early at (today - 1), the
most recent day CENACE would have a real, settled (non-preliminary)
value for.
"""

import argparse
import csv
import gzip
import io
import os
import sys
import zipfile
from datetime import date, timedelta

import pandas as pd
from playwright.sync_api import sync_playwright

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))  # repo root, for xlsx_notes

import xlsx_notes

URL = "https://www.cenace.gob.mx/Paginas/SIM/Reportes/EstimacionDemandaReal.aspx"

DATE_PICKER_IDS = [
    "ctl00_ContentPlaceHolder1_RadDatePickerVisualizarGeneralesPorRetiros",
    "ctl00_ContentPlaceHolder1_RadDatePickerFIVisualizarPorRetiros",
    "ctl00_ContentPlaceHolder1_RadDatePickerFFVisualizarPorRetiros",
]
DOWNLOAD_BUTTON_NAME = "ctl00$ContentPlaceHolder1$DescargarArchivosCsv_PorRetiros"
DOWNLOAD_BUTTON_VALUE = "Descargar en archivo .zip"

DATA_START = date(2016, 1, 1)
DEFAULT_MAX_DAYS = 90

REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DEFAULT_XLSX_OUT = os.path.join(REPO_ROOT, "output", "mexico_demanda_nacional_daily.xlsx")
DEFAULT_CSV_OUT = os.path.join(REPO_ROOT, "output", "mexico_demanda_por_zona_hourly.csv.gz")


def set_all_dates(page, year, month, day):
    for picker_id in DATE_PICKER_IDS:
        page.evaluate(
            """(id) => {
                const picker = window.$find(id);
                if (!picker) return;
                picker.set_selectedDate(new Date(%d, %d, %d));
            }""" % (year, month - 1, day),
            picker_id,
        )


def fetch_one_day(page, context, target_day):
    """Returns a DataFrame of (date, sistema, zona_carga, hora, energia_mwh)
    rows for one day, or None if the download didn't come back as a ZIP
    (e.g. a day CENACE has no data for yet)."""
    set_all_dates(page, target_day.year, target_day.month, target_day.day)
    page.wait_for_timeout(2500)
    set_all_dates(page, target_day.year, target_day.month, target_day.day)
    page.wait_for_timeout(1000)

    fields = page.evaluate(
        """() => {
            const form = document.forms[0];
            const data = {};
            for (const el of form.elements) {
                if (!el.name) continue;
                data[el.name] = el.value;
            }
            return data;
        }"""
    )
    fields[DOWNLOAD_BUTTON_NAME] = DOWNLOAD_BUTTON_VALUE
    for noise in list(fields):
        if noise.endswith((".x", ".y")):
            fields.pop(noise, None)

    resp = context.request.post(URL, form=fields, timeout=30000)
    body = resp.body()
    if body[:2] != b"PK":
        return None

    with zipfile.ZipFile(io.BytesIO(body)) as zf:
        names = zf.namelist()
        if not names:
            return None
        # CENACE sometimes republishes a day with a revision suffix in the
        # filename ("_2", "_3", ...) - the zip's own name order isn't
        # guaranteed sorted, so pick explicitly by that trailing number,
        # highest = latest revision = what the site would show by default.
        def revision(name):
            import re
            m = re.search(r"_(\d+)\s+Dia Operacion", name)
            return int(m.group(1)) if m else 0

        chosen = max(names, key=revision)
        with zf.open(chosen) as f:
            text = f.read().decode("latin-1", errors="replace")

    lines = text.splitlines()
    header_idx = next((i for i, line in enumerate(lines) if "Energia" in line and "Hora" in line), None)
    if header_idx is None:
        return None
    csv_text = "\n".join(lines[header_idx:])
    df = pd.read_csv(io.StringIO(csv_text))
    df.columns = [c.strip() for c in df.columns]
    df = df.rename(columns={"Sistema": "sistema", "Zona de Carga": "zona_carga",
                             "Hora": "hora", "Energia (MWh)": "energia_mwh"})
    df = df.dropna(subset=["sistema"])
    df["energia_mwh"] = pd.to_numeric(df["energia_mwh"], errors="coerce")
    df.insert(0, "date", target_day.isoformat())
    return df[["date", "sistema", "zona_carga", "hora", "energia_mwh"]]


def load_national_archive(path):
    try:
        df = pd.read_excel(path, sheet_name="Data", index_col=0)
    except (FileNotFoundError, ValueError):
        return pd.DataFrame()
    df.index = pd.to_datetime(df.index).date
    df.index.name = "date"
    return df


def append_hourly_csv(path, day_df):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    file_exists = os.path.exists(path)
    mode = "ab" if file_exists else "wb"
    with gzip.open(path, mode) as f:
        text_buf = io.StringIO()
        day_df.to_csv(text_buf, index=False, header=not file_exists, quoting=csv.QUOTE_MINIMAL)
        f.write(text_buf.getvalue().encode("utf-8"))


NOTES_LINES = [
    "UNITS",
    "Total_MWh is Mexico's national real demand for that day - summed across every hourly reading "
    "for every Sistema/Zona de Carga (load zone) that day, i.e. total energy withdrawn from the grid "
    "(a daily energy total, not an average MW).",
    "",
    "DETAIL FILE",
    "This workbook has only the national daily total. Full hourly-by-load-zone detail (Sistema, "
    "Zona de Carga, Hora, Energia (MWh)) is in mexico_demanda_por_zona_hourly.csv.gz alongside this "
    "file - kept as a separate gzipped CSV rather than another tab because a multi-year hourly-by-"
    "zone archive would exceed Excel's 1,048,576-row-per-sheet limit.",
    "",
    "COVERAGE",
    f"Backfilling day by day from {DATA_START.isoformat()} - each run processes a bounded chunk of "
    "new days (see the script's --max-days), so a full multi-year backfill spans several runs rather "
    "than one very long job. Revisions: CENACE sometimes republishes a day under a revised filename "
    "- this always uses the latest revision available at run time.",
    "",
    "SOURCE",
    f"CENACE (Mexico's national grid operator) 'Demanda Real del Sistema - Por Retiros' report: {URL} "
    "- not a documented API, a real ASP.NET WebForms report page automated via Playwright (see "
    "MEXICO_CENACE_HISTORIA_PLAYWRIGHT_PROBE.py's docstring for the full mechanism).",
]
NOTES_SECTION_TITLES = {"UNITS", "DETAIL FILE", "COVERAGE", "SOURCE"}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--start-date", type=date.fromisoformat, default=DATA_START)
    parser.add_argument("--max-days", type=int, default=DEFAULT_MAX_DAYS)
    parser.add_argument("--xlsx-out", default=DEFAULT_XLSX_OUT)
    parser.add_argument("--csv-out", default=DEFAULT_CSV_OUT)
    args = parser.parse_args()

    national = load_national_archive(args.xlsx_out)
    if not national.empty:
        resume_from = max(national.index) + timedelta(days=1)
    else:
        resume_from = args.start_date
    stop_before = date.today()  # today itself is still in progress/preliminary - never fetched

    if resume_from >= stop_before:
        print(f"Archive already current ({resume_from} >= {stop_before}) - nothing to do.", file=sys.stderr)
        return

    days_to_try = []
    d = resume_from
    while d < stop_before and len(days_to_try) < args.max_days:
        days_to_try.append(d)
        d += timedelta(days=1)

    print(f"Fetching {len(days_to_try)} day(s): {days_to_try[0]} to {days_to_try[-1]}", file=sys.stderr)

    new_totals = {}
    with sync_playwright() as p:
        browser = p.chromium.launch()
        context = browser.new_context()
        page = context.new_page()
        page.goto(URL, timeout=60000, wait_until="load")
        page.wait_for_selector(f"#{DATE_PICKER_IDS[0]}", timeout=30000, state="attached")

        for target_day in days_to_try:
            try:
                day_df = fetch_one_day(page, context, target_day)
            except Exception as e:
                print(f"  {target_day}: FAILED ({type(e).__name__}: {e})", file=sys.stderr)
                continue
            if day_df is None or day_df.empty:
                print(f"  {target_day}: no data (not yet published?)", file=sys.stderr)
                continue
            append_hourly_csv(args.csv_out, day_df)
            total = day_df["energia_mwh"].sum()
            new_totals[target_day] = round(total, 1)
            print(f"  {target_day}: OK, {len(day_df)} zone-hours, {total:,.0f} MWh total", file=sys.stderr)

        browser.close()

    if new_totals:
        new_df = pd.DataFrame({"Total_MWh": new_totals})
        new_df.index.name = "date"
        combined = pd.concat([national, new_df]) if not national.empty else new_df
        combined = combined[~combined.index.duplicated(keep="last")].sort_index()

        os.makedirs(os.path.dirname(args.xlsx_out), exist_ok=True)
        xlsx_notes.write_workbook(args.xlsx_out, {"Data": combined}, NOTES_LINES, NOTES_SECTION_TITLES)
        print(f"\nSaved {len(new_totals)} new day(s) to {args.xlsx_out} "
              f"(archive now {len(combined)} days, {combined.index.min()} to {combined.index.max()})")
    else:
        print("\nNo new days saved this run.", file=sys.stderr)


if __name__ == "__main__":
    main()
