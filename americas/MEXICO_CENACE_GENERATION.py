"""
Pull Mexico's power generation by technology from CENACE's "Energia
Generada por Tipo de Tecnologia" report (EnergiaGeneradaTipoTec.aspx,
"Generacion Liquidada"): hourly settled generation (MWh) for the whole
national electricity system (SEN = SIN + Baja California + Baja
California Sur) by technology, kept as a daily workbook in the repo's
standard raw-generation layout plus the hourly detail.

MECHANISM: the report's pickers are Telerik MONTH pickers (FechaConsulta,
FechaInicial, FechaFinal). They are set through Telerik's client API
(window.$find(id).set_selectedDate, twice - setting one fires an async
postback that resets the others) and read back, then the "Descargar en
archivo .zip" button is CLICKED and the download caught. (POSTing the
form as mexico_demanda_real.py does returns the HTML page here, not the
ZIP - discovery_archive/americas/MEXICO_CENACE_GENERATION_PROBE.py.)

An FechaInicial..FechaFinal month range returns one ZIP holding a CSV
per month per settlement:
    "Generacion Liquidada_L<N> SEN <mes> <yyyy> v<published>.csv"
with columns Sistema, Dia (dd/mm/yyyy), Hora (1-24, 25 on the autumn
clock change), then MWh for Eolica, Fotovoltaica, Biomasa,
Carboelectrica, Ciclo Combinado, Combustion Interna, Geotermoelectrica,
Hidroelectrica, Nucleoelectrica, Termica Convencional, Turbo Gas.
N is the settlement: L0 is published about 14 days after the month
ends, L1 ~6 weeks later, L2 ~4 months, L3 ~8 months. 2016-2017 months
also have partial L4-L6 files re-settling single days. Each day uses
the highest settlement that holds it. A multi-month range has been seen
to leave months out (Sep-Dec 2025 missing from a Jan 2025-Aug 2026 ZIP),
so any month a range doesn't return is requested on its own.

COVERAGE: from 2016-04-30 (the pickers' minDate). The latest month is
the pickers' maxDate: a month appears ~14 days after it ends, so data
runs 6-8 weeks behind.

ARCHIVES:
  - mexico_power_generation_daily.xlsx:
      'Daily' - date, Hydro/Gas/Wind/Solar/Coal/Nuclear/Oil/Bioenergy/
        Other/Total _MWh (the repo's standard layout, read by
        add_charts.power_daily and the North America master), Liquidacion
      'Daily by technology' - the same days in CENACE's own technologies
  - mexico_generacion_hourly/YYYY.csv.gz: the hourly rows, one file per
    year (only years with new or revised days are rewritten).

INCREMENTAL: each run fetches every month the archive doesn't hold in
full plus the trailing REVISION_MONTHS (later settlements replace earlier
ones; a day already at an equal or higher settlement is left alone).
Not reachable from the editing sandbox (CENACE blocks it) - run via
.github/workflows/mexico_cenace_generation.yml.
"""

import argparse
import csv
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

URL = "https://www.cenace.gob.mx/Paginas/SIM/Reportes/EnergiaGeneradaTipoTec.aspx"
PICKER_GEN = "ctl00_ContentPlaceHolder1_FechaConsulta"
PICKER_FI = "ctl00_ContentPlaceHolder1_FechaInicial"
PICKER_FF = "ctl00_ContentPlaceHolder1_FechaFinal"
DOWNLOAD_BUTTON = '[name="ctl00$ContentPlaceHolder1$DescargarReportes"]'

DATA_START = date(2016, 4, 1)   # first month in the report (minDate 2016-04-30)
CHUNK_MONTHS = 6                # months per range request
REVISION_MONTHS = 10            # L3 arrives ~8 months after the month
DOWNLOAD_TIMEOUT_MS = 90000     # a whole range downloads in ~1 s; no file by then means none is coming

REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT_DIR = os.path.join(REPO_ROOT, "output", "Data and Chart Outputs")
DEFAULT_XLSX_OUT = os.path.join(OUT_DIR, "mexico_power_generation_daily.xlsx")
DEFAULT_HOURLY_DIR = os.path.join(OUT_DIR, "mexico_generacion_hourly")

MESES = {"enero": 1, "febrero": 2, "marzo": 3, "abril": 4, "mayo": 5, "junio": 6, "julio": 7, "agosto": 8,
         "septiembre": 9, "setiembre": 9, "octubre": 10, "noviembre": 11, "diciembre": 12}
FILE_RE = re.compile(r"Generacion Liquidada_L(\d+) SEN ([a-z]+) (\d{4})", re.I)

# CENACE technology -> column in the standard layout
TECH = {
    "Hidroelectrica": "Hydro",
    "Ciclo Combinado": "Gas",
    "Turbo Gas": "Gas",
    "Eolica": "Wind",
    "Fotovoltaica": "Solar",
    "Carboelectrica": "Coal",
    "Nucleoelectrica": "Nuclear",
    "Termica Convencional": "Oil",   # steam units burning fuel oil and gas (see notes)
    "Combustion Interna": "Oil",     # internal combustion engines, mostly diesel / fuel oil
    "Biomasa": "Bioenergy",
    "Geotermoelectrica": "Other",
}
TECHS = list(TECH)
STD = ["Hydro", "Gas", "Wind", "Solar", "Coal", "Nuclear", "Oil", "Bioenergy", "Other"]
STD_COLUMNS = [f"{f}_MWh" for f in STD] + ["Total_MWh", "Liquidacion"]
TECH_COLUMNS = [f"{t}_MWh" for t in TECHS] + ["Total_MWh", "Liquidacion"]


def log(msg):
    print(msg, file=sys.stderr, flush=True)


def month_start(d):
    return date(d.year, d.month, 1)


def add_months(d, n):
    y, m = divmod(d.month - 1 + n, 12)
    return date(d.year + y, m + 1, 1)


def months_between(first, last):
    out, m = [], month_start(first)
    while m <= month_start(last):
        out.append(m)
        m = add_months(m, 1)
    return out


# ---------------------------------------------------------------- CENACE

def load_page(page):
    page.goto(URL, timeout=90000, wait_until="load")
    page.wait_for_selector(f"#{PICKER_GEN}", timeout=30000, state="attached")
    page.wait_for_timeout(1000)


def latest_published_month(page):
    ymd = page.evaluate(
        """(id) => { const p = window.$find(id); const m = p && p.get_maxDate();
            return m ? [m.getFullYear(), m.getMonth() + 1, m.getDate()] : null; }""",
        PICKER_FI,
    )
    if not ymd:
        raise RuntimeError("Could not read the month pickers' maxDate - page layout changed?")
    return month_start(date(*ymd))


def set_picker(page, picker_id, d):
    page.evaluate(
        """([id, y, m, d]) => { const p = window.$find(id);
            if (p) p.set_selectedDate(new Date(y, m, d)); }""",
        [picker_id, d.year, d.month - 1, d.day],
    )


def read_pickers(page):
    return page.evaluate(
        """(ids) => ids.map(id => { const p = window.$find(id); const s = p && p.get_selectedDate();
            return s ? [s.getFullYear(), s.getMonth() + 1] : null; })""",
        [PICKER_GEN, PICKER_FI, PICKER_FF],
    )


def download_months(page, first, last, attempts=2):
    """ZIP bytes for operating months first..last (month starts, inclusive), or None if CENACE sends no file."""
    want = [[last.year, last.month], [first.year, first.month], [last.year, last.month]]
    dialogs = []
    page.on("dialog", lambda dlg: (dialogs.append(dlg.message), dlg.dismiss()))
    for attempt in range(1, attempts + 1):
        load_page(page)
        for _ in range(2):  # second pass undoes the postback reset
            set_picker(page, PICKER_GEN, last)
            set_picker(page, PICKER_FI, first)
            set_picker(page, PICKER_FF, last)
            page.wait_for_timeout(2500)
        got = read_pickers(page)
        if got != want:
            log(f"  pickers read back {got}, wanted {want} (attempt {attempt}) - retrying")
            continue
        try:
            with page.expect_download(timeout=DOWNLOAD_TIMEOUT_MS) as dl:
                page.click(DOWNLOAD_BUTTON)
            with open(dl.value.path(), "rb") as fh:
                body = fh.read()
        except Exception as e:  # noqa: BLE001 - CENACE answered with the page instead of a file
            text = ""
            try:
                text = page.evaluate("() => document.body.innerText.replace(/\\s+/g, ' ')")
            except Exception:  # noqa: BLE001
                pass
            i = text.find("Mes de Operaci")
            log(f"  {first:%Y-%m}..{last:%Y-%m}: no download ({type(e).__name__}) (attempt {attempt}); "
                f"dialogs {dialogs[-3:]}; page: {text[max(0, i):i + 400]!r}")
            continue
        if body[:2] == b"PK":
            return body
        log(f"  {first:%Y-%m}..{last:%Y-%m}: download is not a ZIP ({len(body):,} bytes) (attempt {attempt})")
    return None


def parse_csv(text, month):
    """{day: DataFrame(hora + TECHS)} from one monthly CSV; every row must fall in `month`."""
    lines = text.splitlines()
    head = next((i for i, l in enumerate(lines) if "Sistema" in l and "Hora" in l), None)
    if head is None:
        raise RuntimeError("no header row with Sistema/Hora")
    reader = csv.reader(lines[head:])
    header = [h.strip() for h in next(reader)]
    missing = [t for t in ["Dia", "Hora", *TECHS] if t not in header]
    if missing:
        raise RuntimeError(f"columns missing: {missing} (header {header})")
    ix = {h: header.index(h) for h in ["Dia", "Hora", *TECHS]}
    rows = []
    for rec in reader:
        if len(rec) < len(header) or not re.fullmatch(r"\d{2}/\d{2}/\d{4}", rec[ix["Dia"]].strip()):
            continue
        dd, mm, yy = (int(x) for x in rec[ix["Dia"]].strip().split("/"))
        day = date(yy, mm, dd)
        if month_start(day) != month:
            raise RuntimeError(f"row dated {day} in the file for {month:%Y-%m}")
        rows.append([day, int(rec[ix["Hora"]])] + [float(rec[ix[t]] or 0) for t in TECHS])
    df = pd.DataFrame(rows, columns=["date", "hora", *TECHS])
    return {d: g.drop(columns="date").reset_index(drop=True) for d, g in df.groupby("date")}


def parse_zip(body, months):
    """{day: (liquidacion, hourly DataFrame)} using, per day, the highest settlement holding it.
    Files for months outside `months` are rejected."""
    best = {}
    seen_months = set()
    with zipfile.ZipFile(io.BytesIO(body)) as zf:
        for name in zf.namelist():
            m = FILE_RE.search(name)
            if not m:
                log(f"  skipping unrecognised file {name!r}")
                continue
            liq, mes, yy = int(m.group(1)), m.group(2).lower(), int(m.group(3))
            if mes not in MESES:
                raise RuntimeError(f"unknown month name in {name!r}")
            month = date(yy, MESES[mes], 1)
            if month not in months:
                raise RuntimeError(f"ZIP holds {name!r} - outside the requested months")
            seen_months.add(month)
            for day, df in parse_csv(zf.read(name).decode("latin-1", errors="replace"), month).items():
                if day not in best or liq > best[day][0]:
                    best[day] = (liq, df)
    return best, seen_months


# --------------------------------------------------------------- archive

def load_archive(path):
    try:
        df = pd.read_excel(path, sheet_name="Daily by technology", index_col=0)
    except (FileNotFoundError, ValueError):
        return pd.DataFrame(columns=TECH_COLUMNS)
    df.index = pd.to_datetime(df.index).date
    df.index.name = "date"
    return df


def day_row(liq, df):
    s = df[TECHS].sum()
    row = {f"{t}_MWh": round(s[t], 1) for t in TECHS}
    row["Total_MWh"] = round(s.sum(), 1)
    row["Liquidacion"] = liq
    return row


def standard_layout(tech):
    out = pd.DataFrame(index=tech.index)
    for f in STD:
        src = [f"{t}_MWh" for t, g in TECH.items() if g == f]
        out[f"{f}_MWh"] = tech[src].sum(axis=1, min_count=1).round(1)
    out["Total_MWh"] = tech["Total_MWh"]
    out["Liquidacion"] = tech["Liquidacion"]
    return out


def write_hourly_year(hourly_dir, year, day_frames):
    """Merge {day: (liq, DataFrame)} into hourly_dir/YYYY.csv.gz (replacing those days)."""
    os.makedirs(hourly_dir, exist_ok=True)
    path = os.path.join(hourly_dir, f"{year}.csv.gz")
    new = pd.concat([df.assign(date=d.isoformat(), liquidacion=liq) for d, (liq, df) in sorted(day_frames.items())])
    new = new[["date", "hora", "liquidacion", *TECHS]]
    if os.path.exists(path):
        old = pd.read_csv(path, dtype={"date": str})
        old = old[~old["date"].isin(new["date"].unique())]
        new = pd.concat([old, new]).sort_values(["date", "hora"], kind="stable")
    tmp = path + f".tmp{os.getpid()}"
    # mtime=0 keeps the gzip bytes identical when the data is unchanged.
    with open(tmp, "wb") as raw, gzip.GzipFile(fileobj=raw, mode="wb", mtime=0) as gz:
        gz.write(new.to_csv(index=False).encode("utf-8"))
    os.replace(tmp, path)


NOTES_LINES = [
    "UNITS",
    "MWh per day: the sum of CENACE's 24 hourly values (25 / 23 on clock-change days) of settled generation "
    "('energia generada liquidada') for the whole national electricity system (SEN: the interconnected SIN plus "
    "Baja California and Baja California Sur). Each hourly MWh value is also that hour's average MW.",
    "",
    "FUEL GROUPS ('Daily' sheet)",
    "Gas = Ciclo Combinado (combined cycle) + Turbo Gas (open-cycle gas turbines). Oil = Termica Convencional "
    "(conventional steam units, which burn fuel oil AND natural gas - CENACE does not split them by fuel, so Gas "
    "is a lower bound for gas-fired generation) + Combustion Interna (internal combustion engines, mostly diesel "
    "and fuel oil). Coal = Carboelectrica. Hydro = Hidroelectrica. Wind = Eolica. Solar = Fotovoltaica (utility "
    "scale; rooftop solar is not in CENACE's settlement). Nuclear = Nucleoelectrica (Laguna Verde). Bioenergy = "
    "Biomasa. Other = Geotermoelectrica (geothermal). 'Daily by technology' keeps CENACE's own categories.",
    "",
    "SETTLEMENT (Liquidacion)",
    "CENACE publishes each month several times: L0 about 14 days after the month ends, then L1 (~6 weeks later), "
    "L2 (~4 months) and L3 (~8 months); 2016-2017 also have partial L4-L6 files re-settling single days. Each day "
    "uses the highest settlement that includes it and is replaced when a later one appears.",
    "",
    "COVERAGE",
    "From 30 April 2016 (the first day CENACE's report offers). The latest month is the last one CENACE has "
    "published, so the data runs about 6-8 weeks behind the pull date. Monthly, not real-time: a whole month "
    "arrives at once.",
    "",
    "DETAIL FILES",
    "Hourly generation by technology (date, hora, liquidacion, one MWh column per technology) is in the "
    "mexico_generacion_hourly folder next to this workbook, one gzipped CSV per year.",
    "",
    "SOURCE",
    f"CENACE (Centro Nacional de Control de Energia), 'Energia Generada por Tipo de Tecnologia' (Generacion "
    f"Liquidada): {URL} - an ASP.NET report page automated with Playwright, downloading month ranges as ZIPs of "
    "monthly CSVs. Each file's month comes from its name and every row's date is checked against it.",
]
NOTES_SECTION_TITLES = {"UNITS", "FUEL GROUPS ('Daily' sheet)", "SETTLEMENT (Liquidacion)", "COVERAGE",
                        "DETAIL FILES", "SOURCE"}


def save(path, tech):
    tech = tech.sort_index()
    tech.index = pd.Index(tech.index, name="date")
    xlsx_notes.write_workbook(path, {"Daily": standard_layout(tech)[STD_COLUMNS],
                                     "Daily by technology": tech[TECH_COLUMNS]},
                              NOTES_LINES, NOTES_SECTION_TITLES)


def months_to_fetch(archive, latest, start):
    """Months the archive doesn't hold in full, plus the trailing REVISION_MONTHS."""
    have = pd.Series(1, index=pd.to_datetime(list(archive.index))) if len(archive) else pd.Series(dtype=float)
    counts = have.resample("MS").size() if len(have) else pd.Series(dtype=int)
    out = []
    for m in months_between(start, latest):
        first_day = max(m, date(2016, 4, 30))
        days = (add_months(m, 1) - first_day).days
        if counts.get(pd.Timestamp(m), 0) < days or m >= add_months(latest, -(REVISION_MONTHS - 1)):
            out.append(m)
    return out


def chunks(months):
    """Runs of consecutive months, each at most CHUNK_MONTHS long."""
    out, run = [], []
    for m in months:
        if run and (m != add_months(run[-1], 1) or len(run) == CHUNK_MONTHS):
            out.append(run)
            run = []
        run.append(m)
    if run:
        out.append(run)
    return out


def main():
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--xlsx-out", default=DEFAULT_XLSX_OUT)
    parser.add_argument("--hourly-dir", default=DEFAULT_HOURLY_DIR)
    args = parser.parse_args()

    archive = load_archive(args.xlsx_out)
    updated, absent = 0, []

    with sync_playwright() as p:
        browser = p.chromium.launch()
        context = browser.new_context(accept_downloads=True)
        page = context.new_page()
        load_page(page)
        latest = latest_published_month(page)
        todo = months_to_fetch(archive, latest, DATA_START)
        log(f"CENACE latest published month: {latest:%Y-%m}. Months to fetch: {len(todo)} "
            f"({todo[0]:%Y-%m}..{todo[-1]:%Y-%m})" if todo else f"CENACE latest month {latest:%Y-%m}; nothing to fetch.")

        for run in chunks(todo):
            body = download_months(page, run[0], run[-1])
            days, seen = parse_zip(body, set(run)) if body else ({}, set())
            for m in [m for m in run if m not in seen]:
                log(f"  {m:%Y-%m} not in the {run[0]:%Y-%m}..{run[-1]:%Y-%m} download - requesting it on its own")
                b2 = download_months(page, m, m)
                d2, s2 = parse_zip(b2, {m}) if b2 else ({}, set())
                days.update(d2)
                if m not in s2:
                    absent.append(m)
            changed = {}
            for day, (liq, df) in days.items():
                if day in archive.index and pd.notna(archive.at[day, "Liquidacion"]) \
                        and int(archive.at[day, "Liquidacion"]) >= liq:
                    continue
                archive.loc[day] = pd.Series(day_row(liq, df))
                changed[day] = (liq, df)
            if changed:
                by_year = {}
                for day, v in changed.items():
                    by_year.setdefault(day.year, {})[day] = v
                for year, frames in sorted(by_year.items()):
                    write_hourly_year(args.hourly_dir, year, frames)
                save(args.xlsx_out, archive)
                updated += len(changed)
            log(f"  {run[0]:%Y-%m}..{run[-1]:%Y-%m}: {len(days)} day(s) returned, {len(changed)} new/revised")

        browser.close()

    if updated:
        print(f"Saved {updated} new/revised day(s). Archive: {len(archive)} days, "
              f"{min(archive.index)} to {max(archive.index)} -> {args.xlsx_out}")
    else:
        print("No new or revised days this run.")
    if absent:   # everything else is saved; fail the job so the gap is seen
        sys.exit(f"CENACE returned no file for {', '.join(f'{m:%Y-%m}' for m in absent)}")


if __name__ == "__main__":
    main()
