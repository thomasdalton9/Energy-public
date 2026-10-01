"""
Diagnostic for mexico_demanda_real.py's broken backfill (Oct 2026).

Symptoms in the committed archive: only 327 days (2016-01-27 to
2017-02-26) despite the checkpoint reaching 2026-09-29; ~all days after
Feb 2017 logged "no data"; and 19 days in 2016 carry an exact copy of
the neighbouring day's hourly data (e.g. 2016-02-14/15/16 identical),
i.e. the file CENACE sent was NOT for the day it was saved under.

This checks, per requested day:
  - what the three Telerik pickers and their posted form fields really
    hold just before the POST (were they reset by the async postback?)
  - what comes back: ZIP names (do they embed the requested date?),
    the CSV preamble, or, for a non-ZIP answer, the page's message
  - single-day vs a fresh page per day vs a multi-day FI..FF range.

FINDINGS (run 2026-10-01): pickers read back correctly after the second
pass, and each ZIP's file names embed the operating day; but production
had still saved 2016-02-15/16 with 2016-02-14's file (Retiro_3 total
591,920 MWh), so the race is intermittent and the date must be validated
from the file name. From 2017-03 the ZIPs also hold Retiro_4/L5/L6 files
in an "Estimacion de Demanda por Retiros" (later fully quoted) layout
that production's parser didn't recognise -> every later day "no data".
The pickers' maxDate is ~15 days behind today (2026-09-16 on 2026-10-01).
An FI..FF range returns every day in the range in one ZIP.

Workflow (archived): discovery_archive/workflows/mexico_cenace_date_diagnostic.yml.
"""

import io
import re
import sys
import zipfile
from datetime import date

from playwright.sync_api import sync_playwright

URL = "https://www.cenace.gob.mx/Paginas/SIM/Reportes/EstimacionDemandaReal.aspx"
GEN, FI, FF = (
    "ctl00_ContentPlaceHolder1_RadDatePickerVisualizarGeneralesPorRetiros",
    "ctl00_ContentPlaceHolder1_RadDatePickerFIVisualizarPorRetiros",
    "ctl00_ContentPlaceHolder1_RadDatePickerFFVisualizarPorRetiros",
)
BUTTON = ("ctl00$ContentPlaceHolder1$DescargarArchivosCsv_PorRetiros", "Descargar en archivo .zip")


def log(*a):
    print(*a, flush=True)


def set_picker(page, pid, d):
    return page.evaluate(
        """([id, y, m, d]) => {
            const p = window.$find(id);
            if (!p) return 'NOT FOUND';
            p.set_selectedDate(new Date(y, m, d));
            const s = p.get_selectedDate();
            return s ? s.toDateString() : null;
        }""",
        [pid, d.year, d.month - 1, d.day],
    )


def read_pickers(page):
    return page.evaluate(
        """(ids) => ids.map(id => { const p = window.$find(id);
            if (!p) return 'NOT FOUND'; const s = p.get_selectedDate();
            return s ? s.toDateString() : null; })""",
        [GEN, FI, FF],
    )


def form_fields(page):
    return page.evaluate(
        """() => { const data = {}; for (const el of document.forms[0].elements) {
            if (el.name) data[el.name] = el.value; } return data; }"""
    )


def text_of(html):
    t = re.sub(r"(?is)<(script|style)[^>]*>.*?</\1>", " ", html)
    t = re.sub(r"(?s)<[^>]+>", " ", t)
    return re.sub(r"\s+", " ", t)


def post_and_report(page, context, label):
    fields = form_fields(page)
    for k, v in fields.items():
        if "RadDatePicker" in k and ("dateInput" in k or k.endswith("PorRetiros")) and "ClientState" not in k:
            log(f"    field {k.split('$')[-2] if k.count('$') > 2 else k} = {v!r}")
        if "RadDatePicker" in k and k.endswith("ClientState"):
            log(f"    field {k} = {v[:160]!r}")
    fields[BUTTON[0]] = BUTTON[1]
    for k in list(fields):
        if k.endswith((".x", ".y")):
            fields.pop(k)
    resp = context.request.post(URL, form=fields, timeout=60000)
    body = resp.body()
    log(f"    -> HTTP {resp.status}, {len(body):,} B, type={resp.headers.get('content-type')}, "
        f"disp={resp.headers.get('content-disposition')}")
    if body[:2] == b"PK":
        with zipfile.ZipFile(io.BytesIO(body)) as zf:
            for n in zf.namelist():
                raw = zf.read(n).decode("latin-1", errors="replace")
                lines = raw.splitlines()
                hdr = next((i for i, l in enumerate(lines) if "Energia" in l and "Hora" in l), None)
                rows = lines[hdr + 1:] if hdr is not None else []
                tot = 0.0
                for r in rows:
                    parts = r.split(",")
                    try:
                        tot += float(parts[3])
                    except (IndexError, ValueError):
                        pass
                log(f"    zip: {n!r}  rows={len(rows)} total={tot:,.0f}")
                for l in lines[: (hdr or 0) + 2][:10]:
                    log(f"       | {l[:150]}")
    else:
        html = body.decode("utf-8", errors="replace")
        alerts = re.findall(r"alert\(([^)]{0,300})\)", html)
        log(f"    non-zip. alerts={alerts[:3]}")
        t = text_of(html)
        for kw in ("No se", "no existe", "No hay", "Error", "error", "fecha", "Fecha", "rango"):
            i = t.find(kw)
            if i >= 0:
                log(f"    text[{kw}]: ...{t[max(0, i - 120): i + 200]}...")
        hidden = [m for m in re.findall(r'name="([^"]+)"', html) if m.startswith("__")][:8]
        log(f"    hidden fields in reply: {hidden}; title={re.findall(r'<title>([^<]*)', html)[:1]}")


def set_day_two_pass(page, d, order=(GEN, FI, FF)):
    for pid in order:
        set_picker(page, pid, d)
    page.wait_for_timeout(2500)
    after1 = read_pickers(page)
    for pid in order:
        set_picker(page, pid, d)
    page.wait_for_timeout(1000)
    after2 = read_pickers(page)
    log(f"    pickers after pass1 wait: {after1}")
    log(f"    pickers after pass2:      {after2}")


def main():
    days = [date(2016, 2, 14), date(2016, 2, 15), date(2016, 2, 16), date(2017, 3, 1),
            date(2017, 3, 2), date(2021, 1, 4), date(2024, 6, 3), date(2026, 9, 29)]
    with sync_playwright() as p:
        browser = p.chromium.launch()
        context = browser.new_context()
        page = context.new_page()
        posts = []
        page.on("request", lambda r: posts.append(r.method) if r.method == "POST" else None)

        page.goto(URL, timeout=90000, wait_until="load")
        page.wait_for_selector(f"#{GEN}", timeout=30000, state="attached")
        log(f"Loaded. Default pickers: {read_pickers(page)}")
        mins = page.evaluate(
            """(ids) => ids.map(id => { const p = window.$find(id); if (!p) return null;
                return [String(p.get_minDate && p.get_minDate()), String(p.get_maxDate && p.get_maxDate())]; })""",
            [GEN, FI, FF])
        log(f"min/max dates: {mins}")

        log("\n=== A: reused session, production's two-pass method ===")
        for d in days:
            posts.clear()
            log(f"\n  [{d}]")
            try:
                set_day_two_pass(page, d)
                log(f"    page POSTs during set: {len(posts)}")
                post_and_report(page, context, str(d))
            except Exception as e:
                log(f"    FAILED {type(e).__name__}: {e}")

        log("\n=== B: reused session, set FI/FF first, then GEN, wait for postbacks ===")
        for d in days[3:6]:
            posts.clear()
            log(f"\n  [{d}]")
            try:
                set_day_two_pass(page, d, order=(FI, FF, GEN))
                post_and_report(page, context, str(d))
            except Exception as e:
                log(f"    FAILED {type(e).__name__}: {e}")

        log("\n=== C: fresh page.goto per day ===")
        for d in [date(2017, 3, 1), date(2021, 1, 4), date(2016, 2, 15)]:
            log(f"\n  [{d}]")
            try:
                page.goto(URL, timeout=90000, wait_until="load")
                page.wait_for_selector(f"#{GEN}", timeout=30000, state="attached")
                set_day_two_pass(page, d)
                post_and_report(page, context, str(d))
            except Exception as e:
                log(f"    FAILED {type(e).__name__}: {e}")

        log("\n=== D: multi-day range FI..FF (one POST) ===")
        for a, b in [(date(2024, 1, 1), date(2024, 1, 7)), (date(2024, 1, 1), date(2024, 1, 31))]:
            log(f"\n  [{a} .. {b}]")
            try:
                page.goto(URL, timeout=90000, wait_until="load")
                page.wait_for_selector(f"#{GEN}", timeout=30000, state="attached")
                for _ in range(2):
                    set_picker(page, GEN, a)
                    set_picker(page, FI, a)
                    set_picker(page, FF, b)
                    page.wait_for_timeout(2500)
                log(f"    pickers: {read_pickers(page)}")
                post_and_report(page, context, f"{a}..{b}")
            except Exception as e:
                log(f"    FAILED {type(e).__name__}: {e}")

        browser.close()


if __name__ == "__main__":
    sys.exit(main())
