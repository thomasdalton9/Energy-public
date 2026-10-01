"""
Follow-up to MEXICO_CENACE_HISTORIA_PLAYWRIGHT_PROBE.py v6, which
finally proved the whole download mechanism end-to-end for ONE
historical day (2016-01-27) - confirmed by the real ZIP's own filename
embedding that exact requested date, not today's.

Before building a real backfill script, checking whether the same
browser page/context can be reused to pull SEVERAL different days in
one run (navigate once, then loop: set dates, read form, POST, repeat)
- much faster than relaunching the browser and reloading the page per
day. The open question is whether the page's own state (Telerik's
ClientState, ASP.NET __VIEWSTATE, etc.) stays valid/settable after a
plain context.request.post() that never actually navigates the page
itself - if not, each day might need its own fresh page.goto().

Testing 3 non-consecutive historical days in a single page session.
"""

import os
import zipfile
import io

from playwright.sync_api import sync_playwright

URL = "https://www.cenace.gob.mx/Paginas/SIM/Reportes/EstimacionDemandaReal.aspx"
TEST_DATES = [(2016, 1, 27), (2018, 6, 15), (2022, 11, 3)]

DATE_PICKER_IDS = [
    "ctl00_ContentPlaceHolder1_RadDatePickerVisualizarGeneralesPorRetiros",
    "ctl00_ContentPlaceHolder1_RadDatePickerFIVisualizarPorRetiros",
    "ctl00_ContentPlaceHolder1_RadDatePickerFFVisualizarPorRetiros",
]
DOWNLOAD_BUTTON_NAME = "ctl00$ContentPlaceHolder1$DescargarArchivosCsv_PorRetiros"
DOWNLOAD_BUTTON_VALUE = "Descargar en archivo .zip"

DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_PATH = os.path.join(DIR, "mexico_multiday_test_output.txt")


def log(text=""):
    print(text, flush=True)
    with open(OUTPUT_PATH, "a", encoding="utf-8") as f:
        f.write(text + "\n")


def set_all_dates(page, year, month, day, label):
    for picker_id in DATE_PICKER_IDS:
        result = page.evaluate(
            """(id) => {
                const picker = window.$find(id);
                if (!picker) return {found: false};
                picker.set_selectedDate(new Date(%d, %d, %d));
                return {found: true, value: picker.get_selectedDate()
                    ? picker.get_selectedDate().toString() : null};
            }""" % (year, month - 1, day),
            picker_id,
        )
        log(f"  [{label}] {picker_id}: {result}")


def fetch_one_day(page, context, year, month, day):
    set_all_dates(page, year, month, day, "pass 1")
    page.wait_for_timeout(2500)
    set_all_dates(page, year, month, day, "pass 2")
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
    is_zip = body[:2] == b"PK"
    log(f"  HTTP {resp.status}, {len(body):,} bytes, is_zip={is_zip}, "
        f"content-disposition={resp.headers.get('content-disposition')}")
    if is_zip:
        with zipfile.ZipFile(io.BytesIO(body)) as zf:
            names = zf.namelist()
            log(f"  zip contains: {names}")
            if names:
                with zf.open(names[0]) as f:
                    preview = f.read(500).decode("latin-1", errors="replace")
                log(f"  preview: {preview[:300]}")
    return is_zip


def main():
    open(OUTPUT_PATH, "w").close()
    log("Mexico CENACE multi-day-in-one-session test\n")

    with sync_playwright() as p:
        browser = p.chromium.launch()
        context = browser.new_context()
        page = context.new_page()

        log("Navigating once...")
        page.goto(URL, timeout=60000, wait_until="load")
        page.wait_for_selector(f"#{DATE_PICKER_IDS[0]}", timeout=30000, state="attached")
        log(f"  loaded, title: {page.title()!r}\n")

        results = []
        for year, month, day in TEST_DATES:
            log(f"=== Day {year:04d}-{month:02d}-{day:02d} ===")
            try:
                ok = fetch_one_day(page, context, year, month, day)
            except Exception as e:
                log(f"  FAILED: {type(e).__name__}: {e}")
                ok = False
            results.append((year, month, day, ok))
            log("")

        browser.close()

    log("=== Summary ===")
    for year, month, day, ok in results:
        log(f"  {year:04d}-{month:02d}-{day:02d}: {'OK' if ok else 'FAILED'}")
    log(f"\nSaved: {OUTPUT_PATH}")


if __name__ == "__main__":
    main()
