"""
mexico_demanda_real.py's first real backfill (2016-01-27 to 2016-02-04)
saved every day with Total_MWh = 0, despite thousands of zone-hour rows
each day - either a genuine data-quality artifact of this very early
period right after CENACE's market went live, or a parsing bug in how
the script reads the CSV. Fetching one of those days (2016-01-29) again
and printing the RAW CSV text plus the parsed DataFrame side by side to
tell which one it is.
"""
import io
import os
import sys
import zipfile

from playwright.sync_api import sync_playwright

URL = "https://www.cenace.gob.mx/Paginas/SIM/Reportes/EstimacionDemandaReal.aspx"
TARGET = (2016, 1, 29)

DATE_PICKER_IDS = [
    "ctl00_ContentPlaceHolder1_RadDatePickerVisualizarGeneralesPorRetiros",
    "ctl00_ContentPlaceHolder1_RadDatePickerFIVisualizarPorRetiros",
    "ctl00_ContentPlaceHolder1_RadDatePickerFFVisualizarPorRetiros",
]
DOWNLOAD_BUTTON_NAME = "ctl00$ContentPlaceHolder1$DescargarArchivosCsv_PorRetiros"
DOWNLOAD_BUTTON_VALUE = "Descargar en archivo .zip"

DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_PATH = os.path.join(DIR, "mexico_demanda_zero_check_output.txt")


def log(text=""):
    print(text, flush=True)
    with open(OUTPUT_PATH, "a", encoding="utf-8") as f:
        f.write(text + "\n")


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


def main():
    open(OUTPUT_PATH, "w").close()
    year, month, day = TARGET
    with sync_playwright() as p:
        browser = p.chromium.launch()
        context = browser.new_context()
        page = context.new_page()
        page.goto(URL, timeout=60000, wait_until="load")
        page.wait_for_selector(f"#{DATE_PICKER_IDS[0]}", timeout=30000, state="attached")

        set_all_dates(page, year, month, day)
        page.wait_for_timeout(2500)
        set_all_dates(page, year, month, day)
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
        log(f"status={resp.status}, is_zip={body[:2] == b'PK'}")

        with zipfile.ZipFile(io.BytesIO(body)) as zf:
            names = zf.namelist()
            log(f"zip names: {names}")
            for name in names:
                with zf.open(name) as f:
                    text = f.read().decode("latin-1", errors="replace")
                log(f"\n=== {name} - RAW first 2000 chars ===")
                log(text[:2000])
                log(f"\n=== {name} - RAW last 500 chars ===")
                log(text[-500:])

        browser.close()
    log(f"\nSaved: {OUTPUT_PATH}")


if __name__ == "__main__":
    main()
