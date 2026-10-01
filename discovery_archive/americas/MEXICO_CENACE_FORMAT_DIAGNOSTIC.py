"""
Second Mexico CENACE diagnostic (Oct 2026), after
MEXICO_CENACE_DATE_DIAGNOSTIC.py showed:
  - an FI..FF date range downloads EVERY day in the range in one ZIP,
    one CSV per day per settlement ("Demanda Real Retiro_<N> Dia
    Operacion YYYY-MM-DD v<publication stamp>.csv");
  - the newer CSVs (2018+?) and the 2017 Retiro_4 / Retiro_L5 / L6
    files are in a different (quoted / "Estimacion de la Demanda Real")
    layout that the production parser never recognised;
  - the pickers' maxDate is ~15 days behind today.

This dumps the raw head of one file of each layout, plus the largest
range one POST will return, so the production parser can be written
against the real formats.

FINDINGS (run 2026-10-01): three layouts - plain "BCA,ENSENADA,1,115.8,"
(2016-2017), fully quoted "\"BCA\",\"ENSENADA\",\"1\",\"117.8\"" (2017 L5/L6,
2018 R0-R2), and quoted with broken preamble lines (2018-08 on). Each
file carries "LIQUIDACION N (Dia de Operacion: dd/mm/yyyy)". 2016-01-27/28
hold only BCA (4 zones); SIN starts 2016-01-29. A 90-day range = 8.4 MB,
a full year = 34 MB / ~20 s, both returned complete.
"""

import io
import re
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


def set_picker(page, pid, d):
    return page.evaluate(
        """([id, y, m, d]) => { const p = window.$find(id); if (!p) return 'NOT FOUND';
            p.set_selectedDate(new Date(y, m, d)); const s = p.get_selectedDate();
            return s ? s.toDateString() : null; }""",
        [pid, d.year, d.month - 1, d.day],
    )


def fetch_range(page, context, a, b):
    page.goto(URL, timeout=90000, wait_until="load")
    page.wait_for_selector(f"#{GEN}", timeout=30000, state="attached")
    for _ in range(2):
        set_picker(page, GEN, a)
        set_picker(page, FI, a)
        set_picker(page, FF, b)
        page.wait_for_timeout(2500)
    fields = page.evaluate(
        """() => { const data = {}; for (const el of document.forms[0].elements) {
            if (el.name) data[el.name] = el.value; } return data; }""")
    print(f"  posted FI={fields.get('ctl00$ContentPlaceHolder1$RadDatePickerFIVisualizarPorRetiros')} "
          f"FF={fields.get('ctl00$ContentPlaceHolder1$RadDatePickerFFVisualizarPorRetiros')}", flush=True)
    fields[BUTTON[0]] = BUTTON[1]
    resp = context.request.post(URL, form=fields, timeout=180000)
    body = resp.body()
    print(f"  HTTP {resp.status} {len(body):,} B {resp.headers.get('content-type')}", flush=True)
    return body


def main():
    with sync_playwright() as p:
        browser = p.chromium.launch()
        context = browser.new_context()
        page = context.new_page()

        seen = set()
        for a, b in [(date(2016, 1, 27), date(2016, 1, 30)), (date(2017, 3, 1), date(2017, 3, 1)),
                     (date(2017, 12, 30), date(2018, 1, 2)), (date(2021, 1, 4), date(2021, 1, 4)),
                     (date(2026, 9, 14), date(2026, 9, 16))]:
            print(f"\n=== {a} .. {b} ===", flush=True)
            body = fetch_range(page, context, a, b)
            if body[:2] != b"PK":
                print("  not a zip", flush=True)
                continue
            with zipfile.ZipFile(io.BytesIO(body)) as zf:
                for n in zf.namelist():
                    raw = zf.read(n)
                    text = raw.decode("latin-1", errors="replace")
                    lines = text.splitlines()
                    kind = re.sub(r"Dia Operacion \d{4}-\d{2}-\d{2} v.*", "", n)
                    key = (kind, text[:60])
                    print(f"  {n}  {len(raw):,} B, {len(lines)} lines", flush=True)
                    if key in seen:
                        continue
                    seen.add(key)
                    print("  ---- head (repr) ----")
                    for l in lines[:14]:
                        print(f"   {l[:180]!r}")
                    print("  ---- tail (repr) ----")
                    for l in lines[-3:]:
                        print(f"   {l[:180]!r}")

        for a, b in [(date(2023, 1, 1), date(2023, 3, 31)), (date(2022, 1, 1), date(2022, 12, 31))]:
            print(f"\n=== range size test {a} .. {b} ===", flush=True)
            try:
                body = fetch_range(page, context, a, b)
                if body[:2] == b"PK":
                    with zipfile.ZipFile(io.BytesIO(body)) as zf:
                        names = zf.namelist()
                        days = sorted({re.search(r"Dia Operacion (\d{4}-\d{2}-\d{2})", n).group(1) for n in names})
                        print(f"  {len(names)} files, {len(days)} days {days[0]}..{days[-1]}", flush=True)
            except Exception as e:
                print(f"  FAILED {type(e).__name__}: {e}", flush=True)
        browser.close()


if __name__ == "__main__":
    main()
