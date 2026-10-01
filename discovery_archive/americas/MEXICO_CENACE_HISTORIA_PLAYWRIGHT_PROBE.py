"""
Follow-up to MEXICO_CENACE_HISTORIA_POST_PROBE.py and this script's
own earlier revisions.

Progress so far:
  - four raw-POST attempts (plain text-field overrides) always got
    back the page's default day - Telerik's RadDatePicker doesn't
    read its posted value from the input's plain value attribute.
  - calling Telerik's own client API (window.$find(id).
    set_selectedDate(...)) via Playwright DOES work - confirmed by
    reading the date back from the widget afterwards.
  - but Locator.click() on either report mode's "Descargar en archivo
    .zip" button times out (30s) - some actionability check (visible/
    stable/not-obscured) never passes, on both PorBalance and
    PorRetiros, even though the underlying ASP.NET postback handler
    for PorRetiros is proven to work (a plain requests.post() with its
    name/value in the body reliably returns a real ZIP).

v3 (set once, read once, POST) DID download a real ZIP/CSV - but for
today's default operating day, not the requested one. Reading the live
DOM back showed why: only the first picker
(RadDatePickerVisualizarGeneralesPorRetiros) actually kept the set
date; FI/FF (the from/to range pickers) had reverted to today's date
by the time the form was read, seconds after being set. The likely
cause is a postback/UpdatePanel refresh triggered by setting the first
picker, which re-renders FI/FF with their server-side defaults and
clobbers the client-side value just set on them - a timing issue, not
a wrong API call (set_selectedDate reported success on all three at
the moment it ran).

v4 fixes this with a second pass: re-applies all three dates after a
longer settle delay, immediately before reading the form and POSTing -
so even if an early postback resets FI/FF once, the second pass
corrects it right before it matters, leaving no window for another
reset to slip in.

v4's first live run failed before ever reaching the date-setting code -
page.goto(..., wait_until="networkidle") timed out at 60s, almost
certainly because this page has some periodic background request (an
UpdatePanel poll or similar) that keeps the network from ever going
fully idle. v5 swaps that for wait_until="load" plus an explicit wait
for the first date picker's DOM element to appear, which doesn't
depend on the network ever going quiet.

v5's first live run got past goto() but then failed at that same
wait_for_selector - the element is there (confirmed by the timeout
error itself showing its live value) but is intentionally hidden,
since Telerik's RadDatePicker keeps its real <input> invisible behind
a decorated widget. wait_for_selector's default state ("visible")
waits forever on it. v6 waits for state="attached" instead - present
in the DOM is all that's actually needed before calling
window.$find() on it.

This combines both working parts instead of clicking anything: sets
the date via Telerik's client API (proven to work), then reads back
every field on the live, now-correctly-synced form via JS - including
whatever ClientState/hidden fields Telerik's own JS updated when the
date was set - adds the download button's name/value manually (since
WebForms identifies "which button was clicked" purely from that
name/value pair being present in the POST body, not from an actual
click event), and submits it as an HTTP POST through Playwright's own
request context (which shares the page's cookies), never calling
click() at all.

Not reachable from this repo's normal editing sandbox - run via GitHub
Actions, see .github/workflows/mexico_cenace_historia_playwright_probe.yml.
"""

import os
import zipfile

from playwright.sync_api import sync_playwright

URL = "https://www.cenace.gob.mx/Paginas/SIM/Reportes/EstimacionDemandaReal.aspx"
TEST_YEAR, TEST_MONTH, TEST_DAY = 2016, 1, 27  # TEST_MONTH is 1-indexed; converted below

DATE_PICKER_IDS = [
    "ctl00_ContentPlaceHolder1_RadDatePickerVisualizarGeneralesPorRetiros",
    "ctl00_ContentPlaceHolder1_RadDatePickerFIVisualizarPorRetiros",
    "ctl00_ContentPlaceHolder1_RadDatePickerFFVisualizarPorRetiros",
]
DOWNLOAD_BUTTON_NAME = "ctl00$ContentPlaceHolder1$DescargarArchivosCsv_PorRetiros"
DOWNLOAD_BUTTON_VALUE = "Descargar en archivo .zip"

DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_PATH = os.path.join(DIR, "mexico_historia_playwright_probe_output.txt")


def log(text=""):
    print(text, flush=True)
    with open(OUTPUT_PATH, "a", encoding="utf-8") as f:
        f.write(text + "\n")


def set_all_dates(page, label):
    for picker_id in DATE_PICKER_IDS:
        result = page.evaluate(
            """(id) => {
                const picker = window.$find(id);
                if (!picker) return {found: false};
                picker.set_selectedDate(new Date(%d, %d, %d));
                return {found: true, value: picker.get_selectedDate()
                    ? picker.get_selectedDate().toString() : null};
            }"""
            % (TEST_YEAR, TEST_MONTH - 1, TEST_DAY),
            picker_id,
        )
        log(f"  [{label}] {picker_id}: {result}")


def main():
    open(OUTPUT_PATH, "w").close()
    log("CENACE EstimacionDemandaReal.aspx - PLAYWRIGHT DOWNLOAD PROBE (v4)")
    log(f"Set date via Telerik API (two passes, to survive a postback reset), "
        f"then POST the live form state directly (no click()): "
        f"{TEST_YEAR:04d}-{TEST_MONTH:02d}-{TEST_DAY:02d}\n")

    with sync_playwright() as p:
        browser = p.chromium.launch()
        context = browser.new_context()
        page = context.new_page()

        log("Step 1: navigating...")
        # v4's first run timed out waiting for "networkidle" - CENACE's page
        # likely has some periodic background request (a UpdatePanel poll or
        # similar) that never lets the network go fully idle. "load" plus an
        # explicit wait for the first date picker's DOM element is a more
        # reliable readiness signal for this kind of page.
        page.goto(URL, timeout=60000, wait_until="load")
        # v5's first live run proved the element exists at this point but is
        # intentionally hidden (Telerik's RadDatePicker keeps its real
        # <input> visually hidden behind a decorated widget) - the default
        # wait_for_selector state ("visible") fails on it forever. "attached"
        # (present in the DOM, visible or not) is what's actually needed
        # before calling window.$find() on it.
        page.wait_for_selector(f"#{DATE_PICKER_IDS[0]}", timeout=30000, state="attached")
        log(f"  loaded, title: {page.title()!r}")

        log("Step 1a: first pass - setting all three dates...")
        set_all_dates(page, "pass 1")

        # v3 found that setting the first picker triggers a postback that
        # resets FI/FF back to their server-side defaults a moment later -
        # give that postback time to actually happen before the corrective
        # second pass, or the second pass could itself get clobbered too.
        page.wait_for_timeout(2500)

        log("Step 1b: second pass - re-applying all three dates "
            "(correcting any postback reset from pass 1)...")
        set_all_dates(page, "pass 2")
        page.wait_for_timeout(1000)

        log("\nStep 2: reading back the live form state...")
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
        log(f"  read {len(fields)} field(s) from the live DOM")
        for picker_id in DATE_PICKER_IDS:
            dotted = picker_id.replace("ctl00_ContentPlaceHolder1_", "ctl00$ContentPlaceHolder1$")
            for key in (dotted, f"{dotted}$dateInput"):
                if key in fields:
                    log(f"    {key} = {fields[key]!r}")

        fields[DOWNLOAD_BUTTON_NAME] = DOWNLOAD_BUTTON_VALUE
        for noise in list(fields):
            if noise.endswith((".x", ".y")):
                fields.pop(noise, None)

        log("\nStep 3: POSTing the live form state via Playwright's request context...")
        resp = context.request.post(URL, form=fields, timeout=30000)
        body = resp.body()
        log(f"  HTTP {resp.status}, {len(body):,} bytes")
        log(f"  Content-Type: {resp.headers.get('content-type')}")
        log(f"  Content-Disposition: {resp.headers.get('content-disposition')}")

        is_zip = body[:2] == b"PK"
        log(f"  looks like a ZIP: {is_zip}")

        if is_zip:
            out_path = os.path.join(DIR, "mexico_historia_playwright_test.zip")
            open(out_path, "wb").write(body)
            log(f"  saved -> {out_path}")
            with zipfile.ZipFile(out_path) as zf:
                names = zf.namelist()
                log(f"\n  zip contains {len(names)} file(s):")
                for name in names:
                    log(f"    {name} ({zf.getinfo(name).file_size:,} bytes)")
                if names:
                    with zf.open(names[0]) as f:
                        preview = f.read(2000).decode("latin-1", errors="replace")
                    log(f"\n  preview of {names[0]!r}:")
                    for line in preview.splitlines()[:15]:
                        log(f"    {line}")
        else:
            out_path = os.path.join(DIR, "mexico_historia_post_response.html")
            open(out_path, "wb").write(body)
            log(f"  not a ZIP - saved response -> {out_path}")

        browser.close()

    log(f"\nSaved: {OUTPUT_PATH}")


if __name__ == "__main__":
    main()
