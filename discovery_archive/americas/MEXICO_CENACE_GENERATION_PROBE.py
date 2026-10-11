"""
Probe for a Mexico power-generation-by-technology pull from CENACE (Oct 2026).

Looks for CENACE's generation report(s) ("Energia Generada por Tipo de
Tecnologia" and any other generation report linked from the SIM report
menu), and for each page that loads:
  - lists every Telerik RadDatePicker (id, selected/min/max date) and
    every download button (name/value);
  - pairs FI/FF pickers with their button (same id suffix), requests a
    short range near the latest published day and one in 2018 (does the
    report keep history?), and prints the ZIP's file names plus the head
    of one file per settlement.

Same mechanism as americas/mexico_demanda_real.py (pickers set through
window.$find(...).set_selectedDate twice, then the live form POSTed with
the button's name/value).

Workflow (manual): discovery_archive/workflows/mexico_cenace_generation_probe.yml.
"""

import io
import re
import zipfile
from datetime import date, timedelta

from playwright.sync_api import sync_playwright

BASE = "https://www.cenace.gob.mx/Paginas/SIM/Reportes/"
CANDIDATES = [
    "EnergiaGeneradaTipoTec.aspx",
    "EnergiaGeneradaTipoTecnologia.aspx",
    "GeneracionLiquidada.aspx",
    "EnergiaGenerada.aspx",
]
MENU_PAGE = BASE + "EstimacionDemandaReal.aspx"


def log(*a):
    print(*a, flush=True)


def menu_links(page):
    page.goto(MENU_PAGE, timeout=90000, wait_until="load")
    links = page.evaluate("""() => Array.from(document.querySelectorAll('a[href]'))
        .map(a => [a.textContent.trim().replace(/\\s+/g, ' '), a.href])""")
    keep = [(t, h) for t, h in links if re.search(r"gener|tecnolog", t + h, re.I)]
    log(f"Menu links mentioning generation/technology ({len(keep)} of {len(links)}):")
    for t, h in keep:
        log(f"  {t!r} -> {h}")
    return [h for _, h in keep if h.lower().endswith(".aspx")]


def describe(page):
    pickers = page.evaluate("""() => {
        const out = [];
        if (!window.Sys || !Sys.Application) return out;
        for (const c of Sys.Application.getComponents()) {
            if (typeof c.get_selectedDate !== 'function' || typeof c.get_maxDate !== 'function') continue;
            const f = d => d ? d.toISOString().slice(0, 10) : null;
            out.push([c.get_id(), f(c.get_selectedDate()), f(c.get_minDate()), f(c.get_maxDate())]);
        }
        return out; }""")
    buttons = page.evaluate("""() => Array.from(document.querySelectorAll('input[type=submit],input[type=button],button'))
        .map(b => [b.name || b.id, b.value || b.textContent.trim()])""")
    headings = page.evaluate("""() => Array.from(document.querySelectorAll('h1,h2,h3,h4,legend,.titulo,span[id*=Titulo]'))
        .map(e => e.textContent.trim().replace(/\\s+/g, ' ')).filter(t => t).slice(0, 20)""")
    log(f"  headings: {headings}")
    log(f"  pickers ({len(pickers)}):")
    for p in pickers:
        log(f"    {p}")
    log(f"  buttons ({len(buttons)}):")
    for b in buttons:
        log(f"    {b}")
    return pickers, buttons


def set_picker(page, pid, d):
    page.evaluate("""([id, y, m, d]) => { const p = window.$find(id);
        if (p) p.set_selectedDate(new Date(y, m, d)); }""", [pid, d.year, d.month - 1, d.day])


def try_download(page, context, url, group, button, first, last):
    page.goto(url, timeout=90000, wait_until="load")
    for _ in range(2):
        for pid in group["gen"]:
            set_picker(page, pid, first)
        set_picker(page, group["fi"], first)
        set_picker(page, group["ff"], last)
        page.wait_for_timeout(2500)
    fields = page.evaluate("""() => { const d = {}; for (const el of document.forms[0].elements) {
        if (el.name) d[el.name] = el.value; } return d; }""")
    fields[button[0]] = button[1]
    for k in [k for k in fields if k.endswith((".x", ".y"))]:
        fields.pop(k)
    resp = context.request.post(url, form=fields, timeout=300000)
    body = resp.body()
    log(f"    {first}..{last} via {button}: HTTP {resp.status} {resp.headers.get('content-type')} {len(body):,} bytes")
    if body[:2] != b"PK":
        text = body.decode("utf-8", errors="replace")
        msg = re.findall(r"(?:alert\(|lblMensaje[^>]*>)([^<)]{0,200})", text)
        log(f"    not a ZIP; messages: {msg[:3]}")
        return
    with zipfile.ZipFile(io.BytesIO(body)) as zf:
        names = zf.namelist()
        log(f"    ZIP holds {len(names)} file(s):")
        for n in names[:40]:
            log(f"      {n}  ({zf.getinfo(n).file_size:,} bytes)")
        seen = set()
        for n in names:
            key = re.sub(r"\d{4}-\d{2}-\d{2}.*", "", n)
            if key in seen:
                continue
            seen.add(key)
            lines = zf.read(n).decode("latin-1", errors="replace").splitlines()
            log(f"    --- head of {n} ({len(lines)} lines) ---")
            for line in lines[:14]:
                log(f"      {line}")
            log(f"      ... {lines[-2:]}")


def groups_for(pickers, buttons):
    """Pair FI/FF pickers (same id with FI->FF) with a download button that shares their suffix."""
    ids = [p[0] for p in pickers]
    out = []
    for pid in ids:
        if "FI" not in pid:
            continue
        ff = pid.replace("FI", "FF", 1)
        if ff not in ids:
            continue
        suffix = pid.split("FI", 1)[1]
        gen = [i for i in ids if i not in (pid, ff) and i.endswith(suffix)]
        btns = [b for b in buttons if b[0] and suffix and b[0].replace("$", "_").endswith(suffix)
                and re.search(r"descarg|zip|csv", str(b[1]) + b[0], re.I)]
        if not btns:
            btns = [b for b in buttons if b[0] and re.search(r"descarg|zip|csv", str(b[1]) + b[0], re.I)]
        maxd = next(p[3] for p in pickers if p[0] == pid)
        out.append({"fi": pid, "ff": ff, "gen": gen, "buttons": btns, "max": maxd, "suffix": suffix})
    return out


def main():
    with sync_playwright() as p:
        browser = p.chromium.launch()
        context = browser.new_context()
        page = context.new_page()
        urls = [BASE + c for c in CANDIDATES]
        try:
            urls += [u for u in menu_links(page) if u not in urls]
        except Exception as e:
            log(f"menu read failed: {e}")
        for url in urls:
            log(f"\n=== {url}")
            try:
                resp = page.goto(url, timeout=90000, wait_until="load")
                log(f"  HTTP {resp.status if resp else '?'}  title {page.title()!r}")
                if not resp or resp.status != 200:
                    continue
                page.wait_for_timeout(2000)
                pickers, buttons = describe(page)
            except Exception as e:
                log(f"  load failed: {e}")
                continue
            for g in groups_for(pickers, buttons):
                log(f"  group suffix {g['suffix']!r}: FI {g['fi']} FF {g['ff']} gen {g['gen']} max {g['max']} "
                    f"buttons {g['buttons']}")
                if not g["buttons"] or not g["max"]:
                    continue
                latest = date.fromisoformat(g["max"])
                for first, last in [(latest - timedelta(days=40), latest - timedelta(days=38)),
                                    (date(2018, 3, 1), date(2018, 3, 2))]:
                    try:
                        try_download(page, context, url, g, g["buttons"][0], first, last)
                    except Exception as e:
                        log(f"    download {first}..{last} failed: {e}")
        browser.close()


if __name__ == "__main__":
    main()
