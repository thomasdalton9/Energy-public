"""
Browser recon of pipeline EBB platforms other than Kinder Morgan's, to
learn each one's page structure before writing a scraper for it (the
KM_OPAVAIL_SCAN.py approach - point-level Total Scheduled Quantity at
LNG terminal delivery points - needs the equivalent page on each
platform).

For each start page: load it in Chromium (these sites are often JS
apps a plain requests.get() can't see), save a screenshot and the
rendered HTML, and print a compact structure summary - tables (with
row counts and header text), form controls (inputs/selects/buttons)
and links whose text or URL mentions capacity / scheduled / flow /
download. Then follow up to FOLLOW_LIMIT of those links one level deep
and summarise each the same way.

Start pages come from FEEDGAS_PIPES_DISCOVERY.py's terminal -> pipeline
map (see its docstring):
  - Enbridge LINK, Texas Eastern: Freeport, Calcasieu Pass, Plaquemines
  - Energy Transfer Messenger+, Trunkline: Sabine Pass, Calcasieu Pass
  - gasnom.com, Cameron Interstate: Cameron LNG
  - Williams 1Line, Transco: Sabine Pass, Cove Point

Usage: python3 EBB_PLATFORM_RECON.py [platform ...]  (default: all)
Outputs: ebb_platform_recon_output/ next to this script.
"""

import os
import re
import sys
from urllib.parse import urljoin, urlparse

try:
    from playwright.sync_api import sync_playwright
except ImportError:
    print("Playwright isn't installed: pip install playwright && playwright install chromium", file=sys.stderr)
    sys.exit(1)

OUTPUT_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "ebb_platform_recon_output")
FOLLOW_LIMIT = 10

# Third pass - platforms for the feedgas gaps (after Kinder Morgan,
# Enbridge TE, Energy Transfer and gasnom were mapped): the terminals'
# own feed pipelines, whose delivery meter into the plant is the whole
# plant (as Cameron Interstate's was for Cameron):
#   - Cheniere LNG Connection: Creole Trail (Sabine Pass) and Cheniere
#     Corpus Christi Pipeline
#   (PipeRiv, a third-party aggregator, was tried once and dropped: data
#   sources here must be the operators' own FERC-mandated postings, not
#   intermediaries that could start charging or disappear.)
#   - TC Energy: ANR (feeds TransCameron/Calcasieu Pass) and Columbia
#     Gulf (feeds Gator Express/Plaquemines)
#   - BHE GT&S: Cove Point
PLATFORMS = {
    "cheniere": [
        "https://lngconnection.cheniere.com/",
    ],
    "tcenergy": [
        "https://ebb.anrpl.com/",
        "https://ebb.tceconnects.com/infopost/",
    ],
    "bhe_covepoint": [
        "https://infopost.bhegts.com/",
    ],
}

LINK_PATTERN = re.compile(r"capacit|operational|avail|schedul|flow|download|csv|nomina", re.I)

SUMMARY_JS = """() => {
  const clean = s => (s || '').replace(/\\s+/g, ' ').trim();
  const tables = [...document.querySelectorAll('table')].map(t => ({
      rows: t.rows.length,
      head: clean((t.rows[0] || {}).innerText).slice(0, 200),
  })).filter(t => t.rows > 2).sort((a, b) => b.rows - a.rows).slice(0, 6);
  const controls = [...document.querySelectorAll('input, select, button')]
      .filter(e => e.type !== 'hidden')
      .map(e => ({
          tag: e.tagName.toLowerCase(), type: e.type || '', id: e.id || '', name: e.name || '',
          text: clean(e.innerText || e.value || e.alt || e.title).slice(0, 60),
          options: e.tagName === 'SELECT' ? [...e.options].map(o => clean(o.text)).slice(0, 12) : undefined,
      })).slice(0, 40);
  const links = [...document.querySelectorAll('a[href]')]
      .map(a => ({text: clean(a.innerText).slice(0, 80), href: a.href}));
  const iframes = [...document.querySelectorAll('iframe')].map(f => f.src);
  return {title: document.title, tables, controls, links, iframes,
          text: clean(document.body ? document.body.innerText : '').slice(0, 800)};
}"""


def log(msg=""):
    print(msg, flush=True)


def slug(url):
    p = urlparse(url)
    return re.sub(r"[^A-Za-z0-9]+", "_", (p.netloc + p.path + ("_" + p.query if p.query else "")))[:120]


def inspect(page, platform, url):
    log(f"\n--- {url}")
    try:
        page.goto(url, wait_until="networkidle", timeout=60000)
        page.wait_for_timeout(6000)  # SPA routes (Williams) render after networkidle
    except Exception as e:
        log(f"  LOAD FAILED: {type(e).__name__}: {str(e)[:200]}")
        return []
    base = os.path.join(OUTPUT_DIR, f"{platform}__{slug(page.url)}")
    try:
        page.screenshot(path=base + ".png", full_page=True)
        with open(base + ".html", "w", encoding="utf-8") as f:
            f.write(page.content())
    except Exception as e:
        log(f"  save failed: {e}")
    s = page.evaluate(SUMMARY_JS)
    log(f"  final URL: {page.url}")
    log(f"  title: {s['title']!r}")
    log(f"  text: {s['text'][:500]!r}")
    for t in s["tables"]:
        log(f"  TABLE rows={t['rows']}: {t['head']!r}")
    for c in s["controls"]:
        extra = f" options={c['options']}" if c.get("options") else ""
        log(f"  CONTROL {c['tag']}[{c['type']}] id={c['id']!r} name={c['name']!r} text={c['text']!r}{extra}")
    for f in s["iframes"]:
        log(f"  IFRAME {f}")
    interesting = []
    seen = set()
    for l in s["links"]:
        if (LINK_PATTERN.search(l["text"]) or LINK_PATTERN.search(l["href"])) and l["href"] not in seen \
                and l["href"].startswith("http"):
            seen.add(l["href"])
            interesting.append(l)
    log(f"  {len(s['links'])} links, {len(interesting)} interesting:")
    for l in interesting[:40]:
        log(f"    {l['text']!r} -> {l['href']}")
    return [l["href"] for l in interesting]


def main():
    wanted = sys.argv[1:] or list(PLATFORMS)
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        context = browser.new_context(ignore_https_errors=True, viewport={"width": 1600, "height": 1200},
                                      user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                                                 "(KHTML, like Gecko) Chrome/124.0 Safari/537.36")
        page = context.new_page()
        for platform in wanted:
            log(f"\n===================== {platform} =====================")
            followed = set()
            for start in PLATFORMS[platform]:
                links = inspect(page, platform, start)
                host = urlparse(start).netloc
                for href in links:
                    if len(followed) >= FOLLOW_LIMIT:
                        break
                    if href in followed or urlparse(href).netloc != host or href.lower().endswith((".pdf", ".doc", ".docx")):
                        continue
                    followed.add(href)
                    inspect(page, platform, href)
        browser.close()
    log(f"\nDONE. Outputs in {OUTPUT_DIR}/")


if __name__ == "__main__":
    main()
