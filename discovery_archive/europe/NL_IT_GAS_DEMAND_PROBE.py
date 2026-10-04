"""
Probe: where do GTS (Netherlands) and Snam (Italy) publish daily national gas consumption?
Loads each page in headless Chromium, logs every JSON/CSV/XLSX/XHR response (url, status, size, head of body), and
lists links to data files. Output only; nothing is written.
"""
import re
import sys

from playwright.sync_api import sync_playwright

PAGES = [
    ("GTS consumption dashboard", "https://www.gasunietransportservices.nl/en/network-operations/dashboard-security-of-supply-gas/gas-consumption-the-netherlands"),
    ("GTS dataport", "https://www.gasunietransportservices.nl/en/network-operations/transparency/dataport"),
    ("Snam figures for today", "https://www.snam.it/en/transportation/operational-data-business/1-Figures-for-today/index.html"),
    ("Snam jarvis", "https://jarvis.snam.it/"),
    ("Snam transparency", "https://www.snam.it/en/transport/transparency/"),
]
KEEP = re.compile(r"json|csv|xls|api|graphql|data|chart", re.I)


def main():
    with sync_playwright() as p:
        b = p.chromium.launch()
        for label, url in PAGES:
            print("=" * 100, f"\n{label}: {url}", flush=True)
            ctx = b.new_context(user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0 Safari/537.36")
            pg = ctx.new_page()
            seen = []

            def on_resp(r):
                u = r.url
                ct = r.headers.get("content-type", "")
                if r.request.resource_type in ("xhr", "fetch") or re.search(r"json|csv|spreadsheet|excel", ct, re.I):
                    try:
                        body = r.text()[:300].replace("\n", " ")
                    except Exception:  # noqa: BLE001
                        body = "<unreadable>"
                    seen.append(f"  [{r.status}] {r.request.method} {u[:200]}  ct={ct[:40]}\n      {body}")

            pg.on("response", on_resp)
            try:
                pg.goto(url, wait_until="networkidle", timeout=60000)
                pg.wait_for_timeout(4000)
            except Exception as e:  # noqa: BLE001
                print("  load problem:", type(e).__name__, str(e)[:200])
            print("  title:", pg.title(), "| final url:", pg.url)
            print("  XHR/JSON responses:", len(seen))
            for s in seen[:40]:
                print(s)
            links = pg.eval_on_selector_all("a[href]", "els => els.map(e => [e.innerText.trim().slice(0,60), e.href])")
            data_links = [(t, h) for t, h in links if re.search(r"\.(csv|xlsx?|json|zip)(\?|$)|download|api", h, re.I)]
            print("  data-like links:", len(data_links))
            for t, h in data_links[:40]:
                print(f"    {t!r} -> {h[:200]}")
            text = pg.inner_text("body")[:1500].replace("\n", " | ")
            print("  page text:", text)
            ctx.close()
        b.close()


if __name__ == "__main__":
    sys.exit(main())
