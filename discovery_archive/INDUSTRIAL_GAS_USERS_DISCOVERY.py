"""
One-off discovery for the Latin America industrial gas users register (south_america/INDUSTRIAL_GAS_USERS.py).
Run in GitHub Actions (usgs.gov and company sites are blocked from the sandbox). Prints to the job log:

  usgs  - USGS Minerals Yearbook country chapters: the myb3 links on each NMIC country page, then every row of
          the latest XLSX tables release that mentions a gas-using industry (structure-of-industry table: owner,
          location, capacity).
  gem   - download links and licence text found on Global Energy Monitor tracker pages.
  pages - capacity sentences from the pages given as further arguments (space-separated URLs).
  url   - document links + ~7,000 characters of page text from the first capacity sentence.
  wiki  - GEM wiki plant pages: coordinates, owner, status and capacity snippets.
  pdf   - capacity sentences from PDFs given as URLs.
  grep  - first argument is a regex; prints up to 6 snippets (+-160 chars) matching it on each later URL.
  verify - for each row of south_america/industrial_gas_users.csv, fetch its source URLs and report whether the
          capacity figure appears in any of them (one line per row: OK / not found / fetch errors).

Usage: python3 discovery_archive/INDUSTRIAL_GAS_USERS_DISCOVERY.py usgs|gem|pages|url [URL ...]
"""
import io
import re
import sys

import requests

H = {"User-Agent": "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124 Safari/537.36",
     "Accept-Language": "en,es;q=0.8,pt;q=0.6"}
NMIC = "https://www.usgs.gov/centers/national-minerals-information-center/"
COUNTRIES = ["argentina", "bolivia", "brazil", "chile", "colombia", "ecuador", "peru", "uruguay", "venezuela",
             "trinidad-and-tobago", "jamaica", "dominican-republic", "guatemala", "honduras", "el-salvador",
             "nicaragua", "costa-rica", "panama", "belize", "cuba", "paraguay", "guyana", "suriname"]
KEY = re.compile(r"ammonia|urea|nitrogen|methanol|steel|direct[- ]reduc|sponge|DRI|HBI|cement|clinker|alumina|"
                 r"aluminum|glass|LNG|liquefied|petrochem|ethylene|polyeth|ceramic|tile|natural gas|fertili", re.I)
GEM = ["https://globalenergymonitor.org/projects/global-steel-plant-tracker/",
       "https://globalenergymonitor.org/projects/global-steel-plant-tracker/download-data/",
       "https://globalenergymonitor.org/projects/global-cement-and-concrete-tracker/",
       "https://globalenergymonitor.org/projects/global-cement-and-concrete-tracker/download-data/"]
CAP = re.compile(r"[^.]{0,220}(?:t/d|tpd|tonnes|toneladas|t/a|mtpa|ton/d|ton/año|t/año|capacity|capacidad|"
                 r"capacidade|m²|mmscfd|metros cuadrados|millones de m|milhões de m)[^.]{0,220}", re.I)


def get(url):
    return requests.get(url, headers=H, timeout=60)


def text_of(html):
    t = re.sub(r"<script.*?</script>|<style.*?</style>", " ", html, flags=re.S | re.I)
    t = re.sub(r"<[^>]+>", " ", t)
    return re.sub(r"\s+", " ", t)


def usgs():
    """Newest XLSX that still carries the structure-of-industry table (recent 'advance releases' only have
    production), all rows of the gas-using commodity sections."""
    import pandas as pd
    done = set()
    for c in COUNTRIES:
        try:
            r = get(NMIC + c)
        except Exception as e:  # noqa: BLE001
            print(f"## {c}: {e}")
            continue
        links = sorted(set(re.findall(r'href="([^"]*myb3[^"]*\.xlsx?)"', r.text, re.I)),
                       key=lambda u: max([int(y) for y in re.findall(r"(20\d\d)", u)] or [0]), reverse=True)
        for lk in links[:8]:
            url = lk if lk.startswith("http") else "https://www.usgs.gov" + lk
            if url in done:
                break
            try:
                book = pd.read_excel(io.BytesIO(get(url).content), sheet_name=None, header=None)
            except Exception as e:  # noqa: BLE001
                print(f"## {c} {url}: {e}")
                continue
            st = [n for n, df in book.items()
                  if re.search(r"structure", " ".join(str(v) for v in df.head(4).fillna("").values.ravel()), re.I)]
            if not st:
                continue
            done.add(url)
            df = book[st[0]]
            print(f"## {c}: {url} [{st[0]}] {str(df.iloc[0, 0])[:90]}")
            sect = ""
            for _, row in df.iterrows():
                cells = ["" if str(v) == "nan" else str(v).strip() for v in row.values]
                if not any(cells):
                    continue
                if cells[0] and cells[0].lower().rstrip(".") != "do":
                    sect = cells[0]
                if KEY.search(sect) or KEY.search(" ".join(cells)):
                    print("   " + " | ".join(cells)[:400])
            break
        else:
            print(f"## {c}: no structure table in {len(links)} xlsx links")


def gem():
    for u in GEM:
        try:
            r = get(u)
            links = sorted(set(re.findall(r'href="([^"]+\.(?:xlsx|csv|zip|geojson)[^"]*)"', r.text, re.I)))
            print(f"## {u}: HTTP {r.status_code}; {links[:20]}")
            m = re.search(r"(licen[cs]e.{0,200})", text_of(r.text), re.I)
            if m:
                print("   ", m.group(1))
        except Exception as e:  # noqa: BLE001
            print(f"## {u}: {e}")


def pages(urls):
    for u in urls:
        try:
            r = get(u)
            t = text_of(r.text)
            print(f"## {u}: HTTP {r.status_code}, {len(t)} chars")
            seen = set()
            for m in CAP.finditer(t):
                s = m.group(0).strip()
                if s[:80] in seen:
                    continue
                seen.add(s[:80])
                print("   >", s[:440])
                if len(seen) > 25:
                    break
        except Exception as e:  # noqa: BLE001
            print(f"## {u}: {e}")


def url(urls):
    for u in urls:
        try:
            r = get(u)
            print(f"## {u}: HTTP {r.status_code}")
            for lk in [x for x in dict.fromkeys(re.findall(r'href="([^"#]+)"', r.text))
                       if re.search(r"\.(pdf|xlsx?|csv)$", x, re.I)][:40]:
                print("   L", lk)
            t = text_of(r.text)
            # skip the site menu: start at the first capacity-like sentence
            m = CAP.search(t)
            start = max(0, m.start() - 600) if m else 0
            print("   T", t[start:start + 7000])
        except Exception as e:  # noqa: BLE001
            print(f"## {u}: {e}")


def wiki(urls):
    """One compact line per GEM wiki plant page: location, coordinates, owner, capacity rows."""
    for u in urls:
        try:
            r = get(u)
            t = text_of(r.text).replace("&#91;", "[").replace("&#93;", "]")
            t = re.sub(r"\[\s*\d+\s*\]", "", t)
            loc = re.search(r"Location: (.{0,110}?) Coordinates", t)
            xy = re.search(r"(-?\d{1,2}\.\d{4,}), ?(-?\d{1,3}\.\d{4,})", t)
            own = re.search(r"Owner GEM entity ID (.{0,90}?\])", t)
            caps = re.findall(r"Nominal (?:iron|crude steel|cement|clinker)[^()]{0,30}capacity \(total\) (\w+) ([\d.,]+)", t)
            caps += re.findall(r"((?:Cement|Clinker) capacity[^.]{0,80})", t)[:3]
            fuel = re.findall(r"(syngas \(reformed methane\)|natural gas|coal)", t)[:3]
            print(f"## {u} [{r.status_code}] loc={loc.group(1) if loc else ''} | xy={xy.groups() if xy else ''} | "
                  f"owner={own.group(1) if own else ''} | cap={caps} | fuel={sorted(set(fuel))}")
        except Exception as e:  # noqa: BLE001
            print(f"## {u}: {e}")


def pdf(urls):
    import pdfplumber
    for u in urls:
        try:
            b = get(u).content
            with pdfplumber.open(io.BytesIO(b)) as doc:
                t = " ".join((pg.extract_text() or "") for pg in doc.pages[:80])
            t = re.sub(r"\s+", " ", t)
            print(f"## {u}: {len(t)} chars")
            seen = set()
            for m in CAP.finditer(t):
                s = m.group(0).strip()
                if s[:80] not in seen:
                    seen.add(s[:80])
                    print("   >", s[:440])
                if len(seen) > 60:
                    break
        except Exception as e:  # noqa: BLE001
            print(f"## {u}: {e}")


def _doc_text(url, cache={}):
    if url in cache:
        return cache[url]
    try:
        r = get(url)
        ct = r.headers.get("content-type", "")
        if url.lower().endswith((".xlsx", ".xls")) or "spreadsheet" in ct or "excel" in ct:
            import pandas as pd
            book = pd.read_excel(io.BytesIO(r.content), sheet_name=None, header=None)
            t = " ".join(" ".join(str(v) for v in df.fillna("").values.ravel()) for df in book.values())
        elif url.lower().endswith(".pdf") or "pdf" in ct:
            import pdfplumber
            with pdfplumber.open(io.BytesIO(r.content)) as doc:
                t = " ".join((pg.extract_text() or "") for pg in doc.pages[:120])
        else:
            t = text_of(r.text)
        cache[url] = (r.status_code, re.sub(r"\s+", " ", t))
    except Exception as e:  # noqa: BLE001
        cache[url] = (0, str(e)[:80])
    return cache[url]


def _variants(v):
    x = float(v)
    out = {str(v)}
    if x == int(x):
        n = int(x)
        out |= {str(n), f"{n:,}", f"{n:,}".replace(",", "."), f"{n:,}".replace(",", " "), f"{n * 1000:,}",
                f"{n * 1000:,}".replace(",", "."), f"{n / 1000:g}", f"{n / 1000:g}".replace(".", ",")}
    else:
        out |= {f"{x:g}", f"{x:g}".replace(".", ",")}
    return {o for o in out if o not in ("0", "")}


def grep(args):
    pat, urls = re.compile(args[0], re.I), args[1:]
    for u in urls:
        code, t = _doc_text(u)
        hits = [t[max(0, m.start() - 160):m.end() + 160] for m in pat.finditer(t)][:6]
        print(f"## {u} [{code}] {len(t)} chars, {len(hits)} hits")
        for h in hits:
            print("   >", h)


def verify():
    import csv
    rows = list(csv.DictReader(open("south_america/industrial_gas_users.csv", encoding="utf-8")))
    for i, rw in enumerate(rows, start=2):
        cap = rw["capacity"].strip()
        urls = [u.strip() for u in rw["sources"].split(";") if u.strip()]
        if not cap:
            print(f"{i:3d} NO-CAPACITY {rw['plant'][:50]}")
            continue
        found, codes = [], []
        for u in urls:
            code, t = _doc_text(u)
            codes.append(code)
            if any(re.search(r"(?<![\d.,])" + re.escape(v) + r"(?![\d])", t) for v in _variants(cap)):
                found.append(u)
        tag = "OK" if found else "NOTFOUND"
        print(f"{i:3d} {tag:8s} {cap:>7s} {rw['plant'][:48]:48s} http={codes} in={[f[:60] for f in found][:2]}")


if __name__ == "__main__":
    what = sys.argv[1] if len(sys.argv) > 1 else "usgs"
    {"usgs": lambda: usgs(), "gem": lambda: gem(), "pages": lambda: pages(sys.argv[2:]),
     "url": lambda: url(sys.argv[2:]), "wiki": lambda: wiki(sys.argv[2:]), "pdf": lambda: pdf(sys.argv[2:]),
     "verify": verify, "grep": lambda: grep(sys.argv[2:])}[what]()
