"""
One-off discovery for the Latin America industrial gas users register (south_america/INDUSTRIAL_GAS_USERS.py).
Run in GitHub Actions (usgs.gov and company sites are blocked from the sandbox). Prints to the job log:

  usgs  - USGS Minerals Yearbook country chapters: the myb3 links on each NMIC country page, then every row of
          the latest XLSX tables release that mentions a gas-using industry (structure-of-industry table: owner,
          location, capacity).
  gem   - download links and licence text found on Global Energy Monitor tracker pages.
  pages - capacity sentences from the pages given as further arguments (space-separated URLs).
  url   - first 200 links + text of the given pages (to find report/PDF links).

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
            for lk in list(dict.fromkeys(re.findall(r'href="([^"#]+)"', r.text)))[:200]:
                print("   L", lk)
            print("   T", text_of(r.text)[:4000])
        except Exception as e:  # noqa: BLE001
            print(f"## {u}: {e}")


if __name__ == "__main__":
    what = sys.argv[1] if len(sys.argv) > 1 else "usgs"
    {"usgs": lambda: usgs(), "gem": lambda: gem(), "pages": lambda: pages(sys.argv[2:]),
     "url": lambda: url(sys.argv[2:])}[what]()
