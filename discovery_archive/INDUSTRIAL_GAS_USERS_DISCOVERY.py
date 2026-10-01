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
    import pandas as pd
    for c in COUNTRIES:
        try:
            r = get(NMIC + c)
        except Exception as e:  # noqa: BLE001
            print(f"## {c}: {e}")
            continue
        links = sorted(set(re.findall(r'href="([^"]*myb3[^"]*\.(?:xlsx?|pdf))"', r.text, re.I)),
                       key=lambda u: max([int(y) for y in re.findall(r"(20\d\d)", u)] or [0]))
        print(f"## {c}: HTTP {r.status_code}, {len(links)} myb3 links")
        for lk in links[-4:]:
            print("   ", lk)
        xl = [lk for lk in links if re.search(r"\.xlsx?$", lk, re.I)]
        if not xl:
            continue
        url = xl[-1] if xl[-1].startswith("http") else "https://www.usgs.gov" + xl[-1]
        try:
            book = pd.read_excel(io.BytesIO(get(url).content), sheet_name=None, header=None)
        except Exception as e:  # noqa: BLE001
            print(f"   xlsx {url}: {e}")
            continue
        print(f"   XLSX {url}: sheets {list(book)}")
        for name, df in book.items():
            head = " ".join(str(v) for v in df.head(4).fillna("").values.ravel())
            if not re.search(r"structure", head, re.I):
                continue
            print(f"   --- {name}: {head[:160]}")
            last = ""
            for _, row in df.iterrows():
                vals = [str(v).strip() for v in row.values if str(v).strip() not in ("", "nan")]
                if not vals:
                    continue
                line = " | ".join(vals)
                if not re.search(r"\d", line) or len(vals) == 1:
                    last = line   # commodity heading rows
                if KEY.search(line) or KEY.search(last):
                    print(f"   {line[:320]}")


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
