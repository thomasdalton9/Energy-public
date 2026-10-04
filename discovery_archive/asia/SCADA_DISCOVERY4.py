"""
Grid-operator data discovery, round 4: Pakistan's hourly plant-wise generation (NEPRA / CPPA-G).
SCADA_DISCOVERY3.py found nepra.org.pk/Admission Notices/2026/09 Sep/01- Plant wise Units.xlsx (hourly MW for 129
plants with fuel, Dec 2025 - May 2026, filed with CPPA-G's FPA application). NEPRA's folders are not listable, so:
  - the NEPRA pages that list admission notices / hearings (links to Plant wise / Hourly / FPA files)
  - guessed paths: the same file names in every month folder 2023-2026 (HEAD requests)
  - CPPA-G pages that may hold the monthly fuel-charge-adjustment data
"""
import re

import requests
import urllib3

urllib3.disable_warnings()
H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 "
                   "Safari/537.36"}
T = (15, 60)
MON = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
NAMES = ["01-%20Plant%20wise%20Units.xlsx", "01-%20Plant%20Wise%20Units.xlsx", "Plant%20wise%20Units.xlsx",
         "01-%20Plant%20wise%20Units.xls", "1-%20Plant%20wise%20Units.xlsx", "03-%20Hourly%20Marginal%20Price.xlsx",
         "Hourly%20Marginal%20Price.xlsx"]


def out(*a):
    print(*a, flush=True)


def main():
    s = requests.Session()
    s.headers.update(H)
    home = s.get("https://nepra.org.pk/", timeout=T, verify=False).text
    pages = sorted(set(l for l in re.findall(r'href="([^"]+)"', home) if re.search(r"admission|hearing|notice|tariff", l, re.I)
                       and not re.search(r"\.(pdf|xlsx?)$", l, re.I)))
    out(f"listing pages: {pages[:40]}")
    for p in pages[:15]:
        u = requests.compat.urljoin("https://nepra.org.pk/", p)
        try:
            t = s.get(u, timeout=T, verify=False).text
        except Exception as e:  # noqa: BLE001
            out(f"  {u}: {e}")
            continue
        f = sorted(set(l for l in re.findall(r'href="([^"]+)"', t) if re.search(r"plant|hourly|units|fpa|fca|cppa", l, re.I)))
        out(f"  {u}: {len(t)} bytes; {f[:40]}")
    # home page itself: all plant/hourly/FPA links
    out("home links: " + str(sorted(set(l for l in re.findall(r'href="([^"]+)"', home)
                                      if re.search(r"plant|hourly|units|fpa|fca", l, re.I)))))
    found = []
    for y in (2023, 2024, 2025, 2026):
        for i, m in enumerate(MON, start=1):
            for n in NAMES:
                u = f"https://nepra.org.pk/Admission%20Notices/{y}/{i:02d}%20{m}/{n}"
                try:
                    r = s.head(u, timeout=T, verify=False, allow_redirects=True)
                except Exception:  # noqa: BLE001
                    continue
                if r.status_code == 200:
                    found.append(u)
                    out(f"  FOUND {u} {r.headers.get('content-length')} {r.headers.get('last-modified')}")
    out(f"{len(found)} found")
    for u in ("https://cppa.gov.pk/fuel-adjustment-notifications", "https://cppa.gov.pk/quarterly-tariff-adjustment",
              "https://cppa.gov.pk/other", "https://cppa.gov.pk/notices", "https://cppa.gov.pk/newsletters"):
        try:
            t = s.get(u, timeout=T, verify=False).text
        except Exception as e:  # noqa: BLE001
            out(f"  {u}: {e}")
            continue
        f = sorted(set(re.findall(r'(?:href|src|data-url)="([^"]+\.(?:pdf|xlsx?|zip|csv))"', t, re.I)))
        api = sorted(set(re.findall(r'["\'](/api/[^"\']+|https?://[^"\']*api[^"\']*)["\']', t)))
        out(f"  {u}: files {f[:20]} api {api[:10]}")
        for m in re.finditer(r"(\$\.(?:ajax|get|post)|fetch\(|axios\.\w+\()\s*\(?\s*[^;]{0,200}", t):
            out("    js: " + m.group(0)[:200])


if __name__ == "__main__":
    main()
