"""
Pakistan power discovery, round 9: K-Electric. NEPRA's K-Electric tariff pages: the monthly KE FCA decisions
(list them), and the text of a few (energy mix: KE own generation by plant/fuel, IPPs on KE's network, CPPA-G).
"""
import io
import re
from urllib.parse import urljoin, unquote

import pdfplumber
import requests
import urllib3

urllib3.disable_warnings()
H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0 Safari/537.36"}
s = requests.Session()
s.headers.update(H)


def get(u):
    try:
        return s.get(u, timeout=(15, 120), verify=False)
    except Exception as e:  # noqa: BLE001
        print(f"  ERR {u}: {e}", flush=True)
        return None


def main():
    home = get("https://nepra.org.pk/").text
    pages = sorted(set(urljoin("https://nepra.org.pk/", l) for l in re.findall(r'href="([^"]+\.php)"', home)
                       if re.search(r"k-electric|kesc|\bke\b", unquote(l), re.I)))
    pages += ["https://nepra.org.pk/tariff/Distribution%20K-Electric.php", "https://nepra.org.pk/tariff/Generation%20K-Electric.php",
              "https://nepra.org.pk/tariff/K-Electric.php"]
    found = {}
    for p in dict.fromkeys(pages):
        r = get(p)
        if r is None:
            continue
        L = [l for l in re.findall(r'href\s*=\s*["\']([^"\']+\.pdf)["\']', r.text, re.I)]
        f = [l for l in L if re.search(r"FCA|FPA|fuel|FCC|MFPA|variation", unquote(l), re.I)]
        print(f"-- {p}: {r.status_code} pdfs {len(L)} fca-ish {len(f)}", flush=True)
        for l in f:
            found[urljoin(p, l)] = p
    for u in sorted(found):
        print("   " + unquote(u).replace("https://nepra.org.pk/", ""))
    # samples: one per year
    picks = {}
    for u in sorted(found):
        y = re.search(r"/(20\d\d)/", u)
        if y and int(y.group(1)) >= 2021 and re.search(r"FCA|FCC|fuel", unquote(u), re.I):
            picks.setdefault(y.group(1), []).append(u)
    for y, us in sorted(picks.items()):
        for u in us[:2]:
            r = get(u)
            if r is None or r.status_code != 200:
                continue
            print(f"\n######## {unquote(u)[-110:]}", flush=True)
            with pdfplumber.open(io.BytesIO(r.content)) as p:
                for i, pg in enumerate(p.pages):
                    t = pg.extract_text() or ""
                    if re.search(r"BQPS|KCCPP|Korangi|CPPA|NTDC|energy mix|generation mix|purchase", t, re.I):
                        print(f"  --- page {i + 1}/{len(p.pages)}:")
                        print("  " + t[:3500].replace("\n", "\n  "))


if __name__ == "__main__":
    main()
