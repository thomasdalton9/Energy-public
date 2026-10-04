"""
Pakistan power discovery, round 14: NEPRA State of Industry Reports - K-Electric monthly tables (fuel-wise own
generation, fuel-wise / source-wise power purchases incl. CPPA-G), every edition linked on the publications page.
Prints the full text of those table pages (pypdfium2) and the download time.
"""
import re
import tempfile
import time
from urllib.parse import urljoin, unquote

import pypdfium2 as pdfium
import requests
import urllib3

urllib3.disable_warnings()
H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0 Safari/537.36"}
s = requests.Session()
s.headers.update(H)
PAGE = "https://nepra.org.pk/publications/State%20of%20Industry%20Reports.php"
PAT = re.compile(r"K-?Electric\s*\{?\(?(Fuel|Source)-?\s*wise|Month-?wise KE Electricity Generation", re.I)


def main():
    t = s.get(PAGE, timeout=(15, 120), verify=False).text
    links = sorted(set(urljoin(PAGE, h) for h in re.findall(r'href\s*=\s*["\']([^"\']+\.pdf)["\']', t, re.I)))
    for u in links:
        print("  " + unquote(u))
    want = [u for u in links if re.search(r"20(2[2-6])", unquote(u).rsplit("/", 1)[-1])]
    for u in sorted(want, reverse=True)[:4]:
        print(f"\n######## {unquote(u)}", flush=True)
        t0 = time.time()
        with tempfile.NamedTemporaryFile(suffix=".pdf") as f:
            try:
                with s.get(u.replace(" ", "%20"), timeout=(15, 900), verify=False, stream=True) as r:
                    print(f"  {r.status_code} {r.headers.get('content-length')}")
                    for ch in r.iter_content(1 << 20):
                        f.write(ch)
            except Exception as e:  # noqa: BLE001
                print(f"  ERR {e}")
                continue
            f.flush()
            print(f"  downloaded in {time.time() - t0:.0f}s", flush=True)
            doc = pdfium.PdfDocument(f.name)
            for i in range(len(doc)):
                tx = doc[i].get_textpage().get_text_range()
                if PAT.search(tx[:600]):
                    print(f"  --- page {i + 1}")
                    print("  " + tx[:5000].replace("\r", "").replace("\n", "\n  "))


if __name__ == "__main__":
    main()
