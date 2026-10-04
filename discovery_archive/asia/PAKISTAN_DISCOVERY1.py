"""
Pakistan power discovery, round 1: NEPRA / CPPA-G monthly fuel-charge-adjustment (FCA) filings.
  a) NEPRA HTML index pages (.php) that list the Ex-WAPDA DISCOs FCA decisions and the CPPA-G FCA requests /
     "XWDISCOs Energy Purchase Data" admission notices
  b) text of sample PDFs (decision, request, energy purchase data) to see the generation-by-source table
  c) cppa.gov.pk page scripts (API for the file lists)
"""
import io
import re
from urllib.parse import urljoin

import pdfplumber
import requests
import urllib3

urllib3.disable_warnings()
H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 "
                   "Safari/537.36"}
T = (15, 90)
s = requests.Session()
s.headers.update(H)
PAT = re.compile(r"fca|fpa|mfpa|fuel|energy%20purchase|energy purchase|xwdisco|ex-wapda|plant", re.I)


def out(*a):
    print(*a, flush=True)


def get(u):
    try:
        r = s.get(u, timeout=T, verify=False)
        return r
    except Exception as e:  # noqa: BLE001
        out(f"  ERR {u}: {e}")
        return None


def links(html):
    return re.findall(r'href\s*=\s*["\']([^"\']+)["\']', html, re.I)


def pages():
    home = get("https://nepra.org.pk/").text
    php = sorted(set(l for l in links(home) if ".php" in l.lower()))
    out(f"home .php links ({len(php)}): {php}")
    cands = set(urljoin("https://nepra.org.pk/", l) for l in php)
    for g in ["tariff/Distribution XWDISCOs.php", "tariff/Distribution Ex-WAPDA DISCOs.php", "tariff/Distribution.php",
              "tariff/Distribution GEPCO.php", "tariff/tariff.php", "admission-notices.php", "Admission Notices.php",
              "admission_notices.php", "tariff/Distribution FCA.php", "tariff/FCA.php", "tariff/Distribution LESCO.php",
              "tariff/Distribution XWDISCOs FCA.php", "hearings.php", "tariff/Distribution K-Electric.php"]:
        cands.add(urljoin("https://nepra.org.pk/", g))
    hits = {}
    for u in sorted(cands):
        r = get(u)
        if r is None:
            continue
        L = links(r.text)
        f = [l for l in L if PAT.search(l)]
        fa = [l for l in L if re.search(r"fca|fpa|fuel|energy%20purchase|energy purchase|plant", l, re.I)]
        out(f"-- {u}: {r.status_code} {len(r.text)}b links {len(L)} matching {len(f)} fca-ish {len(fa)}")
        if fa:
            hits[u] = fa
            for l in fa[:8]:
                out(f"     {l}")
    # second level: php pages linked from tariff pages
    sub = set()
    for u in list(cands):
        r = get(u) if "tariff" in u.lower() else None
        if r is None:
            continue
        for l in links(r.text):
            if ".php" in l.lower():
                sub.add(urljoin(u, l))
    out(f"second-level php ({len(sub)}): {sorted(sub)[:120]}")
    for u in sorted(sub - cands):
        if not re.search(r"distribution|tariff|admission|notice|hearing|wapda|fca|fuel", u, re.I):
            continue
        r = get(u)
        if r is None:
            continue
        fa = [l for l in links(r.text) if re.search(r"fca|fpa|fuel|energy%20purchase|energy purchase|plant", l, re.I)]
        out(f"-- {u}: {r.status_code} {len(r.text)}b fca-ish {len(fa)}")
        for l in fa[:12]:
            out(f"     {l}")
        if fa:
            hits[u] = fa
    for u, fa in hits.items():
        out(f"\n==== ALL fca-ish links on {u} ({len(fa)})")
        for l in fa:
            out(f"   {l}")


def pdf(u, maxpages=8):
    out(f"\n######## PDF {u}")
    r = get(u)
    if r is None or r.status_code != 200:
        out(f"  status {getattr(r, 'status_code', None)}")
        return
    out(f"  {len(r.content)} bytes")
    try:
        with pdfplumber.open(io.BytesIO(r.content)) as p:
            out(f"  {len(p.pages)} pages")
            for i, pg in enumerate(p.pages[:maxpages]):
                t = pg.extract_text() or ""
                out(f"  --- page {i + 1} ({len(t)} chars)")
                out(t[:3500])
    except Exception as e:  # noqa: BLE001
        out(f"  pdf error {e}")


def cppa():
    for u in ("https://cppa.gov.pk/fuel-adjustment-notifications", "https://cppa.gov.pk/xwdiscos-energy-purchase-data"):
        r = get(u)
        if r is None:
            continue
        t = r.text
        scripts = re.findall(r'<script[^>]+src="([^"]+)"', t)
        out(f"\n-- {u}: {len(t)}b scripts {scripts}")
        for m in re.finditer(r"storage/uploads[^\"' )]*|/api/[^\"' )]*|wire:[a-z\-]+=\"[^\"]*\"|livewire|inertia|data-page=\"[^\"]{0,300}", t):
            out("   " + m.group(0)[:300])
        i = t.find("storage")
        out("   snippet: " + t[max(0, i - 1500):i + 1500].replace("\n", " ")[:3000])
        for sc in scripts:
            if "jquery" in sc or "bootstrap" in sc:
                continue
            js = get(urljoin(u, sc))
            if js is None:
                continue
            for m in re.finditer(r"[\"'`](/?api/[^\"'`]+|https?://[^\"'`]*cppa[^\"'`]*)[\"'`]", js.text):
                out(f"   js {sc}: {m.group(1)[:200]}")


def main():
    pages()
    for u in ["https://nepra.org.pk/tariff/Tariff/Ex-WAPDA%20DISCOS/2026/TRF-100%20XWDISCOS%20FCA%20JUN%202026%2007-08-2026%2017536-54.pdf",
              "https://nepra.org.pk/Admission%20Notices/2026/08%20Aug/XWDISCOs%20FCA%20-%20Jul-26%20-%20Request%20Jul%202026.pdf",
              "https://nepra.org.pk/Admission%20Notices/2025/08%20Aug/XWDISCOs'%20ENERGY%20PURCHASE%20DATA%20FOR%20MONTH%20OF%20JULY%202025.PDF",
              "https://www.nepra.org.pk/Admission%20Notices/2021/09%20Sep/XWDISCOs%20Energy%20Purchase%20Data%20for%20the%20month%20of%20August%202021.pdf"]:
        pdf(u, maxpages=6)
    cppa()


if __name__ == "__main__":
    main()
