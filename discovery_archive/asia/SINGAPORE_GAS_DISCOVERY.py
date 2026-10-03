"""
Singapore gas discovery: is there any daily / monthly metered gas data (city gas / town gas, natural gas
transportation, receipt or offtake points)? Probes:
  1. data.gov.sg: every dataset whose name mentions gas / LNG / energy / electricity (id, name, frequency, coverage)
  2. EMA gas market page and statistics page: links mentioning gas, plus the Gas Network Code PDF lines on what the
     Gas Transporter must publish
  3. SP Group (PowerGas, the Gas Transporter): sitemap / pages with gas transportation information
  4. City Energy (town gas producer and retailer): sitemap / pages with data or statistics
"""
import io
import re

import requests

H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0 Safari/537.36",
     "Accept": "*/*"}
T = (20, 90)
KEY = re.compile(r"\bgas\b|town gas|piped|lng|natural gas|city energy|powergas|transport", re.I)


def out(*a):
    print(*a, flush=True)


def get(u, **kw):
    try:
        r = requests.get(u, headers=H, timeout=T, **kw)
        out(f"GET {u} -> {r.status_code} {r.headers.get('content-type')} {len(r.content)}")
        return r
    except Exception as e:  # noqa: BLE001
        out(f"GET {u} ERROR {e!r}")
        return None


def links(r, base):
    if r is None or r.status_code != 200:
        return []
    found = set()
    for href, text in re.findall(r'<a[^>]+href="([^"#]+)"[^>]*>(.*?)</a>', r.text, re.S):
        text = re.sub(r"<[^>]+>|\s+", " ", text).strip()
        u = href if href.startswith("http") else base.rstrip("/") + "/" + href.lstrip("/")
        found.add((u, text[:90]))
    return sorted(found)


def data_gov():
    out("\n=== data.gov.sg datasets")
    seen = 0
    for page in range(1, 60):
        r = get(f"https://api-production.data.gov.sg/v2/public/api/datasets?page={page}")
        if r is None or r.status_code != 200:
            break
        d = r.json().get("data", {})
        ds = d.get("datasets") or []
        if not ds:
            break
        for x in ds:
            seen += 1
            name = x.get("name", "")
            if KEY.search(name) or re.search(r"electric|energy", name, re.I):
                out(f"  {x.get('datasetId')} | {name} | {x.get('managedByAgencyName')} | {x.get('coverageStart')}"
                    f"..{x.get('coverageEnd')} | {x.get('frequency') or ''}")
    out(f"  {seen} datasets scanned")


def ema():
    out("\n=== EMA")
    for u in ("https://www.ema.gov.sg/our-energy-story/energy-market-landscape/gas",
              "https://www.ema.gov.sg/resources/statistics",
              "https://www.ema.gov.sg/resources/codes-and-guidelines"):
        for lu, t in links(get(u), "https://www.ema.gov.sg"):
            if KEY.search(lu + " " + t):
                out(f"  {t} -> {lu}")
    # the Gas Network Code: what the Transporter must publish
    r = get("https://www.ema.gov.sg/resources/codes-and-guidelines")
    pdfs = [lu for lu, t in links(r, "https://www.ema.gov.sg") if re.search(r"gas.network.code|gnc", lu + t, re.I)]
    out(f"  GNC candidates: {pdfs[:5]}")
    try:
        import pdfplumber
        for u in pdfs[:2]:
            p = get(u)
            if p is None or p.status_code != 200 or not p.content.startswith(b"%PDF"):
                continue
            with pdfplumber.open(io.BytesIO(p.content)) as pdf:
                for i, page in enumerate(pdf.pages):
                    t = page.extract_text() or ""
                    for line in t.split("\n"):
                        if re.search(r"publish|website|bulletin|make available", line, re.I):
                            out(f"    p{i + 1}: {line[:200]}")
    except Exception as e:  # noqa: BLE001
        out(f"  GNC pdf: {e!r}")


def site(base, paths, sitemap_key):
    out(f"\n=== {base}")
    for sm in ("/sitemap.xml", "/sitemap_index.xml", "/wp-sitemap.xml"):
        r = get(base + sm)
        if r is not None and r.status_code == 200:
            locs = re.findall(r"<loc>([^<]+)</loc>", r.text)
            subs = [x for x in locs if x.endswith(".xml")]
            for s in subs[:30]:
                rs = get(s)
                if rs is not None and rs.status_code == 200:
                    locs += re.findall(r"<loc>([^<]+)</loc>", rs.text)
            hits = sorted({x for x in locs if re.search(sitemap_key, x, re.I)})
            out(f"  sitemap {len(locs)} urls; {len(hits)} matching: " + "\n    ".join([""] + hits[:80]))
    for p in paths:
        for lu, t in links(get(base + p), base):
            if re.search(sitemap_key, lu + " " + t, re.I):
                out(f"  {t} -> {lu}")


def main():
    data_gov()
    ema()
    site("https://www.spgroup.com.sg", ["/", "/our-services/gas", "/our-services/gas/gas-transportation",
                                         "/our-services/utilities/gas"],
         r"gas.(transport|network|data|information|flow|nomination|shipper|transporter|statistic)|powergas|linepack")
    site("https://www.cityenergy.com.sg", ["/", "/about-us/", "/sustainability/", "/investor-relations/"],
         r"data|statistic|report|sendout|send-out|supply|production|annual|investor|meter")


if __name__ == "__main__":
    main()
