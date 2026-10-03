"""
Singapore gas discovery, round 2 (after SINGAPORE_GAS_DISCOVERY.py):
  1. data.gov.sg keyword search (gas, town gas, piped gas, natural gas, LNG) - the plain listing stopped at 590 datasets
  2. SingStat TableBuilder keyword search for gas tables (any monthly / higher-frequency ones)
  3. SP Group 'Gas Works' resources (Gas Network Code etc.): list the documents; in the Gas Network Code, every line
     on what the Gas Transporter publishes (website / bulletin / information to shippers)
  4. City Energy piped-gas-supply and gas-production-plant pages: any numbers / data links
"""
import io
import re

import requests

H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0 Safari/537.36",
     "Accept": "*/*"}
T = (20, 120)


def out(*a):
    print(*a, flush=True)


def get(u, **kw):
    try:
        r = requests.get(u, headers=H, timeout=T, **kw)
        out(f"GET {r.url} -> {r.status_code} {r.headers.get('content-type')} {len(r.content)}")
        return r
    except Exception as e:  # noqa: BLE001
        out(f"GET {u} ERROR {e!r}")
        return None


def data_gov():
    out("\n=== data.gov.sg search")
    seen = set()
    for q in ("gas", "town gas", "piped gas", "natural gas", "LNG", "energy"):
        for url in (f"https://api-production.data.gov.sg/v2/public/api/datasets?query={q}",
                    f"https://api-production.data.gov.sg/v2/public/api/datasets/search?query={q}",
                    f"https://data.gov.sg/api/action/package_search?q={q}&rows=100"):
            r = get(url)
            if r is None or r.status_code != 200:
                continue
            try:
                j = r.json()
            except ValueError:
                continue
            ds = (j.get("data") or {}).get("datasets") or (j.get("result") or {}).get("results") or []
            for x in ds:
                name = x.get("name") or x.get("title") or ""
                did = x.get("datasetId") or x.get("id")
                if did in seen or not re.search(r"gas|lng|energy|electric", name, re.I):
                    continue
                seen.add(did)
                out(f"  {did} | {name} | {x.get('managedByAgencyName') or x.get('organization', {}).get('title', '')}"
                    f" | {x.get('coverageStart')}..{x.get('coverageEnd')}")
    # total catalogue size, for reference
    r = get("https://api-production.data.gov.sg/v2/public/api/datasets?page=1")
    if r is not None and r.status_code == 200:
        out(f"  listing keys: {list((r.json().get('data') or {}).keys())}")


def singstat():
    out("\n=== SingStat TableBuilder search")
    for kw in ("gas", "town gas", "natural gas", "piped gas", "LNG"):
        r = get("https://tablebuilder.singstat.gov.sg/api/table/resourceid",
                params={"keyword": kw, "searchOption": "all"})
        if r is None or r.status_code != 200:
            continue
        try:
            recs = (r.json().get("Data") or {}).get("records") or []
        except ValueError:
            continue
        for x in recs:
            out(f"  {x.get('id')} | {x.get('title')} | {x.get('frequency')} | {x.get('tableType')}")


def sp_group():
    out("\n=== SP Group Gas Works resources")
    r = get("https://www.spgroup.com.sg/resources", params={"category": "Gas Works"})
    docs = set()
    if r is not None and r.status_code == 200:
        docs |= set(re.findall(r'["\'](/dam/[^"\']+\.pdf)["\']', r.text))
        docs |= set(re.findall(r'["\'](https://www\.spgroup\.com\.sg/dam/[^"\']+\.pdf)["\']', r.text))
        for m in re.finditer(r'.{0,120}(?:Network Code|Transporter|Shipper|nomination|linepack|flow).{0,160}', r.text, re.I):
            out("  ctx: " + re.sub(r"<[^>]+>|\s+", " ", m.group(0))[:260])
        # JSON API behind the resources page
        for api in re.findall(r'["\'](/api/[^"\']+)["\']', r.text)[:20]:
            out(f"  api path in page: {api}")
    docs = sorted("https://www.spgroup.com.sg" + d if d.startswith("/") else d for d in docs)
    out(f"  {len(docs)} pdfs:")
    for d in docs:
        out(f"    {d}")
    import pdfplumber
    for d in [x for x in docs if re.search(r"network.?code|gnc", x, re.I)][:3]:
        p = get(d)
        if p is None or p.status_code != 200 or not p.content.startswith(b"%PDF"):
            continue
        with pdfplumber.open(io.BytesIO(p.content)) as pdf:
            out(f"  {d}: {len(pdf.pages)} pages")
            for i, page in enumerate(pdf.pages):
                t = page.extract_text() or ""
                for line in t.split("\n"):
                    if re.search(r"publish|website|bulletin board|information exchange|made available to (all )?shippers"
                                 r"|daily (gas|flow|quantity)|linepack", line, re.I):
                        out(f"    p{i + 1}: {line[:220]}")


def city_energy():
    out("\n=== City Energy")
    for u in ("https://www.cityenergy.com.sg/piped-gas-supply/",
              "https://www.cityenergy.com.sg/about-us/our-gas-production-plant/"):
        r = get(u)
        if r is None or r.status_code != 200:
            continue
        text = re.sub(r"<script.*?</script>|<style.*?</style>", " ", r.text, flags=re.S)
        text = re.sub(r"<[^>]+>", " ", text)
        text = re.sub(r"\s+", " ", text)
        for m in re.finditer(r".{0,160}\d[\d,.]*\s*(?:million|mmscfd|m3|GWh|kWh|km|tonnes|households|customers).{0,120}",
                             text, re.I):
            out("  " + m.group(0))
        for f in sorted(set(re.findall(r'href="([^"]+\.(?:pdf|xlsx?|csv))"', r.text))):
            out(f"  file: {f}")


def main():
    data_gov()
    singstat()
    sp_group()
    city_energy()


if __name__ == "__main__":
    main()
