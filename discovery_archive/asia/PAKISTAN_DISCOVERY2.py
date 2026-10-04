"""
Pakistan power discovery, round 2.
  a) NEPRA home .php pages: which list the CPPA-G FCA requests ("Energy Purchase Data", Admission Notices)
  b) cppa.gov.pk page bodies (fuel-adjustment-notifications, xwdiscos-energy-purchase-data): links, ajax in site.js
  c) for every CPPA-G request PDF found, whether the "Summary" fuel table is in the text layer
"""
import io
import re
from urllib.parse import urljoin, quote, unquote

import pdfplumber
import requests
import urllib3

urllib3.disable_warnings()
H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 "
                   "Safari/537.36"}
T = (15, 90)
s = requests.Session()
s.headers.update(H)
REQ = re.compile(r"energy%20purchase|energy purchase|fca%20-|fca data|fpa data|request|fpa%20data|fca%20data|"
                 r"energy%20data|cppa", re.I)


def out(*a):
    print(*a, flush=True)


def get(u):
    try:
        return s.get(u, timeout=T, verify=False)
    except Exception as e:  # noqa: BLE001
        out(f"  ERR {u}: {e}")
        return None


def links(html):
    return re.findall(r'href\s*=\s*["\']([^"\']+)["\']', html, re.I)


def nepra_pages():
    home = get("https://nepra.org.pk/").text
    php = sorted(set(urljoin("https://nepra.org.pk/", l) for l in links(home) if ".php" in l.lower()))
    out(f"home php ({len(php)}): {php}")
    found = {}
    for u in php:
        if "tariff/" in u and ("Distribution" in u or "Generation" in u):
            continue
        r = get(u)
        if r is None:
            continue
        L = [l for l in links(r.text) if re.search(r"\.pdf|\.xlsx?", l, re.I)]
        q = [l for l in L if REQ.search(l)]
        out(f"-- {u}: {r.status_code} files {len(L)} cppa-request-ish {len(q)}")
        for l in q[:10]:
            out(f"     {l}")
        for l in q:
            found[urljoin(u, l)] = u
    return found


def cppa():
    for u in ("https://cppa.gov.pk/fuel-adjustment-notifications", "https://cppa.gov.pk/xwdiscos-energy-purchase-data",
              "https://cppa.gov.pk/policies-reports"):
        r = get(u)
        if r is None:
            continue
        t = r.text
        i, j = t.find("</header>"), t.find("<footer")
        body = t[i if i > 0 else 0:j if j > 0 else len(t)]
        out(f"\n==== {u} body {len(body)}b")
        out(re.sub(r"\n\s*\n+", "\n", body)[:6000])
    for js in ("https://cppa.gov.pk/js/site.js", "https://cppa.gov.pk/js/home.js"):
        r = get(js)
        if r is None:
            continue
        out(f"\n==== {js} {len(r.text)}b")
        for m in re.finditer(r".{0,200}(ajax|fetch\(|\.get\(|\.post\(|url\s*:|/api/|storage).{0,250}", r.text):
            out("   " + m.group(0)[:450])
    for api in ("https://cppa.gov.pk/api/documents", "https://cppa.gov.pk/storage/uploads/"):
        r = get(api)
        if r is not None:
            out(f"  {api}: {r.status_code} {r.headers.get('content-type')} {r.text[:300]!r}")


def check_pdf(u):
    r = get(u)
    if r is None or r.status_code != 200:
        out(f"  {getattr(r, 'status_code', None)} {unquote(u)[-90:]}")
        return
    try:
        with pdfplumber.open(io.BytesIO(r.content)) as p:
            txt = ""
            for pg in p.pages:
                t = pg.extract_text() or ""
                if "Summary" in t or "Coal-Local" in t or "Grand Total" in t:
                    txt += t
            m = re.search(r"Summary.*?Grand Totals? for the month[^\n]*", txt, re.S)
            out(f"  {len(p.pages)}p {unquote(u)[-90:]}: summary={'YES' if m else 'no'}")
            if m:
                out("     " + m.group(0)[:1800].replace("\n", "\n     "))
    except Exception as e:  # noqa: BLE001
        out(f"  pdf error {e} {u}")


def main():
    found = nepra_pages()
    out(f"\n{len(found)} request-ish files")
    for u in sorted(found):
        out(f"  {u}")
    cppa()
    out("\n==== text-layer check")
    for u in sorted(found):
        if re.search(r"energy%20purchase|energy purchase|fca%20-|request", u, re.I) and u.lower().endswith(".pdf"):
            check_pdf(u)


if __name__ == "__main__":
    main()
