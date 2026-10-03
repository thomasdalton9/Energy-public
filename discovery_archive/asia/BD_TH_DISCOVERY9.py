"""
Sabah / Sarawak discovery, round 4:
  1. MEIH legacy portal: POST /STOASPublicPortlet/energystatistic/searchStatistic.oas with the parameter form
     (installed capacity by region = Sabah / Sarawak; electricity generation) and print the returned table text
  2. myenergystats.st.gov.my (the site MEIH now redirects to): pages, API / data / download endpoints
"""
import re

import requests
import urllib3

urllib3.disable_warnings()
H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0 Safari/537.36"}
T = (20, 90)
BASE = "https://meih.st.gov.my"


def out(*a):
    print(*a, flush=True)


def table_text(html, n=40):
    rows = [" | ".join(re.sub(r"<[^>]+>|\s+", " ", c).strip() for c in re.findall(r"<t[hd][^>]*>(.*?)</t[hd]>", tr, re.S))
            for tr in re.findall(r"<tr.*?</tr>", html, re.S)]
    return "\n    ".join(r for r in rows[:n] if r.strip(" |"))


def form_fields(html):
    form = re.search(r'<form[^>]*id="parameterForm".*?</form>', html, re.S)
    body = form.group(0) if form else html
    fields = []
    for name, val in re.findall(r'<input[^>]*name="([^"]+)"[^>]*value="([^"]*)"', body):
        fields.append((name, val))
    years = re.findall(r'<option[^>]*value="(\d{4})"', body)
    return fields, sorted(set(years))


def main():
    s = requests.Session()
    s.headers.update(H)
    r = s.get(f"{BASE}/statistics", timeout=T, verify=False)
    links = {re.sub(r"<[^>]+>|\s+", " ", t).strip(): h.replace("&amp;", "&")
             for h, t in re.findall(r'<a[^>]+href="([^"]+)"[^>]*>(.*?)</a>', r.text, re.S)}
    for name, extra in (("Installed Capacity", [("optionBy", "1"), ("region", "2")]),
                        ("Installed Capacity", [("optionBy", "1"), ("region", "3")]),
                        ("Electricity Generation", [])):
        p = s.get(links[name], timeout=T, verify=False)
        fields, years = form_fields(p.text)
        out(f"\n######## {name} {extra}: years {years[:2]}..{years[-2:]}; fields {fields[:12]}")
        data = [(k, v) for k, v in fields if k not in ("yearFrom", "yearTo")] + extra + \
               [("yearFrom", years[max(0, len(years) - 12)] if years else "2010"), ("yearTo", years[-1] if years else "2019")]
        q = s.post(f"{BASE}/STOASPublicPortlet/energystatistic/searchStatistic.oas", data=data, timeout=T, verify=False,
                   headers={"X-Requested-With": "XMLHttpRequest", "Referer": p.url})
        out(f"  POST -> {q.status_code} {q.headers.get('content-type')} {len(q.text)}")
        out("  " + (table_text(q.text) or q.text[:800]))
    out("\n==================== myenergystats.st.gov.my")
    for u in ("https://myenergystats.st.gov.my/", "https://myenergystats.st.gov.my/sitemap.xml"):
        try:
            g = s.get(u, timeout=T, verify=False)
        except Exception as e:  # noqa: BLE001
            out(f"GET {u} ERROR {e!r}")
            continue
        out(f"GET {g.url} -> {g.status_code} {g.headers.get('content-type')} {len(g.text)}")
        for m in sorted(set(re.findall(r'["\'](/[a-zA-Z0-9_\-/]*(?:api|data|stat|download|chart|electric|generation|capacity)[^"\'\s]*)["\']', g.text, re.I)))[:60]:
            out(f"  path: {m}")
        for m in sorted(set(re.findall(r'(https?://[a-z0-9.\-]+/[^"\'\s]*(?:api|graphql)[^"\'\s]*)', g.text, re.I)))[:20]:
            out(f"  api: {m}")
        for js in re.findall(r'<script[^>]+src="([^"]+)"', g.text)[:8]:
            jsu = js if js.startswith("http") else "https://myenergystats.st.gov.my" + ("" if js.startswith("/") else "/") + js
            try:
                t = s.get(jsu, timeout=T, verify=False).text
            except Exception:  # noqa: BLE001
                continue
            hits = sorted(set(re.findall(r'["\'`](/?api/[^"\'`\s]{2,80}|https?://[^"\'`\s]*api[^"\'`\s]*)["\'`]', t)))
            out(f"  js {jsu[-60:]}: {len(t)} bytes; api strings {hits[:30]}")


if __name__ == "__main__":
    main()
