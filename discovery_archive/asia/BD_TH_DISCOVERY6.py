"""
Sabah / Sarawak discovery (runs via bd_th_discovery.yml, which takes the newest BD_TH_DISCOVERY*.py). GSO covers
Peninsular Malaysia only; look for generation / demand data for Sabah (SESB grid) and Sarawak (Sarawak Energy grid):
  - data.gov.my open API: electricity datasets (supply / consumption, by state?)
  - Energy Commission Malaysia Energy Information Hub (meih.st.gov.my): statistics by region
  - SESB (Sabah Electricity) and Sarawak Energy websites: any system data / demand pages
"""
import io
import re

import pandas as pd
import requests
import urllib3

urllib3.disable_warnings()
H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0 Safari/537.36",
     "Accept": "*/*"}
T = (20, 60)
KEY = re.compile(r"sabah|sarawak|region|state|negeri|generation|demand|statistic|electric|janaan|beban", re.I)


def out(*a):
    print(*a, flush=True)


def get(u, **kw):
    try:
        r = requests.get(u, headers=H, timeout=T, verify=False, **kw)
        out(f"GET {r.url} -> {r.status_code} {r.headers.get('content-type')} {len(r.content)}")
        return r
    except Exception as e:  # noqa: BLE001
        out(f"GET {u} ERROR {type(e).__name__}: {str(e)[:150]}")
        return None


def links(r, key=KEY):
    if r is None or r.status_code != 200:
        return []
    found = []
    for href, text in re.findall(r'<a[^>]+href=["\']([^"\'#]+)["\'][^>]*>(.*?)</a>', r.text, re.S | re.I):
        text = re.sub(r"<[^>]+>|\s+", " ", text).strip()
        if key.search(href + " " + text):
            found.append((requests.compat.urljoin(r.url, href), text[:90]))
    return sorted(set(found))


def main():
    out("==================== data.gov.my")
    for ds in ("electricity_supply", "electricity_consumption", "electricity_generation", "electricity_state",
               "electricity_access", "energy_balance"):
        r = get("https://api.data.gov.my/data-catalogue", params={"id": ds, "limit": 5})
        if r is not None and r.status_code == 200:
            out("  " + r.text[:600])
    for u in ("https://data.gov.my/data-catalogue?search=electricity", "https://open.dosm.gov.my/data-catalogue?search=electricity"):
        r = get(u)
        if r is not None and r.status_code == 200:
            ids = sorted(set(re.findall(r'data-catalogue/([a-z0-9_]+)', r.text)))
            out(f"  catalogue ids: {[i for i in ids if re.search('electric|energy|power', i)]}")
    out("\n==================== MEIH (Energy Commission)")
    for u in ("https://meih.st.gov.my/statistics", "https://meih.st.gov.my/", "https://www.st.gov.my/en/web/industry/details/2/26",
              "https://www.st.gov.my/"):
        for lu, t in links(get(u))[:40]:
            out(f"  [{t}] -> {lu}")
    out("\n==================== SESB (Sabah)")
    for u in ("https://www.sesb.com.my/", "https://www.sesb.com.my/sitemap.xml", "https://gso.sesb.com.my/"):
        r = get(u)
        if r is not None and r.status_code == 200 and u.endswith(".xml"):
            locs = re.findall(r"<loc>([^<]+)</loc>", r.text)
            out(f"  sitemap {len(locs)}: {[x for x in locs if KEY.search(x)][:40]}")
        else:
            for lu, t in links(r)[:40]:
                out(f"  [{t}] -> {lu}")
    out("\n==================== Sarawak Energy")
    for u in ("https://www.sarawakenergy.com/", "https://www.sarawakenergy.com/sitemap.xml",
              "https://www.sarawakenergy.com/what-we-do/generation", "https://www.sarawakenergy.com/investor-relations"):
        r = get(u)
        if r is not None and r.status_code == 200 and u.endswith(".xml"):
            locs = re.findall(r"<loc>([^<]+)</loc>", r.text)
            out(f"  sitemap {len(locs)}: {[x for x in locs if KEY.search(x)][:40]}")
        else:
            for lu, t in links(r)[:40]:
                out(f"  [{t}] -> {lu}")


if __name__ == "__main__":
    main()
