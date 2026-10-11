"""
Discovery: do grid operators / utilities / regulators / ministries in Cameroon,
Equatorial Guinea, Sierra Leone, Liberia, Gambia, Guinea-Bissau and Cape Verde
publish generation / load / dispatch data in structured form?
Probes reachability, then lists data-like links (csv/xls/json/api/pdf/statistics)
on each home page and a few common sub-paths. Output: stdout + west_africa_other_output.txt.
"""
import re
import sys
from urllib.parse import urljoin

import requests

H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                   "(KHTML, like Gecko) Chrome/124.0 Safari/537.36"}
SITES = {
    "Cameroon": [
        "https://www.eneocameroon.cm/", "https://www.arsel.cm/", "https://arsel-cm.org/",
        "https://www.sonatrel.cm/", "https://www.minee.gov.cm/", "https://www.snh.cm/",
        "https://www.golarlng.com/", "https://www.perenco.com/", "https://www.cameroon-info.net/",
        "https://www.ins.cm/", "https://www.stat.cm/"],
    "Equatorial Guinea": [
        "https://www.segesa.com/", "https://www.segesa.gq/", "https://www.eglng.com/",
        "https://www.mmie.gob.gq/", "https://www.marathonoil.com/", "https://inege.gob.gq/",
        "https://www.guineaecuatorialpress.com/"],
    "Sierra Leone": [
        "https://www.edsa.gov.sl/", "https://edsa.sl/", "https://www.npa.gov.sl/",
        "https://www.energy.gov.sl/", "https://www.epra.gov.sl/", "https://www.statistics.sl/",
        "https://www.statsl.org/"],
    "Liberia": [
        "https://www.lec.com.lr/", "https://www.lerc.gov.lr/", "https://www.mme.gov.lr/",
        "https://www.lisgis.gov.lr/", "https://www.lec.gov.lr/"],
    "Gambia": [
        "https://www.nawec.gm/", "https://www.pura.gm/", "https://www.moepe.gm/",
        "https://www.gbos.gov.gm/", "https://www.moe.gov.gm/"],
    "Guinea-Bissau": [
        "https://www.eagb.gw/", "https://eagb.gw/", "https://www.ine-gb.org/", "https://www.mrn.gov.gw/"],
    "Cape Verde": [
        "https://www.electra.cv/", "https://www.arme.cv/", "https://www.ine.cv/",
        "https://www.governo.cv/", "https://www.dgenergia.cv/", "https://www.cabeolica.com/"],
    "Regional": [
        "https://www.ecowapp.org/", "https://www.ecreee.org/", "https://www.erera.org/", "https://www.arreec.org/",
        "https://www.ecowrex.org/"],
}
KEY = re.compile(r"(\.csv|\.xlsx?|\.json|\.xml|api|open.?data|dashboard|statist|estatist|bulletin|"
                 r"annual|report|relat|generation|geracao|g[ée]n[ée]ration|produc|load|dispatch|"
                 r"demand|energy|energia|[ée]lectric|boletim|dados)", re.I)
out = []


def p(*a):
    s = " ".join(str(x) for x in a)
    print(s, flush=True)
    out.append(s)


def get(url):
    try:
        return requests.get(url, headers=H, timeout=(10, 25), allow_redirects=True)
    except requests.RequestException as e:
        p(f"  ERR {type(e).__name__}: {str(e)[:100]}")
        return None


for country, urls in SITES.items():
    p(f"\n##### {country}")
    for u in urls:
        p(f"\n== {u}")
        r = get(u)
        if r is None:
            continue
        t = re.search(r"<title[^>]*>(.*?)</title>", r.text, re.I | re.S)
        p(f"  status={r.status_code} final={r.url} bytes={len(r.content)} title={(t.group(1).strip()[:100] if t else '')}")
        if r.status_code != 200:
            continue
        seen = set()
        for m in re.finditer(r'href=["\']([^"\'#]+)["\']', r.text, re.I):
            link = urljoin(r.url, m.group(1))
            if link in seen or not KEY.search(link):
                continue
            seen.add(link)
        for l in sorted(seen)[:60]:
            p("   link:", l)
        # common API/sitemap paths
        for sub in ["wp-json/wp/v2/pages?per_page=5", "sitemap.xml", "api/3/action/package_list"]:
            rr = get(urljoin(r.url, sub))
            if rr is not None and rr.status_code == 200:
                p(f"   PATH {sub}: 200 {rr.headers.get('content-type')} {rr.text[:200]!r}")

open("west_africa_other_output.txt", "w").write("\n".join(out))
