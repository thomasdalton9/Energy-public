"""
South & Southeast Asia gas discovery, round 1 (run from GitHub Actions; the sandbox is blocked).
Looks for monthly-or-better official gas data for Pakistan, Indonesia, Malaysia, Vietnam, Philippines:
  PK  PBS Monthly Bulletin of Statistics (natural gas production), OGRA, Petroleum Division, SNGPL/SSGC, PLL/PSO, DGPC
  ID  ESDM homepage cards (daily gas), SKK Migas, Ditjen Migas, Satu Data ESDM
  MY  OpenDOSM / data.gov.my catalogue, DOSM petroleum & natural gas release, MEIH
  VN  NSO/GSO monthly socio-economic report tables (natural gas output), PV Gas
  PH  DOE natural gas production & consumption pages
Prints status, size, page text around gas keywords and data-like links. Nothing is written.
"""
import re
from urllib.parse import urljoin

import requests
import urllib3

urllib3.disable_warnings()
H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) "
                   "Chrome/124.0 Safari/537.36", "Accept-Language": "en,id;q=0.8,vi;q=0.6"}
T = (15, 60)
KW = re.compile(r"(natural gas|gas bumi|gas alam|MMSCFD|MMCFD|mmcft|mmscf|RLNG|LNG|lifting|khí|bcf|Malampaya)", re.I)
LINK = re.compile(r"""(?:href|src)=["']([^"']+)["']""", re.I)
DATA = re.compile(r"\.(xlsx?|csv|json|pdf|zip)(\?|$)|api|statisti|bulletin|data|download|gas|khi|thang|report", re.I)


def out(s=""):
    print(s, flush=True)


def probe(name, url, show=300, links=80, kw=8, pat=DATA, **kwa):
    out(f"\n{'=' * 100}\n{name}: {url}")
    try:
        r = requests.get(url, headers=H, timeout=T, verify=False, **kwa)
    except Exception as e:  # noqa: BLE001
        out(f"  ERROR {type(e).__name__}: {str(e)[:200]}")
        return None
    out(f"  -> {r.status_code} {r.headers.get('content-type')} {len(r.content)}b final={r.url} "
        f"LM={r.headers.get('last-modified')}")
    ct = r.headers.get("content-type", "")
    if any(x in ct for x in ("html", "json", "text", "xml")):
        txt = r.text
        t = re.search(r"<title>(.*?)</title>", txt, re.S | re.I)
        out(f"  TITLE: {t.group(1).strip()[:150] if t else None}")
        if show:
            out("  HEAD: " + re.sub(r"\s+", " ", txt[:show]))
        plain = re.sub(r"<script.*?</script>|<style.*?</style>", " ", txt, flags=re.S | re.I)
        plain = re.sub(r"<[^>]+>", " ", plain)
        plain = re.sub(r"\s+", " ", plain)
        n = 0
        for m in KW.finditer(plain):
            if n >= kw:
                break
            out(f"  KW: ...{plain[max(0, m.start() - 150):m.end() + 200]}...")
            n += 1
        ls = []
        for l in LINK.findall(txt):
            if pat.search(l) and not l.startswith(("javascript", "#", "mailto")):
                a = urljoin(r.url, l)
                if a not in ls:
                    ls.append(a)
        out(f"  LINKS ({len(ls)}):")
        for l in ls[:links]:
            out(f"    {l}")
    return r


def main():
    # ---------------- Pakistan
    for u in ("https://www.pbs.gov.pk/monthly-bulletin-of-statistics/", "https://www.pbs.gov.pk/monthly-bulletin-of-statistics",
              "https://www.pbs.gov.pk/content/monthly-bulletin-statistics", "https://www.pbs.gov.pk/publications/"):
        probe("PBS MBS", u, links=120, pat=re.compile(r"MBS|bulletin|\.pdf|\.xls", re.I))
    probe("PBS MBS Sep2024 pdf", "https://www.pbs.gov.pk/wp-content/uploads/2020/07/MBS_Sep_2024.pdf", show=0)
    for u in ("https://www.ogra.org.pk/", "https://www.ogra.org.pk/annual-reports", "https://www.ogra.org.pk/publications",
              "https://www.ogra.org.pk/state-of-the-regulated-petroleum-industry"):
        probe("OGRA", u)
    for u in ("https://mope.gov.pk/", "https://petroleum.gov.pk/", "https://www.mopnr.gov.pk/",
              "https://petroleum.gov.pk/Detail/YmJiMjcwMWYtYmI5Zi00ZDU1LWEyNmQtZmE0NDE3YzIwN2Ux",
              "https://dgpc.gov.pk/", "https://www.dgpc.gov.pk/", "https://ppis.gov.pk/", "https://www.ppisonline.com/"):
        probe("PetDiv/DGPC", u)
    for u in ("https://www.sngpl.com.pk/", "https://www.ssgc.com.pk/", "https://www.ssgc.com.pk/web/?page_id=1063",
              "https://pll.com.pk/", "https://www.pll.com.pk/", "https://psopk.com/en/", "https://www.isgs.com.pk/"):
        probe("PK gas cos", u)
    # ---------------- Indonesia
    for u in ("https://www.esdm.go.id/id", "https://www.esdm.go.id/en", "https://www.esdm.go.id/"):
        probe("ESDM home", u, show=0, kw=15)
    for u in ("https://www.skkmigas.go.id/", "https://www.skkmigas.go.id/en", "https://skkmigas.go.id/kinerja",
              "https://www.skkmigas.go.id/en/data", "https://www.skkmigas.go.id/laporan-tahunan"):
        probe("SKK Migas", u, kw=12)
    for u in ("https://migas.esdm.go.id/", "https://migas.esdm.go.id/post/statistik",
              "https://migas.esdm.go.id/cms/statistik", "https://satudata.esdm.go.id/",
              "https://satudata.esdm.go.id/api/3/action/package_search?q=gas&rows=50",
              "https://data.esdm.go.id/", "https://www.esdm.go.id/id/publikasi/handbook-of-energy-economic-statistics-of-indonesia"):
        probe("Migas/SatuData", u, kw=10)
    # ---------------- Malaysia
    probe("data.gov.my catalogue", "https://api.data.gov.my/data-catalogue?meta=true", show=2000, links=0)
    for q in ("gas", "petroleum", "mining", "lng"):
        probe(f"OpenDOSM search {q}", f"https://open.dosm.gov.my/data-catalogue?search={q}", show=0, kw=10, links=60,
              pat=re.compile(r"data-catalogue/", re.I))
    for u in ("https://www.dosm.gov.my/portal-main/release-content/mining-of-petroleum-and-natural-gas-statistics-q22026",
              "https://www.dosm.gov.my/portal-main/release-content/mining-of-petroleum-and-natural-gas-statistics-third-quarter-2025"):
        probe("DOSM petroleum & gas", u, kw=15)
    for u in ("https://meih.st.gov.my/", "https://meih.st.gov.my/statistics", "https://meih.st.gov.my/database",
              "https://meih.st.gov.my/statistics?p_p_id=Eng_Statistic_WAR_STOASPublicPortlet"):
        probe("MEIH", u, kw=10)
    # ---------------- Vietnam
    for u in ("https://www.nso.gov.vn/bao-cao-tinh-hinh-kinh-te-xa-hoi-hang-thang/",
              "https://www.nso.gov.vn/so-lieu-thong-ke/", "https://www.nso.gov.vn/en/data-and-statistics/",
              "https://www.nso.gov.vn/en/monthly-socio-economic-situation/", "https://www.nso.gov.vn/?p=62262",
              "https://www.gso.gov.vn/bao-cao-tinh-hinh-kinh-te-xa-hoi-hang-thang/"):
        probe("VN NSO", u, kw=6, links=100)
    for u in ("https://www.pvgas.com.vn/", "https://www.pvn.vn/"):
        probe("PV Gas/PVN", u, kw=6)
    # ---------------- Philippines
    for u in ("https://doe.gov.ph/natgas", "https://doe.gov.ph/natgas?page=4", "https://legacy.doe.gov.ph/natgas",
              "https://doe.gov.ph/natural-gas-production-and-consumption-million-standard-cubic-feet-mmcsf",
              "https://legacy.doe.gov.ph/energy-statistics", "https://doe.gov.ph/energy-statistics",
              "https://legacy.doe.gov.ph/natgas/gas-markets-power"):
        probe("PH DOE", u, kw=10, links=100)


if __name__ == "__main__":
    main()
