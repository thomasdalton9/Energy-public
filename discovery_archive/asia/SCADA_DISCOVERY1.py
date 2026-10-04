"""
Grid-operator (SCADA / dispatch) data discovery across South & Southeast Asia, round 1: for each country's system
operator, open the home / data pages and list links, data files, iframes and api-looking endpoints that carry
real-time or daily generation by source, demand or reservoir data.
  India        MERIT India dashboard, Grid-India reports (daily PSP report: all-India generation incl. wind/solar)
  Vietnam      NSMO (system & market operator), EVN
  Pakistan     NTDC / NPCC, NEPRA, WAPDA (Tarbela / Mangla levels)
  Nepal        NEA Load Dispatch Centre daily operation reports (NDOR pdfs)
  Indonesia    PLN, ESDM Gatrik
  Philippines  NGCP, DOE electric power industry statistics (prod-cms / legacy pdfs)
  Sri Lanka    PUCSL dispatch dashboard (reservoirs), CEB
  Laos / Cambodia / Myanmar  EDL, EDC, MOEE
"""
import re
import sys

import requests
import urllib3

urllib3.disable_warnings()
H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0 Safari/537.36",
     "Accept": "text/html,application/json,*/*", "Accept-Language": "en-US,en;q=0.9"}
T = (15, 60)
KEY = re.compile(r"generat|dispatch|demand|load|daily|report|scada|real.?time|statistic|data|operation|reservoir|"
                 r"level|psp|merit|huy|phu.?tai|san.?luong|van.?hanh|ndor|ldc|beban|pembangkit|api|json|\.xlsx?|\.pdf|"
                 r"\.csv", re.I)

SITES = {
    "india": ["https://meritindia.in/", "https://grid-india.in/", "https://grid-india.in/en/reports/daily-psp-report",
              "https://report.grid-india.in/", "https://vidyutpravah.in/"],
    "vietnam": ["https://www.nsmo.vn/", "https://nsmo.vn/", "https://www.evn.com.vn/", "https://en.evn.com.vn/"],
    "pakistan": ["https://ntdc.gov.pk/", "https://www.ntdc.gov.pk/", "https://nepra.org.pk/", "https://www.wapda.gov.pk/",
                 "https://www.wapda.gov.pk/index.php/river-flow-data", "https://cppa.gov.pk/"],
    "nepal": ["https://www.nea.org.np/", "https://nea.org.np/", "https://cms.nea.org.np/",
              "https://www.nea.org.np/ldc", "https://www.nea.org.np/page/ldc"],
    "indonesia": ["https://web.pln.co.id/", "https://www.pln.co.id/", "https://gatrik.esdm.go.id/",
                  "https://www.esdm.go.id/"],
    "philippines": ["https://www.ngcp.ph/", "https://www.ngcp.ph/operations",
                    "https://doe.gov.ph/data-and-prices/energy-statistics/electric-power-industry",
                    "https://doe.gov.ph/sites/default/files/pdf/energy_statistics/03_Gross%20Power%20Generation_2023.pdf"],
    "sri_lanka": ["https://gendata.pucsl.gov.lk/", "https://www.pucsl.gov.lk/", "https://www.ceb.lk/"],
    "laos_cambodia_myanmar": ["https://www.edl.com.la/", "https://edlgen.com.la/", "https://edc.com.kh/",
                              "https://www.eac.gov.kh/", "https://www.moee.gov.mm/"],
}


def out(*a):
    print(*a, flush=True)


def scan(u):
    try:
        r = requests.get(u, headers=H, timeout=T, verify=False)
    except Exception as e:  # noqa: BLE001
        out(f"  {u}: ERROR {type(e).__name__}: {str(e)[:120]}")
        return
    ct = r.headers.get("content-type", "")
    out(f"  {u} -> {r.status_code} {ct[:40]} {len(r.content)} final {r.url}")
    if r.status_code != 200:
        return
    if r.content[:4] == b"%PDF":
        try:
            import io
            import pdfplumber
            with pdfplumber.open(io.BytesIO(r.content)) as pdf:
                out(f"    PDF {len(pdf.pages)} pages: " + (pdf.pages[0].extract_text() or "")[:1500].replace("\n", " | "))
        except Exception as e:  # noqa: BLE001
            out(f"    pdf: {e!r}")
        return
    if "html" not in ct and "json" not in ct:
        return
    t = r.text
    links = sorted(set(re.findall(r'(?:href|src|action)\s*=\s*["\']([^"\'#]+)["\']', t)))
    hits = [l for l in links if KEY.search(l) and not re.search(r"\.(png|jpe?g|gif|svg|webp|css|woff2?)($|\?)", l, re.I)]
    out(f"    {len(links)} links; {len(hits)} relevant: {hits[:70]}")
    frames = re.findall(r'<iframe[^>]+src=["\']([^"\']+)', t, re.I)
    if frames:
        out(f"    iframes: {frames[:10]}")
    apis = sorted(set(re.findall(r'["\'`]((?:https?://[^"\'`\s]+)?/(?:api|Api|API|ajax|Ajax|services?|ws|json|Handler|'
                                 r'[A-Za-z]+\.aspx/|[A-Za-z]+\.ashx)[^"\'`\s]{0,120})["\'`]', t)))
    if apis:
        out(f"    api-like: {apis[:30]}")
    text = re.sub(r"<script.*?</script>|<style.*?</style>", " ", t, flags=re.S | re.I)
    text = re.sub(r"<[^>]+>|\s+", " ", text)
    for m in list(re.finditer(r".{0,60}(\d[\d,.]*\s*(?:MW|GWh|MU|MWh|kWh)).{0,80}", text))[:8]:
        out("    num: " + m.group(0)[:160])


def main():
    for country in sys.argv[1:] or SITES:
        out(f"\n==================== {country}")
        for u in SITES[country]:
            scan(u)


if __name__ == "__main__":
    main()
