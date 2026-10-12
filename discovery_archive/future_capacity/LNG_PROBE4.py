"""Probe 4: GIIGNL 2026 annual report (cross-check + existing plants), Woodside half-year, Cheniere, Mozambique LNG releases, LNG Canada."""
import sys, os, re
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from LNG_PROBE_COMMON import probe, hrefs, get, OUT, slug

probe("https://giignl-documents.s3.fr-par.scw.cloud/public/giignl-2026-annual-report.pdf", maxkeep=2500000)
probe("https://www.woodside.com/docs/default-source/investor-documents/quarterly-and-half-yearly-pdfs-and-data-tables/2026/half-year-2026-report.pdf?sfvrsn=df6d6ddc_6", maxkeep=1500000)
probe("https://lngir.cheniere.com/_assets/_75dc2c9a1fb706f14e67cee93b0e9e5a/cheniere/news/2026-08-06_Cheniere_Reports_Second_Quarter_2026_Results_and_343.pdf", maxkeep=800000)
probe("https://www.sempra.com/newsroom/press-releases/sempra-reports-strong-second-quarter-2026-results")
probe("https://investor.sempra.com/static-files/d9cdcb67-8e35-40d1-82e0-269b905b8315", maxkeep=800000)
for u in hrefs("https://mozambiquelng.co.mz/press-releases/", r"press-release/|restart|force"):
    print("MZ", u)
for u in hrefs("https://mozambiquelng.co.mz/press-releases/", r"restart|force-majeure|resum")[:3]:
    probe(u)
for u in ["https://www.lngcanada.ca/about-lng-canada/", "https://www.lngcanada.ca/who-we-are/overview/", "https://www.lngcanada.ca/what-we-do/overview/",
          "https://www.lngcanada.ca/what-we-do/facility-operations/", "https://www.lngcanada.ca/what-we-do/construction/"]:
    probe(u, kw=r"phase|train|capacity")
# raw html for the Shell release (script-rendered page)
r = get("https://www.shell.com/news-and-insights/newsroom/news-and-media-releases/2026/shell-takes-final-investment-decision-to-double-lng-canada-capacity.html")
if r is not None:
    open(os.path.join(OUT, "shell_lngcanada_raw.html.txt"), "w", encoding="utf-8").write(r.text[:300000])
    print("shell raw", r.status_code, len(r.text))
for u, kw in [("https://www.santos.com/investors/", r"report|quarter|pdf"), ("https://www.santos.com/our-business/papua-lng/", None),
              ("https://www.inpex.com/ir/library/", r"pdf|quarter|presentation"), ("https://www.petronas.com/lng", None),
              ("https://www.exxonmobil.com/en/lng/papua-new-guinea", None), ("https://www.tangguh.com/", None)]:
    probe(u, kw=kw)
