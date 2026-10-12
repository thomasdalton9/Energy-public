"""Probe 5: remaining project documents (Cedar, Mozambique LNG, Papua LNG, Pluto Train 2, QatarEnergy statements, LNG Canada Phase 2)."""
import sys, os, re
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from LNG_PROBE_COMMON import probe, hrefs

for u in ["https://www.hellenicshippingnews.com/shell-takes-final-investment-decision-to-double-lng-canada-capacity",
          "https://www.woodside.com/docs/default-source/investor-documents/major-reports-(static-pdfs)/2025-annual-report/full-year-2025-results-briefing-pdf.pdf?sfvrsn=5c83be1c_4",
          "https://www.woodside.com/what-we-do/our-businesses/scarborough",
          "https://www.woodside.com/what-we-do/projects-and-operations/scarborough",
          "https://mozambiquelng.co.mz/about-the-mozambique-liquefied-natural-gas-project/development/",
          "https://www.mozambiquelng.co.mz/about-the-mozambique-liquefied-natural-gas-project/",
          "https://www.cedarlng.com/project/",
          "https://www.cedarlng.com/2026/04/30/spring-2026-newsletter/",
          "https://woodfibrelng.ca/",
          "https://www.kosmosenergy.com/",
          ]:
    probe(u, kw=r"capacity|mtpa|tonnes")
for base, pat in [("https://www.qatarenergy.qa/en/MediaCenter/Pages/News.aspx", r"newsdetails|NewsDetail|news"),
                  ("https://www.santos.com/investors/company-reporting/", r"quarter|report.*2026|pdf")]:
    for u in hrefs(base, pat)[:60]:
        print("LINK", base[-40:], u)
for u in hrefs("https://www.qatarenergy.qa/en/MediaCenter/Pages/News.aspx", r"newsdetails")[:25]:
    probe(u, links=False)
