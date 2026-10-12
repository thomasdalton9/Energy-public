"""Probe 2: EIA Q3 file, DISR REQ Sept 2026 files, QatarEnergy decks and news, FERC attachments, operator releases."""
import sys, os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from LNG_PROBE_COMMON import probe, hrefs

EIA = "https://www.eia.gov/naturalgas/importsexports/liquefactioncapacity/"
probe(EIA + "U.S.liquefactioncapacity_2026_Q3.xlsx")
for u in hrefs("https://www.industry.gov.au/publications/resources-and-energy-quarterly-september-2026", r"\.(pdf|xlsx)"):
    print("REQ file", u)
    if u.endswith(".pdf") or "forecast" in u:
        probe(u, maxkeep=1500000)
for u in hrefs("https://www.ferc.gov/media/north-american-lng-export-terminals-existing-approved-not-yet-built-and-proposed", r"\.(pdf|xlsx|xls)|lng"):
    print("FERC link", u)
    if u.lower().endswith((".pdf", ".xlsx", ".xls")):
        probe(u)
for u in ["https://www.qatarenergy.qa/en/MediaCenter/Publications/QatarEnergy_Investors_Presentation_November_2025.pdf",
          "https://www.qatarenergy.qa/en/MediaCenter/Publications/%E2%80%8BQatarEnergy%20Annual%20Review%202024.pdf",
          "https://www.qatarenergy.qa/en/Pages/vHome.aspx"]:
    probe(u, kw=r"ras laffan|north field|lng|incident|force")
for u in hrefs("https://www.qatarenergy.qa/en/Pages/vHome.aspx", r"newsdetails|MediaCenter/Pages/News"):
    print("QE news", u)
for u in ["https://www.businesswire.com/news/home/20260129402488/en/Mozambique-LNG-Announces-the-Full-Restart-of-All-Its-Activities-Onshore-and-Offshore-in-Mozambique",
          "https://mozambiquelng.co.mz/",
          "https://www.woodside.com/investors/reports-investor-briefings/annual-and-half-year-reports",
          "https://www.eni.com/en-IT/media/press-release.html",
          "https://www.golarlng.com/investors/reports-and-presentations/",
          "https://www.lngcanada.ca/news/news-stories/",
          "https://woodfibrelng.ca/news/",
          "https://woodfibrelng.ca/construction/",
          "https://www.cedarlng.com/project/",
          "https://www.cedarlng.com/2026/04/30/spring-2026-newsletter/"]:
    probe(u, kw=r"pdf|lng|quarter|report|phase|train")
