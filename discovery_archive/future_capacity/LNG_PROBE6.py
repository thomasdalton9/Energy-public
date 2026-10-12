"""Probe 6: Qatar timing (press copies of CEO statements), Cedar (Pembina), Papua LNG (Santos), Rovuma, Woodside annual report."""
import sys, os, re
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from LNG_PROBE_COMMON import probe, hrefs

for u in ["https://boereport.com/2026/10/07/qatarenergys-lng-expansion-project-poised-to-start-operations-in-november-sources-say/amp/",
          "https://www.qatar-tribune.com/article/254953/business/north-field-east-lng-output-expected-in-h1-2027-al-kaabi",
          "https://www.bairdmaritime.com/shipping/tankers/gas/launch-of-new-qatar-lng-train-moves-to-counter-iranian-attack-losses",
          "https://www.roic.ai/news/qatarenergys-first-north-field-east-lng-train-said-ready-but-november-startup-timing-clashes-with-2027-guidance-10-07-2026",
          "https://www.pembina.com/media-centre/news/details/fa13b9e4-bbc0-44d0-9ab6-88c72cbd54e5",
          "https://www.sec.gov/Archives/edgar/data/0001546066/000154606626000064/q22026interimreport.htm",
          "https://www.santos.com/news/2026-second-quarter-report/",
          "https://www.cnbcafrica.com/wire/653161",
          "https://www.offshore-technology.com/news/exxonmobil-delays-mozambique-lng-project/",
          ]:
    probe(u, links=False)
for u in hrefs("https://www.woodside.com/investors/reports-investor-briefings/annual-and-half-year-reports", r"\.pdf")[:40]:
    print("WDS", u)
for u in hrefs("https://www.woodside.com/investors/reports-investor-briefings/annual-and-half-year-reports", r"annual-report.*\.pdf|2025-annual-report")[:2]:
    probe(u, links=False, maxkeep=2500000)
