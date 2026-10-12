"""Probe 1: reachability of the official sources for global liquefaction (EIA, FERC, CER, DISR, QatarEnergy, operators)."""
import sys
sys.path.insert(0, __import__("os").path.dirname(__import__("os").path.abspath(__file__)))
from LNG_PROBE_COMMON import probe

EIA = "https://www.eia.gov/naturalgas/importsexports/liquefactioncapacity/"
URLS = [
    (EIA + "U.S.liquefactioncapacity_2026_Q3.xlsx", None),
    (EIA + "U.S.liquefactioncapacity_2026_Q2.xlsx", None),
    (EIA, r"xlsx|liquef"),
    ("https://www.ferc.gov/industries-data/natural-gas/overview/liquefied-natural-gas-lng", r"lng|terminal|export"),
    ("https://www.ferc.gov/media/north-american-lng-export-terminals-existing-approved-not-yet-built-proposed", None),
    ("https://www.ferc.gov/media/north-american-lng-export-terminals-existing-approved-not-yet-built-and-proposed", None),
    ("https://www.cer-rec.gc.ca/en/data-analysis/energy-commodities/natural-gas/report/lng-export-licences/index.html", r"lng|licen"),
    ("https://www.cer-rec.gc.ca/en/applications-hearings/view-applications-projects/lng-projects.html", r"lng|licen"),
    ("https://www.industry.gov.au/publications/resources-and-energy-quarterly-september-2026", r"xlsx|pdf|gas|lng"),
    ("https://www.industry.gov.au/publications/resources-and-energy-quarterly-june-2026", r"xlsx|pdf|gas|lng"),
    ("https://www.industry.gov.au/publications/resources-and-energy-quarterly", r"quarterly-.*20(25|26)"),
    ("https://www.qatarenergy.qa/en/MediaCenter/Pages/default.aspx", r"north field|lng"),
    ("https://www.qatarenergy.qa/en/OurBusiness/Pages/NorthFieldExpansion.aspx", None),
    ("https://www.lngcanada.ca/", r"phase|train|capacity|news"),
    ("https://www.woodfibrelng.ca/", r"news|project"),
    ("https://www.cedarlng.com/", r"news|project"),
    ("https://giignl.org/resources/", r"annual|report|pdf"),
    ("https://www.igu.org/resources/world-lng-report-2025/", r"pdf|report"),
    ("https://www.igu.org/resources/", r"lng|pdf"),
    ("https://www.woodside.com/investors/reports-investor-briefings", r"pdf|quarter|report"),
    ("https://www.nlng.com/Pages/Train7.aspx", None),
    ("https://www.novatek.ru/en/business/arctic-lng/", None),
    ("https://www.santos.com/investors/reports/", r"pdf|quarter"),
    ("https://www.totalenergies.com/projects/liquefied-natural-gas/mozambique-lng", None),
]
for u, kw in URLS:
    probe(u, kw)
