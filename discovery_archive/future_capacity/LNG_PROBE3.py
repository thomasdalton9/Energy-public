"""Probe 3: operator and company releases for global liquefaction projects (opened in full, text saved)."""
import sys, os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from LNG_PROBE_COMMON import probe, hrefs

URLS = [
 ("https://www.shell.com/news-and-insights/newsroom/news-and-media-releases/2026/shell-takes-final-investment-decision-to-double-lng-canada-capacity.html", None),
 ("https://totalenergies.com/system/files/documents/totalenergies_pr-results-2q26_2026_en.pdf", None),
 ("https://www.woodside.com/docs/default-source/investor-documents/quarterly-and-half-yearly-pdfs-and-data-tables/2026/second-quarter-2026-report.pdf", None),
 ("https://www.eni.com/en-IT/media/press-release/2025/10/pr-eni-announces-final-investment-decision-mozambique-coral-north-project.html", None),
 ("https://mfn.se/one/a/golar-lng/golar-lng-limited-interim-results-for-the-period-ended-june-30-2026-b168e773", None),
 ("https://www.golar.com/investors/results-centre/highlights.aspx", None),
 ("https://www.sec.gov/Archives/edgar/data/0001032208/000103220826000043/ex99_1x20260630xearningsta.htm", None),
 ("https://www.nlng.com/operations/train-7", None),
 ("https://www.nlng.com/documents/2024_NLNG_FACTS_FIGURESB_compressed.pdf", None),
 ("https://www.bp.com/api/files/6cqieuqhq4no/master/3m82HmOzrcUhzGfOQ2Uajc/9fbef18373195a44947e3a797d6cc0cf/bp-second-quarter-2026-results.pdf", None),
 ("https://www.bp.com/en/global/corporate/news-and-insights/press-releases/bp-announces-final-investment-decision-for-phase-1-of-the-greater-tortue-ahmeyim-lng-development.html", None),
 ("https://investors.kosmosenergy.com/news-releases/news-release-details/kosmos-energy-announces-first-lng-cargo-greater-tortue-ahmeyim", None),
 ("https://www.sec.gov/Archives/edgar/data/3570/000000357026000030/cei20262ndqtrerex991.htm", None),
 ("https://www.novatek.ru/common/upload/press/FID_Arctic_LNG_2_Eng.pdf", None),
 ("https://www.giignl.org/annual-report/2026", r"pdf|download|annual"),
 ("https://giignl.org/annual-report/", r"pdf|download|annual"),
 ("https://woodfibrelng.ca/construction/", None),
 ("https://www.cedarlng.com/project/", None),
 ("https://www.lngcanada.ca/news/news-stories/", r"phase|news-stories/"),
 ("https://www.novatek.ru/en/press/releases/", r"arctic|lng"),
 ("https://www.adnocgas.ae/en/news-and-media", r"ruwais|lng"),
 ("https://www.omanlng.com/", r"train|capacity|lng"),
]
for u, kw in URLS:
    probe(u, kw=kw)
