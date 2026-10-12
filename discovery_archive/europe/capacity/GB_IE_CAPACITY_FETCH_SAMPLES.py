"""Fetch the GB/IE capacity source files once (12 Oct 2026) so the parsers can be developed locally against the real
layouts. Files are saved into discovery_archive/results/europe_capacity/gb_ie/ (public government/TSO data)."""
import os, re, requests
S = requests.Session(); S.headers["User-Agent"] = "Mozilla/5.0 energy-research"
OUT = "discovery_archive/results/europe_capacity/gb_ie"; os.makedirs(OUT, exist_ok=True)
B = "https://assets.publishing.service.gov.uk/media/"
for name, url in [("ET_6.1.xlsx", B+"6aba724dfceb6fb3a65012d2/ET_6.1_SEP_26.xlsx"),
                  ("ET_6.3_solar.xlsx", B+"6aba729a734e2435f202f19e/ET_6.3_SOLAR_PV_SEP_26.xlsx"),
                  ("REPD.csv", B+"6a6cbdc00c36759b5ccaa305/REPD_Publication_Q2_2026.csv"),
                  ("DUKES_5.12.xlsx", B+"6a6a3668862aaf18d9c629ed/DUKES_5.12.xlsx"),
                  ("DUKES_5.16.xlsx", B+"6a6a36a0bf4a7aca8f8976c6/DUKES_5.16.xlsx"),
                  ("eirgrid_summary.xlsx", "https://cms.eirgrid.ie/sites/default/files/publications/System-and-Renewable-Data-Summary-Report-V30.xlsx")]:
    r = S.get(url, timeout=240); print(name, r.status_code, len(r.content))
    open(os.path.join(OUT, name), "wb").write(r.content)
