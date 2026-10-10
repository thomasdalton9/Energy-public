"""
Pull China's monthly clean-energy-related output from the National
Bureau of Statistics (NBS), industrial enterprises above designated
size, from 2021 on:
  - solar cells (PV cells) and power generation equipment manufactured
    (GW of capacity), motor vehicles and new energy vehicles produced;
  - electricity GENERATION by source (thermal, hydro, nuclear, wind,
    solar) and total - NBS publishes no wind-turbine manufacturing
    figure, so wind generation is the closest wind metric;
  - coal, coke, cement and natural gas output (coal/gas-adjacent:
    cement kilns are overwhelmingly coal-fired).
Heat pumps are not reported by NBS at all.

Previously parsed from the English "Industrial Production Operation"
release: that list only reaches back to Apr 2024, and its product labels
changed case in Sep 2025 ("Thermal power" -> "Thermal Power"), which left
generation by source blank (except hydro) before Sep 2025. Now read
from the Chinese release's table (same NBS numbers) - see
china_nbs_common.py and china_nbs_output.py.

    python3 asia/CHINA_NBS_CLEAN_ENERGY_PRODUCTS.py --out "output/Data and Chart Outputs/china_nbs_clean_energy_products_monthly.xlsx"
"""

import argparse
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))                   # asia/, for china_nbs_*
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))  # repo root, for xlsx_notes

import china_nbs_output as out  # noqa: E402

REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DEFAULT_OUT = os.path.join(REPO_ROOT, "output", "Data and Chart Outputs",
                           "china_nbs_clean_energy_products_monthly.xlsx")
COLUMNS = ["Solar_Cells_GW", "Power_Generation_Equipment_GW", "New_Energy_Vehicles_k_units", "Motor_Vehicles_k_units",
           "Total_Generation_TWh", "Thermal_Generation_TWh", "Hydro_Generation_TWh", "Nuclear_Generation_TWh",
           "Wind_Generation_TWh", "Solar_Generation_TWh", "Raw_Coal_Mt", "Coke_Mt", "Natural_Gas_Bcm", "Cement_Mt"]

NOTES_LINES = [
    "UNITS",
    "Every value is that calendar month's output. " + out.units_line(COLUMNS) + ". "
    "Solar cells / power generation equipment are manufactured capacity (GW), not generation.",
    "",
] + out.COMMON_NOTES
NOTES_SECTION_TITLES = {"UNITS", "SCOPE", "JANUARY AND FEBRUARY", "Y/Y AND YEAR-TO-DATE COLUMNS", "SOURCE", "UPDATES"}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--out", default=DEFAULT_OUT)
    args = parser.parse_args()
    out.pull(args.out, COLUMNS, NOTES_LINES, NOTES_SECTION_TITLES)


if __name__ == "__main__":
    main()
