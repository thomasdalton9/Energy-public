"""
Pull China's monthly ENERGY production from the National Bureau of
Statistics (NBS), industrial enterprises above designated size, from
2021 on: raw coal, coke, crude oil, crude oil processed (refinery runs),
natural gas, and electricity generation - total and by source (thermal,
hydro, nuclear, wind, solar).

Previously parsed from the English "Energy Production in {Month}
{Year}" press release (prose, rounded figures, English list only reaches
back to Apr 2024, no generation by source). Now read from the Chinese
monthly industrial production release's product-output table - same
NBS numbers, unrounded, with generation by source and history back to
late 2021. NBS's Chinese "能源生产情况" release itself has no table (prose
with the same monthly figures plus daily averages = monthly / days in
month), so it adds nothing pullable. See china_nbs_common.py and
china_nbs_output.py for the source and parsing.

    python3 asia/CHINA_NBS_ENERGY_PRODUCTION.py --out "output/Data and Chart Outputs/china_nbs_energy_production_monthly.xlsx"
"""

import argparse
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))                   # asia/, for china_nbs_*
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))  # repo root, for xlsx_notes

import china_nbs_output as out  # noqa: E402

REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DEFAULT_OUT = os.path.join(REPO_ROOT, "output", "Data and Chart Outputs", "china_nbs_energy_production_monthly.xlsx")
COLUMNS = ["Raw_Coal_Mt", "Coke_Mt", "Crude_Oil_Mt", "Crude_Oil_Processing_Mt", "Natural_Gas_Bcm",
           "Total_Generation_TWh", "Thermal_Generation_TWh", "Hydro_Generation_TWh", "Nuclear_Generation_TWh",
           "Wind_Generation_TWh", "Solar_Generation_TWh"]

NOTES_LINES = [
    "UNITS",
    "Every value is that calendar month's output. " + out.units_line(COLUMNS) + ". "
    "Daily averages (as quoted in NBS's energy production release) = monthly value / days in the month.",
    "",
] + out.COMMON_NOTES
NOTES_SECTION_TITLES = {"UNITS", "SCOPE", "JANUARY AND FEBRUARY", "SOURCE", "UPDATES"}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--out", default=DEFAULT_OUT)
    args = parser.parse_args()
    out.pull(args.out, COLUMNS, NOTES_LINES, NOTES_SECTION_TITLES)


if __name__ == "__main__":
    main()
