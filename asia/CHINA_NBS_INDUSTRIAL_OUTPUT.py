"""
Pull EVERY product in China NBS's monthly "Output of Major Industrial
Products" table (规模以上工业主要产品产量, industrial enterprises above
designated size) into one workbook, from 2021 on:

  energy      raw coal, coke, crude oil, crude oil processed (refinery
              runs), natural gas, electricity generation total and by
              source (thermal, hydro, nuclear, wind, solar)
  industry    crude steel, pig iron, finished steel, cement, plate glass,
              ten non-ferrous metals, primary aluminium, ethylene,
              sulfuric acid, caustic soda, chemical fibre, cloth
  equipment   solar cells, power generation equipment, motor vehicles,
              cars, SUVs, new energy vehicles, metal-cutting machine
              tools, industrial and service robots, microcomputers,
              mobile phones, smartphones, integrated circuits

Source, history and parsing: see china_nbs_common.py / china_nbs_output.py
(Chinese release list www.stats.gov.cn/sj/zxfb/, which reaches back to
late 2021; data.stats.gov.cn is blocked to automated requests).

    python3 asia/CHINA_NBS_INDUSTRIAL_OUTPUT.py --out "output/Data and Chart Outputs/china_nbs_industrial_output_monthly.xlsx"
"""

import argparse
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))                   # asia/, for china_nbs_*
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))  # repo root, for xlsx_notes

import china_nbs_output as out  # noqa: E402

REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DEFAULT_OUT = os.path.join(REPO_ROOT, "output", "Data and Chart Outputs", "china_nbs_industrial_output_monthly.xlsx")
COLUMNS = [p[1] for p in out.PRODUCTS]

NOTES_LINES = [
    "UNITS",
    "Every value is that calendar month's output. " + out.units_line(COLUMNS) + ".",
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
