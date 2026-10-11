"""
West Africa electricity generation by fuel from Ember (FALLBACK: no raw national feed is used here).

Reuses south_america/EMBER_POWER_BY_TYPE.py (Ember's whole-file releases, no key, re-downloaded each run - about a
minute). Ember's MONTHLY release covers only Nigeria in West Africa; every other country is in Ember's YEARLY
release only (2000-2024), so those go to the annual workbook. Raw feeds for Nigeria and Ghana are handled
separately; their Ember series are kept here as a labelled fallback.

  west_africa_power_by_type.xlsx          monthly (Nigeria)
  west_africa_power_by_type_annual.xlsx   yearly, all countries below

Usage: python3 africa/WEST_AFRICA_EMBER.py [monthly|annual|both]
"""
import os
import runpy
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
SCRIPT = os.path.join(os.path.dirname(HERE), "south_america", "EMBER_POWER_BY_TYPE.py")
OUT = "output/Data and Chart Outputs/"

# Ember area name : sheet label
COUNTRIES = ("Senegal,Cote d'Ivoire,Mauritania,Mali,Burkina Faso,Guinea,Sierra Leone,Liberia,Benin,Togo,"
             "Niger (the):Niger,Gambia (the):Gambia,Guinea-Bissau,Cabo Verde:Cape Verde,Cameroon,"
             "Equatorial Guinea,Nigeria,Ghana")
MONTHLY_COUNTRIES = "Nigeria"

SOURCES = {
    "west_africa_power_by_type.xlsx": ("Ember monthly electricity data (fallback: Nigeria only; raw NISO feed handled separately)",
                                       "https://ember-energy.org/data/monthly-electricity-data/"),
    "west_africa_power_by_type_annual.xlsx": ("Ember yearly electricity data (fallback: no monthly Ember data for these countries)",
                                              "https://ember-energy.org/data/yearly-electricity-data/"),
}


def run(args):
    sys.argv = [SCRIPT] + args
    runpy.run_path(SCRIPT, run_name="__main__")


if __name__ == "__main__":
    mode = sys.argv[1] if len(sys.argv) > 1 else "both"
    if mode in ("monthly", "both"):
        run(["--countries", MONTHLY_COUNTRIES, "--out", OUT + "west_africa_power_by_type.xlsx"])
    if mode in ("annual", "both"):
        run(["--yearly", "--start-date", "2000-01-01", "--countries", COUNTRIES,
             "--out", OUT + "west_africa_power_by_type_annual.xlsx"])
