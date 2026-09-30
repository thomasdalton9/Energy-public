"""
Your API keys/credentials, as plain Python constants - edit this file
directly and save; every script that needs a key imports from here.

This file is in .gitignore, so it never gets committed or pushed to
GitHub - only you have it, on your own machine. Keep it that way: do
not remove it from .gitignore, and never paste real key values into
any other file in this repo (they'd end up in git history).

Leave a value as "" for any service you don't have a key for yet -
the script that needs it will tell you clearly what's missing and
where to get one when you run it.
"""

ENTSOE_API_KEY = ""       # https://transparency.entsoe.eu - My Account Settings -> API access
GIE_API_KEY = ""          # https://agsi.gie.eu/#/registration - same key works for AGSI and ALSI
OPENELECTRICITY_TOKEN = ""  # https://openelectricity.org.au
TR_USERNAME = ""          # EPIAS Transparency Platform (Turkey)
TR_PASSWORD = ""
EIA_API_KEY = ""          # https://www.eia.gov/opendata/register.php - used by all the americas/ scripts (EIA-930, storage, demand, production, ERCOT)
EMBER_API_KEY = ""        # https://ember-energy.org/data/ember-api/ - used by the europe/ Ember_*.py scripts
CEN_USER_KEY = ""         # https://sipub.coordinador.cl - Chile hydro reservoir cota (CHILE_CEN_HYDRO.py)
DATAGOVINDIA_API_KEY = "" # https://www.data.gov.in - My Account -> Generate API Key - India gas demand by sector (india_gas_demand_by_sector.py)
FRED_API_KEY = ""         # https://fred.stlouisfed.org/docs/api/api_key.html - free, instant - Dallas Fed Energy Survey breakeven prices (PERMIAN_BREAKEVEN_DALLAS_FED.py)
