"""
ERCOT natural gas burn for power (Bcf/d): ERCOT's daily gas MWh (EIA-930, eia930_fuel_mix_daily.xlsx sheet ERCOT) x a heat
rate calibrated on EIA-923 (balancing authority ERCO, fuel NG). Same method and code as MISO_GAS_BURN.py (see its docstring).
"""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from MISO_GAS_BURN import main  # noqa: E402

if __name__ == "__main__":
    main(default_region="ERCOT")
