"""Report which of the named environment variables (API keys) are set -
names and lengths only, never values. Usage: check_keys.py NAME1 NAME2 ..."""
import os
import sys

names = sys.argv[1:]
if not names:
    print("Pass the key names to check, e.g. check_keys.py ENTSOE_API_KEY EIA_API_KEY")
for name in names:
    value = os.environ.get(name)
    print(f"{name}: {'set (' + str(len(value)) + ' characters)' if value else 'NOT SET'}")
