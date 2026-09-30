"""
Follow-up to ARGENTINA_GAS_DISCOVERY.py: that script's search results got
truncated in the job log before every series id could be read (each hit
repeats the full dataset/publisher block, so a handful of hits fills the
log quickly). All the fields we need belong to one dataset ("Produccion
y consumo de gas natural" - Secretaria de Energia, Ministerio de
Economia), so this searches for each expected column name individually
and prints ONLY id/title/frequency/date range per hit (no repeated
dataset metadata) to stay well under any truncation limit, then filters
to the dataset by publisher name to avoid pulling in unrelated series
that happen to share a word.

Usage: python3 ARGENTINA_GAS_DISCOVERY2.py
"""
import sys

import requests

HEADERS = {"User-Agent": "gas-demand-scripts/1.0"}
TIMEOUT = (10, 60)
SEARCH_API = "https://apis.datos.gob.ar/series/api/search/"
DATASET_TITLE = "Producción y consumo de gas natural"

# Every column name we know exists in the underlying dataset (from the
# actividad-gas.csv column list found earlier this session), used one at
# a time as a search query since the API doesn't expose a plain
# "list every series in dataset X" call.
QUERIES = [
    "produccion de gas natural", "residencial", "comercial", "entes oficiales",
    "industria gas", "centrales electricas gas", "sdb", "gnc",
    "metrogas", "gas natural fenosa", "distrib gas del centro ecogas",
    "distrib gas cuyana ecogas", "litoral gas", "gasnea", "redengas",
    "gasnor", "camuzzi gas",
]


def search(q):
    r = requests.get(SEARCH_API, headers=HEADERS, timeout=TIMEOUT, params={"q": q, "limit": 20})
    r.raise_for_status()
    return r.json().get("data", [])


def main():
    found = {}  # title (column name) -> list of (id, frequency, start, end)
    for q in QUERIES:
        hits = search(q)
        for hit in hits:
            field = hit.get("field", {})
            dataset = hit.get("dataset", {})
            if dataset.get("title") != DATASET_TITLE:
                continue
            key = field.get("title")
            found.setdefault(key, []).append(
                (field.get("id"), field.get("frequency"), field.get("time_index_start"), field.get("time_index_end"))
            )
        print(f"query {q!r}: {len(hits)} hits total", file=sys.stderr)

    print("\n=== matched series in the target dataset, by column name ===", file=sys.stderr)
    for title, entries in sorted(found.items()):
        for sid, freq, start, end in entries:
            print(f"  {title:45s} {sid:35s} {freq:8s} {start} - {end}", file=sys.stderr)


if __name__ == "__main__":
    main()
