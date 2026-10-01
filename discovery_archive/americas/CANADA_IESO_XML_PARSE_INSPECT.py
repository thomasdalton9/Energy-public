"""
Follow-up to CANADA_IESO_XML_INSPECT.py, which only printed raw text
slices (first/last N chars) of the Generator Output and Capability
XML - never actually parsed it, so it's unclear whether each Generator
element carries its own fuel-type tag (needed to aggregate to a
fuel-mix series) or whether that requires cross-referencing a separate
report. Parsing the real XML tree and printing one full Generator
element plus a summary of every distinct child tag name seen across
all Generators, to settle that before building a production script.
"""
import sys
import xml.etree.ElementTree as ET

import requests

HEADERS = {"User-Agent": "Mozilla/5.0"}
TIMEOUT = (10, 45)
URL = "https://reports-public.ieso.ca/public/GenOutputCapability/PUB_GenOutputCapability.xml"

r = requests.get(URL, headers=HEADERS, timeout=TIMEOUT)
print(f"status={r.status_code} bytes={len(r.content)}", file=sys.stderr)

root = ET.fromstring(r.content)


def strip_ns(tag):
    return tag.split("}")[-1] if "}" in tag else tag


def local_findall(elem, tag):
    return [child for child in elem if strip_ns(child.tag) == tag]


print(f"root tag: {strip_ns(root.tag)}", file=sys.stderr)


def walk(elem, path="", depth=0, max_depth=4):
    if depth > max_depth:
        return
    tag = strip_ns(elem.tag)
    print("  " * depth + tag, file=sys.stderr)
    seen_children = set()
    for child in elem:
        ctag = strip_ns(child.tag)
        if ctag not in seen_children:
            seen_children.add(ctag)
            walk(child, path + "/" + ctag, depth + 1, max_depth)


walk(root)

# Find the Generators container and print one full generator's direct children/tags.
generators_container = None
for elem in root.iter():
    if strip_ns(elem.tag) == "Generators":
        generators_container = elem
        break

if generators_container is not None:
    gens = local_findall(generators_container, "Generator")
    print(f"\n{len(gens)} Generator elements found", file=sys.stderr)
    if gens:
        first = gens[0]
        print("\nFirst Generator's direct child tags (with text for simple ones):", file=sys.stderr)
        for child in first:
            ctag = strip_ns(child.tag)
            text = (child.text or "").strip()
            if text:
                print(f"  {ctag}: {text!r}", file=sys.stderr)
            else:
                grandchildren = sorted(set(strip_ns(gc.tag) for gc in child))
                print(f"  {ctag} (container, {len(list(child))} children, tags: {grandchildren})",
                      file=sys.stderr)

        fuel_types = sorted(set(
            (g.find(".//{*}FuelType").text if g.find(".//{*}FuelType") is not None else None)
            for g in gens
        ))
        print(f"\ndistinct FuelType values across all generators: {fuel_types}", file=sys.stderr)

        names_sample = [(g.find(".//{*}GeneratorName").text if g.find(".//{*}GeneratorName") is not None else None,
                          g.find(".//{*}FuelType").text if g.find(".//{*}FuelType") is not None else None)
                         for g in gens[:10]]
        print(f"\nfirst 10 (name, fuel type) pairs: {names_sample}", file=sys.stderr)
