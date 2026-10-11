"""
Offline plumbing test for europe/EUROSTAT_GAS_BY_SECTOR.py. The sandbox cannot reach ec.europa.eu, so this
feeds the script SYNTHETIC JSON-stat 2.0 documents (made-up numbers, shaped like Eurostat's response) through
get_json and checks: parsing (sparse dict and list "value"), incremental upsert, workbook write, and the
add_charts.py registry chart build. It proves the code path, NOT the live Eurostat codes/units.

    python3 discovery_archive/world/TEST_EUROSTAT_JSONSTAT.py /path/to/scratch/dir
"""
import importlib.util
import os
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
sys.path.insert(0, ROOT)
spec = importlib.util.spec_from_file_location("eu", os.path.join(ROOT, "europe", "EUROSTAT_GAS_BY_SECTOR.py"))
eu = importlib.util.module_from_spec(spec)
spec.loader.exec_module(eu)

CODES = {"TI_EHG_E": "Transformation input - electricity and heat generation - energy use",
         "FC_IND_E": "Final consumption - industry sector - energy use",
         "FC_OTH_HH_E": "Final consumption - other sectors - households - energy use",
         "FC_OTH_CP_E": "Final consumption - other sectors - commercial and public services - energy use",
         "FC_TRA_E": "Final consumption - transport sector - energy use",
         "GIC": "Gross inland consumption"}
MCODES = {"GID_CAL": "Gross inland deliveries - calendar adjusted", "TI_EHG_E": "Transformation input - power and heat",
          "IMP": "Imports"}


def doc(dataset, geo, unit_list, since):
    monthly = dataset == eu.MONTHLY_DS
    codes = MCODES if monthly else CODES
    if monthly:
        start = since or "2008-01"
        times = [str(p) for p in eu.pd.period_range(start, "2025-12", freq="M")]
    else:
        times = [str(y) for y in range(int(since or 1990), 2024)]
    units = unit_list
    ids = ["freq", "nrg_bal", "siec", "unit", "geo", "time"]
    sizes = [1, len(codes), 1, len(units), 1, len(times)]
    value = {}
    n = 0
    for ci in range(len(codes)):
        for ui in range(len(units)):
            for ti in range(len(times)):
                flat = ((ci * len(units)) + ui) * len(times) + ti
                value[str(flat)] = 100.0 + ci * 10 + ti % 12
                n += 1
    cat = lambda d: {"index": {c: i for i, c in enumerate(d)}, "label": {c: c for c in d}}  # noqa: E731
    dims = {"freq": cat(["M" if monthly else "A"]), "nrg_bal": {"category": {"index": list(codes), "label": codes}},
            "siec": cat(["G3000"]), "unit": cat(units), "geo": cat([geo]), "time": cat(times)}
    dims = {k: (v if "category" in v else {"category": v}) for k, v in dims.items()}
    return {"version": "2.0", "class": "dataset", "id": ids, "size": sizes, "dimension": dims, "value": value}


def fake_get_json(dataset, params):
    if params["geo"] in ("XK", "UA"):
        return None
    units = ["MIO_M3", "TJ_GCV"] if dataset == eu.MONTHLY_DS else ["GWH", "KTOE", "TJ"]
    use = [params["unit"]] if "unit" in params else units
    return doc(dataset, params["geo"], use, params.get("sinceTimePeriod"))


def main():
    scratch = sys.argv[1]
    os.makedirs(scratch, exist_ok=True)
    path = os.path.join(scratch, "eurostat_gas_by_sector.xlsx")
    eu.get_json = fake_get_json
    eu.time.sleep = lambda s: None
    sys.argv = ["x", "--out", path]
    eu.main()
    first = eu.pd.read_excel(path, sheet_name="Annual raw")
    eu.main()   # second run exercises the incremental path
    second = eu.pd.read_excel(path, sheet_name="Annual raw")
    assert len(first) == len(second), (len(first), len(second))
    import add_charts
    add_charts.add_charts(path)
    wb = eu.pd.ExcelFile(path)
    print(wb.sheet_names)


if __name__ == "__main__":
    main()
