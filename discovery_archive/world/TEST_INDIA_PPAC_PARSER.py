"""
Offline plumbing test for asia/INDIA_PPAC_GAS_BY_SECTOR.py: a SYNTHETIC workbook (made-up numbers, an assumed
header-row-of-months layout) is served through a fake requests session to exercise parsing, the
Last-Modified skip, the workbook write and the add_charts.py registry chart. It does not show what PPAC's
real files look like.

    python3 discovery_archive/world/TEST_INDIA_PPAC_PARSER.py /path/to/scratch/dir
"""
import importlib.util
import io
import os
import sys

import pandas as pd

ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
sys.path.insert(0, ROOT)
spec = importlib.util.spec_from_file_location("ppac", os.path.join(ROOT, "asia", "INDIA_PPAC_GAS_BY_SECTOR.py"))
m = importlib.util.module_from_spec(spec)
spec.loader.exec_module(m)

df = pd.DataFrame([["Sector", "Apr-24", "May-24", "Jun-24", "Jul-24"], ["Fertilizer", 1, 2, 3, 4],
                   ["Power", 5, 6, 7, 8], ["City Gas (CGD)", 9, 9, 9, 9], ["Total", 15, 17, 19, 21]])
buf = io.BytesIO()
df.to_excel(buf, header=False, index=False)


class Resp:
    status_code = 200
    text = '<a href="/files/gas_consumption.xlsx">Consumption</a>'
    content = buf.getvalue()
    headers = {"Last-Modified": "x", "Content-Length": "1"}

    def raise_for_status(self):
        pass


class Session:
    def get(self, u, **k):
        return Resp()

    def head(self, u, **k):
        return Resp()


m.requests.Session = Session
path = os.path.join(sys.argv[1], "india_ppac_gas_by_sector.xlsx")
os.makedirs(sys.argv[1], exist_ok=True)
sys.argv = ["x", "--out", path]
m.main()
m.main()   # second run: file unchanged -> skipped, history kept
import add_charts  # noqa: E402

add_charts.add_charts(path)
print(pd.ExcelFile(path).sheet_names)
