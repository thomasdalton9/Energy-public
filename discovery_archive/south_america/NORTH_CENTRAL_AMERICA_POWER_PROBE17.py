"""
Round 17 (Oct-2026), El Salvador. PROBE16 found SIGET's public Power BI
report 'Visualizador dinamico de Estadisticas Electricas' (cluster
wabi-paas-1-scus-api) with model table 'DATOS GENERACION PLANTA (3)':
ID_ANO, ID_MES, ID_RECURSO, ID_TECNOLOGIA, ID_MERCADO, ID_TRANSACCION ... MWh,
i.e. monthly generation per plant. Here: run the anonymous 'querydata' call
the browser makes, dump the dimension tables, and sum MWh by year / month /
resource (and the other IDs) to see coverage and what the dimensions mean.

Usage: python3 NORTH_CENTRAL_AMERICA_POWER_PROBE17.py pbi
"""

print("STARTING", flush=True)

import base64
import collections
import json
import sys

import requests

TOKEN = "eyJrIjoiYmM0MDJhMTktNmQ2YS00ZDkxLTgyYWUtMDUxYjQ5MjQxYWY5IiwidCI6IjE1OTg0YzZmLTNiMmMtNDZmMi1iNmU0LTUzNDQ3MGY2MDVmNyJ9"
BASE = "https://wabi-paas-1-scus-api.analysis.windows.net"
FACT = "DATOS GENERACIÓN PLANTA (3)"
H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36",
     "Accept": "application/json", "Content-Type": "application/json"}


def decode(resp):
    """Power BI DSR (compressed rows) -> (column names, rows)."""
    ds = resp["results"][0]["result"]["data"]["dsr"]["DS"][0]
    vd = ds.get("ValueDicts", {})
    rows = ds["PH"][0]["DM0"]
    if not rows:
        return [], [], ds
    schema = rows[0]["S"]
    out, prev = [], [None] * len(schema)
    for r in rows:
        rep, nul = r.get("R", 0), r.get("Ø", 0)
        c = iter(r.get("C", []))
        cur = []
        for i, s in enumerate(schema):
            if rep >> i & 1:
                v = prev[i]
            elif nul >> i & 1:
                v = None
            else:
                v = next(c)
                if "DN" in s and isinstance(v, int):
                    v = vd[s["DN"]][v]
            cur.append(v)
        out.append(cur)
        prev = cur
    return [s["N"] for s in schema], out, ds


class PBI:
    def __init__(self):
        tok = json.loads(base64.b64decode(TOKEN + "=" * (-len(TOKEN) % 4)))
        self.key = tok["k"]
        self.h = dict(H, **{"X-PowerBI-ResourceKey": self.key})
        j = requests.get(f"{BASE}/public/reports/{self.key}/modelsAndExploration?preferReadOnlySession=true",
                         headers=self.h, timeout=60).json()
        m = j["models"][0]
        self.model_id, self.db = m["id"], m["dbName"]
        self.report = j["exploration"]["report"]["objectId"]
        print("model", self.model_id, self.db, "report", self.report, flush=True)

    def query(self, entity, group_cols, sum_col=None, count=30000):
        sel = [{"Column": {"Expression": {"SourceRef": {"Source": "t"}}, "Property": c}, "Name": f"t.{c}"} for c in group_cols]
        if sum_col:
            sel.append({"Aggregation": {"Expression": {"Column": {"Expression": {"SourceRef": {"Source": "t"}}, "Property": sum_col}},
                                        "Function": 0}, "Name": f"Sum(t.{sum_col})"})
        q = {"Version": 2, "From": [{"Name": "t", "Entity": entity, "Type": 0}], "Select": sel}
        body = {"version": "1.0.0", "cancelQueries": [], "modelId": self.model_id,
                "queries": [{"Query": {"Commands": [{"SemanticQueryDataShapeCommand": {
                    "Query": q,
                    "Binding": {"Primary": {"Groupings": [{"Projections": list(range(len(sel)))}]},
                                "DataReduction": {"DataVolume": 4, "Primary": {"Window": {"Count": count}}}, "Version": 1}}}]},
                    "QueryId": "", "ApplicationContext": {"DatasetId": self.db, "Sources": [{"ReportId": self.report, "VisualId": ""}]}}]}
        r = requests.post(f"{BASE}/public/reports/querydata?synchronous=true", headers=self.h, data=json.dumps(body), timeout=120)
        if not r.ok:
            print("  querydata", r.status_code, r.text[:400], flush=True)
            return [], []
        try:
            names, rows, ds = decode(r.json())
        except (KeyError, IndexError) as e:
            print("  decode failed", e, r.text[:800], flush=True)
            return [], []
        print(f"  {entity} {group_cols} -> {len(rows)} rows, complete={ds.get('IC')}, restart={'RT' in ds}", flush=True)
        return names, rows


def pbi():
    p = PBI()
    dims = {"AÑO (2)": ["ID_AÑO", "N_AÑO"], "MES": ["ID_MES", "N_MES"], "RECURSO": ["ID_RECURSO", "N_RECURSO"],
            "TECNOLOGÍA (3)": ["ID_TECNOLOGÍA", "N_TECNOLOGÍA"], "GENERACIÓN (2)": ["ID_GENERACIÓN", "N_GENERACIÓN"],
            "TRANSACCIONES": ["ID_TRANSACCIÓN", "N_TRANSACCIÓN"], "MERCADO": ["ID_MERCADO", "N_MERCADO"],
            "PÉRDIDAS": ["ID_PÉRDIDAS", "N_PÉRDIDAS"], "TIPO DE GENERADOR (2)": ["ID_TIPO", "N_TIPO GENERADOR"],
            "CATEGORÍA (2)": ["ID_CATEGORÍA", "N_CATEGORIA"]}
    names = {}
    for ent, cols in dims.items():
        _, rows = p.query(ent, cols)
        names[cols[0]] = {r[0]: r[1] for r in rows}
        print(f"   DIM {ent}: {names[cols[0]]}", flush=True)
    ids = ["ID_AÑO", "ID_MES", "ID_RECURSO", "ID_TECNOLOGÍA", "ID_GENERACIÓN", "ID_TRANSACCIÓN", "ID_MERCADO", "ID_PÉRDIDAS",
           "ID_TIPO", "ID_CATEGORÍA"]
    _, rows = p.query(FACT, ids, "MWh")
    if not rows:
        return
    year = lambda r: names["ID_AÑO"].get(r[0], r[0])  # noqa: E731
    by_y = collections.defaultdict(float)
    by_y_res = collections.defaultdict(float)
    for r in rows:
        by_y[year(r)] += r[-1] or 0
        by_y_res[(year(r), names["ID_RECURSO"].get(r[2], r[2]))] += r[-1] or 0
    print("  ANNUAL GWh:", {k: round(v / 1000) for k, v in sorted(by_y.items(), key=lambda x: str(x[0]))}, flush=True)
    for k, v in sorted(by_y_res.items(), key=lambda x: str(x[0])):
        if str(k[0]) >= "2019":
            print(f"   {k[0]} {k[1]}: {v / 1000:,.1f} GWh", flush=True)
    for i, col in enumerate(ids[3:], start=3):
        agg = collections.defaultdict(float)
        for r in rows:
            if str(year(r)) == "2024":
                agg[names.get(col, {}).get(r[i], r[i])] += r[-1] or 0
        print(f"  2024 by {col}: {{{', '.join(f'{k}: {v / 1000:,.0f}' for k, v in agg.items())}}}", flush=True)
    ym = collections.defaultdict(float)
    for r in rows:
        ym[(year(r), r[1])] += r[-1] or 0
    last = sorted(ym, key=lambda x: (str(x[0]), x[1] if isinstance(x[1], int) else 0))[-30:]
    print("  LATEST MONTHS GWh:", [(k, round(ym[k] / 1000, 1)) for k in last], flush=True)
    print("  sample rows:", rows[:5], flush=True)
    _, drows = p.query("DEMANDA totales mes año", ["ID_AÑO", "ID_MES", "ID_TRANSACCIÓN", "ID_RECURSO"], "GWh")
    agg = collections.defaultdict(float)
    for r in drows:
        if str(names["ID_AÑO"].get(r[0], r[0])) == "2024":
            agg[(names["ID_TRANSACCIÓN"].get(r[2], r[2]), names["ID_RECURSO"].get(r[3], r[3]))] += r[-1] or 0
    print("  DEMANDA 2024 by transaccion/recurso GWh:", {k: round(v) for k, v in agg.items()}, flush=True)


if __name__ == "__main__":
    {"pbi": pbi}[sys.argv[1]]()
    print("DONE", flush=True)
