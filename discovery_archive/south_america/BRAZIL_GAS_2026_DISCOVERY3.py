"""
Round 3 (see BRAZIL_GAS_2026_DISCOVERY.py / 2). No public source of national
demand BY END-USE SEGMENT exists after MME's annex (2025-06): the 2026
bulletin pages 404, the Observatorio "Boletim do Gas" is login-gated, ABEGAS
publishes annual totals only, EPE has no monthly gas data. What IS current
(to Aug-2026) is ANP open data:
  - "Movimentacao de Gas Natural em Gasodutos de Transporte": one CSV per
    month, daily values per receipt/delivery point, per variable;
  - production by state (gas disponivel / consumo proprio / queima);
  - gas imports (total).
This round works out how to turn the movement CSVs into monthly supply by
source and grid deliveries by consumer type:
  1. variables, units, delivery/receipt point names and Aug-2026 / Jun-2025
     monthly means of "Volume Realizado" per point;
  2. the ANP Power BI unit dimension (DM_SIMP_UNIDADE: point name, type,
     classification, consumer type, codes) to classify points;
  3. heads/tails of the production and import CSVs.
"""
import base64
import io
import json
import re

import pandas as pd
import requests

H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0 Safari/537.36"}
T = (10, 120)
S = requests.Session()
S.headers.update(H)
pd.set_option("display.width", 250)
pd.set_option("display.max_rows", 500)
pd.set_option("display.max_colwidth", 45)


def out(*a):
    print(*a, flush=True)


MOV = ("https://www.gov.br/anp/pt-br/centrais-de-conteudo/dados-abertos/arquivos/"
       "arquivos-movimentacao-de-gas-natural-em-gasodutos-de-transporte/")


def read_mov(path):
    r = S.get(MOV + path, timeout=T)
    out(f"GET {path} -> {r.status_code} {len(r.content)}B")
    c = r.content
    txt = c.decode("utf-8-sig") if b"\xc3" in c[:50000] else c.decode("latin-1")
    df = pd.read_csv(io.StringIO(txt), sep=";", dtype=str)
    days = [x for x in df.columns if re.match(r"\d{2}/\d{2}/\d{4}$", x)]
    meta = [x for x in df.columns if x not in days]
    out(f"  cols meta={meta}\n  {len(days)} day columns {days[:1]}..{days[-1:]}")
    vals = df[days].apply(lambda s: pd.to_numeric(s.str.replace(".", "", regex=False).str.replace(",", ".", regex=False),
                                                  errors="coerce"))
    df["mean"] = vals.mean(axis=1)
    df["n"] = vals.notna().sum(axis=1)
    return df, meta


for path in ["2026/gn_agosto_2026.csv", "2025/gn_junho_2025.csv"]:
    out(f"\n######## 1. {path} ########")
    df, meta = read_mov(path)
    var = [c for c in meta if "vari" in c.lower()][0]
    tipo = [c for c in meta if c.startswith("Tipo")][0]
    out(df[var].value_counts().to_string())
    v = df[df[var].str.strip() == "Volume Realizado"]
    out(f"\n  Volume Realizado rows: {len(v)}; sum of means by {tipo}:")
    out(v.groupby(tipo, dropna=False)["mean"].sum().round(1).to_string())
    keys = [c for c in meta if c.startswith(("Nome da Instala", "Código da Instalação de Gasoduto", "Nome da UF", "Nome do Carregador"))]
    g = v.groupby([tipo] + keys, dropna=False)["mean"].sum().reset_index().sort_values([tipo, "mean"], ascending=[True, False])
    out(g.round(1).to_string(index=False))

out("\n######## 2. ANP Power BI unit dimension ########")
page = S.get("https://www.gov.br/anp/pt-br/centrais-de-conteudo/paineis-dinamicos-da-anp/"
             "painel-dinamico-de-movimentacao-de-gas-natural-em-gasodutos-de-transporte", timeout=T).text
tok = re.findall(r"app\.powerbi\.com/view\?r=([A-Za-z0-9=_%-]+)", page)[0]
cfg = json.loads(base64.b64decode(tok + "=" * (-len(tok) % 4)))
k = cfg["k"]
BASE = "https://wabi-brazil-south-api.analysis.windows.net"
hdr = {"X-PowerBI-ResourceKey": k, "Content-Type": "application/json;charset=UTF-8"}
me = S.get(f"{BASE}/public/reports/{k}/modelsAndExploration?preferReadOnlySession=true", headers=hdr, timeout=T).json()
model = me["models"][0]
out(f"model id {model['id']} dbName {model.get('dbName')}; report {me['exploration']['report'].get('objectId')}")


def dsr_rows(resp, ncols):
    """Decode a Power BI DSR (data shape result) table into a list of rows."""
    ds = resp["results"][0]["result"]["data"]["dsr"]["DS"][0]
    dicts = ds.get("ValueDicts", {})
    rows, prev, schema = [], [None] * ncols, None
    for r in ds["PH"][0].get("DM0", []):
        if "S" in r:
            schema = r["S"]
        c = list(r.get("C", []))
        rep = r.get("R", 0)
        nul = r.get("Ø", 0)
        row = []
        for i in range(ncols):
            if rep >> i & 1:
                row.append(prev[i])
            elif nul >> i & 1:
                row.append(None)
            else:
                val = c.pop(0) if c else None
                dn = schema[i].get("DN") if schema and i < len(schema) else None
                if dn and isinstance(val, int):
                    val = dicts[dn][val]
                row.append(val)
        prev = row
        rows.append(row)
    return rows


def query(entity, props, filt=None, count=30000):
    sel = [{"Column": {"Expression": {"SourceRef": {"Source": "u"}}, "Property": p}, "Name": f"{entity}.{p}"} for p in props]
    q = {"Version": 2, "From": [{"Name": "u", "Entity": entity, "Type": 0}], "Select": sel}
    body = {"version": "1.0.0", "queries": [{"Query": {"Commands": [{"SemanticQueryDataShapeCommand": {
        "Query": q, "Binding": {"Primary": {"Groupings": [{"Projections": list(range(len(props)))}]},
                                "DataReduction": {"DataVolume": 4, "Primary": {"Window": {"Count": count}}},
                                "Version": 1}}}]},
        "ApplicationContext": {"DatasetId": model["dbName"],
                               "Sources": [{"ReportId": me["exploration"]["report"]["objectId"]}]}}],
            "cancelQueries": [], "modelId": model["id"]}
    r = S.post(f"{BASE}/public/reports/querydata?synchronous=true", headers=hdr, data=json.dumps(body), timeout=T)
    out(f"  querydata {entity} -> {r.status_code} {len(r.content)}B")
    return dsr_rows(r.json(), len(props))


try:
    props = ["NK_COD_UNIDADE", "COD_PBLCO_UNIDADE", "DSC_UNIDADE", "DSC_TIPO_UNIDADE", "DSC_CLASSIFICACAO",
             "DSC_TIPO_CONSUMIDOR", "SIG_UF", "UNIDADE_ATIVA"]
    rows = query("DM_SIMP_UNIDADE", props)
    u = pd.DataFrame(rows, columns=props)
    out(f"  {len(u)} units")
    out(u.groupby(["DSC_TIPO_UNIDADE", "DSC_CLASSIFICACAO", "DSC_TIPO_CONSUMIDOR"], dropna=False).size().to_string())
    out(u.sort_values(["DSC_TIPO_UNIDADE", "DSC_CLASSIFICACAO", "DSC_UNIDADE"]).to_string(index=False))
except Exception as e:
    out(f"  PBI ERR {type(e).__name__} {str(e)[:300]}")
try:
    vrows = query("DM_CMGN_VARIAVEL_ALARME", ["NOM_VARIAVEL", "NOM_VAR_UNID_MED", "SIG_UNIDADE_MEDIDA"])
    out(pd.DataFrame(vrows, columns=["var", "var_unit", "unit"]).to_string(index=False))
except Exception as e:
    out(f"  PBI vars ERR {type(e).__name__} {str(e)[:300]}")

out("\n######## 3. ANP production / import CSVs ########")
DA = "https://www.gov.br/anp/pt-br/centrais-de-conteudo/dados-abertos/"
r = S.get(DA + "producao-de-petroleo-e-gas-natural-por-estado-e-localizacao", timeout=T)
links = sorted(set(re.findall(r'href="([^"]+\.csv)"', r.text)))
out("\n".join("  " + l for l in links))
for u_ in [l for l in links if re.search(r"gn|gas", l, re.I)] + [DA + "arquivos/ie/gn/importacao-gas-natural-2000-2025.csv"]:
    try:
        c = S.get(u_, timeout=T).content
        txt = c.decode("utf-8-sig") if b"\xc3" in c[:50000] else c.decode("latin-1")
        d = pd.read_csv(io.StringIO(txt), sep=";", dtype=str)
        out(f"\n  {u_.rsplit('/', 1)[-1]} shape {d.shape} cols {list(d.columns)}")
        for col in d.columns[:6]:
            if d[col].nunique() < 15:
                out(f"    distinct {col}: {sorted(d[col].dropna().unique())}")
        out("    " + d.tail(4).to_string(index=False).replace("\n", "\n    "))
    except Exception as e:
        out(f"  {u_} ERR {type(e).__name__} {str(e)[:200]}")
out("\ndone")
