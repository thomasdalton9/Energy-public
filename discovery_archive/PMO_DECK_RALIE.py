"""Step 6 (GitHub Actions): RALIE forward schedule by construction status.
Unit-level history carries the expected commercial-operation date (DatPrevisaoOpComercialSFG) and unit MW;
plant-level history carries DscSituacaoObra / DscViabilidade / DscSituacaoCronograma. Join on (DatRalie, CodCEG)
and aggregate per snapshot: MW expected within 12 / 24 months, all projects vs projects with works started
vs 'on schedule', by source. Writes discovery_archive/pmo_deck/ralie_forward_by_status.csv."""
import re, io, os, time, requests, pandas as pd, numpy as np, unicodedata
UA = {"User-Agent": "Mozilla/5.0 (X11; Linux x86_64) energy-data-probe/1.0"}
OUT = "discovery_archive/pmo_deck"; os.makedirs(OUT, exist_ok=True)
def norm(s): return unicodedata.normalize("NFKD", str(s)).encode("ascii", "ignore").decode()
def get(url, tries=3):
    for i in range(tries):
        try: r = requests.get(url, headers=UA, timeout=600, stream=True); r.raise_for_status(); return r.content
        except Exception as e: print("retry", i + 1, repr(e)[:120]); time.sleep(10 * (i + 1))
    raise RuntimeError("download failed")
base = "https://dadosabertos.aneel.gov.br/dataset/57e4b8b5-a5db-40e6-9901-27ca629d0477/resource/"
P = pd.read_parquet(io.BytesIO(get(base + "4cd32195-55fb-4f3f-83e9-2168369e818d/download/ralie-usina-historico.parquet")))
print("plant rows", len(P)); print("plant columns:", list(P.columns))
U = pd.read_parquet(io.BytesIO(get(base + "896a51b2-6d40-4b0a-b2b2-b460f6a5b7ed/download/ralie-unidade-geradora-historico.parquet")),
                    columns=["DatRalie", "CodCEG", "DscOrigemCombustivel", "SigTipoGeracao", "MdaPotenciaUnitaria", "DatPrevisaoOpComercialSFG", "DatUGInicioOpComerOutorgado", "DatLiberOpTesteRealizado"])
print("unit rows", len(U))
for df in (P, U):
    for c in [c for c in df.columns if c.lower().startswith("dat")]: df[c] = pd.to_datetime(df[c], errors="coerce")
U["mw"] = pd.to_numeric(U.MdaPotenciaUnitaria.astype(str).str.replace(".", "", regex=False).str.replace(",", ".", regex=False), errors="coerce") / 1000
scols = [c for c in P.columns if re.search(r"SituacaoObra|Viabilidade|SituacaoCronograma", c)]
print("status columns:", scols)
last = P[P.DatRalie == P.DatRalie.max()]
for c in scols: print(f"\n{c} (latest plant snapshot, plants by category):"); print(last[c].value_counts().head(10).to_string())
S = P[["DatRalie", "CodCEG"] + scols].drop_duplicates(["DatRalie", "CodCEG"])
M = U.merge(S, on=["DatRalie", "CodCEG"], how="left"); print("units with status:", M[scols[0]].notna().mean().round(3) if scols else "none")
M = M[M.DatLiberOpTesteRealizado.isna()]   # units not yet released for test operation = still to come
obra = M[[c for c in scols if "Obra" in c][0]].map(norm).str.lower() if any("Obra" in c for c in scols) else pd.Series("", index=M.index)
cron = M[[c for c in scols if "Cronograma" in c][0]].map(norm).str.lower() if any("Cronograma" in c for c in scols) else pd.Series("", index=M.index)
M["started"] = obra.str.contains(r"andamento|iniciad|conclu|avanc|montagem|comission", regex=True)
M["onschedule"] = cron.str.contains(r"dentro|no prazo|adiantad|conforme", regex=True)
M["src"] = M.DscOrigemCombustivel.map(lambda x: re.sub(r"\W+", "_", norm(x))[:10])
out = []
for s, d in M.dropna(subset=["DatRalie", "DatPrevisaoOpComercialSFG"]).groupby("DatRalie"):
    lead = (d.DatPrevisaoOpComercialSFG - s).dt.days; w12 = (lead >= 0) & (lead <= 365); w24 = (lead >= 0) & (lead <= 730)
    row = dict(snapshot=s.date(), n_units=len(d), all_12m=d.loc[w12, "mw"].sum(), all_24m=d.loc[w24, "mw"].sum(), started_12m=d.loc[w12 & d.started, "mw"].sum(), started_24m=d.loc[w24 & d.started, "mw"].sum(), onsched_12m=d.loc[w12 & d.onschedule, "mw"].sum())
    for k, v in d.loc[w12 & d.started].groupby("src").mw.sum().items(): row["started12_" + k] = v
    for k, v in d.loc[w12].groupby("src").mw.sum().items(): row["all12_" + k] = v
    out.append(row)
O = pd.DataFrame(out).sort_values("snapshot"); O.to_csv(f"{OUT}/ralie_forward_by_status.csv", index=False)
print("\nforward schedule rows:", len(O)); pd.set_option("display.width", 250); print(O[["snapshot", "n_units", "all_12m", "started_12m", "onsched_12m", "all_24m", "started_24m"]].iloc[::12].round(1).to_string())
print("DONE")
