"""Step 7 (GitHub Actions): RALIE forward schedule by construction status, with by-source splits for every
status class. Unit-level history has the expected commercial-operation date and unit MW; plant-level history
has DscSituacaoObra / DscViabilidade / DscSituacaoCronograma. Join on (DatRalie, CodCEG). Also prints the
category labels so the status classes can be read. Writes discovery_archive/pmo_deck/ralie_forward_by_status.csv."""
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
U = pd.read_parquet(io.BytesIO(get(base + "896a51b2-6d40-4b0a-b2b2-b460f6a5b7ed/download/ralie-unidade-geradora-historico.parquet")),
                    columns=["DatRalie", "CodCEG", "DscOrigemCombustivel", "MdaPotenciaUnitaria", "DatPrevisaoOpComercialSFG", "DatLiberOpTesteRealizado"])
for df in (P, U):
    for c in [c for c in df.columns if c.lower().startswith("dat")]: df[c] = pd.to_datetime(df[c], errors="coerce")
U["mw"] = pd.to_numeric(U.MdaPotenciaUnitaria.astype(str).str.replace(".", "", regex=False).str.replace(",", ".", regex=False), errors="coerce") / 1000
scols = [c for c in P.columns if re.search(r"SituacaoObra|Viabilidade|SituacaoCronograma", c)]; print("status columns:", scols)
last = P[P.DatRalie == P.DatRalie.max()]
for c in scols: print(f"\n{c} (latest snapshot, plants by label):"); print(last[c].value_counts(dropna=False).head(12).to_string())
S = P[["DatRalie", "CodCEG"] + scols].drop_duplicates(["DatRalie", "CodCEG"])
M = U.merge(S, on=["DatRalie", "CodCEG"], how="left"); M = M[M.DatLiberOpTesteRealizado.isna()]
ob = M[[c for c in scols if "Obra" in c][0]].map(norm).str.lower().fillna("")
cr = M[[c for c in scols if "Cronograma" in c][0]].map(norm).str.lower().fillna("")
vi = M[[c for c in scols if "Viabilidade" in c][0]].map(norm).str.lower().fillna("") if any("Viabilidade" in c for c in scols) else pd.Series("", index=M.index)
print("\nobra labels:", ob.value_counts().head(8).to_dict()); print("cronograma labels:", cr.value_counts().head(8).to_dict()); print("viabilidade labels:", vi.value_counts().head(8).to_dict())
M["cls_started"] = ob.str.contains(r"andamento|iniciad|conclu|avanc|montagem|comission", regex=True) & ~ob.str.contains(r"nao iniciad|n.o iniciad|sem previs|paralis|suspens", regex=True)
M["cls_notstarted"] = ob.str.contains(r"nao iniciad|n.o iniciad|sem previs", regex=True)
M["cls_onsched"] = cr.str.contains(r"dentro|no prazo|adiantad|conforme|sem atraso", regex=True)
M["cls_viable"] = vi.str.contains(r"alta|media|viavel", regex=True) & ~vi.str.contains(r"baixa|inviavel|sem", regex=True)
M["src"] = M.DscOrigemCombustivel.map(lambda x: re.sub(r"\W+", "_", norm(x))[:10])
out = []
for s, d in M.dropna(subset=["DatRalie", "DatPrevisaoOpComercialSFG"]).groupby("DatRalie"):
    lead = (d.DatPrevisaoOpComercialSFG - s).dt.days; w12 = (lead >= 0) & (lead <= 365); w24 = (lead >= 0) & (lead <= 730)
    row = dict(snapshot=s.date(), n_units=len(d))
    for cls, mask in [("all", pd.Series(True, index=d.index)), ("started", d.cls_started), ("notstarted", d.cls_notstarted), ("onsched", d.cls_onsched), ("viable", d.cls_viable)]:
        row[f"{cls}_12m"] = d.loc[w12 & mask, "mw"].sum(); row[f"{cls}_24m"] = d.loc[w24 & mask, "mw"].sum()
        for k, v in d.loc[w12 & mask].groupby("src").mw.sum().items(): row[f"{cls}12_{k}"] = v
    out.append(row)
O = pd.DataFrame(out).sort_values("snapshot"); O.to_csv(f"{OUT}/ralie_forward_by_status.csv", index=False)
pd.set_option("display.width", 250); print("\nrows:", len(O)); print(O[["snapshot", "all_12m", "started_12m", "notstarted_12m", "onsched_12m", "viable_12m"]].iloc[::12].round(0).to_string())
print("DONE")
