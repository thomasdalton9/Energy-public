"""Step 5 (GitHub Actions): RALIE unit-level history -> forward schedule per snapshot, split by construction
status and viability, so the unfiltered 180 GW 'expected within 12 months' can be reduced to what the deck
would count. Writes discovery_archive/pmo_deck/ralie_forward_by_status.csv and prints the category counts."""
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
b = get(base + "896a51b2-6d40-4b0a-b2b2-b460f6a5b7ed/download/ralie-unidade-geradora-historico.parquet"); R = pd.read_parquet(io.BytesIO(b))
print("rows", len(R)); print("columns:", list(R.columns))
for c in [c for c in R.columns if c.lower().startswith("dat")]: R[c] = pd.to_datetime(R[c], errors="coerce")
snap = "DatRalie" if "DatRalie" in R.columns else [c for c in R.columns if "ralie" in c.lower()][0]
exp = [c for c in R.columns if re.search(r"previs", c, re.I) and c.lower().startswith("dat")]; print("expected-date cols:", exp)
pw = [c for c in R.columns if re.search(r"potencia", c, re.I)][0]; R[pw] = pd.to_numeric(R[pw].astype(str).str.replace(".", "", regex=False).str.replace(",", ".", regex=False), errors="coerce")
cat = [c for c in R.columns if re.search(r"situacao|viabilidade|cronograma|fase|status", c, re.I)]; print("status columns:", cat)
last = R[R[snap] == R[snap].max()]
for c in cat:
    print(f"\n{c} (latest snapshot, MW by category):"); print((last.groupby(c)[pw].sum() / 1000).round(0).sort_values(ascending=False).head(12).to_string())
src = [c for c in R.columns if re.search(r"origem|tipogeracao|fonte", c, re.I)]; sc = src[0] if src else None
obra = [c for c in cat if "obra" in c.lower()]; viab = [c for c in cat if "viabil" in c.lower()]
ec = exp[0]; out = []
for s, d in R.dropna(subset=[snap, ec]).groupby(snap):
    lead = (d[ec] - s).dt.days; w12 = (lead >= 0) & (lead <= 365); w24 = (lead >= 0) & (lead <= 730)
    row = dict(snapshot=s.date(), n_units=len(d), all_12m=d.loc[w12, pw].sum() / 1000, all_24m=d.loc[w24, pw].sum() / 1000)
    if obra:
        oc = obra[0]; st = d[oc].map(norm).str.lower()
        started = st.str.contains(r"andamento|iniciad|conclu|comission|operac", regex=True)
        row["started_12m"] = d.loc[w12 & started, pw].sum() / 1000; row["started_24m"] = d.loc[w24 & started, pw].sum() / 1000
        if sc:
            for k, v in d.loc[w12 & started].groupby(sc)[pw].sum().items(): row["started12_" + re.sub(r"\W+", "_", norm(k))[:12]] = v / 1000
    if viab:
        vc = viab[0]; vv = d[vc].map(norm).str.lower(); ok = vv.str.contains(r"viavel|viab|alta|media", regex=True) & ~vv.str.contains(r"inviavel|baixa|sem", regex=True)
        row["viable_12m"] = d.loc[w12 & ok, pw].sum() / 1000; row["viable_24m"] = d.loc[w24 & ok, pw].sum() / 1000
    out.append(row)
O = pd.DataFrame(out); O.to_csv(f"{OUT}/ralie_forward_by_status.csv", index=False); print("\nforward schedule rows:", len(O)); print(O.head(4).to_string()[:2500]); print(O.tail(4).to_string()[:2500])
print("DONE")
