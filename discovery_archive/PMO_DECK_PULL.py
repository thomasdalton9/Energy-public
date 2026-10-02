"""Step 4 (GitHub Actions): build the deck's forward view from public sources.
  A. EPE/ONS/CCEE load forecasts: every 'Previsoes (mensais) de energia' PDF and spreadsheet on EPE's
     quadrimestral-review page -> discovery_archive/pmo_deck/epe_load_forecasts.csv
     columns: revision_file, plan_years, revision_kind (PLAN / 1RQ / 2RQ / extra), mmgd (com/sem/na),
              subsystem, year, month, mwmed
  B. ANEEL RALIE unit-level history -> discovery_archive/pmo_deck/ralie_forward_schedule.csv
     (per monthly snapshot: MW expected to enter within 12 / 24 months, by source) + columns listing.
"""
import re, io, os, html, time, requests, pandas as pd, numpy as np, unicodedata
UA = {"User-Agent": "Mozilla/5.0 (X11; Linux x86_64) energy-data-probe/1.0"}
OUT = "discovery_archive/pmo_deck"; os.makedirs(OUT, exist_ok=True)
def get(url, tries=3, **kw):
    for i in range(tries):
        try:
            r = requests.get(url, headers=UA, timeout=300, **kw); r.raise_for_status(); return r
        except Exception as e:
            print("   retry", i + 1, url[-60:], repr(e)[:120]); time.sleep(5 * (i + 1))
    raise RuntimeError("download failed: " + url)
def norm(s): return unicodedata.normalize("NFKD", str(s)).encode("ascii", "ignore").decode()

# ---------- A. EPE
r = get("https://www.epe.gov.br/pt/publicacoes-dados-abertos/publicacoes/revisoes-quadrimestrais-da-carga"); page = html.unescape(r.text)
files = sorted(set(m.group(1) for m in re.finditer(r'href="([^"]+PublicacoesArquivos[^"]+)"', page)))
def meta(name):
    n = norm(name); yrs = re.search(r"(20\d\d)\s*[-_]\s*(20\d\d)", n); plan = f"{yrs.group(1)}-{yrs.group(2)}" if yrs else ""
    if re.search(r"extraordin", n, re.I): kind = "extra"
    elif re.search(r"\b2\s*(a|\.|\?|o)?\s*rev|2RQ|2 Rev|2Rev", n, re.I): kind = "2RQ"
    elif re.search(r"\b1\s*(a|\.|\?|o)?\s*rev|1RQ|1 Rev|1Rev", n, re.I): kind = "1RQ"
    else: kind = "PLAN"
    mm = "sem" if re.search(r"sem MMGD", n, re.I) else ("com" if re.search(r"com MMGD", n, re.I) else "na")
    return plan, kind, mm
ROW = re.compile(r"^\s*(20\d\d)\s+((?:[\d.]+(?:,\d+)?\s+){12})")
def parse_text(txt, src):
    rows = []; sub = None; plan, kind, mm = meta(src)
    for line in txt.split("\n"):
        l = norm(line).strip()
        if re.match(r"(?i)^subsistema\b|^sin\b|^sistema interligado|^(norte|nordeste|sul|sudeste)\b", l): sub = l; continue
        m = ROW.match(l)
        if m and sub:
            yr = int(m.group(1)); vals = m.group(2).split()
            for mi, v in enumerate(vals[:12], 1):
                try: x = float(v.replace(".", "").replace(",", "."))
                except Exception: continue
                rows.append(dict(revision_file=src, plan_years=plan, revision_kind=kind, mmgd=mm, subsystem=sub, year=yr, month=mi, mwmed=x))
    return rows
def parse_sheet(df, src):
    rows = []; sub = None; plan, kind, mm = meta(src)
    for i in range(len(df)):
        v0 = df.iat[i, 0]; s0 = norm(v0).strip() if pd.notna(v0) else ""
        if re.match(r"(?i)^subsistema\b|^sin\b|^(norte|nordeste|sul|sudeste)\b", s0): sub = s0
        try: yr = int(float(v0))
        except Exception: continue
        if 2000 <= yr <= 2040 and sub:
            for mi in range(1, 13):
                try: x = float(df.iat[i, mi])
                except Exception: continue
                if np.isfinite(x): rows.append(dict(revision_file=src, plan_years=plan, revision_kind=kind, mmgd=mm, subsystem=sub, year=yr, month=mi, mwmed=x))
    return rows
import pdfplumber
tidy = []
cand = [f for f in files if re.search(r"previs|prev mensais|valores mensais", norm(f), re.I) and not re.search(r"apresenta|boletim|nota|\bNT\b|carga global|patamar|mmgd\.pdf$", norm(f), re.I)]
print("EPE candidate files:", len(cand))
for f in cand:
    url = "https://www.epe.gov.br" + f if f.startswith("/") else f; name = f.split("/")[-1]
    try:
        b = get(url).content; n0 = len(tidy)
        if re.search(r"\.xlsx?$", name, re.I):
            x = pd.ExcelFile(io.BytesIO(b))
            for s in x.sheet_names: tidy += parse_sheet(x.parse(s, header=None), name)
        elif name.lower().endswith(".pdf"):
            with pdfplumber.open(io.BytesIO(b)) as pdf:
                txt = "\n".join((pg.extract_text() or "") for pg in pdf.pages[:6])
            tidy += parse_text(txt, name)
        print(f"  {len(tidy) - n0:5d} rows  {name}")
    except Exception as e: print("  FAILED", name, repr(e)[:150])
T = pd.DataFrame(tidy)
if len(T):
    T.to_csv(f"{OUT}/epe_load_forecasts.csv", index=False)
    print("EPE tidy rows:", len(T), "| revisions:", T.revision_file.nunique(), "| subsystems:", sorted(T.subsystem.unique()))
    print(T.groupby(["plan_years", "revision_kind", "mmgd"]).size().to_string())

# ---------- B. RALIE
base = "https://dadosabertos.aneel.gov.br/dataset/57e4b8b5-a5db-40e6-9901-27ca629d0477/resource/"
srcs = [("unit-parquet", base + "896a51b2-6d40-4b0a-b2b2-b460f6a5b7ed/download/ralie-unidade-geradora-historico.parquet"),
        ("plant-parquet", base + "4cd32195-55fb-4f3f-83e9-2168369e818d/download/ralie-usina-historico.parquet"),
        ("unit-zip", base + "f08f3f8f-db6f-4c3b-9f4f-7179d099c255/download/ralie-unidade-geradora-historico.zip")]
R = None
for lab, u in srcs:
    try:
        print("RALIE: trying", lab); rr = get(u, tries=2, stream=True); b = rr.content; print("   bytes", len(b))
        if lab.endswith("parquet"): R = pd.read_parquet(io.BytesIO(b))
        else:
            import zipfile; z = zipfile.ZipFile(io.BytesIO(b)); nm = [n for n in z.namelist() if n.lower().endswith(".csv")][0]
            R = pd.read_csv(z.open(nm), sep=";", encoding="latin-1", low_memory=False)
        print("   rows", len(R)); print("   columns:", list(R.columns)); break
    except Exception as e: print("   failed:", repr(e)[:200])
if R is not None:
    dcols = [c for c in R.columns if c.lower().startswith("dat")]
    for c in dcols: R[c] = pd.to_datetime(R[c], errors="coerce")
    snap = [c for c in dcols if "ralie" in c.lower()]; snap = snap[0] if snap else dcols[0]
    print("   snapshot col", snap, R[snap].min(), "..", R[snap].max(), "distinct", R[snap].nunique())
    pcols = [c for c in R.columns if re.search(r"potencia", c, re.I)]; exp = [c for c in dcols if re.search(r"prev|previs", c, re.I)]
    print("   power cols", pcols, "| expected-date cols", exp)
    print(R.head(2).to_string()[:2500])
    if pcols and exp:
        pw = pcols[0]; R[pw] = pd.to_numeric(R[pw].astype(str).str.replace(".", "", regex=False).str.replace(",", ".", regex=False), errors="coerce")
        src = [c for c in R.columns if re.search(r"origem|tipogeracao|fonte", c, re.I)]; sc = src[0] if src else None
        out = []
        for ec in exp:
            g = R.dropna(subset=[snap, ec])
            for s, d in g.groupby(snap):
                lead = (d[ec] - s).dt.days
                row = dict(snapshot=s.date(), expected_col=ec, n_units=len(d), mw_within_12m=d.loc[(lead >= 0) & (lead <= 365), pw].sum() / 1000, mw_within_24m=d.loc[(lead >= 0) & (lead <= 730), pw].sum() / 1000, mw_overdue=d.loc[lead < 0, pw].sum() / 1000)
                if sc:
                    for k, v in d.loc[(lead >= 0) & (lead <= 365)].groupby(sc)[pw].sum().items(): row["m12_" + re.sub(r"\W+", "_", norm(k))[:14]] = v / 1000
                out.append(row)
        O = pd.DataFrame(out); O.to_csv(f"{OUT}/ralie_forward_schedule.csv", index=False); print("RALIE forward rows:", len(O)); print(O.head(5).to_string()[:2000]); print(O.tail(3).to_string()[:2000])
print("DONE")
