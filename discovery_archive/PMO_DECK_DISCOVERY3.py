"""Discovery step 3 (GitHub Actions):
  A. EPE quadrimestral load reviews: full file list (compact), parse the spreadsheet revisions into a tidy
     table (revision, subsystem, year, month, MWmed) and print it as CSV; test pdfplumber on one PDF year.
  B. ANEEL RALIE unit-level history (parquet): columns, snapshot dates, and - if an expected-operation date
     column exists - the forward schedule by snapshot: MW expected within 12 and 24 months, by source.
"""
import re, io, html, json, requests, pandas as pd, numpy as np, unicodedata
UA = {"User-Agent": "Mozilla/5.0 (X11; Linux x86_64) energy-data-probe/1.0"}
def get(url, **kw): return requests.get(url, headers=UA, timeout=180, **kw)
def norm(s): return unicodedata.normalize("NFKD", str(s)).encode("ascii", "ignore").decode()
MON = {m: i + 1 for i, m in enumerate(["Jan", "Fev", "Mar", "Abr", "Mai", "Jun", "Jul", "Ago", "Set", "Out", "Nov", "Dez"])}

print("===== A. EPE files (compact: topic / name)")
r = get("https://www.epe.gov.br/pt/publicacoes-dados-abertos/publicacoes/revisoes-quadrimestrais-da-carga"); page = html.unescape(r.text)
files = sorted(set(m.group(1) for m in re.finditer(r'href="([^"]+PublicacoesArquivos[^"]+)"', page)))
for f in files:
    t = re.search(r"topico-(\d+)/(.+)$", f); print(f"  {t.group(1)} | {t.group(2)}" if t else f"  {f}")

def parse_sheet(df, revision):
    """EPE layout: blocks 'Subsistema X' then a header row 'ANO Jan..Dez', then year rows."""
    rows = []; sub = None
    for i in range(len(df)):
        v0 = df.iat[i, 0]
        s0 = norm(v0).strip() if pd.notna(v0) else ""
        if re.match(r"(?i)^(subsistema|sistema|sin)\b", s0) or re.match(r"(?i)^(norte|nordeste|sul|sudeste|sudeste/centro|sin)\b", s0): sub = s0
        try: yr = int(float(v0))
        except Exception: continue
        if 2000 <= yr <= 2040 and sub:
            for mi in range(1, 13):
                val = df.iat[i, mi] if mi < df.shape[1] else np.nan
                try: val = float(val)
                except Exception: continue
                if np.isfinite(val): rows.append(dict(revision=revision, subsystem=sub, year=yr, month=mi, mwmed=val))
    return rows

print("\n===== A2. spreadsheet revisions -> tidy")
xls = [f for f in files if re.search(r"\.xlsx?$", f, re.I) and not re.search(r"sensibilidade", f, re.I)]
tidy = []
for f in xls:
    url = "https://www.epe.gov.br" + f if f.startswith("/") else f
    try:
        x = pd.ExcelFile(io.BytesIO(get(url).content)); rev = re.search(r"topico-\d+/(.+)$", f).group(1)
        n0 = len(tidy)
        for s in x.sheet_names:
            df = x.parse(s, header=None); rows = parse_sheet(df, rev + " | " + s); tidy += rows
        print(f"  {rev}: sheets {x.sheet_names} -> {len(tidy) - n0} rows")
        if len(tidy) - n0 == 0:
            df = x.parse(x.sheet_names[0], header=None); print(df.head(15).to_string()[:3000])
    except Exception as e: print("  failed:", url, e)
T = pd.DataFrame(tidy)
if len(T):
    print("\nsubsystems found:", sorted(T.subsystem.unique())); print("revisions:", T.revision.nunique())
    print("\n=====BEGIN_EPE_CSV====="); print(T.round(1).to_csv(index=False)); print("=====END_EPE_CSV=====")

print("\n===== A3. PDF test (pdfplumber) on the newest 'Previsoes mensais' PDF")
pdfs = [f for f in files if re.search(r"\.pdf$", f, re.I) and re.search(r"previs|mensa", norm(f), re.I)]
print("candidate PDFs:", len(pdfs))
for f in pdfs[-3:]:
    url = "https://www.epe.gov.br" + f if f.startswith("/") else f
    try:
        import pdfplumber
        with pdfplumber.open(io.BytesIO(get(url).content)) as pdf:
            print("\n", f.split("/")[-1], "pages:", len(pdf.pages))
            for pg in pdf.pages[:2]:
                print("--- page text (first 1200 chars):"); print((pg.extract_text() or "")[:1200])
                tb = pg.extract_tables(); print("--- tables on page:", len(tb))
                for t in tb[:2]: print(pd.DataFrame(t).head(10).to_string()[:1500])
    except Exception as e: print("  failed:", url, e)

print("\n===== B. ANEEL RALIE unit-level history")
try:
    u = "https://dadosabertos.aneel.gov.br/dataset/57e4b8b5-a5db-40e6-9901-27ca629d0477/resource/896a51b2-6d40-4b0a-b2b2-b460f6a5b7ed/download/ralie-unidade-geradora-historico.parquet"
    b = get(u).content; print("parquet bytes:", len(b)); R = pd.read_parquet(io.BytesIO(b)); print("rows", len(R)); print("columns:", list(R.columns))
    print(R.head(3).to_string()[:3000])
    dcols = [c for c in R.columns if c.lower().startswith("dat")]; print("date columns:", dcols)
    for c in dcols: R[c] = pd.to_datetime(R[c], errors="coerce")
    snap = [c for c in dcols if "ralie" in c.lower()][0]
    print("snapshots:", R[snap].min(), "..", R[snap].max(), "n distinct", R[snap].nunique())
    pcols = [c for c in R.columns if re.search(r"potencia|mda", c, re.I)]; print("power columns:", pcols)
    exp = [c for c in dcols if re.search(r"prev|previs|operacao|comercial", c, re.I)]; print("expected-date columns:", exp)
    if exp and pcols:
        pw = pcols[0]; R[pw] = pd.to_numeric(R[pw].astype(str).str.replace(".", "", regex=False).str.replace(",", ".", regex=False), errors="coerce")
        src = [c for c in R.columns if re.search(r"origem|tipogeracao|fonte", c, re.I)]; sc = src[0] if src else None
        out = []
        for ec in exp[:2]:
            g = R.dropna(subset=[snap, ec])
            for s, d in g.groupby(snap):
                lead = (d[ec] - s).dt.days
                row = dict(snapshot=s.date(), expected_col=ec, mw_within_12m=d.loc[(lead >= 0) & (lead <= 365), pw].sum() / 1000, mw_within_24m=d.loc[(lead >= 0) & (lead <= 730), pw].sum() / 1000, mw_overdue=d.loc[lead < 0, pw].sum() / 1000)
                if sc:
                    for k, v in d.loc[(lead >= 0) & (lead <= 365)].groupby(sc)[pw].sum().items(): row["m12_" + norm(k)[:12]] = v / 1000
                out.append(row)
        O = pd.DataFrame(out); print("\n=====BEGIN_RALIE_CSV====="); print(O.round(3).to_csv(index=False)); print("=====END_RALIE_CSV=====")
except Exception as e: print("RALIE failed:", repr(e))
print("\nDONE")
