"""Probe 6: France register vs snapshots (pandas), SDES quarterly xlsx, Terna download-centre JS, DGEG archive, REData doc example.
Probe 5 findings (12 Oct 2026): ODRE register puismaxinstallee is in kW; year-end snapshots (-311220 ... -311225) and the current register
('au 31/07/2026') exist; date-based rebuild from the current register matches snapshots for wind/nuclear but overstates solar (15.1 GW at end-2020
against 10.4 GW in the 31/12/2020 snapshot, because <36 kW aggregates carry one old date) and misses undated old hydro/thermal rows;
Terna download centre has an API /api/sitecore/dati/downloadcenter/records?f=xlsx with datasets CapacityRenewableSources / CapacityThermal /
GenerationPlants; DGEG monthly PDFs slug estatisticas-rapidas-das-renovaveis-<mes>-<ano> exist for 2024-12 onwards, older slugs fall back."""
import io, json, re, subprocess, sys
from html import unescape
UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0"

def get(u, n=800, show=True, hdr=(), out=None):
    cmd = ["curl", "-sL", "-m", "120", "-A", UA] + [x for h in hdr for x in ("-H", h)]
    if out:
        cmd += ["-o", out]
    cmd += ["-w", "\n#HTTP %{http_code} %{content_type} %{size_download}", u]
    o = subprocess.run(cmd, capture_output=True, text=True, timeout=150)
    t = o.stdout
    head = t[t.rfind("#HTTP"):].strip()
    body = t[:t.rfind("#HTTP")]
    if show:
        print("\n=====", u[:230], "|", head, flush=True)
        print(body[:n].replace("\n", " "), flush=True)
    return body

def anchors(html, pat, lim=80):
    k = 0
    for m in re.finditer(r'<a [^>]*href="([^"]+)"[^>]*>(.*?)</a>', html, re.S | re.I):
        h, tx = m.group(1), re.sub(r"<[^>]+>", " ", m.group(2))
        tx = " ".join(unescape(tx).split())
        if re.search(pat, h + " " + tx, re.I):
            print("   ", h[:170], "|", tx[:100]); k += 1
            if k >= lim: break

import pandas as pd
pd.set_option("display.width", 250); pd.set_option("display.max_columns", 30)
B = "https://odre.opendatasoft.com/api/explore/v2.1/catalog/datasets/"
R = "registre-national-installation-production-stockage-electricite-agrege"

print("#### A. France register: pandas comparison")
def load(ds):
    p = f"/tmp/{ds[-8:]}.csv"
    get(f"{B}{ds}/exports/csv?delimiter=%3B&use_labels=false", show=False, out=p)
    d = pd.read_csv(p, sep=";", low_memory=False, usecols=lambda c: c in ("nominstallation", "filiere", "combustible", "puismaxinstallee", "dateraccordement", "datemiseenservice", "datederaccordement", "nbinstallations", "regime", "codeeicresourceobject", "codefiliere"))
    d["agg"] = d["nominstallation"].astype(str).str.startswith("Agr")
    d["mw"] = pd.to_numeric(d["puismaxinstallee"], errors="coerce") / 1000
    d["d_ms"] = pd.to_datetime(d["datemiseenservice"], dayfirst=True, errors="coerce")
    d["d_rac"] = pd.to_datetime(d["dateraccordement"], dayfirst=True, errors="coerce")
    print(ds, len(d), "agg rows", int(d["agg"].sum()), "missing ms", int(d["d_ms"].isna().sum()), "missing rac", int(d["d_rac"].isna().sum()))
    return d
cur = load(R)
snaps = {y: load(f"{R}-3112{str(y)[2:]}") for y in (2020, 2021, 2022, 2023, 2024, 2025)}
for y, s in snaps.items():
    ye = pd.Timestamp(f"{y}-12-31")
    rows = []
    for f in sorted(cur["filiere"].dropna().unique()):
        c = cur[cur.filiere == f]
        row = dict(filiere=f, snap=s[s.filiere == f].mw.sum(), snap_agg=s[(s.filiere == f) & s.agg].mw.sum(),
                   cur_ms_all=c[(c.d_ms <= ye)].mw.sum(), cur_ms_nonagg=c[(c.d_ms <= ye) & ~c.agg].mw.sum(),
                   cur_undated=c[c.d_ms.isna()].mw.sum(), cur_rac_all=c[(c.d_rac <= ye)].mw.sum(),
                   cur_rac_agg=c[(c.d_rac <= ye) & c.agg].mw.sum(), cur_total=c.mw.sum())
        rows.append(row)
    print("\n== year-end", y, "(GW)")
    t = pd.DataFrame(rows).set_index("filiere") / 1000
    print(t.round(3).to_string())
s = snaps[2024]
print("\nsnapshot 2024: rows by agg and filiere (solar)", s[s.filiere == "Solaire"].groupby("agg").agg(n=("mw", "size"), mw=("mw", "sum"), ni=("nbinstallations", "sum")).to_string())
print("current: agg solar by date-year (MW):")
c = cur[(cur.filiere == "Solaire") & cur.agg]
print(c.groupby(c.d_rac.dt.year).mw.sum().round(0).to_string())
print("current non-agg solar by connection year (MW):")
c = cur[(cur.filiere == "Solaire") & ~cur.agg]
print(c.groupby(c.d_rac.dt.year).mw.sum().round(0).to_string())

print("\n#### A2. ODRE catalogue: trimestre / panorama / raccord")
for term in ["trimestr", "panorama", "raccord", "puissance", "parc"]:
    b = get(f"{B}?limit=40&select=dataset_id&where=search(dataset_id,%22{term}%22)", show=False)
    try:
        print(term, [r["dataset_id"] for r in json.loads(b)["results"]])
    except Exception as e:
        print(term, "ERR", b[:100])

print("\n#### B. SDES quarterly xlsx")
for mid in (9714, 9715):
    get(f"https://www.statistiques.developpement-durable.gouv.fr/media/{mid}/download?inline", show=False, out=f"/tmp/sdes{mid}.xlsx")
    try:
        x = pd.ExcelFile(f"/tmp/sdes{mid}.xlsx")
        print(mid, x.sheet_names)
        for sh in x.sheet_names[:8]:
            d = x.parse(sh, header=None, nrows=14)
            print("--", sh, d.shape); print(d.iloc[:14, :9].to_string()[:1500])
    except Exception as e:
        print(mid, "ERR", e)

print("\n#### C. Terna JS")
h = get("https://dati.terna.it/en/download-center", show=False)
for m in list(re.finditer(r"downloadcenter/records", h))[:3]:
    print("CTX:", h[max(0, m.start() - 2500):m.start() + 800])
for kw in ["CapacityRenewableSources", "CapacityThermal", "GenerationPlants"]:
    idx = [m.start() for m in re.finditer(kw, h)]
    print(kw, len(idx), [h[i-100:i+160].replace("\n", " ") for i in idx[-2:]])

print("\n#### D. DGEG archive")
h = get("https://www.dgeg.gov.pt/pt/estatistica/energia/publicacoes/estatisticas-rapidas-das-renovaveis/", show=False)
print(len(h)); anchors(h, r"\.pdf|destaques|estatisticas-rapidas|page=", 60)
for p in (2, 3):
    h = get(f"https://www.dgeg.gov.pt/pt/destaques/?page={p}", show=False)
    print("destaques page", p, len(h)); anchors(h, r"estatisticas-rapidas-das-renovaveis", 20)
get("https://www.dgeg.gov.pt/media/axom32dy/dgeg-rpe-1995-2024-v1.xlsx", show=False, out="/tmp/rpe.xlsx")
try:
    x = pd.ExcelFile("/tmp/rpe.xlsx"); print("rpe sheets", x.sheet_names)
    for sh in x.sheet_names[:3]:
        d = x.parse(sh, header=None, nrows=8); print("--", sh, d.shape); print(d.iloc[:8, :8].to_string()[:900])
except Exception as e:
    print("rpe ERR", e)

print("\n#### E. REData doc examples")
H = ("Accept: application/json",)
for q in ["start_date=2019-01-01T00:00&end_date=2019-01-31T23:59&time_trunc=month",
          "start_date=2024-01-01T00:00&end_date=2024-01-31T23:59&time_trunc=month",
          "start_date=2024-01-01T00:00&end_date=2024-01-31T23:59&time_trunc=day",
          "start_date=2024-01-01T00:00&end_date=2024-12-31T23:59&time_trunc=month&geo_trunc=electric_system&geo_limit=peninsular&geo_ids=8741"]:
    get("https://apidatos.ree.es/es/datos/generacion/potencia-instalada?" + q, 300, hdr=H)
get("https://apidatos.ree.es/es/datos/balance/balance-electrico?start_date=2024-01-01T00:00&end_date=2024-01-31T23:59&time_trunc=month", 200, hdr=H)
