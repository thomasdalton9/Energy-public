"""
Discovery probe, round 6, for south_america/SA_POWER_PRICES_DAILY.py (compact output).

Round 2 (SA_POWER_PRICES_DISCOVERY2.py) found:
  - Brazil: every CCEE host (www, dadosabertos, pda-download) answers 403 "Acesso bloqueado ... nao atender as
    politicas de seguranca da CCEE" to GitHub runners, whatever the user agent. ONS open data (S3, no key) has
    'cmo-semi-horario' (CMO_SEMIHORARIO_YYYY.csv, from 2020) and 'cmo-semanal'.
  - Peru: COES menu links /Portal/mercadomayorista/costosmarginales/index (and /revisados).
  - Chile: cne.cl media has monthly Precio_Medio_de_Mercado.xlsx (2024-2026) and old BarCMg zips (node-price
    studies); no marginal-cost statistics files.
  - Uruguay: ADME menu 'Precio Spot sancionado' (/mme_admin/sancionado.php), /detalleejecucionhoraria/,
    /datosabiertos.html.
  - Bolivia: CNDC API /cm-diario (arg dias, default 30), /precmp (arg fecha), /dashboard/precios (modo, anio, mes),
    /historico/monomicos (desde, hasta).
Round 3 (SA_POWER_PRICES_DISCOVERY3.py) found:
  - Peru: COES page /Portal/mercadomayorista/costosmarginales/index, script costomarginal.js: controller
    siteRoot + 'mercadomayorista/costosmarginales/' with actions lista, mapa, Exportar?fecha=, ExportarMasivo?fechaInicio=.
  - Uruguay: spot price page https://adme.com.uy/mmee/spot/sancionado.php?remota=1; /detalleejecucionhoraria/ lists
    monthly YYYYMM_despachoEjecutado.xlsx files from 2015.
  - Bolivia: /cm-diario?dias=N -> {"fechas": ["dd/mm"...], "valores": [...]} (only ~6 months back, no year);
    /dashboard/precios?modo=mes&anio=&mes= -> monthly precio_energia / potencia / monomico (US$/MWh) back to 1996;
    /historico/monomicos/detalle?anio= monthly energy/power/transport components.
Round 4 (SA_POWER_PRICES_DISCOVERY4.py) found:
  - Peru: GET .../costosmarginales/ExportarMasivo?fechaInicio=dd/mm/yyyy&fechaFin=dd/mm/yyyy returns
    CostosMarginalesNodales.xlsx, sheet COSTOMARGINAL: FECHA HORA (half-hourly), NODO EMD, NOMBRE BARRA, ENERGIA,
    CONGESTION, TOTAL (3 days = 0.9 MB, 28.7k rows).
  - Chile: Precio_Medio_de_Mercado.xlsx sheets 'PMM SEN', 'PMM SIC', 'PMM SING' (monthly, $/kWh, 4-month windows).
    reportediario.cne.cl does not resolve; energiaabierta.cl fails TLS verification.
  - Uruguay: sancionado.php?remota=1 links spotSancionadoSeleccion.php?anio=YYYY&remota=1 (hourly spot sancionado);
    YYYYMM_despachoEjecutado.xlsx sheet 'CMO' = daily rows of 24 hourly CMO values (USD/MWh) + 'Total' (their sum).
  - Bolivia: cm-diario?dias=400 returns 180 days (28/03-29/09), values ~15-23.
Round 5 (SA_POWER_PRICES_DISCOVERY5.py) found:
  - Brazil: ONS CMO_SEMIHORARIO_YYYY.csv = id_subsistema;nom_subsistema;din_instante;val_cmo (SE, S, NE, N,
    half-hourly, R$/MWh), 2026 file runs to the current day (1.8 MB).
  - Peru: ExportarMasivo works for Jan-2021 (0.5 MB / 2 days) and a whole month (7.2 MB, 36 s); data start in
    column B (column A empty).
  - Uruguay: spotSancionadoSeleccion.php?anio=YYYY lists months -> spotSancionadoDetalle.php?remota=1&a=YYYY&m=MM.
"""
import io
import re
import sys
import traceback

import requests

UA = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) "
                    "Chrome/124.0 Safari/537.36"}
ONLY = set(sys.argv[1:])


def flat(s, n):
    return re.sub(r"\s+", " ", str(s))[:n]


def get(url, n=300, quiet=False, **kw):
    kw.setdefault("timeout", 90)
    kw.setdefault("headers", UA)
    try:
        r = requests.get(url, **kw)
        if not quiet:
            body = r.text if "html" not in str(r.headers.get("content-type")) else "[html]"
            print(f"  {r.status_code} {str(r.headers.get('content-type'))[:25]} {len(r.content)}B {r.url[:150]}"
                  f"\n     {flat(body, n)}")
        return r
    except Exception as e:  # noqa: BLE001
        print(f"  !! {url[:150]}: {type(e).__name__}: {str(e)[:200]}")
        return None


def hrefs(html, pat, n=25):
    out = sorted(set(re.findall(r"""(?:href|src|action)\s*=\s*["']([^"']+)["']""", html, re.I)))
    out = [u for u in out if re.search(pat, u, re.I)]
    for u in out[:n]:
        print("     link", u[:200])
    return out


def run(name, fn):
    if ONLY and name not in ONLY:
        return
    print("\n==== " + name, flush=True)
    try:
        fn()
    except Exception:  # noqa: BLE001
        traceback.print_exc(limit=2)


def peru():
    import pandas as pd
    base = "https://www.coes.org.pe/Portal/mercadomayorista/costosmarginales/"
    r = requests.get(base + "ExportarMasivo?fechaInicio=01/09/2026&fechaFin=01/09/2026", headers=UA, timeout=600)
    df = pd.read_excel(io.BytesIO(r.content), header=None)
    print("  shape", df.shape)
    print("  rows0-3:", [flat(list(df.iloc[i].values), 200) for i in range(3)])
    hdr = next(i for i in range(10) if "NOMBRE BARRA" in [str(v).strip() for v in df.iloc[i].values])
    df.columns = [str(c).strip() for c in df.iloc[hdr].values]
    df = df.iloc[hdr + 1:]
    print("  cols", list(df.columns), "n barras", df["NOMBRE BARRA"].nunique())
    m = df[df["NOMBRE BARRA"].astype(str).str.contains("ROSA|CHAVARR|SAN JUAN|CARABAYLLO", case=False)]
    print(m.groupby(["NODO EMD", "NOMBRE BARRA"])["TOTAL"].agg(["count", "mean", "min", "max"]).to_string()[:1500])
    print("  all-node mean:", pd.to_numeric(df["TOTAL"], errors="coerce").mean(), "times:", df["FECHA HORA"].iloc[:3].tolist(),
          df["FECHA HORA"].nunique())


def chile():
    import pandas as pd
    r = get("https://www.cne.cl/wp-content/uploads/2026/09/Precio_Medio_de_Mercado.xlsx", quiet=True)
    df = pd.read_excel(io.BytesIO(r.content), sheet_name="PMM SEN", header=None)
    print("  shape", df.shape)
    for i in list(range(0, 4)) + list(range(len(df) - 4, len(df))):
        print("   ", i, flat([v for v in df.iloc[i].values if str(v) != "nan"], 400))


def uruguay():
    r = get("https://adme.com.uy/mmee/spot/spotSancionadoDetalle.php?remota=1&a=2026&m=08", quiet=True)
    print("  detalle:", r.status_code, len(r.content), r.headers.get("content-type"))
    t = re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " | ", re.sub(r"(?is)<(script|style).*?</\1>", " ", r.text)))
    print("     text:", t[:1500])
    hrefs(r.text, r"\.(xls|xlsx|ods|csv|zip)|php|cgi", n=15)
    r = get("https://adme.com.uy/mmee/spot/spotSancionadoDetalle.php?remota=1&a=2021&m=01", quiet=True)
    t = re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " | ", re.sub(r"(?is)<(script|style).*?</\1>", " ", r.text)))
    print("  2021-01:", r.status_code, len(r.content), t[:400])


run("peru", peru)
run("chile", chile)
run("uruguay", uruguay)
print("\nDONE")
