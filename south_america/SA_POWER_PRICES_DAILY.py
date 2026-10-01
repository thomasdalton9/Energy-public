"""
South America daily wholesale power prices, in local currency and USD, from
2021 - a separate workbook (NOT part of the South & Central America master).

Output: output/Data and Chart Outputs/south_america_power_prices_daily.xlsx
  Units        notes: definitions, units, sources, FX, coverage, caveats
  USD daily    one column per market, USD/MWh (daily average of the hourly / half-hourly prices)
  Brazil, Colombia, Peru, Argentina, Uruguay, Bolivia
               one row per day: the prices in local currency, the FX rate used and the prices in USD
  Chile        monthly (CNE 'Precio Medio de Mercado', the only key-free Chilean market price series)
  Bolivia monthly  CNDC monthly energy / capacity / monomic prices
Charts are added by add_charts.py (registry entry 'south_america_power_prices_daily.xlsx').

Sources (all public, no key; found with discovery_archive/south_america/SA_POWER_PRICES_DISCOVERY*.py):
  Brazil     ONS open data 'CMO semi-horario' (https://dados.ons.org.br/dataset/cmo-semi-horario), half-hourly
             marginal operating cost by subsystem from the DESSEM day-ahead model, R$/MWh. CCEE's PLD (the
             settlement price, CMO bounded by ANEEL's floor/cap) is the preferred series, but every CCEE host
             (www, dadosabertos, pda-download) answers 403 'Acesso bloqueado' to GitHub's runners.
  Colombia   XM API (servapibi.xm.com.co): PrecBolsNaci hourly (Precio de Bolsa Nacional, COP/kWh), PrecEsca and
             PrecEscaAct daily (Precio de Escasez, Precio de Escasez de Activacion, COP/kWh).
  Peru       COES 'Costos Marginales Nodales' export (www.coes.org.pe/Portal/mercadomayorista/costosmarginales/
             ExportarMasivo), half-hourly marginal cost per bar, S/./MWh; barra SANTA ROSA 220 (node STAROSA220).
  Argentina  CAMMESA 'Parte control post-operativo' (nemo PARTE_POST_OPERATIVO, one POyymmdd.zip a day with an
             Access database; table PRECIOS_AREA): hourly CMO (costo marginal operado) and PRECIO_AREA (the
             sanctioned spot price, capped by the Secretaria de Energia), ARS/MWh. Needs mdbtools (mdb-export).
  Uruguay    ADME 'Precio Spot sancionado horario' (adme.com.uy/mmee/spot/spotSancionadoDetalle.php), USD/MWh,
             one page per month, published ~8 days after month end.
  Bolivia    CNDC API (www.cndc.bo/wp-json/cndc/v1/cm-diario): daily marginal cost, US$/MWh, last ~180 days only
             (kept and accumulated by this workbook); /dashboard/precios: monthly energy, capacity and monomic
             prices (US$/MWh) from 2021.
  Chile      CNE 'Precio Medio de Mercado' (PMM SEN, monthly, CLP/kWh) from cne.cl's media library. Chile's
             marginal costs (Coordinador Electrico Nacional) are behind a browser challenge / a user key.
  Ecuador    not included: no spot market (regulated contracts) and cenace.gob.ec fails TLS verification.
FX (official daily rates; weekends/holidays carry the last published rate):
  BRL  Banco Central do Brasil PTAX sell rate (SGS series 1, api.bcb.gov.br)
  COP  TRM (Tasa Representativa del Mercado, certified by the Superintendencia Financiera) from datos.gov.co
       dataset 32sa-8pi3 (Banco de la Republica's own site sits behind a bot captcha)
  PEN  BCRP series PD04640PD (TC sistema bancario SBS, venta), estadisticas.bcrp.gob.pe
  ARS  BCRA Comunicacion A 3500 wholesale reference rate (api.bcra.gob.ar estadisticas v4.0, variable 5)
  CLP  not available key-free (Banco Central de Chile's API needs credentials) - Chile is shown in CLP only.

Incremental: every day already in the workbook is kept; each run fetches only missing days since 2021-01-01
plus the last REVISION_DAYS days. Argentina's daily files are 5-13 MB, so its history is built newest-first
within --budget-min minutes per run.

Usage: python3 SA_POWER_PRICES_DAILY.py [--out PATH] [--markets brazil,colombia,...] [--start YYYY-MM-DD]
                                        [--end YYYY-MM-DD] [--budget-min N]
"""

print("STARTING", flush=True)

import argparse
import csv
import datetime as dt
import io
import os
import re
import subprocess
import sys
import tempfile
import time
import zipfile
from concurrent.futures import ThreadPoolExecutor, as_completed

import pandas as pd
import requests

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))  # repo root, for xlsx_notes
import xlsx_notes  # noqa: E402

OUT = os.path.join("output", "Data and Chart Outputs", "south_america_power_prices_daily.xlsx")
START = dt.date(2021, 1, 1)
REVISION_DAYS = 7
UA = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) "
                    "Chrome/124.0 Safari/537.36"}
MONTHS_ES = {"ene": 1, "feb": 2, "mar": 3, "abr": 4, "may": 5, "jun": 6, "jul": 7, "ago": 8, "set": 9, "sep": 9,
             "oct": 10, "nov": 11, "dic": 12}

# ------------------------------------------------------------------ markets
# Each country sheet: the local price columns, the FX column, then the same prices in USD/MWh.
# 'to_mwh' converts the local unit to per-MWh before dividing by the FX rate (COP/kWh -> COP/MWh).
COUNTRIES = {
    "Brazil": {"fx": "BRL per USD (PTAX)", "to_mwh": 1, "cols": {
        "SE/CO CMO (BRL/MWh)": "Brazil SE/CO (CMO)", "S CMO (BRL/MWh)": "Brazil S (CMO)",
        "NE CMO (BRL/MWh)": "Brazil NE (CMO)", "N CMO (BRL/MWh)": "Brazil N (CMO)"}},
    "Colombia": {"fx": "COP per USD (TRM)", "to_mwh": 1000, "cols": {
        "Precio de Bolsa Nacional (COP/kWh)": "Colombia (Precio de Bolsa)",
        "Precio de Escasez (COP/kWh)": None, "Precio de Escasez de Activacion (COP/kWh)": None}},
    "Peru": {"fx": "PEN per USD (SBS venta)", "to_mwh": 1, "cols": {
        "CMg Santa Rosa 220 kV (PEN/MWh)": "Peru (CMg Santa Rosa 220 kV)"}},
    "Argentina": {"fx": "ARS per USD (BCRA A3500)", "to_mwh": 1, "cols": {
        "CMO costo marginal operado (ARS/MWh)": "Argentina (CMO marginal cost)",
        "Precio spot sancionado (ARS/MWh)": "Argentina (spot price, capped)"}},
    "Uruguay": {"fx": None, "to_mwh": 1, "cols": {
        "Precio spot sancionado (USD/MWh)": "Uruguay (spot sancionado)"}},
    "Bolivia": {"fx": None, "to_mwh": 1, "cols": {
        "Costo marginal (USD/MWh)": "Bolivia (CNDC marginal cost)"}},
}
EXTRA_COLS = {"Argentina": ["Hours with more than one price area"],
              "Peru": ["Half-hours"], "Brazil": ["Half-hours"]}


def usd_name(col):
    return re.sub(r"\((BRL|COP|PEN|ARS)/(k?W?h|MWh)\)", "(USD/MWh)", col)


def session():
    s = requests.Session()
    s.headers.update(UA)
    return s


def get(s, url, tries=4, timeout=120, **kw):
    for i in range(tries):
        try:
            r = s.get(url, timeout=timeout, **kw)
            r.raise_for_status()
            return r
        except requests.RequestException as e:
            if i == tries - 1:
                raise
            print(f"    retry {i + 1} {url[:90]}: {type(e).__name__}", flush=True)
            time.sleep(5 * (i + 1))


def post(s, url, tries=4, timeout=180, **kw):
    for i in range(tries):
        try:
            r = s.post(url, timeout=timeout, **kw)
            r.raise_for_status()
            return r
        except requests.RequestException as e:
            if i == tries - 1:
                raise
            print(f"    retry {i + 1} {url[:90]}: {type(e).__name__}", flush=True)
            time.sleep(5 * (i + 1))


def months(days):
    """Sorted first-of-month dates covering `days`."""
    return sorted({d.replace(day=1) for d in days})


def month_end(m):
    return (pd.Timestamp(m) + pd.offsets.MonthEnd(0)).date()


# ------------------------------------------------------------------ Brazil (ONS CMO)
def fetch_brazil(days):
    s = session()
    frames = []
    for year in sorted({d.year for d in days}):
        url = f"https://ons-aws-prod-opendata.s3.amazonaws.com/dataset/cmo_tm/CMO_SEMIHORARIO_{year}.csv"
        r = get(s, url)
        df = pd.read_csv(io.BytesIO(r.content), sep=";", dtype={"id_subsistema": str})
        df["din_instante"] = pd.to_datetime(df["din_instante"])
        df["val_cmo"] = pd.to_numeric(df["val_cmo"], errors="coerce")
        df["date"] = df["din_instante"].dt.normalize()
        piv = df.pivot_table(index="date", columns="id_subsistema", values="val_cmo", aggfunc="mean")
        n = df[df["id_subsistema"] == "SE"].groupby("date").size()
        out = pd.DataFrame({"SE/CO CMO (BRL/MWh)": piv.get("SE"), "S CMO (BRL/MWh)": piv.get("S"),
                            "NE CMO (BRL/MWh)": piv.get("NE"), "N CMO (BRL/MWh)": piv.get("N"),
                            "Half-hours": n})
        print(f"  Brazil {year}: {len(out)} days ({len(r.content) / 1e6:.1f} MB)", flush=True)
        frames.append(out)
    return pd.concat(frames) if frames else pd.DataFrame()


# ------------------------------------------------------------------ Colombia (XM)
XM = "https://servapibi.xm.com.co/"


def fetch_colombia(days):
    s = session()
    s.headers.update({"Connection": "close"})
    rows = {}
    for m in months(days):
        a, b = m, month_end(m)
        hourly = post(s, XM + "hourly", json={"MetricId": "PrecBolsNaci", "Entity": "Sistema",
                                              "StartDate": a.isoformat(), "EndDate": b.isoformat()}).json()
        for it in hourly.get("Items", []):
            for e in it.get("HourlyEntities", []):
                vals = [float(v) for k, v in e.get("Values", {}).items() if k.startswith("Hour") and v not in (None, "")]
                if len(vals) >= 20:
                    rows.setdefault(it["Date"], {})["Precio de Bolsa Nacional (COP/kWh)"] = sum(vals) / len(vals)
        for metric, col in (("PrecEsca", "Precio de Escasez (COP/kWh)"),
                            ("PrecEscaAct", "Precio de Escasez de Activacion (COP/kWh)")):
            try:
                d = post(s, XM + "daily", json={"MetricId": metric, "Entity": "Sistema", "StartDate": a.isoformat(),
                                                "EndDate": b.isoformat()}).json()
            except requests.RequestException as e:
                print(f"  Colombia {metric} {m:%Y-%m}: FAILED {e}", flush=True)
                continue
            for it in d.get("Items", []):
                for e in it.get("DailyEntities", []):
                    if e.get("Value") not in (None, ""):
                        rows.setdefault(it["Date"], {})[col] = float(e["Value"])
        print(f"  Colombia {m:%Y-%m}: {sum(1 for k in rows if k.startswith(m.strftime('%Y-%m')))} days", flush=True)
        time.sleep(0.3)
    out = pd.DataFrame.from_dict(rows, orient="index")
    out.index = pd.to_datetime(out.index)
    return out


# ------------------------------------------------------------------ Peru (COES)
COES = "https://www.coes.org.pe/Portal/mercadomayorista/costosmarginales/ExportarMasivo"
PERU_NODE = "STAROSA220"   # 'SANTA ROSA 220'


def fetch_peru(days):
    s = session()
    frames = []
    for m in months(days):
        a = max(m, START)
        b = month_end(m) + dt.timedelta(days=1)   # the 00:00 value closing the month's last day
        r = get(s, f"{COES}?fechaInicio={a:%d/%m/%Y}&fechaFin={b:%d/%m/%Y}", timeout=600)
        if r.content[:2] != b"PK":
            print(f"  Peru {m:%Y-%m}: not an xlsx ({r.headers.get('content-type')})", flush=True)
            continue
        raw = pd.read_excel(io.BytesIO(r.content), header=None)
        hdr = next(i for i in range(15) if "NOMBRE BARRA" in [str(v).strip() for v in raw.iloc[i].values])
        raw.columns = [str(c).strip() for c in raw.iloc[hdr].values]
        raw = raw.iloc[hdr + 1:]
        node = raw[raw["NODO EMD"].astype(str).str.strip() == PERU_NODE].copy()
        node["t"] = pd.to_datetime(node["FECHA HORA"].astype(str), format="%d/%m/%Y %H:%M", errors="coerce")
        node["TOTAL"] = pd.to_numeric(node["TOTAL"], errors="coerce")
        node = node.dropna(subset=["t", "TOTAL"]).drop_duplicates("t")
        # FECHA HORA marks the END of each half hour: 00:30 .. 24:00 (shown as 00:00 next day) is one day
        node["date"] = (node["t"] - pd.Timedelta(minutes=1)).dt.normalize()
        g = node.groupby("date")["TOTAL"]
        out = pd.DataFrame({"CMg Santa Rosa 220 kV (PEN/MWh)": g.mean(), "Half-hours": g.size()})
        out = out[(out.index >= pd.Timestamp(m)) & (out.index <= pd.Timestamp(month_end(m)))]
        print(f"  Peru {m:%Y-%m}: {len(out)} days, {raw['NOMBRE BARRA'].nunique()} bars "
              f"({len(r.content) / 1e6:.1f} MB)", flush=True)
        frames.append(out)
    return pd.concat(frames) if frames else pd.DataFrame()


# ------------------------------------------------------------------ Argentina (CAMMESA)
CAMMESA_LOOKUP = "https://api.cammesa.com/pub-svc/public/findDocumentosByNemoRango"
CAMMESA_ATTACH = "https://api.cammesa.com/pub-svc/public/findAttachmentByNemoId"
NEMO = "PARTE_POST_OPERATIVO"


def cammesa_month_docs(s, first):
    nxt = (pd.Timestamp(first) + pd.offsets.MonthBegin(1)).date()
    fmt = "%Y-%m-%dT%H:%M:%S.000Z"
    docs = get(s, CAMMESA_LOOKUP, params={"fechadesde": first.strftime(fmt), "fechahasta": nxt.strftime(fmt),
                                          "nemo": NEMO}, timeout=60).json()
    out = {}
    for doc in docs if isinstance(docs, list) else []:
        for att in doc.get("adjuntos", []):
            if att.get("id", "").startswith("PO") and att["id"].endswith(".zip"):
                out[att["id"]] = (doc, att)
    return out


def cammesa_day(s, doc, att):
    r = get(s, CAMMESA_ATTACH, params={"attachmentId": att["id"], "docId": doc["id"], "nemo": doc.get("nemo", NEMO)},
            timeout=300)
    zf = zipfile.ZipFile(io.BytesIO(r.content))
    name = next(n for n in zf.namelist() if n.lower().endswith(".mdb"))
    with tempfile.NamedTemporaryFile(suffix=".mdb", delete=False) as f:
        f.write(zf.read(name))
        path = f.name
    try:
        res = subprocess.run(["mdb-export", path, "PRECIOS_AREA"], capture_output=True, text=True, timeout=300)
    finally:
        os.unlink(path)
    if res.returncode != 0:
        raise RuntimeError(f"mdb-export PRECIOS_AREA: {res.stderr[:200]}")
    rows = list(csv.DictReader(io.StringIO(res.stdout)))
    by_hour = {}
    for x in rows:
        try:
            h, area = int(x["HORA"]), int(float(x["ARL_ECON"] or 0))
            cmo, pa = float(x["CMO"]), float(x["PRECIO_AREA"])
        except (ValueError, KeyError, TypeError):
            continue
        by_hour.setdefault(h, []).append((area, cmo, pa))
    if len(by_hour) < 23:
        raise RuntimeError(f"only {len(by_hour)} hours in PRECIOS_AREA")
    main = [min(v) for v in by_hour.values()]   # the lowest-numbered (main) area when the grid splits
    return {"CMO costo marginal operado (ARS/MWh)": sum(x[1] for x in main) / len(main),
            "Precio spot sancionado (ARS/MWh)": sum(x[2] for x in main) / len(main),
            "Hours with more than one price area": sum(len(v) > 1 for v in by_hour.values())}, len(r.content)


def fetch_argentina(days, budget_min=40, workers=4, checkpoint=None):
    s = session()
    s.headers.update({"User-Agent": "gas-demand-scripts/1.0"})
    t0 = time.time()
    rows, listings = {}, {}
    todo = sorted(days, reverse=True)   # newest first
    batch_size = workers * 10
    for i in range(0, len(todo), batch_size):
        if time.time() - t0 > budget_min * 60:
            print(f"  Argentina: time budget ({budget_min} min) used - {len(todo) - i} day(s) left for later runs",
                  flush=True)
            break
        jobs = {}
        with ThreadPoolExecutor(workers) as pool:
            for day in todo[i:i + batch_size]:
                first = day.replace(day=1)
                if first not in listings:
                    try:
                        listings[first] = cammesa_month_docs(s, first)
                    except (requests.RequestException, ValueError) as e:
                        print(f"  Argentina listing {first:%Y-%m}: FAILED {type(e).__name__}", flush=True)
                        listings[first] = {}
                hit = listings[first].get(day.strftime("PO%y%m%d.zip"))
                if hit is None:
                    continue
                jobs[pool.submit(cammesa_day, s, *hit)] = day
            for fut in as_completed(jobs):
                day = jobs[fut]
                try:
                    row, size = fut.result()
                    rows[pd.Timestamp(day)] = row
                    print(f"  Argentina {day}: CMO {row['CMO costo marginal operado (ARS/MWh)']:,.0f}, spot "
                          f"{row['Precio spot sancionado (ARS/MWh)']:,.0f} ARS/MWh ({size / 1e6:.1f} MB, "
                          f"{time.time() - t0:.0f}s)", flush=True)
                except Exception as e:  # noqa: BLE001 - one bad day must not stop the backfill
                    print(f"  Argentina {day}: FAILED ({type(e).__name__}: {str(e)[:150]})", flush=True)
        if checkpoint and rows:
            checkpoint(pd.DataFrame.from_dict(rows, orient="index"))
    return pd.DataFrame.from_dict(rows, orient="index")


# ------------------------------------------------------------------ Uruguay (ADME)
ADME = "https://adme.com.uy/mmee/spot/spotSancionadoDetalle.php"


def fetch_uruguay(days):
    s = session()
    rows = {}
    for m in months(days):
        r = get(s, ADME, params={"remota": 1, "a": m.year, "m": f"{m.month:02d}"})
        html = r.content.decode("utf-8", "replace")
        n = 0
        for tr in re.findall(r"(?is)<tr[^>]*>(.*?)</tr>", html):
            cells = [re.sub(r"<[^>]+>|&nbsp;", " ", c).strip() for c in re.findall(r"(?is)<td[^>]*>(.*?)</td>", tr)]
            if not cells or not re.fullmatch(r"\d{2}/\d{2}/\d{4}", cells[0]):
                continue
            vals = []
            for c in cells[1:25]:
                try:
                    vals.append(float(c.replace(",", ".")))
                except ValueError:
                    pass
            if len(vals) >= 23:
                rows[pd.to_datetime(cells[0], format="%d/%m/%Y")] = {
                    "Precio spot sancionado (USD/MWh)": sum(vals) / len(vals)}
                n += 1
        print(f"  Uruguay {m:%Y-%m}: {n} days", flush=True)
    return pd.DataFrame.from_dict(rows, orient="index")


# ------------------------------------------------------------------ Bolivia (CNDC)
CNDC = "https://www.cndc.bo/wp-json/cndc/v1/"


def fetch_bolivia(days):
    """The daily marginal cost: CNDC's API serves only the last ~180 days, as 'dd/mm' without a year."""
    s = session()
    j = get(s, CNDC + "cm-diario", params={"dias": 400}).json()
    today = dt.date.today()
    rows, year, prev = {}, today.year, None
    for f, v in reversed(list(zip(j.get("fechas", []), j.get("valores", [])))):
        d_, m_ = (int(x) for x in f.split("/"))
        if prev is not None and (m_, d_) > prev:     # walked back past 1 January
            year -= 1
        if year == today.year and (m_, d_) > (today.month, today.day):
            year -= 1
        prev = (m_, d_)
        if v is not None:
            rows[pd.Timestamp(year, m_, d_)] = {"Costo marginal (USD/MWh)": float(v)}
    out = pd.DataFrame.from_dict(rows, orient="index").sort_index()
    print(f"  Bolivia: {len(out)} days {out.index.min():%Y-%m-%d}..{out.index.max():%Y-%m-%d}" if len(out)
          else "  Bolivia: no data", flush=True)
    return out


def fetch_bolivia_monthly():
    s = session()
    rows = {}
    for year in range(START.year, dt.date.today().year + 1):
        j = get(s, CNDC + "dashboard/precios", params={"modo": "anio", "anio": year}).json()
        for x in (j.get("monomicos") or {}).get("serie") or []:
            if x.get("mes"):
                rows[pd.Timestamp(year, int(x["mes"]), 1)] = {
                    "Precio de energia (USD/MWh)": x.get("precio_energia"),
                    "Precio de potencia (USD/MWh)": x.get("precio_potencia"),
                    "Precio de transporte (USD/MWh)": x.get("precio_transporte"),
                    "Precio monomico (USD/MWh)": x.get("precio_monomico")}
    out = pd.DataFrame.from_dict(rows, orient="index").sort_index()
    out.index.name = "month"
    print(f"  Bolivia monthly: {len(out)} months", flush=True)
    return out


# ------------------------------------------------------------------ Chile (CNE PMM)
def fetch_chile_monthly():
    s = session()
    r = get(s, "https://www.cne.cl/wp-json/wp/v2/media",
            params={"search": "Precio_Medio_de_Mercado", "per_page": 50, "_fields": "date,source_url"})
    items = [i for i in r.json() if re.search(r"/Precio_Medio_de_Mercado(-\d+)?\.xlsx$", i.get("source_url", ""))]
    if not items:
        raise RuntimeError("no Precio_Medio_de_Mercado.xlsx on cne.cl")
    url = max(items, key=lambda i: i["date"])["source_url"]
    raw = pd.read_excel(io.BytesIO(get(s, url).content), sheet_name="PMM SEN", header=None)
    hdr = next(i for i in range(15) if any("PMM SEN" in str(v) and "kWh" in str(v) for v in raw.iloc[i].values))
    raw.columns = [re.sub(r"\s+", " ", str(c)).strip() for c in raw.iloc[hdr].values]
    raw = raw.iloc[hdr + 1:]
    pub_col = next(c for c in raw.columns if c.lower().startswith("fecha de publ"))
    win_col = next(c for c in raw.columns if c.lower().startswith("ventana"))
    pmm_col = next(c for c in raw.columns if re.fullmatch(r"PMM SEN \[\$/kWh\] ?\*?", c))
    out = {}
    for _, x in raw.iterrows():
        pub = str(x[pub_col]).strip().lower()
        m = re.match(r"(\d{1,2}) de ([a-z]+) de (\d{4})", pub)
        val = pd.to_numeric(x[pmm_col], errors="coerce")
        if not m or pd.isna(val):
            continue
        mon = next((v for k, v in MONTHS_ES.items() if m.group(2).startswith(k)), None)
        if mon is None:
            continue
        out[pd.Timestamp(int(m.group(3)), mon, 1)] = {"PMM SEN (CLP/kWh)": float(val),
                                                      "4-month window": str(x[win_col]).strip()}
    df = pd.DataFrame.from_dict(out, orient="index").sort_index()
    df = df[df.index >= pd.Timestamp(START)]
    df.index.name = "month (of publication)"
    print(f"  Chile PMM: {len(df)} months from {url}", flush=True)
    return df


# ------------------------------------------------------------------ FX
def _daily(series, a, b):
    s = series.sort_index()
    s = s[~s.index.duplicated(keep="last")]
    idx = pd.date_range(min(s.index.min(), pd.Timestamp(a)) if len(s) else a, b, freq="D")
    return s.reindex(idx).ffill(limit=10)


def fx_brl(a, b):
    s = session()
    vals = {}
    for y in range(a.year, b.year + 1):
        lo, hi = max(a, dt.date(y, 1, 1)) - dt.timedelta(days=10), min(b, dt.date(y, 12, 31))
        r = get(s, "https://api.bcb.gov.br/dados/serie/bcdata.sgs.1/dados",
                params={"formato": "json", "dataInicial": lo.strftime("%d/%m/%Y"), "dataFinal": hi.strftime("%d/%m/%Y")})
        for x in r.json():
            vals[pd.to_datetime(x["data"], format="%d/%m/%Y")] = float(x["valor"])
    return pd.Series(vals)


def fx_cop(a, b):
    s = session()
    lo = (a - dt.timedelta(days=10)).isoformat()
    r = get(s, "https://www.datos.gov.co/resource/32sa-8pi3.json",
            params={"$where": f"vigenciahasta >= '{lo}T00:00:00' AND vigenciadesde <= '{b.isoformat()}T00:00:00'",
                    "$order": "vigenciadesde", "$limit": 50000})
    vals = {}
    for x in r.json():
        d0, d1 = pd.Timestamp(x["vigenciadesde"][:10]), pd.Timestamp(x["vigenciahasta"][:10])
        for d in pd.date_range(d0, d1, freq="D"):
            vals[d] = float(x["valor"])
    return pd.Series(vals)


def fx_pen(a, b):
    s = session()
    vals = {}
    for y in range(a.year, b.year + 1):
        lo, hi = max(a, dt.date(y, 1, 1)) - dt.timedelta(days=10), min(b, dt.date(y, 12, 31))
        r = get(s, f"https://estadisticas.bcrp.gob.pe/estadisticas/series/api/PD04640PD/json/{lo}/{hi}")
        j = r.json() if r.text.strip().startswith("{") else {}
        for p in j.get("periods", []):
            m = re.match(r"(\d{2})\.([A-Za-z]{3})\.(\d{2})", p["name"])
            try:
                v = float(p["values"][0])
            except (ValueError, IndexError, TypeError):
                continue
            if m and m.group(2).lower() in MONTHS_ES:
                vals[pd.Timestamp(2000 + int(m.group(3)), MONTHS_ES[m.group(2).lower()], int(m.group(1)))] = v
    return pd.Series(vals)


def fx_ars(a, b):
    s = session()
    vals = {}
    for y in range(a.year, b.year + 1):
        lo, hi = max(a, dt.date(y, 1, 1)) - dt.timedelta(days=10), min(b, dt.date(y, 12, 31))
        r = get(s, "https://api.bcra.gob.ar/estadisticas/v4.0/monetarias/5",
                params={"desde": lo.isoformat(), "hasta": hi.isoformat(), "limit": 3000})
        for res in r.json().get("results", []):
            for x in res.get("detalle", []):
                vals[pd.Timestamp(x["fecha"])] = float(x["valor"])
    return pd.Series(vals)


FX = {"Brazil": fx_brl, "Colombia": fx_cop, "Peru": fx_pen, "Argentina": fx_ars}
FETCH = {"Brazil": fetch_brazil, "Colombia": fetch_colombia, "Peru": fetch_peru, "Argentina": fetch_argentina,
         "Uruguay": fetch_uruguay, "Bolivia": fetch_bolivia}
# how far back each source publishes: days missing before this are not asked for again
SOURCE_START = {"Bolivia": None}   # Bolivia: only the API's rolling window (handled in fetch_bolivia)


# ------------------------------------------------------------------ workbook
def load(path):
    sheets = {}
    try:
        xl = pd.ExcelFile(path)
    except (FileNotFoundError, ValueError) as e:
        print(f"No workbook yet ({type(e).__name__}) - full backfill from {START}", flush=True)
        return sheets
    for name in list(COUNTRIES) + ["Chile", "Bolivia monthly"]:
        if name in xl.sheet_names:
            df = pd.read_excel(xl, sheet_name=name, index_col=0)
            df.index = pd.to_datetime(df.index)
            sheets[name] = df.drop(columns=[c for c in df.columns if str(c).startswith("Unnamed")])
    return sheets


def local_cols(country):
    return list(COUNTRIES[country]["cols"]) + EXTRA_COLS.get(country, [])


def build_country(country, df):
    cfg = COUNTRIES[country]
    df = df.sort_index()
    out = df[[c for c in local_cols(country) if c in df.columns]].copy()
    if cfg["fx"]:
        out[cfg["fx"]] = df[cfg["fx"]] if cfg["fx"] in df else float("nan")
        for c in cfg["cols"]:
            if c in out:
                out[usd_name(c)] = (out[c] * cfg["to_mwh"] / out[cfg["fx"]]).round(2)
    out.index.name = "date"
    return out.round(4)


def notes(sheets):
    def cover(name):
        d = sheets.get(name)
        if d is None or d.empty:
            return "no data yet"
        first_col = next(iter(COUNTRIES.get(name, {"cols": {d.columns[0]: 0}})["cols"]))
        s = d[first_col].dropna() if first_col in d else d.iloc[:, 0].dropna()
        return f"{s.index.min():%Y-%m-%d} to {s.index.max():%Y-%m-%d} ({len(s):,} days)" if len(s) else "no data yet"
    return [
        "SOUTH AMERICA DAILY WHOLESALE POWER PRICES",
        "Daily values are the simple average of the published hourly (or half-hourly) prices of each day. USD "
        "columns = local price / that day's official FX rate (local currency per USD); weekends and holidays use "
        "the last published rate. Charts: monthly averages of the daily USD prices, and one chart per country "
        "in local currency.",
        "Separate workbook - deliberately not part of the South & Central America master.",
        "",
        "BRAZIL (sheet 'Brazil')",
        "SE/CO, S, NE, N = the four subsystems (Sudeste/Centro-Oeste, Sul, Nordeste, Norte). CMO = custo marginal de "
        "operacao, half-hourly, from ONS's DESSEM day-ahead model, R$/MWh - ONS open data 'CMO semi-horario' "
        "(https://dados.ons.org.br/dataset/cmo-semi-horario; files ons-aws-prod-opendata.s3.amazonaws.com/dataset/"
        "cmo_tm/CMO_SEMIHORARIO_YYYY.csv). Half-hours = number of SE/CO values that day.",
        "NOT the PLD: CCEE's PLD (the settlement price - the hourly CMO of CCEE's own DESSEM run, bounded by ANEEL's "
        "yearly floor and caps) could not be read - every CCEE host (www.ccee.org.br, dadosabertos.ccee.org.br, "
        "pda-download.ccee.org.br) returns 403 'Acesso bloqueado' to GitHub's runners. The CMO is not bounded, so "
        "it can sit below the PLD floor (e.g. near zero in wet-season surplus hours) or above the cap.",
        f"Coverage: {cover('Brazil')}. FX: BCB PTAX sell rate (SGS series 1, https://api.bcb.gov.br).",
        "",
        "COLOMBIA (sheet 'Colombia')",
        "Precio de Bolsa Nacional: XM metric PrecBolsNaci (hourly, COP/kWh; the offer price of the last flexible "
        "plant needed for national demand plus the CREG delta), daily average of the 24 hours. Precio de Escasez "
        "(PrecEsca) and Precio de Escasez de Activacion (PrecEscaAct, CREG 140/2017): daily XM metrics, COP/kWh. "
        "Source: XM API https://servapibi.xm.com.co (hourly / daily endpoints). USD/MWh = COP/kWh x 1000 / TRM.",
        f"Coverage: {cover('Colombia')}. FX: TRM certified by the Superintendencia Financiera, dataset 32sa-8pi3 on "
        "https://www.datos.gov.co (Banco de la Republica's site blocks automated access with a captcha).",
        "",
        "PERU (sheet 'Peru')",
        "CMg Santa Rosa 220 kV: COES nodal marginal cost (energy + congestion, column TOTAL) at barra SANTA ROSA 220 "
        "(node STAROSA220, Lima), half-hourly, S/./MWh - COES 'Costos Marginales Nodales' "
        "(https://www.coes.org.pe/Portal/mercadomayorista/costosmarginales/index, export ExportarMasivo). Each value "
        "is stamped at the END of its half hour, so 00:30 to 24:00 make one day. Half-hours = values that day.",
        f"Coverage: {cover('Peru')}. FX: BCRP series PD04640PD - tipo de cambio sistema bancario SBS, venta "
        "(https://estadisticas.bcrp.gob.pe).",
        "",
        "ARGENTINA (sheet 'Argentina')",
        "CMO costo marginal operado: hourly marginal cost of the operated dispatch, ARS/MWh. Precio spot "
        "sancionado: PRECIO_AREA, the spot price CAMMESA settles at, capped by the Secretaria de Energia (it sits at "
        "the cap almost every hour, so it moves only when the cap is changed). Both from table PRECIOS_AREA of "
        "CAMMESA's daily 'Parte control post-operativo' (nemo PARTE_POST_OPERATIVO, POyymmdd.zip with an Access "
        "database; https://api.cammesa.com/pub-svc/public/findDocumentosByNemoRango). When the grid splits into "
        "separate price areas CAMMESA reports one row per area; the main area (lowest ARL_ECON) is used and the "
        "number of split hours is shown. CAMMESA's monthly 'precio monomico' is not in a public feed and is not "
        "included.",
        f"Coverage: {cover('Argentina')}. FX: BCRA Comunicacion A 3500 wholesale reference rate "
        "(https://api.bcra.gob.ar/estadisticas/v4.0/monetarias/5).",
        "",
        "URUGUAY (sheet 'Uruguay')",
        "Precio spot sancionado: ADME's hourly sanctioned spot price, USD/MWh (capped at 250 USD/MWh; 0 in hours "
        "of surplus renewable/hydro), daily average of 24 hours - https://adme.com.uy/mmee/spot/"
        "spotSancionadoDetalle.php (one page per month, published about 8 days after month end).",
        f"Coverage: {cover('Uruguay')}. Already in USD - no FX needed.",
        "",
        "BOLIVIA (sheets 'Bolivia', 'Bolivia monthly')",
        "Costo marginal: CNDC daily marginal cost of generation, US$/MWh (https://www.cndc.bo/wp-json/cndc/v1/"
        "cm-diario). CNDC's API only serves the last ~180 days, so daily history starts in 2026 and grows from "
        "each run. 'Bolivia monthly': CNDC monthly energy, capacity, transmission and monomic prices, US$/MWh, "
        "from 2021 (https://www.cndc.bo/wp-json/cndc/v1/dashboard/precios).",
        f"Coverage (daily): {cover('Bolivia')}.",
        "",
        "CHILE (sheet 'Chile', monthly)",
        "PMM SEN: CNE 'Precio Medio de Mercado' of the Sistema Electrico Nacional, CLP/kWh (the average price of "
        "the market's contracts and short-term sales over a 4-month window, published monthly; the row date is the "
        "publication month) - Precio_Medio_de_Mercado.xlsx from https://www.cne.cl (media library). Chile's "
        "marginal costs at Quillota 220 kV / Crucero 220 kV come only from the Coordinador Electrico Nacional, "
        "whose site sits behind a browser challenge and whose API (sipub.api.coordinador.cl) needs a user key. "
        "CLP is not converted to USD: the Banco Central de Chile's API needs credentials.",
        "",
        "ECUADOR: not included - Ecuador has no spot market (regulated contracts through the single buyer) and "
        "CENACE publishes no marginal-cost series; its site also fails TLS certificate verification.",
        "",
        "UPDATES",
        "Scheduled daily (.github/workflows/south_america_power_prices.yml). Incremental: stored days are kept; "
        f"each run fetches only missing days since {START} plus the last {REVISION_DAYS} days (revisions). "
        "Argentina's daily files are large, so its history is filled newest-first within a time budget per run.",
    ]


SECTIONS = ["SOUTH AMERICA DAILY WHOLESALE POWER PRICES", "BRAZIL (sheet 'Brazil')", "COLOMBIA (sheet 'Colombia')",
            "PERU (sheet 'Peru')", "ARGENTINA (sheet 'Argentina')", "URUGUAY (sheet 'Uruguay')",
            "BOLIVIA (sheets 'Bolivia', 'Bolivia monthly')", "CHILE (sheet 'Chile', monthly)", "UPDATES"]


def save(path, sheets):
    built = {c: build_country(c, sheets[c]) for c in COUNTRIES if c in sheets and not sheets[c].empty}
    usd = pd.DataFrame()
    for c, df in built.items():
        for col, label in COUNTRIES[c]["cols"].items():
            if label is None:
                continue
            src = usd_name(col) if COUNTRIES[c]["fx"] else col
            if src in df:
                usd[label] = df[src]
    usd = usd.sort_index().round(2)
    usd.index.name = "date"
    out = {"USD daily": usd}
    out.update(built)
    for extra, idx_name in (("Chile", "month (of publication)"), ("Bolivia monthly", "month")):
        if extra in sheets and not sheets[extra].empty:
            out[extra] = sheets[extra].sort_index()
            out[extra].index.name = idx_name
    os.makedirs(os.path.dirname(os.path.abspath(path)), exist_ok=True)
    xlsx_notes.write_workbook(path, out, notes(sheets), SECTIONS)
    print(f"  saved {path}", flush=True)


def merge(old, new):
    if old is None or old.empty:
        return new.sort_index()
    if new is None or new.empty:
        return old
    new = new.copy()
    new.index = pd.to_datetime(new.index)
    return new.combine_first(old).sort_index()


def wanted_days(df, cols, start, end):
    """Days to (re)fetch: those missing since `start` plus the last REVISION_DAYS stored."""
    have = set()
    if df is not None and not df.empty:
        ok = df[[c for c in cols if c in df.columns]].notna().any(axis=1)
        have = {d.date() for d in df.index[ok]}
    allday = [start + dt.timedelta(days=i) for i in range((end - start).days + 1)]
    missing = [d for d in allday if d not in have]
    recent = sorted(have)[-REVISION_DAYS:] if have else []
    return sorted(set(missing) | {d for d in recent if d >= start})


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", default=OUT)
    ap.add_argument("--markets", default="brazil,colombia,peru,argentina,uruguay,bolivia,chile")
    ap.add_argument("--start", default=START.isoformat())
    ap.add_argument("--end", default=None, help="last day to fetch (default: today)")
    ap.add_argument("--budget-min", type=float, default=40, help="Argentina download time budget per run")
    args = ap.parse_args()
    start = dt.date.fromisoformat(args.start)
    end = dt.date.fromisoformat(args.end) if args.end else dt.date.today()
    wanted = [m.strip().lower() for m in args.markets.split(",") if m.strip()]
    sheets = load(args.out)
    failures = []

    for country in COUNTRIES:
        if country.lower() not in wanted:
            continue
        print(f"\n== {country}", flush=True)
        old = sheets.get(country)
        days = wanted_days(old, list(COUNTRIES[country]["cols"])[:1], start, end)
        if country == "Uruguay":   # monthly pages, published ~8 days after the month ends
            days = [d for d in days if d < dt.date.today().replace(day=1)]
        if not days:
            print("  up to date", flush=True)
            continue
        print(f"  {len(days)} day(s) to fetch: {days[0]} .. {days[-1]}", flush=True)
        try:
            if country == "Argentina":
                def checkpoint(part, _old=old):
                    sheets["Argentina"] = merge(_old, part)
                    add_fx("Argentina", sheets, failures)
                    save(args.out, sheets)
                new = fetch_argentina(days, budget_min=args.budget_min, checkpoint=checkpoint)
            else:
                new = FETCH[country](days)
        except Exception as e:  # noqa: BLE001 - one market failing must not stop the others
            print(f"  {country}: FAILED ({type(e).__name__}: {str(e)[:300]})", flush=True)
            failures.append(country)
            continue
        if new is None or new.empty:
            print(f"  {country}: nothing new", flush=True)
            continue
        new.index = pd.to_datetime(new.index)
        new = new[(new.index >= pd.Timestamp(start)) & (new.index <= pd.Timestamp(end))]
        sheets[country] = merge(old, new)
        add_fx(country, sheets, failures)
        save(args.out, sheets)

    for name, fn in (("Chile", fetch_chile_monthly), ("Bolivia monthly", fetch_bolivia_monthly)):
        if name.split()[0].lower() not in wanted:
            continue
        print(f"\n== {name}", flush=True)
        try:
            sheets[name] = merge(sheets.get(name), fn())
        except Exception as e:  # noqa: BLE001
            print(f"  {name}: FAILED ({type(e).__name__}: {str(e)[:300]})", flush=True)
            failures.append(name)

    if not any(not v.empty for v in sheets.values()):
        print("No data at all.", flush=True)
        sys.exit(1)
    save(args.out, sheets)
    summary(sheets)
    if failures:
        print(f"\nFAILED: {', '.join(failures)}", flush=True)
        sys.exit(1 if len(failures) == len(wanted) else 0)


FX_CACHE = {}


def fx_series(country, a):
    """Daily FX (local currency per USD) from `a` to today, fetched once per run and reused."""
    have = FX_CACHE.get(country)
    if have is None or have.index.min() > pd.Timestamp(a):
        b = dt.date.today()
        FX_CACHE[country] = _daily(FX[country](a, b), a, b)
    return FX_CACHE[country]


def add_fx(country, sheets, failures):
    """Fill the FX column for stored days that lack it, and refresh the last REVISION_DAYS (official rate,
    carried over weekends/holidays)."""
    fxcol = COUNTRIES[country]["fx"]
    if not fxcol:
        return
    df = sheets[country]
    if fxcol not in df:
        df[fxcol] = float("nan")
    need = df.index[df[fxcol].isna()]
    need = need.union(df.index[df.index >= df.index.max() - pd.Timedelta(days=REVISION_DAYS)])
    if len(need) == 0:
        return
    try:
        series = fx_series(country, need.min().date())
    except Exception as e:  # noqa: BLE001
        print(f"  {country} FX: FAILED ({type(e).__name__}: {str(e)[:200]}) - USD columns left blank", flush=True)
        failures.append(f"{country} FX")
        return
    df.loc[need, fxcol] = series.reindex(need).values
    print(f"  {country} FX {fxcol}: {series.reindex(need).notna().sum()} of {len(need)} days filled", flush=True)
    sheets[country] = df


def summary(sheets):
    print("\n== Summary (monthly averages, last 3 months)", flush=True)
    for c in COUNTRIES:
        if c not in sheets or sheets[c].empty:
            continue
        b = build_country(c, sheets[c])
        cols = [x for x in b.columns if "Hours" not in x and "Half-hours" not in x]
        m = b[cols].resample("MS").mean().tail(3)
        print(f"-- {c}: {b.index.min():%Y-%m-%d}..{b.index.max():%Y-%m-%d}, {len(b)} days")
        print(m.round(2).to_string(), flush=True)
    for extra in ("Chile", "Bolivia monthly"):
        if extra in sheets and not sheets[extra].empty:
            print(f"-- {extra}:\n{sheets[extra].tail(3).to_string()}", flush=True)


if __name__ == "__main__":
    main()
