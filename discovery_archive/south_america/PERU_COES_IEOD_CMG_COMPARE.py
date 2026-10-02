"""
Round 4 of the Peru marginal-cost gap probe: the IEOD day folder's marginal-cost zip (CMgYYYYMMDD.zip, earlier
Anexo6_CMgCP_DDMM.zip) holds CMgCP<DDMM>.xlsx with sheets Cmg_Barra / Cmg_Ener / Cmg_Cong (half-hourly marginal
cost per bar, S/./MWh). For each test day: half-hours present for SANTA ROSA 220 in each sheet, and how each sheet
compares with the costos-marginales export (ExportarMasivo TOTAL / ENERGIA / CONGESTION) on the half-hours both have.

Usage: python3 PERU_COES_IEOD_CMG_COMPARE.py [YYYY-MM-DD ...]
"""
import html
import io
import re
import sys
import time
import zipfile
from urllib.parse import quote

import pandas as pd
import requests

PORTAL = "https://www.coes.org.pe/Portal/"
EXPORT = PORTAL + "mercadomayorista/costosmarginales/ExportarMasivo"
S = requests.Session()
S.headers.update({"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) "
                                "Chrome/124.0 Safari/537.36"})
DAYS = sys.argv[1:] or ["2026-03-05", "2026-03-11", "2023-09-16", "2024-04-19", "2026-07-25", "2025-07-10",
                        "2021-05-03"]


def req(method, url, **kw):
    for i in range(4):
        try:
            r = S.request(method, url, timeout=180, **kw)
            r.raise_for_status()
            return r
        except requests.RequestException as e:
            print(f"    retry {i + 1}: {type(e).__name__}", flush=True)
            time.sleep(5 * (i + 1))
    return None


def browse(path):
    r = req("POST", PORTAL + "browser/vistadatos", data={"baseDirectory": path, "url": path, "indicador": "",
                                                          "initialLink": "", "orderFolder": ""})
    return [] if r is None else [(html.unescape(m.group(1)), m.group(2))
                                 for m in re.finditer(r"openBlob\('([^']+)',\s*'(\w)'", r.text)]


def export_day(day):
    r = req("GET", f"{EXPORT}?fechaInicio={day:%d/%m/%Y}&fechaFin={day + pd.Timedelta(days=1):%d/%m/%Y}")
    raw = pd.read_excel(io.BytesIO(r.content), header=None)
    hdr = next(i for i in range(15) if "NOMBRE BARRA" in [str(v).strip() for v in raw.iloc[i].values])
    raw.columns = [str(c).strip() for c in raw.iloc[hdr].values]
    raw = raw.iloc[hdr + 1:]
    sr = raw[raw["NODO EMD"].astype(str).str.strip() == "STAROSA220"].copy()
    sr.index = pd.to_datetime(sr["FECHA HORA"].astype(str), format="%d/%m/%Y %H:%M", errors="coerce")
    cols = {c: c for c in sr.columns if c.upper().startswith(("ENERG", "CONGES", "TOTAL"))}
    sr = sr[list(cols)].apply(pd.to_numeric, errors="coerce")
    sr = sr[(sr.index > day) & (sr.index <= day + pd.Timedelta(days=1))]
    return sr


def ieod_sheets(day):
    months = browse(f"Post Operación/Reportes/IEOD/{day.year}/")
    mpath = next((p for p, k in months if k == "D" and re.match(rf"0?{day.month}_", p.rstrip('/').split('/')[-1])), None)
    dpath = next((p for p, k in browse(mpath) if k == "D" and p.rstrip('/').split('/')[-1] in (f"{day.day:02d}", str(day.day)))
                 , None) if mpath else None
    if not dpath:
        print(f"  no IEOD day folder ({mpath})", flush=True)
        return {}
    zf = next((p for p, k in browse(dpath) if k == "F" and re.search(r"cmg.*\.zip$", p, re.I)), None)
    if not zf:
        print("  no CMg zip in the day folder", flush=True)
        return {}
    z = zipfile.ZipFile(io.BytesIO(req("GET", PORTAL + "browser/download?url=" + quote(zf)).content))
    member = next(i.filename for i in z.infolist() if re.search(r"cmgcp.*\.xlsx?$", i.filename, re.I))
    xl = pd.ExcelFile(io.BytesIO(z.read(member)))
    out = {}
    for sh in xl.sheet_names:
        df = pd.read_excel(xl, sheet_name=sh, header=None)
        # header row = the one holding 'SANTA ROSA 220'; time stamps in the column left of the first bar name
        loc = [(i, j) for i in range(min(10, len(df))) for j in range(df.shape[1])
               if str(df.iat[i, j]).strip().upper() == "SANTA ROSA 220"]
        if not loc:
            print(f"  {member} sheet {sh}: no 'SANTA ROSA 220' header", flush=True)
            continue
        i, j = loc[0]
        tcol = next(c for c in range(df.shape[1]) if pd.to_datetime(df.iloc[i + 1:i + 5, c], errors="coerce").notna().all())
        s = pd.Series(pd.to_numeric(df.iloc[i + 1:, j], errors="coerce").values,
                      index=pd.to_datetime(df.iloc[i + 1:, tcol], errors="coerce")).dropna()
        s = s[s.index.notna()]
        out[sh] = s
    return out


for d in DAYS:
    day = pd.Timestamp(d)
    print(f"\n===== {d}", flush=True)
    ex = export_day(day)
    print(f"  export STAROSA220: {ex['TOTAL'].notna().sum() if 'TOTAL' in ex else 0} half-hours; columns {list(ex.columns)}; "
          f"TOTAL mean {ex['TOTAL'].mean():.2f}", flush=True)
    sheets = ieod_sheets(day)
    for sh, s in sheets.items():
        print(f"  IEOD {sh}: {len(s)} half-hours ({s.index.min()} .. {s.index.max()}), mean {s.mean():.2f}", flush=True)
        for col in ex.columns:
            both = ex[col].dropna().index.intersection(s.index)
            if len(both):
                diff = (s.reindex(both) - ex[col].reindex(both)).abs()
                print(f"     vs export {col}: {len(both)} common half-hours, mean |diff| {diff.mean():.3f}, max {diff.max():.3f}",
                      flush=True)
    if "Cmg_Barra" in sheets and "TOTAL" in ex:
        filled = ex["TOTAL"].combine_first(sheets["Cmg_Barra"])
        filled = filled[(filled.index > day) & (filled.index <= day + pd.Timedelta(days=1))]
        print(f"  export + IEOD Cmg_Barra: {filled.notna().sum()} of 48 half-hours, daily mean {filled.mean():.2f}", flush=True)
