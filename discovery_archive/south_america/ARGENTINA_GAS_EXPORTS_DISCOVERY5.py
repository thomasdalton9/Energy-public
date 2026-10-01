"""
Round 5 for Argentina gas exports by destination (see rounds 1-4).

Round 4 found the ENARGAS daily export report's AJAX endpoint:
  POST secciones/transporte-y-distribucion/partes-diarios-exp-imp-consulta-listado.php
       data: fecha_desde, fecha_hasta (JS fecha_yyyymmdd format), tipo_list
             (exp_dentro / exp_fuera); max 365 days per request
  GET  same URL ?tipo_file=xls&tipo_list=..&titulo=..&fecha_desde=dd/mm/yyyy&fecha_hasta=dd/mm/yyyy
This checks the date format, how far back the daily reports go, the
column headers per year (export points come and go), and the monthly
sums per export point for comparison with ENARGAS's monthly
Exportaciones.xlsx (dentro del sistema only).
"""
import io
import re

import pandas as pd
import requests

HEADERS = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                         "(KHTML, like Gecko) Chrome/124.0 Safari/537.36"}
TIMEOUT = (10, 120)
BASE = "https://www.enargas.gob.ar/secciones/transporte-y-distribucion/"
URL = BASE + "partes-diarios-exp-imp-consulta-listado.php"
PAGE = BASE + "dod-partes-exp-imp-consulta.php"

pd.set_option("display.width", 250)
pd.set_option("display.max_columns", 30)


def parse(html):
    """Header cells '<b>Chile</b><br>GasAndes' -> 'Chile | GasAndes'; rows -> DataFrame."""
    heads = [re.sub(r"\s+", " ", re.sub(r"<br\s*/?>", " | ", h)) for h in re.findall(r"<th[^>]*>(.*?)</th>", html, re.S | re.I)]
    heads = [re.sub(r"<[^>]+>", "", h).strip() for h in heads]
    rows = []
    for tr in re.findall(r"<tr[^>]*>(.*?)</tr>", html, re.S | re.I):
        cells = [re.sub(r"<[^>]+>|\s+", " ", c).strip() for c in re.findall(r"<td[^>]*>(.*?)</td>", tr, re.S | re.I)]
        if cells and re.match(r"\d{2}/\d{2}/\d{4}", cells[0]):
            rows.append(cells)
    if not rows:
        return heads, pd.DataFrame()
    df = pd.DataFrame(rows, columns=heads[:len(rows[0])] if len(heads) >= len(rows[0]) else None)
    return heads, df


def main():
    s = requests.Session()
    s.headers.update(HEADERS)
    s.get(PAGE, params={"tipo": "exp_dentro"}, timeout=TIMEOUT)
    fmt_ok = None
    for fmt in ("%Y%m%d", "%Y-%m-%d", "%Y/%m/%d", "%d/%m/%Y"):
        d0, d1 = pd.Timestamp("2021-01-01").strftime(fmt), pd.Timestamp("2021-01-31").strftime(fmt)
        r = s.post(URL, data={"fecha_desde": d0, "fecha_hasta": d1, "tipo_list": "exp_dentro"}, timeout=TIMEOUT)
        heads, df = parse(r.text)
        print(f"POST fmt={fmt}: status={r.status_code} bytes={len(r.text)} rows={len(df)} "
              f"first={df.iloc[0, 0] if len(df) else None} last={df.iloc[-1, 0] if len(df) else None}")
        if len(df) and fmt_ok is None:
            fmt_ok = fmt
    if fmt_ok is None:
        print("no POST format worked; text sample:", re.sub(r"\s+", " ", r.text)[:1500])
    # xls GET
    r = s.get(URL, params={"tipo_file": "xls", "tipo_list": "exp_dentro", "titulo": "x",
                           "fecha_desde": "01/01/2021", "fecha_hasta": "31/01/2021"}, timeout=TIMEOUT)
    print(f"GET xls: status={r.status_code} ct={r.headers.get('content-type')} bytes={len(r.content)} "
          f"head={r.content[:200]!r}")

    if fmt_ok is None:
        return
    for tipo in ("exp_dentro", "exp_fuera"):
        for year in range(2019, 2027):
            d0 = pd.Timestamp(f"{year}-01-01")
            d1 = min(pd.Timestamp(f"{year}-12-31"), pd.Timestamp("2026-09-30"))
            r = s.post(URL, data={"fecha_desde": d0.strftime(fmt_ok), "fecha_hasta": d1.strftime(fmt_ok),
                                  "tipo_list": tipo}, timeout=TIMEOUT)
            heads, df = parse(r.text)
            print(f"\n== {tipo} {year}: rows={len(df)} heads={heads}")
            if df.empty:
                continue
            df["Fecha"] = pd.to_datetime(df.iloc[:, 0], format="%d/%m/%Y", errors="coerce")
            num = df.drop(columns=[df.columns[0]]).set_index("Fecha").apply(
                lambda c: pd.to_numeric(c.str.replace(".", "", regex=False).str.replace(",", ".", regex=False),
                                        errors="coerce"))
            print(f"   first={df['Fecha'].min().date()} last={df['Fecha'].max().date()} "
                  f"missing days={(d1 - d0).days + 1 - len(df)}; sample raw row: {df.iloc[0, :].tolist()}")
            print(num.resample("MS").sum().astype(int).to_string())


if __name__ == "__main__":
    main()
