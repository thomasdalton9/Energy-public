"""
Discovery round 8 (after CAPACITY_PBUE_DISCOVERY7.py).

Round 7 found:
  Peru: COES monthly bulletins as Excel for every month: Publicaciones/Boletines/<Y>/<MM>_<MES>/BOLETIN_<MES> <Y>.xlsx
        (accented 'BOLETÍN', September folder '09_SETIEMBRE'), 2011 to Aug 2026.
  Ecuador: older BNEE files exist at wp-content/uploads/downloads/<upload Y>/<upload M>/BNEE_<mes>_<anio>[_revACH].xls,
        upload month = the month of the matching PNG in the WP media library (found 6/6 tested, Apr 2024 - May 2026).
        ARCONEL's statistics publications page links download-monitor ids (25, 26, 272-318, 439, 452, 1274, 1285).

This round: COES bulletin layout (capacity sheets); titles/file names behind ARCONEL's download ids; BNEE before
Apr 2024.
"""
import io
import itertools
import re
import sys
from urllib.parse import quote

import pandas as pd
import requests

sys.path.insert(0, __file__.rsplit("/", 1)[0])
from CAPACITY_PBUE_DISCOVERY import get  # noqa: E402

pd.set_option("display.width", 300)
pd.set_option("display.max_columns", 30)
COES = "https://www.coes.org.pe/Portal/browser/download?url="


def bulletin(year, mm, mes):
    path = f"Publicaciones/Boletines/{year}/{mm:02d}_{mes}/BOLETÍN_{mes} {year}.xlsx"
    r = get(COES + quote(path))
    if r is None or r.status_code != 200 or len(r.content) < 1000:
        return
    xl = pd.ExcelFile(io.BytesIO(r.content))
    print(f"     {year}-{mm:02d} sheets: {xl.sheet_names}")
    for s in xl.sheet_names:
        d = pd.read_excel(xl, sheet_name=s, header=None)
        txt = d.astype(str).apply(lambda c: c.str.upper())
        hit = txt.apply(lambda c: c.str.contains("POTENCIA EFECTIVA|POTENCIA INSTALADA|CAPACIDAD INSTALADA")).any().any()
        if hit:
            print(f"     --- sheet {s!r} {d.shape} mentions capacity")
            rows = d.dropna(how="all")
            print(rows.iloc[:70, :16].to_string(max_colwidth=28)[:7000])


def peru():
    print("\n================ PERU")
    bulletin(2026, 8, "AGOSTO")
    bulletin(2021, 1, "ENERO")


def ecuador():
    print("\n================ ECUADOR")
    r = get("https://arconel.gob.ec/publicaciones-estadistica-del-sector-electrico-2/")
    if r is not None:
        for m in re.finditer(r'<a[^>]+href=["\']([^"\']*download\.php\?id=(\d+)&(?:amp;)?force=0)["\'][^>]*>(.*?)</a>',
                             r.text, re.S):
            text = re.sub(r"<[^>]+>|\s+", " ", m.group(3)).strip()
            ctx = re.sub(r"<[^>]+>|\s+", " ", r.text[max(0, m.start() - 400):m.start()])[-160:]
            print(f"     id {m.group(2)}: {text!r} (before: {ctx!r})")
    for i in [25, 26, 272, 300, 318, 439, 452, 1274, 1285]:
        try:
            h = requests.get(f"https://arconel.gob.ec/wp-content/plugins/download-monitor/download.php?id={i}&force=0",
                             headers={"User-Agent": "Mozilla/5.0"}, timeout=60, allow_redirects=True, stream=True)
            print(f"     id {i}: {h.status_code} {h.url} {h.headers.get('content-type')} "
                  f"{h.headers.get('content-disposition')} {h.headers.get('content-length')}")
            h.close()
        except requests.RequestException as e:
            print(f"     id {i}: {type(e).__name__}")
    for (mes, year), folder in itertools.product([("diciembre", 2023), ("marzo", 2024), ("diciembre", 2022),
                                                  ("diciembre", 2021)], ["2024/07", "2024/06", "2024/05", "2024/04"]):
        for suffix in ["", "_revACH", "-1"]:
            for base in ["https://arconel.gob.ec/wp-content/uploads/downloads/", "https://arconel.gob.ec/wp-content/uploads/"]:
                u = f"{base}{folder}/BNEE_{mes}_{year}{suffix}.xls"
                try:
                    if requests.head(u, timeout=20, headers={"User-Agent": "Mozilla/5.0"}).status_code == 200:
                        print("     found", u)
                except requests.RequestException:
                    pass


if __name__ == "__main__":
    for w in sys.argv[1:] or ["peru", "ecuador"]:
        try:
            globals()[w]()
        except Exception as e:  # noqa: BLE001
            import traceback
            traceback.print_exc()
            print(f"!! {w} failed: {type(e).__name__}: {e}")
        sys.stdout.flush()
