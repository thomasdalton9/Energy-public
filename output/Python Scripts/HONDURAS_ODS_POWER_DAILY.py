"""
Honduras daily power generation by technology from the ODS / CND
(Operador del Sistema, Centro Nacional de Despacho - ENEE), in the standard
raw grid-operator layout (see power_daily_std.py): sheet "Daily" with date,
Hydro_MWh, Gas_MWh, Wind_MWh, Solar_MWh, Coal_MWh, Nuclear_MWh, Oil_MWh,
Bioenergy_MWh, Other_MWh, Total_MWh.

Source (found via discovery_archive/south_america/NORTH_CENTRAL_AMERICA_POWER_PROBE7..10.py):
  ODS real-time app (Oracle APEX), page 'Produccion Horaria'
  https://appcnd.enee.hn:3200/odsprd/r/ods_prd/operador-del-sistema-ods/producci%C3%B3n-horaria?p8_indx=<n>
  = hourly MW per plant for one technology (P8_INDX 1..10) over a date range,
  downloaded as the report's CSV. A day's MWh = sum over plants of the 24
  hourly values. The app keeps hourly data from 2026-06-01 only
  (P8_MIN_DATE), so the Daily sheet starts there and grows every run.
  Before that, sheet "Monthly_GWh" holds ODS's monthly production by
  technology (GWh) from its annual market reports ('PRODUCCION GENERAL DE
  ENERGIA <year> (GWh)', 2021-2025) and the latest monthly report (Tabla 8,
  current year); thermal there is not split by fuel.

The server does not send its intermediate certificate; the script fetches
it from the leaf certificate's AIA URL and verifies normally against it.

Usage: python3 HONDURAS_ODS_POWER_DAILY.py [--out PATH]
"""

print("STARTING", flush=True)

import argparse
import datetime as dt
import html
import io
import json
import os
import re
import ssl
import sys
import time

import pandas as pd
import requests

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import power_daily_std as std  # noqa: E402

OUT = "output/Data and Chart Outputs/honduras_power_generation_daily.xlsx"
BASE = "https://appcnd.enee.hn:3200/odsprd/"
APP = BASE + "r/ods_prd/operador-del-sistema-ods/"
FUELS = std.FUELS
COLS = [f"{f}_MWh" for f in FUELS]
# ODS technology index (P8_INDX) -> (ODS name, standard fuel); None = not generation
TECH = {
    1: ("Hidroelectrica de paso", "Hydro"),
    2: ("Eolica", "Wind"),
    3: ("Solar fotovoltaica", "Solar"),
    4: ("Geotermica", "Other"),
    5: ("Biomasa", "Bioenergy"),
    6: ("Carbon", "Coal"),
    7: ("Interconexion", None),
    8: ("Bunker", "Oil"),
    9: ("Diesel", "Oil"),
    10: ("Hidroelectrica regulada (embalse)", "Hydro"),
}
CHUNK_DAYS = 5          # the report stops at 250 rows; ~45 plants x 5 days stays under it
MONTHLY_COLS = ["Hydro_GWh", "Thermal_GWh", "Wind_GWh", "Solar_GWh", "Bioenergy_GWh", "Other_GWh", "Total_GWh"]
HEADERS = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) "
                         "Chrome/124.0 Safari/537.36", "Accept-Language": "es-ES,es;q=0.9"}
MESES = ["ENE", "FEB", "MAR", "ABR", "MAY", "JUN", "JUL", "AGO", "SEP", "OCT", "NOV", "DIC"]

NOTES = [
    "UNITS",
    "Sheet 'Daily': MWh per day (gross generation in Honduras' national grid, SIN). Interconnection (imports) excluded.",
    "Sheet 'Monthly_GWh': GWh per month, ODS's own monthly production by technology (history before the daily data).",
    "",
    "COVERAGE",
    "Daily from 2026-06-01: the ODS app keeps hourly plant data only from that date (P8_MIN_DATE), so the daily "
    "series starts there and grows on every run (missing days plus the last 14 are fetched; ODS revises recent days). "
    "Monthly_GWh: 2021-01 to 2025-12 from ODS annual market reports, current-year months from the latest monthly "
    "market report. Updated daily by GitHub Actions (honduras_power_generation_daily.yml).",
    "",
    "SOURCE",
    "ODS / CND (Centro Nacional de Despacho, ENEE), https://www.ods.org.hn/ -> https://cnd.enee.hn/",
    "Daily: ODS app page 'Produccion Horaria', https://appcnd.enee.hn:3200/odsprd/r/ods_prd/operador-del-sistema-ods/"
    "producci%C3%B3n-horaria?p8_indx=<technology> (hourly MW per plant, CSV download of the report); day MWh = sum of "
    "the 24 hourly values over all plants of the technology.",
    "Monthly: ODS 'Informe Anual Operacion del Mercado <year>' (table 'PRODUCCION GENERAL DE ENERGIA <year> (GWh)') and "
    "'Informe Mensual Operacion del Mercado' (Tabla 8), listed at https://cnd.enee.hn/informe-anual-operacion-del-mercado/ "
    "and https://cnd.enee.hn/operacion-del-mercado/.",
    "",
    "MAPPING",
    "Hydro_MWh = Hidroelectrica de paso (run-of-river, P8_INDX 1) + Hidroelectrica regulada (reservoir: El Cajon, "
    "Patuca III, Rio Lindo, ..., P8_INDX 10).",
    "Wind_MWh = Eolica (2). Solar_MWh = Solar fotovoltaica (3).",
    "Other_MWh = Geotermica (4; Platanares - geothermal is reported in Other).",
    "Bioenergy_MWh = Biomasa (5; sugar-mill bagasse cogeneration and palm-oil biomass).",
    "Coal_MWh = Carbon (6). Oil_MWh = Bunker (8) + Diesel (9).",
    "Gas_MWh = 0 and Nuclear_MWh = 0 (none in Honduras). Interconexion (7) is imports, excluded.",
    "Total_MWh = sum of the fuel columns.",
    "Monthly_GWh: Hydro = public + private hydro; Thermal = public + private thermal (bunker, diesel and coal together - "
    "ODS's monthly table does not split them); Bioenergy = Biomasa; Other = Geotermica.",
]

S = requests.Session()
S.headers.update(HEADERS)


def aia_bundle(host, port=443):
    """certifi roots + the intermediate certificate(s) named in the server certificate's AIA field.
    The server omits its intermediate; verification stays fully on (chains must still end at a certifi root)."""
    import certifi
    from cryptography import x509
    from cryptography.hazmat.primitives.serialization import Encoding, pkcs7
    from cryptography.x509.oid import AuthorityInformationAccessOID, ExtensionOID
    pem = ssl.get_server_certificate((host, port))  # only read, to find the AIA URL
    cert = x509.load_pem_x509_certificate(pem.encode())
    extra = []
    for _ in range(3):
        try:
            aia = cert.extensions.get_extension_for_oid(ExtensionOID.AUTHORITY_INFORMATION_ACCESS).value
        except x509.ExtensionNotFound:
            break
        urls = [d.access_location.value for d in aia if d.access_method == AuthorityInformationAccessOID.CA_ISSUERS]
        if not urls:
            break
        der = requests.get(urls[0], timeout=30).content
        certs = pkcs7.load_der_pkcs7_certificates(der) if urls[0].endswith(".p7c") else [x509.load_der_x509_certificate(der)]
        extra += [c.public_bytes(Encoding.PEM).decode() for c in certs]
        cert = certs[0]
        if cert.issuer == cert.subject:
            break
    path = os.path.join(os.environ.get("RUNNER_TEMP", "/tmp"), f"{host}_ca_bundle.pem")
    with open(path, "w") as f:
        f.write(open(certifi.where()).read() + "\n" + "\n".join(extra))
    return path


def get(url, **kw):
    for attempt in range(4):
        try:
            r = S.get(url, timeout=120, **kw)
            if r.status_code >= 500:
                raise requests.RequestException(f"HTTP {r.status_code}")
            return r
        except requests.RequestException as e:
            print(f"  retry {url[:120]}: {e}", flush=True)
            time.sleep(5 * (attempt + 1))
    return None


def ctx(t):
    g = lambda p: (re.search(p, t).group(1) if re.search(p, t) else "")  # noqa: E731
    return {"sess": g(r'name="p_instance" value="(\d+)"'), "step": g(r'name="p_flow_step_id" value="(\d+)"'),
            "flow": g(r'name="p_flow_id" value="(\d+)"'), "salt": g(r'value="(\d+)" id="pSalt"'),
            "prot": g(r'id="pPageItemsProtected" value="([^"]+)"').replace("&#x2F;", "/"),
            "ajax": g(r'"ajaxIdentifier":"([^"]+)"').encode().decode("unicode_escape"),
            "ws": g(r'_worksheet_id" value="(\d+)"'), "rep": g(r'_report_id" value="(\d+)"'),
            "min": html.unescape(g(r'id="P8_MIN_DATE" value="([^"]*)"'))}


# ---------- daily: Produccion Horaria CSV ----------
def hourly_csv(c, indx, a, b):
    q = (f"{BASE}f?p={c['flow']}:{c['step']}:{c['sess']}:CSV:::P8_INDX,P8_FECHA_INICIAL,P8_FECHA_FINAL:"
         f"{indx},{a.month}/{a.day}/{a.year},{b.month}/{b.day}/{b.year}")
    r = get(q)
    if r is None or "csv" not in (r.headers.get("content-type") or ""):
        return None
    df = pd.read_csv(io.StringIO(r.content.decode("utf-8", "replace")))
    return df


def fetch_days(days, c):
    """date -> {fuel: MWh} and per-technology detail, for complete days only."""
    out, detail = {}, {}
    for a, b in std.ranges(days, CHUNK_DAYS):
        print(f"  ODS {a}..{b}", flush=True)
        per = {}
        for indx, (name, fuel) in TECH.items():
            df = hourly_csv(c, indx, a, b)
            if df is None:
                print(f"    {name}: no CSV", flush=True)
                continue
            if len(df) >= 249:
                print(f"    WARNING {name}: {len(df)} rows - report row limit reached, days may be incomplete", flush=True)
            if df.empty:
                continue
            hours = [h for h in df.columns if re.fullmatch(r"\d{2}", str(h))]
            df["date"] = pd.to_datetime(df["Fecha"], format="%m/%d/%Y", errors="coerce")
            vals = df[hours].apply(pd.to_numeric, errors="coerce")
            df["mwh"] = vals.sum(axis=1, min_count=1)
            df["nh"] = vals.notna().sum(axis=1)
            for d, g in df.groupby("date"):
                per.setdefault(d.date(), {})[name] = (float(g["mwh"].sum()), int(g["nh"].max()), fuel)
        for d, techs in per.items():
            if max(nh for _, nh, _ in techs.values()) < 24:
                continue  # day not complete yet
            row = dict.fromkeys(FUELS, 0.0)
            for name, (mwh, _, fuel) in techs.items():
                if fuel:
                    row[fuel] += mwh
            out[d] = row
            detail[d] = {name: round(mwh, 1) for name, (mwh, _, _) in techs.items()}
    return out, detail


# ---------- monthly: annual + monthly market reports ----------
def report_list(p4_id):
    r = get(f"{BASE}f?p=110:4:::::p4_id:{p4_id}")
    c = ctx(r.text)
    ck = re.search(r'data-for="P4_ID" value="([^"]+)"', r.text)
    data = {"p_flow_id": "110", "p_flow_step_id": "4", "p_instance": c["sess"], "p_debug": "",
            "p_request": "PLUGIN=" + c["ajax"], "p_widget_name": "worksheet", "p_widget_mod": "ACTION",
            "p_widget_action": "LAZY_LOAD", "p_widget_num_return": "250", "x01": c["ws"], "x02": c["rep"],
            "p_json": json.dumps({"pageItems": {"itemsToSubmit": [{"n": "P4_ID", "v": str(p4_id), "ck": ck.group(1) if ck else ""}],
                                                "protected": c["prot"], "rowVersion": "", "formRegionChecksums": []},
                                  "salt": c["salt"]})}
    htm = S.post(BASE + "wwv_flow.ajax", data=data, timeout=60).text
    rows = []
    for m in re.finditer(r"<tr[^>]*>(.*?)</tr>", htm, re.S):
        cells = [html.unescape(re.sub(r"<[^>]+>|\s+", " ", x)).strip() for x in re.findall(r"<td[^>]*>(.*?)</td>", m.group(1), re.S)]
        links = [BASE + html.unescape(h) for h in re.findall(r'href="([^"]+)"', m.group(1)) if "get_blob" in h]
        if cells and links:
            rows.append((cells[0], links[0]))
    return rows


def monthly_from_table(year, rows):
    """rows: list of (month index 1..12, 11 numbers) in ODS column order ->
    Pub Hidro, Pub Termica, Pub Subtotal, Priv Hidro, Priv Termica, Biomasa, Solar, Eolica, Geotermica, Priv Subtotal, Total."""
    out = {}
    for mi, v in rows:
        tot = v[10]
        if abs(v[0] + v[1] + v[3] + v[4] + v[5] + v[6] + v[7] + v[8] - tot) > 0.01 * max(tot, 1):
            print(f"    {year}-{mi:02d}: parts do not add up to total ({v}); skipped", flush=True)
            continue
        out[pd.Timestamp(year, mi, 1)] = {"Hydro_GWh": v[0] + v[3], "Thermal_GWh": v[1] + v[4], "Bioenergy_GWh": v[5],
                                          "Solar_GWh": v[6], "Wind_GWh": v[7], "Other_GWh": v[8], "Total_GWh": tot}
    return out


def parse_annual(pdf_bytes, year):
    import pdfplumber
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for p in pdf.pages[:80]:
            txt = p.extract_text() or ""
            if not re.search(rf"PRODUCCI[OÓ]N GENERAL DE ENERG[IÍ]A {year}", txt, re.I):
                continue
            rows = []
            for line in txt.splitlines():
                m = re.match(r"\s*(ENE|FEB|MAR|ABR|MAY|JUN|JUL|AGO|SEP|OCT|NOV|DIC)\s+(.*)", line)
                if not m:
                    continue
                nums = [float(x.replace(",", "")) for x in re.findall(r"-?[\d,]+\.\d+", m.group(2))]
                if len(nums) == 11:
                    rows.append((MESES.index(m.group(1)) + 1, nums))
            if rows:
                return monthly_from_table(year, rows)
    return {}


def parse_monthly_report(pdf_bytes, year):
    """Tabla 8 of the monthly report: a rotated table whose text comes out one value per line with spaces
    between characters. Values run column by column: 11 groups of (months so far + accumulated)."""
    import pdfplumber
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for p in pdf.pages[:40]:
            txt = p.extract_text() or ""
            if "Tabla 8 Producci" not in txt:
                continue
            lines = [ln.replace(" ", "") for ln in txt.splitlines()]
            months = [ln for ln in lines if ln in MESES]
            nums = [float(ln.replace(",", "")) for ln in lines if re.fullmatch(r"-?[\d,]+\.\d+", ln)]
            n = len(months) + 1          # + accumulated
            if not months:
                continue                 # the list of tables also names 'Tabla 8'
            if len(nums) < 11 * n:
                print(f"  Tabla 8: {len(months)} months, {len(nums)} values - layout not recognised", flush=True)
                return {}
            groups = [nums[k * n:(k + 1) * n] for k in range(11)]
            rows = [(MESES.index(mo) + 1, [g[i] for g in groups]) for i, mo in enumerate(months)]
            return monthly_from_table(year, rows)
    return {}


def monthly_history(existing):
    have = set() if existing is None or existing.empty else {i.year for i in existing.index}
    out = {}
    this_year = dt.date.today().year
    try:
        annual = report_list(41)
    except Exception as e:  # noqa: BLE001
        print(f"  annual report list failed: {e}", flush=True)
        annual = []
    for title, link in annual:
        m = re.search(r"[Aa]ño (\d{4})", title)
        if not m or "potencia" in title.lower():
            continue
        year = int(m.group(1))
        if year < 2021 or year in have:
            continue
        r = get(link)
        if r is None or r.content[:4] != b"%PDF":
            continue
        rows = parse_annual(r.content, year)
        print(f"  annual report {year}: {len(rows)} months", flush=True)
        out.update(rows)
    try:
        monthly = report_list(10)
    except Exception as e:  # noqa: BLE001
        print(f"  monthly report list failed: {e}", flush=True)
        monthly = []
    for title, link in monthly[:1]:
        m = re.search(r"(\d{4})", title)
        year = int(m.group(1)) if m else this_year
        r = get(link)
        if r is not None and r.content[:4] == b"%PDF":
            rows = parse_monthly_report(r.content, year)
            print(f"  monthly report '{title}': {len(rows)} months", flush=True)
            out.update(rows)
    if not out:
        return existing
    new = pd.DataFrame.from_dict(out, orient="index").reindex(columns=MONTHLY_COLS)
    new.index.name = "month"
    if existing is None or existing.empty:
        return new.sort_index().round(3)
    return std.merge(new, existing).reindex(columns=MONTHLY_COLS).round(3)


def to_frame(rows):
    df = pd.DataFrame.from_dict(rows, orient="index").reindex(columns=FUELS).fillna(0.0)
    df.index = pd.to_datetime(df.index)
    df = df.rename(columns={f: f"{f}_MWh" for f in FUELS})
    df["Total_MWh"] = df[COLS].sum(axis=1)
    df.index.name = "date"
    return df.sort_index().round(1)


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--out", default=OUT)
    args = ap.parse_args()
    print("MAPPING (ODS technology -> column):", {v[0]: v[1] for v in TECH.values()}, flush=True)

    S.verify = aia_bundle("cnd.enee.hn")
    daily = std.load_sheet(args.out, "Daily")
    if not daily.empty:
        daily = daily.reindex(columns=COLS + ["Total_MWh"]).fillna(0.0)
    detail = std.load_sheet(args.out, "By_technology_MWh")
    monthly = std.load_sheet(args.out, "Monthly_GWh")

    r = get(f"{APP}producci%C3%B3n-horaria?p8_indx=1")
    c = ctx(r.text)
    first = dt.datetime.strptime(c["min"], "%m/%d/%Y").date() if c["min"] else dt.date(2026, 6, 1)
    print(f"ODS session {c['sess']} (app {c['flow']} page {c['step']}); hourly data from {first}", flush=True)
    end = dt.date.today() - dt.timedelta(days=1)
    days = std.missing_days(daily, max(first, std.HISTORY_START), end, refresh_days=14)
    print(f"{0 if daily.empty else len(daily):,} days saved; fetching {len(days)}", flush=True)
    rows, det = fetch_days(days, c)
    if rows:
        daily = std.merge(to_frame(rows), daily)
        d = pd.DataFrame.from_dict(det, orient="index")
        d.index = pd.to_datetime(d.index)
        d.index.name = "date"
        detail = std.merge(d.sort_index(), detail)
    monthly = monthly_history(monthly)
    if daily.empty:
        print("No daily data returned.", flush=True)
        sys.exit(1)
    extra = {"By_technology_MWh": detail}
    if monthly is not None and not monthly.empty:
        extra["Monthly_GWh"] = monthly
    std.write(args.out, daily.reindex(columns=COLS + ["Total_MWh"]), NOTES, extra)
    if monthly is not None and not monthly.empty:
        print("Monthly_GWh (last 6):", flush=True)
        print(monthly.tail(6).to_string(), flush=True)
        m = daily.resample("MS").sum() / 1000
        n = daily["Total_MWh"].resample("MS").count()
        for i in m.index:
            if n[i] == i.days_in_month and i in monthly.index:
                print(f"  CHECK {i:%Y-%m}: daily-sum total {m.loc[i, 'Total_MWh']:,.1f} GWh vs ODS monthly "
                      f"{monthly.loc[i, 'Total_GWh']:,.1f} GWh", flush=True)


if __name__ == "__main__":
    main()
