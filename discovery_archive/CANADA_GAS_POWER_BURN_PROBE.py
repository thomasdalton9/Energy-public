"""
Canada gas demand by sector and province: what do the raw sources carry?  Read-only probe, runs in GitHub Actions only
(the editing sandbox blocks these hosts).

1. StatCan 25-10-0086-01 (gas supply and disposition, monthly): which provinces / items / periods exist, so the
   residential / commercial / industrial pull can keep every province instead of only GEO == "Canada".
2. StatCan 25-10-0015-01 (electric power generation, monthly): which types of generation exist per province
   (does a combustible-fuels / non-renewable total exist for each province?).
3. StatCan 25-10-0029-01 (annual energy supply and demand, TJ): does it carry natural gas used for electricity
   generation by province (the annual anchor for calibrating a gas-generation x heat-rate estimate)?
4. Fuel-level generation feeds: AESO (Alberta) and IESO (Ontario) public report pages - status, headers, file links.
5. The earlier checks: old 25-10-0055-01 and CER / NRCan pages.
"""
import io
import re
import zipfile

import pandas as pd
import requests

H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0 Safari/537.36"}
FLAG = re.compile(r"electric|thermal|power|generat|utilit|natural gas|combustible|non-renewable", re.I)


def out(*a):
    print(*a, flush=True)


def load(pid):
    r = requests.get(f"https://www150.statcan.gc.ca/n1/tbl/csv/{pid}-eng.zip", headers=H, timeout=(10, 600))
    z = zipfile.ZipFile(io.BytesIO(r.content))
    name = next(n for n in z.namelist() if n.endswith(".csv") and "MetaData" not in n)
    return pd.read_csv(z.open(name), dtype=str, encoding="utf-8-sig")


def dims(d):
    return list(d.columns[d.columns.get_loc("DGUID") + 1:d.columns.get_loc("UOM")])


def members(d):
    for c in dims(d):
        for v in d[c].dropna().unique():
            out(f"  {c}: {v}" + ("   <-- FLAG" if FLAG.search(v) else ""))


def per_geo(d, item_col, items, uom_filter=None):
    """Rows and period per GEO for the named items."""
    x = d[d[item_col].isin(items)]
    if uom_filter:
        x = x[x["UOM"].str.contains(uom_filter, case=False, na=False)]
    x = x.assign(ok=x["VALUE"].notna() & (x["VALUE"] != ""))
    g = x.groupby("GEO").agg(rows=("ok", "size"), with_value=("ok", "sum"), first=("REF_DATE", "min"),
                             last=("REF_DATE", "max"))
    out(g.to_string())


def table(pid, title, fn):
    out(f"=========== {pid}  {title}")
    try:
        d = load(pid)
        out("columns", list(d.columns), "rows", len(d), "period", d["REF_DATE"].min(), d["REF_DATE"].max())
        out("GEO values:", sorted(d["GEO"].dropna().unique()))
        members(d)
        fn(d)
    except Exception as e:  # noqa: BLE001
        out("  failed", type(e).__name__, str(e)[:300])


def gas0086(d):
    col = next(c for c in dims(d) if c != "GEO")
    for item in ("Residential consumption", "Commercial consumption", "Industrial consumption",
                 "Industrial deliveries from transmission pipelines",
                 "Industrial deliveries from distribution pipelines"):
        out(f"--- {item}  (Cubic metres) rows per GEO")
        per_geo(d, col, [item], "Cubic")


def gen0015(d):
    cols = dims(d)
    typ = next((c for c in cols if "type" in c.lower() and "electric" in c.lower()), cols[-1])
    out("type column:", typ)
    for t in d[typ].dropna().unique():
        if re.search(r"combustible|non-renewable|total all", t, re.I):
            out(f"--- {t}: rows per GEO")
            per_geo(d, typ, [t])


def energy0029(d):
    for c in dims(d):
        for v in d[c].dropna().unique():
            if re.search(r"natural gas", v, re.I):
                out(f"  natural-gas member in {c}: {v}")
    gas = d[d.apply(lambda r: any(re.search(r"natural gas", str(r[c]), re.I) for c in dims(d)), axis=1)]
    out("natural-gas rows:", len(gas))
    elec = gas[gas.apply(lambda r: any(re.search(r"electric", str(r[c]), re.I) for c in dims(d)), axis=1)]
    out("natural-gas rows that mention electric:", len(elec))
    for key, s in elec.groupby(dims(d)[0 if dims(d)[0] != "GEO" else 1]).size().head(40).items():
        out(f"    {key}: {s} rows")
    if len(elec):
        out(elec[["REF_DATE", "GEO"] + dims(d) + ["UOM", "VALUE"]].tail(25).to_string())


table("25100086", "gas supply and disposition (monthly)", gas0086)
table("25100015", "electric power generation (monthly)", gen0015)
table("25100029", "energy supply and demand (annual, TJ)", energy0029)
table("25100055", "old gas table", lambda d: None)

out("=========== generation-by-fuel feeds (AESO / IESO) and CER / NRCan pages")
URLS = (
    "https://www.aeso.ca/market/market-and-system-reporting/",
    "https://ets.aeso.ca/ets_web/ip/Market/Reports/CSDReportServlet",
    "https://www.aeso.ca/market/market-and-system-reporting/data-requests/",
    "https://reports-public.ieso.ca/public/GenOutputbyFuelHourly/",
    "https://reports-public.ieso.ca/public/GenOutputCapability/",
    "https://www.ieso.ca/power-data/data-directory",
    "https://www.cer-rec.gc.ca/en/data-analysis/energy-commodities/natural-gas/index.html",
    "https://www.cer-rec.gc.ca/en/data-analysis/canada-energy-future/index.html",
    "https://open.canada.ca/data/en/dataset?q=natural+gas+electricity+generation+consumption",
    "https://natural-resources.canada.ca/energy-facts/natural-gas-facts/20066",
)
for u in URLS:
    try:
        r = requests.get(u, headers=H, timeout=(10, 60))
        out(f"GET {u} -> {r.status_code} {len(r.content)}B  {r.headers.get('Content-Type', '')}")
        txt = r.text
        for m in sorted(set(re.findall(r'href="([^"]+\.(?:csv|xlsx?|zip|xml))"', txt)))[:40]:
            out("    file:", m)
        if "ets_web" in u or "GenOutput" in u:
            out("    head:", re.sub(r"\s+", " ", txt[:600]))
    except Exception as e:  # noqa: BLE001
        out(f"GET {u} -> ERR {type(e).__name__}")
