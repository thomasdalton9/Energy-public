"""
Pakistan natural gas, monthly, from the Pakistan Bureau of Statistics (PBS):
  1. Monthly Bulletin of Statistics (MBS), table 3.2 'Production of Natural Gas' - production by field, company and
     province, MMCFt (million cubic feet) per month, compiled from the Directorate General of Petroleum Concessions
     (Petroleum Division). Each issue holds the fiscal year to date (July-June).
  2. Monthly 'Statement showing imports of selected commodities' (Import_<Month>-<Year>.xlsx / .pdf), item
     24 'Natural gas, liquified' - LNG imports by VALUE (Rs million and US$ thousand); PBS publishes no LNG quantity.
Both are found through the PBS WordPress media library (https://www.pbs.gov.pk/wp-json/wp/v2/media). Found via
discovery_archive/asia/GAS_SSEA_DISCOVERY2.py .. 5.py. OGRA (403), Petroleum Division (yearbook to 2019-20), PLL and
DGPC (no DNS from GitHub runners) and SNGPL / SSGC (no flow data, only capacity declarations) gave nothing usable.

Writes output/Data and Chart Outputs/pakistan_gas.xlsx:
  Production   monthly MMCFD (= MMCFt in the month / days): Pakistan_total, Punjab, Sindh, KPK, Balochistan
  By field     monthly MMCFD per field (as named in the MBS; company in brackets)
  LNG imports  monthly LNG import value: LNG_USD_thousand, LNG_Rs_million (no volumes published)
  Files        the PBS files already read (the incremental record)

Incremental: the sheets are the history store. Each run lists the media library and downloads only files not read yet;
a later issue's figure for a month replaces an earlier one (revisions). Runs on the 1st and 15th.
Caveat: PBS stopped posting the MBS after the September 2024 issue (data to March 2024); the pull picks up new issues
automatically if they reappear.

    python3 asia/PAKISTAN_PBS_GAS.py [--out path]
"""
import argparse
import io
import os
import re
import sys
import time

import pandas as pd
import pdfplumber
import requests
import urllib3

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)
import xlsx_notes  # noqa: E402

urllib3.disable_warnings()
MEDIA = "https://www.pbs.gov.pk/wp-json/wp/v2/media"
H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0 Safari/537.36"}
T = (20, 240)
OUT = os.path.join(ROOT, "output", "Data and Chart Outputs", "pakistan_gas.xlsx")
MON = {m: i for i, m in enumerate(["jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec"], 1)}
COMPANIES = ("OGDCL", "OPL", "POL", "PPL", "MPCL", "DEWAN", "OPPL", "UEPL", "Eni (Pak)", "Eni(Pak)", "MOL", "PEL", "POGC",
             "Spud Energy", "OMV", "UEP", "Prime", "SPUD")
PROVINCES = {"punjab total": "Punjab", "sindh total": "Sindh", "kpk total": "KPK", "k.p.k total": "KPK",
             "balochistan total": "Balochistan", "pakistan total": "Pakistan_total"}


def out(*a):
    print(*a, flush=True)


def get(url, **kw):
    for i in range(4):
        try:
            r = requests.get(url, headers=H, timeout=T, verify=False, **kw)
            r.raise_for_status()
            return r
        except requests.RequestException as e:
            if i == 3:
                raise
            out(f"  retry {url[:100]}: {e}")
            time.sleep(5 * (i + 1))


def media(search):
    res, page = [], 1
    while True:
        try:
            r = get(MEDIA, params={"search": search, "per_page": 100, "page": page, "_fields": "date,source_url"})
        except requests.RequestException:
            break
        js = r.json()
        res += [(m["date"], m["source_url"]) for m in js]
        if not js or page >= int(r.headers.get("X-WP-TotalPages", page)):
            break
        page += 1
    return res


def list_files():
    mbs = {u for _, u in media("MBS") + media("Bulletin_of_Statistics") + media("Monthly-Bulletin") + media("MSB_")
           if u.lower().endswith(".pdf") and re.search(r"MBS|MSB|Bulletin_of_Stat|Bulletin-of-Stat", u, re.I)
           and not re.search(r"News|Current_Contents", u, re.I)}
    imp = {u for _, u in media("Import") if re.search(r"/Import[_-]", u) and re.search(r"\.(xlsx|pdf)$", u, re.I)}
    return sorted(mbs), sorted(imp)


def _month(tok_m, tok_y):
    m = MON.get(tok_m[:3].lower())
    return pd.Timestamp(int(tok_y), m, 1) if m else None


NUM = re.compile(r"^-?[\d,]*\.?\d+$")


def parse_mbs(content):
    """({month: {province/total: MMCFt}}, {month: {field: MMCFt}}) from table 3.2 of one MBS issue."""
    totals, fields = {}, {}
    with pdfplumber.open(io.BytesIO(content)) as pdf:
        for p in pdf.pages[:60]:
            tx = p.extract_text() or ""
            if not re.search(r"3\.2\s+Production of Natural Gas", tx) or "MMCF" not in tx.upper():
                continue
            lines = tx.splitlines()
            hdr = next((l for l in lines if re.search(r"Company\s+Field", l)), "")
            months = [_month(m, y) for m, y in re.findall(r"([A-Z][a-z]{2,5})\.?,?\s*(20\d\d)", hdr)]
            months = [m for m in months if m is not None]
            if not months:
                continue
            company = None
            for l in lines:
                toks = l.split()
                vals = []
                for t in reversed(toks):
                    if NUM.match(t):
                        vals.append(float(t.replace(",", "")))
                    else:
                        break
                vals = vals[::-1]
                name = " ".join(toks[:len(toks) - len(vals)])
                if len(vals) < 2 or not name:
                    continue
                mvals = vals[:-1][:len(months)]          # the last figure is the year-to-date total
                low = name.lower()
                if low in PROVINCES:
                    for d, v in zip(months, mvals):
                        totals.setdefault(d, {})[PROVINCES[low]] = v
                    continue
                if "total" in low or low in ("punjab", "sindh", "kpk", "balochistan"):
                    continue
                for c in COMPANIES:
                    if name.startswith(c + " "):
                        company, name = c.replace("Eni(Pak)", "Eni (Pak)"), name[len(c) + 1:]
                        break
                key = f"{name} ({company})" if company else name
                for d, v in zip(months, mvals):
                    fields.setdefault(d, {})[key] = v
    return totals, fields


def parse_import_xlsx(content):
    """{month: (USD thousand, Rs million)} for LNG from a monthly imports statement."""
    df = pd.read_excel(io.BytesIO(content), header=None)
    res = {}
    rows = [i for i in range(len(df)) if re.search(r"NATURAL\s*GAS,?\s*LIQUI", " ".join(map(str, df.iloc[i, :3])), re.I)]
    if not rows:
        return res
    i = rows[0]
    # the period header row: month names with years, one per column group (current, previous month, last year)
    hdr = None
    for k in range(min(i, 12)):
        cells = [(j, str(df.iat[k, j])) for j in range(df.shape[1]) if pd.notna(df.iat[k, j])]
        found = [(j, re.search(r"([A-Z]{3,9}),?\s*(20\d\d)", c, re.I)) for j, c in cells]
        found = [(j, m) for j, m in found if m and m.group(1)[:3].lower() in MON]
        if len(found) >= 2:
            hdr = found
            break
    if not hdr:
        return res
    starts = [j for j, _ in hdr]
    for n, (j, m) in enumerate(hdr):
        end = min(starts[n + 1] if n + 1 < len(starts) else df.shape[1], j + 3)   # (quantity, Rs, US$)
        nums = [pd.to_numeric(df.iat[i, c], errors="coerce") for c in range(j, end)]
        nums = [x for x in nums if pd.notna(x)]
        # each group is (quantity, Rs million, US$ thousand); LNG has no quantity, so the last two numbers
        if len(nums) >= 2:
            d = _month(m.group(1), m.group(2))
            if d is not None and d not in res:
                res[d] = (float(nums[-1]), float(nums[-2]))
    return res


def parse_import_pdf(content):
    res = {}
    with pdfplumber.open(io.BytesIO(content)) as pdf:
        tx = "\n".join((p.extract_text() or "") for p in pdf.pages[:4])
    hdr = re.findall(r"\b(JANUARY|FEBRUARY|MARCH|APRIL|MAY|JUNE|JULY|AUGUST|SEPTEMBER|OCTOBER|NOVEMBER|DECEMBER),?\s*"
                     r"(20\d\d)", tx, re.I)
    seen = []
    for m, y in hdr:
        d = _month(m, y)
        if d not in seen:
            seen.append(d)
    line = next((l for l in tx.splitlines() if re.search(r"NATURAL\s*GAS,?\s*LIQUI", l, re.I)), None)
    if not line or not seen:
        return res
    nums = [float(t.replace(",", "")) for t in line.split() if NUM.match(t)]
    for k, d in enumerate(seen[:3]):
        if len(nums) >= 2 * k + 2:
            res[d] = (nums[2 * k + 1], nums[2 * k])
    return res


def load(path, sheet):
    try:
        d = pd.read_excel(path, sheet_name=sheet, index_col=0)
        if sheet != "Files":
            d.index = pd.to_datetime(d.index)
        return d
    except (FileNotFoundError, ValueError):
        return pd.DataFrame()


def merge(old, new):
    if new.empty:
        return old
    if old.empty:
        return new.sort_index()
    m = new.combine_first(old)
    return m[sorted(m.columns, key=lambda c: list(new.columns).index(c) if c in new.columns else 999)].sort_index()


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", default=OUT)
    args = ap.parse_args()
    prod, field, lng, files = (load(args.out, s) for s in ("Production", "By field", "LNG imports", "Files"))
    done = set(files.index) if not files.empty else set()
    mbs, imp = list_files()
    out(f"{len(mbs)} MBS issues, {len(imp)} import statements listed; {len(done)} files already read")
    tot_new, fld_new, lng_new, read = {}, {}, {}, []
    # oldest issue first (by the period in the file name is unreliable; the upload order is close enough, and later
    # issues of the same fiscal year carry the same months revised) - so sort by the latest month each one holds
    parsed = []
    for u in mbs:
        if u in done:
            continue
        try:
            t, f = parse_mbs(get(u).content)
        except Exception as e:  # noqa: BLE001
            out(f"  {u.rsplit('/', 1)[-1]}: {type(e).__name__}: {e}")
            continue
        out(f"  {u.rsplit('/', 1)[-1]}: {len(t)} months" + (f" {min(t):%Y-%m}..{max(t):%Y-%m}" if t else ""))
        read.append({"url": u, "kind": "MBS", "months": len(t)})
        if t:
            parsed.append((max(t), t, f))
    for _, t, f in sorted(parsed, key=lambda x: x[0]):
        for d, v in t.items():
            tot_new.setdefault(d, {}).update(v)
        for d, v in f.items():
            fld_new.setdefault(d, {}).update(v)
    pi = []
    for u in imp:
        if u in done:
            continue
        try:
            c = get(u).content
            res = parse_import_xlsx(c) if u.lower().endswith(".xlsx") else parse_import_pdf(c)
        except Exception as e:  # noqa: BLE001
            out(f"  {u.rsplit('/', 1)[-1]}: {type(e).__name__}: {e}")
            continue
        out(f"  {u.rsplit('/', 1)[-1]}: " + ", ".join(f"{d:%Y-%m}=${v[0] / 1000:.0f}m" for d, v in sorted(res.items())))
        read.append({"url": u, "kind": "Imports", "months": len(res)})
        if res:
            pi.append((max(res), res))
    for _, res in sorted(pi, key=lambda x: x[0]):
        for d, v in res.items():
            if d == max(res) or d not in lng_new:      # a statement's own month (revised later) wins over carry-overs
                lng_new[d] = v
    if tot_new:
        t = pd.DataFrame(tot_new).T.sort_index()
        t = t.div(t.index.days_in_month, axis=0).round(1)
        prod = merge(prod, t[[c for c in ("Pakistan_total", "Punjab", "Sindh", "KPK", "Balochistan") if c in t]])
    if fld_new:
        f = pd.DataFrame(fld_new).T.sort_index()
        f = f.div(f.index.days_in_month, axis=0).round(2)
        field = merge(field, f[sorted(f.columns, key=lambda c: -f[c].tail(12).mean())])
    if lng_new:
        n = pd.DataFrame({d: {"LNG_USD_thousand": v[0], "LNG_Rs_million": v[1]} for d, v in lng_new.items()}).T
        lng = merge(lng, n.sort_index())
    if read:
        nf = pd.DataFrame(read).set_index("url")
        files = pd.concat([files, nf]) if not files.empty else nf
    for v in (prod, field, lng):
        if not v.empty:
            v.index.name = "date"
    files.index.name = "url"
    if prod.empty and lng.empty:
        raise SystemExit("No Pakistan gas data")
    if not prod.empty:
        out(prod.tail(6).to_string())
        out(f"last 12 months production: {prod['Pakistan_total'].tail(12).mean():.0f} MMCFD")
    if not lng.empty:
        out(lng.tail(8).to_string())
    notes = [
        "UNITS",
        "Production / By field: MMCFD = million cubic feet per day, monthly average (PBS publishes MMCFt per month; "
        "divided by the days in the month here). 1 MMCFD = 0.0283 million m3 per day (charts show mcm/d). "
        "Pakistan_total is domestic production only (system gas); RLNG (imported LNG) is not included.",
        "LNG imports: value of liquefied natural gas imported in the month (PBS external trade), US$ thousand and "
        "Rs million. PBS publishes no LNG quantity; at about $10/MMBtu, $100 million is roughly 10 million MMBtu "
        "(about 0.3 bcf/d over a month).",
        "",
        "COVERAGE",
        (f"Production monthly {prod.index.min():%Y-%m}..{prod.index.max():%Y-%m}; " if not prod.empty else "") +
        (f"LNG imports monthly {lng.index.min():%Y-%m}..{lng.index.max():%Y-%m}. " if not lng.empty else "") +
        "PBS has not posted a Monthly Bulletin of Statistics after the September 2024 issue (data to March 2024); the "
        "import statements continue monthly (about 18 days after the month).",
        f"{len(files)} PBS files read so far (Files sheet).",
        "",
        "SOURCE",
        "Pakistan Bureau of Statistics: Monthly Bulletin of Statistics, table 3.2 Production of Natural Gas (source: "
        "Directorate General of Petroleum Concessions, Petroleum Division); monthly statement of imports of selected "
        "commodities (item 24, natural gas liquified). https://www.pbs.gov.pk/publication-2/",
    ]
    sheets = {k: v for k, v in (("Production", prod), ("By field", field), ("LNG imports", lng)) if not v.empty}
    sheets["Files"] = files
    xlsx_notes.write_workbook(args.out, sheets, notes, {"UNITS", "COVERAGE", "SOURCE"})
    out(f"Saved {args.out}")


if __name__ == "__main__":
    main()
