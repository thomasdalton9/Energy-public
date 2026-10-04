"""
Indonesia natural gas from the Directorate General of Oil and Gas (Ditjen Migas, Ministry of Energy and Mineral
Resources / ESDM), "Buku Statistik Minyak dan Gas Bumi" (Oil and Gas Statistics), published twice a year (Semester I
and full year): https://migas.esdm.go.id/post/buku-statistik-migas
Found via discovery_archive/asia/GAS_SSEA_DISCOVERY3.py / 4.py / 5.py. (SKK Migas' own site blocks GitHub runners;
the ESDM homepage 'Highlight' card is only a year-to-date snapshot and its time series returns no data.)

  Table 1.6  Monitoring Produksi Gas Bumi Indonesia - natural gas production by production-sharing contractor (KKKS),
             MMSCFD, monthly (books from Semester I 2021 on; data from SKK Migas)
  Table 1.7  Pemanfaatan Gas Bumi Dalam Negeri - natural gas utilisation by sector, BBTUD, annual from 2016
             (domestic: fertiliser, electricity, industry, city gas, gas fuel, lifting, domestic LNG / LPG;
             exports: pipeline gas, LNG)

Writes output/Data and Chart Outputs/indonesia_gas.xlsx:
  Production   monthly MMSCFD: Total (national), then one column per KKKS
  Utilisation  annual BBTUD by sector (latest full-year book; the year shown is the calendar year)
  Books        the books read (url, Last-Modified, months found) - the incremental record

Whole-file source: each run lists the books and downloads only a book whose URL is new or whose Last-Modified has
changed; values from a full-year book replace the Semester I book's (the later book wins). Runs on the 1st and 15th.

    python3 asia/INDONESIA_MIGAS_GAS.py [--out path]
"""
import argparse
import io
import os
import re
import sys
import time
from urllib.parse import urljoin

import pandas as pd
import pdfplumber
import requests
import urllib3

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)
import xlsx_notes  # noqa: E402

urllib3.disable_warnings()
PAGE = "https://migas.esdm.go.id/post/buku-statistik-migas"
H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0 Safari/537.36"}
T = (20, 240)
OUT = os.path.join(ROOT, "output", "Data and Chart Outputs", "indonesia_gas.xlsx")
MON = {"jan": 1, "feb": 2, "mar": 3, "apr": 4, "mei": 5, "may": 5, "jun": 6, "jul": 7, "agt": 8, "ags": 8, "aug": 8,
       "agu": 8, "sep": 9, "okt": 10, "oct": 10, "nov": 11, "nop": 11, "des": 12, "dec": 12}
UTIL = [("Total_utilisation", r"total pemanfaatan"), ("Domestic", r"realisasi domestik"),
        ("Export", r"realisasi ekspor"), ("Gas_fuel_BBG", r"\bbbg\b"), ("City_gas", r"city gas"),
        ("Lifting_own_use", r"\blifting\b"), ("Fertiliser", r"pupuk"), ("Electricity", r"kelistrikan"),
        ("Industry", r"industri"), ("Domestic_LNG", r"domestik lng"), ("Domestic_LPG", r"domestik lpg"),
        ("Pipeline_export", r"ekspor gas pipa"), ("LNG_export", r"ekspor lng")]


def out(*a):
    print(*a, flush=True)


def get(url, method="get"):
    for i in range(4):
        try:
            r = requests.request(method, url, headers=H, timeout=T, verify=False, allow_redirects=True)
            r.raise_for_status()
            return r
        except requests.RequestException as e:
            if i == 3:
                raise
            out(f"  retry {url[:100]}: {e}")
            time.sleep(5 * (i + 1))


def book_links():
    r = get(PAGE)
    links = re.findall(r'href="([^"]+\.pdf)[^"]*"', r.text, re.I)
    return sorted({urljoin(r.url, l) for l in links if re.search(r"Statisti", l, re.I)})


def number(tok, mode):
    """mode 'dot': 1.234,5 style (dot thousands); 'comma': 1,234.5 style."""
    tok = tok.strip()
    if tok in ("-", "–", "—", ""):
        return 0.0
    if mode == "dot":
        tok = tok.replace(".", "").replace(",", ".")
    else:
        tok = tok.replace(",", "")
    try:
        return float(tok)
    except ValueError:
        return None


def pick_mode(tok, lo, hi):
    for mode in ("dot", "comma"):
        v = number(tok, mode)
        if v is not None and lo <= v <= hi:
            return mode
    return "comma"


NUM = re.compile(r"^(?:-|–|\d[\d.,]*)$")


def _num_tail(toks):
    vals = []
    for t in reversed(toks):
        if NUM.match(t):
            vals.append(t)
        else:
            break
    return vals[::-1]


def parse_production(pdf):
    """{month: {KKKS or 'Total': MMSCFD}} from Table 1.6 (1.3 in older books) and its continuation page.
    Handles names wrapped above / below a values-only line, and two-column page layouts (TOTAL found anywhere)."""
    rows, total, year, months, active = [], None, None, None, False
    for p in pdf.pages:
        tx = p.extract_text() or ""
        title = re.search(r"Monitoring (?:Produksi Gas Bumi|of Indonesian Natural Gas Production)", tx, re.I)
        if not (title or (active and "MMSCFD" in tx)):
            active = False
            continue
        lines = tx.splitlines()
        hdr = next((l for l in lines if len([t for t in l.split() if t[:3].lower() in MON]) >= 6), None)
        if not hdr or "MMSCFD" not in tx:
            active = False
            continue
        active = True
        y = re.search(r"Monitoring (?:Produksi Gas Bumi Indonesia|of Indonesian Natural Gas Production)\s+"
                      r"(?:Semester\s+I+\s+)?(\d{4})", tx, re.I)
        year = int(y.group(1)) if y else year
        mon = []
        for t in hdr.split():
            m = MON.get(t[:3].lower())
            if m and m not in mon:
                mon.append(m)
        months = mon
        n = len(months)
        for k, l in enumerate(lines):
            toks = l.split()
            if not toks:
                continue
            up = [t.upper() for t in toks]
            for word in ("TOTAL", "JUMLAH"):
                if word in up:
                    v = [t for t in toks[up.index(word) + 1:] if NUM.match(t)]
                    if len(v) >= n:
                        total = v[:n]
            if "TOTAL" in up or "JUMLAH" in up:
                continue
            vals = _num_tail(toks)
            head = toks[:len(toks) - len(vals)]
            if len(vals) < n:
                continue
            if head and re.match(r"^\d{1,3}$", head[0]):
                head = head[1:]
            elif not head and re.match(r"^\d{1,3}$", vals[0]) and len(vals) > n + 1:
                vals = vals[1:]          # row number printed on the values line
            if len(vals) > n + 1:        # two rows side by side (two-column page): keep the first
                vals = vals[:n + 1]
            name = " ".join(head)
            if not re.search(r"[A-Za-z]", name):
                # values-only line: the name sits on the text line(s) above and/or below
                above = lines[k - 1] if k and not _num_tail(lines[k - 1].split()) else ""
                below = lines[k + 1] if k + 1 < len(lines) and not _num_tail(lines[k + 1].split()) else ""
                name = re.sub(r"^\d{1,3}\s+", "", f"{above} {below}".strip())
            if not re.search(r"[A-Za-z]", name):
                continue
            rows.append((name, vals[:n]))
    if not year or not months:
        return {}
    probe = total[0] if total else (rows[0][1][0] if rows else "0")
    mode = pick_mode(probe, 2500, 12000) if total else ("dot" if re.search(r",\d{1,2}$", probe) else "comma")
    res = {}
    for name, vals in rows:
        for m, t in zip(months, vals):
            v = number(t, mode)
            if v is not None:
                res.setdefault(pd.Timestamp(year, m, 1), {})[clean(name)] = v
    for d in res:
        res[d]["Total"] = None
    if total:
        for m, t in zip(months, total):
            res.setdefault(pd.Timestamp(year, m, 1), {})["Total"] = number(t, mode)
    for d, v in res.items():   # no printed total: the sum of the contractors
        if not v.get("Total"):
            v["Total"] = round(sum(x for k, x in v.items() if k != "Total" and x), 2)
            v["Total_is_sum"] = 1
    return res


def clean(name):
    """Contractor names differ in case and punctuation between books: upper-case, no dots / commas."""
    name = re.sub(r"[.,]", " ", name.upper().replace("MONTD'OR", "MONTDOR"))
    return re.sub(r"\s+", " ", name).strip(" -/")


def parse_utilisation(pdf):
    """DataFrame (year x sector, BBTUD) from Table 1.7 of a full-year book."""
    for p in pdf.pages:
        tx = p.extract_text() or ""
        if not re.search(r"Pemanfaatan Gas Bumi Dalam Negeri", tx, re.I) or "BBTUD" not in tx:
            continue
        lines = tx.splitlines()
        yl = next((l for l in lines if len(re.findall(r"\b20[12]\d\b", l)) >= 4), None)
        if not yl:
            continue
        years = [int(y) for y in re.findall(r"\b(20[12]\d)\b", yl)]
        semester = bool(re.search(r"SM\s*T?\s*I\b|Semester", yl, re.I))
        data, seen = {}, set()
        for i, l in enumerate(lines):
            if "BBTUD" not in l:
                continue
            label, rest = l.split("BBTUD", 1)
            j = i - 1
            while not re.search(r"[A-Za-z]{3}", label.replace("-", "")) or not label.strip().startswith("-"):
                # the label wrapped onto the line(s) above: walk back to the line that starts the row ('- ...')
                if j < 0 or j < i - 3 or "BBTUD" in lines[j]:
                    break
                label = lines[j] + " " + label
                j -= 1
            # ... or continues on the next line ('- Total' / 'Pemanfaatan Gas Bumi')
            if i + 1 < len(lines) and "BBTUD" not in lines[i + 1] and not re.search(r"\d", lines[i + 1]):
                label = label + " " + lines[i + 1]
            vals = [t for t in rest.split() if NUM.match(t)]
            for key, pat in UTIL:
                if re.search(pat, label, re.I) and key not in seen:
                    seen.add(key)
                    data[key] = vals
                    break
        if "Total_utilisation" not in data:
            continue
        mode = pick_mode(data["Total_utilisation"][0], 3000, 12000)
        n = len(years) - (1 if semester else 0)
        df = pd.DataFrame({k: [number(t, mode) for t in v[:n]] + [None] * (n - len(v[:n])) for k, v in data.items()},
                          index=pd.Index(years[:n], name="year"))
        return df
    return pd.DataFrame()


def load(path):
    try:
        prod = pd.read_excel(path, sheet_name="Production", index_col=0)
        prod.index = pd.to_datetime(prod.index)
        books = pd.read_excel(path, sheet_name="Books", index_col=0)
        util = pd.read_excel(path, sheet_name="Utilisation", index_col=0)
    except (FileNotFoundError, ValueError):
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame()
    return prod, books, util


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", default=OUT)
    args = ap.parse_args()
    prod, books, util = load(args.out)
    links = book_links()
    out(f"{len(links)} books listed")
    new_rows, new_prod, parsed, latest_util = [], {}, [], None
    for u in links:
        try:
            lm = get(u, "head").headers.get("last-modified")
        except Exception as e:  # noqa: BLE001
            out(f"  {u}: {type(e).__name__}")
            continue
        if not books.empty and u in books.index and str(books.at[u, "last_modified"]) == str(lm):
            continue
        name = u.rsplit("/", 1)[-1]
        try:
            r = get(u)
            with pdfplumber.open(io.BytesIO(r.content)) as pdf:
                first = (pdf.pages[0].extract_text() or "") + (pdf.pages[2].extract_text() or "" if len(pdf.pages) > 2 else "")
                res = parse_production(pdf)
                ut = parse_utilisation(pdf) if not re.search(r"Semester", first + name, re.I) else pd.DataFrame()
        except Exception as e:  # noqa: BLE001
            out(f"  {name}: {type(e).__name__}: {e}")
            continue
        out(f"  {name}: {len(res)} months" + (f" {min(res):%Y-%m}..{max(res):%Y-%m}, total "
                                              + ", ".join(f"{d:%m}={v.get('Total', 0):.0f}" for d, v in sorted(res.items()))
                                              if res else "") + (f"; utilisation {list(ut.index)}" if not ut.empty else ""))
        new_rows.append({"url": u, "last_modified": lm, "months": len(res),
                         "period": f"{min(res):%Y-%m}..{max(res):%Y-%m}" if res else ""})
        semester = bool(re.search(r"Semester|SMT|SM I", name + first, re.I))
        if res:
            parsed.append(((max(res), not semester), res))
        if not ut.empty and (latest_util is None or ut.index.max() >= latest_util.index.max()):
            latest_util = ut
    # apply oldest book first, so a later book - and a full-year book over its Semester I book - wins
    for _, res in sorted(parsed, key=lambda x: x[0]):
        new_prod.update(res)
    if new_prod:
        nd = pd.DataFrame(new_prod).T.sort_index()
        lead = [c for c in ("Total", "Total_is_sum") if c in nd]
        nd = nd[lead + sorted(c for c in nd.columns if c not in lead)]
        prod = pd.concat([prod[~prod.index.isin(nd.index)], nd]) if not prod.empty else nd
        lead = [c for c in ("Total", "Total_is_sum") if c in prod]
        prod = prod[lead + [c for c in prod.columns if c not in lead]].sort_index()
    if prod.empty:
        raise SystemExit("No Indonesia gas production table parsed")
    prod.index.name = "date"
    if latest_util is not None:
        util = pd.concat([util[~util.index.isin(latest_util.index)], latest_util]).sort_index() if not util.empty \
            else latest_util
    if new_rows:
        nb = pd.DataFrame(new_rows).set_index("url")
        books = pd.concat([books[~books.index.isin(nb.index)], nb]) if not books.empty else nb
    books.index.name = "url"
    # sanity: KKKS sum vs Total
    kk = prod.drop(columns=[c for c in ("Total", "Total_is_sum") if c in prod]).sum(axis=1, min_count=1)
    chk = pd.DataFrame({"Total": prod["Total"], "Sum_KKKS": kk.round(0)})
    out(chk.tail(14).to_string())
    out(f"last 12 months average: {prod['Total'].tail(12).mean():.0f} MMSCFD = {prod['Total'].tail(12).mean() / 1000:.2f} "
        "bcf/d")
    if not util.empty:
        out(util.tail(4).to_string())
    notes = [
        "UNITS",
        "Production: MMSCFD = million standard cubic feet per day, monthly average (Ditjen Migas / SKK Migas unit). "
        "1 MMSCFD = 0.0283 million m3 per day (the charts show mcm/d). Total = national natural gas production as "
        "printed in the table (the KKKS columns may not sum exactly to it); where a book prints no total row "
        "(Total_is_sum = 1) it is the sum of the contractors read. Contractor names are as printed (upper-cased), so a "
        "contractor renamed between books appears under both names.",
        "Utilisation: BBTUD = billion British thermal units per day, annual average. At about 1,000-1,100 BTU/scf, "
        "1 BBTUD is roughly 0.9-1.0 MMSCFD. Sectors: Fertiliser (pupuk), Electricity (kelistrikan), Industry, City_gas, "
        "Gas_fuel_BBG (transport), Lifting_own_use (oil lifting), Domestic_LNG / Domestic_LPG (LNG and LPG plants "
        "supplying the domestic market), Pipeline_export, LNG_export; Domestic + Export = Total_utilisation.",
        "",
        "COVERAGE",
        (f"Production monthly {prod.index.min():%Y-%m}..{prod.index.max():%Y-%m}" if not prod.empty else "") +
        (f"; utilisation annual {util.index.min()}..{util.index.max()}" if not util.empty else "") +
        ". The books appear about 6 months after each half-year (Semester I book around December, full-year book "
        "around mid-year), so the latest months lag well behind.",
        "Books read: " + "; ".join(f"{u.rsplit('/', 1)[-1]} (Last-Modified {r['last_modified']})"
                                   for u, r in books.iterrows()),
        "",
        "SOURCE",
        "Direktorat Jenderal Minyak dan Gas Bumi (Ditjen Migas), Kementerian ESDM, Buku Statistik Minyak dan Gas Bumi, "
        "Table 1.6 (Monitoring Produksi Gas Bumi Indonesia, data from SKK Migas) and Table 1.7 (Pemanfaatan Gas Bumi "
        "Dalam Negeri): https://migas.esdm.go.id/post/buku-statistik-migas",
    ]
    sheets = {"Production": prod.round(1)}
    if not util.empty:
        sheets["Utilisation"] = util.round(2)
    sheets["Books"] = books
    xlsx_notes.write_workbook(args.out, sheets, notes, {"UNITS", "COVERAGE", "SOURCE"})
    out(f"Saved {args.out}")


if __name__ == "__main__":
    main()
