"""
Philippines discovery, round 3: a resource -> fuel map for IEMOP's WESM resources (raw daily generation mix).
  iemop_pages  every market-data slug with post_id (find DIPC Energy Results FINAL, offers, must-run lists)
  energy       one complete day of DIPCER: MWh per resource (positive / negative), region, joined to CAPEG
               (TP_NAME = trading participant, MAXIMUM_CAPACITY); printed as CSV lines 'RES;...'
  doe          DOE 'List of Existing Power Plants' per grid (prod-cms / doe.gov.ph PDFs): table rows printed
  wiki         Wikipedia 'List of power stations in the Philippines' raw wikitext (plant / fuel / capacity)
"""
import base64
import io
import re
import sys
import zipfile
from concurrent.futures import ThreadPoolExecutor

import pandas as pd
import requests
import urllib3

urllib3.disable_warnings()
H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0 Safari/537.36",
     "Accept": "*/*", "Accept-Language": "en-US,en;q=0.9"}
T = (20, 120)
AJAX = "https://www.iemop.ph/wp-admin/admin-ajax.php"


def out(*a):
    print(*a, flush=True)


def get(u, quiet=False, **kw):
    try:
        r = requests.get(u, headers=H, timeout=T, verify=False, **kw)
        if not quiet:
            out(f"GET {r.url} -> {r.status_code} {r.headers.get('content-type')} {len(r.content)}")
        return r
    except Exception as e:  # noqa: BLE001
        out(f"GET {u} ERROR {type(e).__name__}: {str(e)[:150]}")
        return None


def ajax(post_id, page=1):
    x = requests.post(AJAX, data={"action": "display_filtered_market_data_files", "sort": "", "datefilter": "",
                                  "page": page, "post_id": post_id}, headers=H, timeout=T)
    j = x.json()
    return [base64.b64decode(s).decode() for s in (j.get("source") or [])], int(j.get("count") or 0)


def iemop_pages():
    r = get("https://www.iemop.ph/market-data/")
    slugs = sorted(set(re.findall(r'https://www\.iemop\.ph/market-data/([a-z0-9\-]+)/', r.text)))
    out(f"  {len(slugs)} slugs")

    def one(slug):
        p = get(f"https://www.iemop.ph/market-data/{slug}/", quiet=True)
        if p is None or not p.ok:
            return f"  {slug}: fail"
        m = re.search(r'"post_id":"(\d+)","min_date":"([^"]+)"', p.text)
        if not m:
            return f"  {slug}: no post_id"
        try:
            names, n = ajax(m.group(1))
            return f"  {slug}: post_id {m.group(1)} min {m.group(2)} count {n} first {[x.rsplit('/', 1)[-1] for x in names[:2]]}"
        except Exception as e:  # noqa: BLE001
            return f"  {slug}: post_id {m.group(1)} {e!r}"
    with ThreadPoolExecutor(8) as ex:
        for line in ex.map(one, slugs):
            out(line)


def listing(post_id, pages=6):
    files = []
    for p in range(1, pages + 1):
        names, n = ajax(post_id, p)
        files += names
        if not names or len(files) >= n:
            break
    return files


def energy():
    files = listing(5754, 4)
    by_day = {}
    for f in files:
        stamp = f.rsplit("_", 1)[-1].split(".")[0]
        d = pd.Timestamp(stamp[:8]).date()
        if stamp[8:12] == "0000":
            d -= pd.Timedelta(days=1)
        by_day.setdefault(d, []).append("https://www.iemop.ph" + f[f.find("/wp-content/"):])
    full = sorted(d for d, v in by_day.items() if len(v) >= 24)
    out(f"  DIPCER days listed {sorted(by_day)[:3]}..{sorted(by_day)[-3:]}; full {len(full)}")
    day = full[-3]

    def zipdf(u):
        r = get(u, quiet=True)
        with zipfile.ZipFile(io.BytesIO(r.content)) as zf:
            return pd.read_csv(zf.open(zf.namelist()[0]))
    with ThreadPoolExecutor(8) as ex:
        d = pd.concat(list(ex.map(zipdf, by_day[day])))
    d.columns = [c.strip() for c in d.columns]
    out(f"  day {day}: {d.shape}, columns {list(d.columns)}, intervals {d['TIME_INTERVAL'].nunique()}")
    for c in d.columns:
        if d[c].dtype == object and d[c].nunique() < 20:
            out(f"  {c}: {d[c].value_counts().to_dict()}")
    d["SCHED_MW"] = pd.to_numeric(d["SCHED_MW"], errors="coerce")
    d["pos"] = d["SCHED_MW"].clip(lower=0) * 5 / 60
    d["neg"] = d["SCHED_MW"].clip(upper=0) * 5 / 60
    g = d.groupby("RESOURCE_NAME").agg(region=("REGION_NAME", "first"), pos=("pos", "sum"), neg=("neg", "sum"),
                                       maxmw=("SCHED_MW", "max"))
    extra = [c for c in d.columns if "TYPE" in c.upper() or "FLAG" in c.upper()]
    for c in extra:
        g[c] = d.groupby("RESOURCE_NAME")[c].agg(lambda s: "/".join(sorted(set(map(str, s))))[:30])
    try:
        names, _ = ajax(302634)
        cap = pd.read_csv(io.BytesIO(get("https://www.iemop.ph" + names[0][names[0].find("/wp-content/"):]).content))
        cap.columns = [c.strip() for c in cap.columns]
        cap = cap.drop_duplicates("RESOURCE_NAME").set_index("RESOURCE_NAME")
        g = g.join(cap[["TP_NAME", "MAXIMUM_CAPACITY"]], how="outer")
    except Exception as e:  # noqa: BLE001
        out(f"  CAPEG {e!r}")
    g = g.sort_values("pos", ascending=False)
    tot = g["pos"].sum()
    out(f"  total positive {tot:,.0f} MWh, negative {g['neg'].sum():,.0f} MWh; resources {len(g)}")
    out("RES;name;region;MWh_pos;share_pct;MWh_neg;max_sched_MW;" + ";".join(extra) + ";TP_NAME;MAX_CAP")
    for n, x in g.iterrows():
        out(f"RES;{n};{x.get('region')};{x['pos']:.0f};{100 * (x['pos'] or 0) / tot:.3f};{x['neg']:.0f};"
            f"{x['maxmw']};" + ";".join(str(x[c]) for c in extra) + f";{x.get('TP_NAME')};{x.get('MAXIMUM_CAPACITY')}")


def pdf_rows(content, label, maxpages=60):
    import pdfplumber
    with pdfplumber.open(io.BytesIO(content)) as pdf:
        out(f"  PDF {label}: {len(pdf.pages)} pages")
        for i, p in enumerate(pdf.pages[:maxpages]):
            tabs = p.extract_tables()
            if not tabs:
                out(f"  p{i + 1} text: " + (p.extract_text() or "")[:3000].replace("\n", " | "))
            for t in tabs:
                for row in t:
                    out("DOE;" + label + ";" + " | ".join("" if c is None else str(c).replace("\n", " ") for c in row))


def doe():
    cands = set()
    for page in ("https://doe.gov.ph/sitemap.xml", "https://doe.gov.ph/electric-power",
                 "https://doe.gov.ph/data-and-prices/energy-statistics/electric-power-industry",
                 "https://doe.gov.ph/data-and-prices/energy-statistics/electric-power-industry/list-of-existing-power-plants",
                 "https://doe.gov.ph/data-and-prices/energy-statistics/electric-power-industry/2025-power-statistics",
                 "https://doe.gov.ph/data-and-prices/energy-statistics/electric-power-industry/2026-power-statistics",
                 "https://doe.gov.ph/list-existing-power-plants",
                 "https://legacy.doe.gov.ph/list-existing-power-plants"):
        r = get(page)
        if r is None or not r.ok:
            continue
        t = r.text.replace("\\u002F", "/").replace("\\/", "/")
        urls = set(re.findall(r'https?://[^"\'\s<>\\]+', t)) | {requests.compat.urljoin(r.url, h) for h in
                                                                   re.findall(r'href="([^"]+)"', t)}
        hit = sorted(u for u in urls if re.search(r"exist|power-plant|luzon|visayas|mindanao|lvm|loepp", u, re.I))
        out(f"  {len(urls)} urls; matching {len(hit)}: {hit[:80]}")
        cands |= {u for u in hit if re.search(r"prod-cms|\.pdf|\.xlsx?|/documents/", u, re.I)}
    cands |= {"https://doe.gov.ph/sites/default/files/pdf/electric_power/01_Luzon_Grid_31_aug_2024.pdf",
              "https://legacy.doe.gov.ph/sites/default/files/pdf/electric_power/01_Luzon%20Grid_1.pdf",
              "https://legacy.doe.gov.ph/sites/default/files/pdf/electric_power/03_Mindanao%20Grid_1.pdf",
              "https://doe.gov.ph/sites/default/files/pdf/electric_power/01_Luzon%20Grid_1.pdf"}
    out(f"  candidates: {sorted(cands)}")
    done = 0
    for u in sorted(cands, key=lambda s: ("prod-cms" not in s, s)):
        if done >= 8:
            break
        x = get(u)
        if x is None or not x.ok:
            continue
        if x.content[:4] == b"%PDF":
            pdf_rows(x.content, u.rsplit("/", 1)[-1][:40])
            done += 1
        elif x.content[:2] == b"PK":
            xl = pd.ExcelFile(io.BytesIO(x.content))
            for sh in xl.sheet_names:
                df = xl.parse(sh, header=None)
                for row in df.itertuples(index=False):
                    out("DOE;" + sh + ";" + " | ".join("" if pd.isna(c) else str(c) for c in row))
            done += 1


def wiki():
    for t in ("List_of_power_stations_in_the_Philippines",):
        r = get(f"https://en.wikipedia.org/w/index.php?title={t}&action=raw")
        if r is not None and r.ok:
            for line in r.text.splitlines():
                if line.strip():
                    out("WIKI;" + line[:300])


if __name__ == "__main__":
    for f in sys.argv[1:] or ["iemop_pages", "energy", "doe", "wiki"]:
        out(f"\n==================== {f}")
        try:
            globals()[f]()
        except Exception as e:  # noqa: BLE001
            out(f"!! {f}: {e!r}")
