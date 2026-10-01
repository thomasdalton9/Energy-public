"""
Paraguay power generation (monthly), Paraguay's own consumption from its
plants, and exports - Itaipu energy ceded to Brazil and Yacyreta energy to
Argentina. Writes the standard power layout (sheet "Daily", one row per
MONTH dated the 1st, MWh in the month) plus the full-plant figures, the
export split and the series the South America total uses.

Paraguay's power is almost all hydro from two binational plants it owns half
of (Itaipu with Brazil, Yacyreta with Argentina) plus ANDE's Acaray. Under the
treaties each country owns 50% of each plant's output; what one side does not
use is ceded (sold) to the other. So:
    Paraguay share   = 50% of Itaipu + 50% of Yacyreta   (sheet "Daily")
    Paraguay use     = energy ANDE takes from Itaipu + Yacyreta's supply to the
                       Paraguayan grid (SINP)
    Ceded to Brazil  = 50% of Itaipu's supply - ANDE's take
    Ceded to Arg.    = 50% of Yacyreta's supply - SINP

Sources (found with discovery_archive/south_america/PARAGUAY_POWER_DISCOVERY.py):
  1. ITAIPU Binacional monthly production reports, published as news posts on
     itaipu.gov.py (WordPress REST API /wp-json/wp/v2/posts): each month's total
     generation, the 50 Hz sector's generation and the energy supplied to ANDE,
     plus ANDE's year-to-date total. From 2021 (posts go back to 2019).
  2. ONS Brazil open data, geracao_usina_2_ho (hourly, per plant): 'ITAIPU 60 HZ'
     + 'ITAIPU 50 HZ' = Itaipu's energy delivered to Brazil (ENBPar). Checked
     against Itaipu's 2025 annual report: within 0.1% every month. Daily.
  3. Entidad Binacional Yacyreta (eby.gov.py) monthly 'Datos oficiales sobre
     generacion' posts: energy supplied to Argentina's SADI and to Paraguay's
     SINP, metered at the SR1 / SR2 substations (2018 - Nov 2023, not every month).
  4. CAMMESA post-operation daily database (PARTE_POST_OPERATIVO, .mdb): Yacyreta's
     groups YACYHI + YACYHIPY (energy delivered to the SADI; Nov-2023 sum is within
     0.2% of EBY's SADI figure) and ANDE's market-node sales to Argentina. Daily.
  5. EBY Argentina home page (eby.org.ar): Yacyreta's total net generation of the
     latest month - stored each run (no archive exists, so history starts when this
     pull started reading it).
  6. VMME (Viceministerio de Minas y Energia) Balance Energetico Nacional (annual
     PDF): exports to Argentina (Yacyreta cession) and Brazil (Itaipu cession) -
     anchors the annual SINP total where EBY's monthly figures are missing, and
     validation.
  ANDE's own statistics (ande.gov.py) are behind a Radware bot captcha, so ANDE's
  Acaray plant (210 MW) and its small thermal/solar units have no monthly source
  and are NOT included (about 1-2% of Paraguay's share; see Validation sheet).

Incremental: each source is archived in the workbook; ONS months are fetched
only when missing (plus the last two), CAMMESA days only when missing (plus the
last three, newest first, within --budget-min), Itaipu/EBY posts only after the
last month already parsed.

Usage: python3 PARAGUAY_POWER.py [--out PATH] [--budget-min 40] [--workers 4]
"""

print("STARTING", flush=True)

import argparse
import csv
import datetime as dt
import html
import io
import os
import re
import ssl
import subprocess
import sys
import tempfile
import time
import zipfile
from collections import defaultdict
from concurrent.futures import ThreadPoolExecutor, as_completed

import pandas as pd
import requests

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))  # repo root

import xlsx_notes  # noqa: E402

OUT = "output/Data and Chart Outputs/paraguay_power_generation_daily.xlsx"
START = pd.Timestamp("2021-01-01")
UA = {"User-Agent": "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36",
      "Accept-Language": "es-PY,es;q=0.9,en;q=0.7"}
MONTHS = {"enero": 1, "febrero": 2, "marzo": 3, "abril": 4, "mayo": 5, "junio": 6, "julio": 7, "agosto": 8,
          "septiembre": 9, "setiembre": 9, "octubre": 10, "noviembre": 11, "diciembre": 12}

ITAIPU_WP = "https://www.itaipu.gov.py/wp-json/wp/v2/posts"
ITAIPU_QUERIES = ["suministró", "inyectó", "suministro de energía"]
ONS_BASE = "https://ons-aws-prod-opendata.s3.amazonaws.com/dataset/geracao_usina_2_ho/"
EBY_SEARCH = "https://www.eby.gov.py/{page}?s={q}"
EBY_QUERIES = ["Datos+oficiales+generaci%C3%B3n", "Informe+generaci%C3%B3n+Yacyret%C3%A1",
               "generaci%C3%B3n+de+energ%C3%ADa+Yacyret%C3%A1"]
EBYAR_HOME = "https://www.eby.org.ar/"
SECTIGO_OV_R36 = "http://crt.sectigo.com/SectigoPublicServerAuthenticationCAOVR36.crt"
BEN_URLS = {2025: "https://minasyenergia.mopc.gov.py/pdf/balance2025/BEN 2025_Paraguay_final.pdf",
            2024: "https://minasyenergia.mopc.gov.py/pdf/balance2024/BEN 2024_Paraguay_final.pdf",
            2023: "https://minasyenergia.mopc.gov.py/pdf/balance2023/BEN 2023_Paraguay_final.pdf",
            2022: "https://minasyenergia.mopc.gov.py/pdf/balance2022/BEN 2022_Paraguay_final.pdf"}

S_ITAIPU, S_ONS, S_EBY, S_EBYAR, S_CAM, S_BEN = ("Itaipu reports", "Itaipu ONS daily", "Yacyreta EBY reports",
                                                 "Yacyreta eby.org.ar", "Yacyreta CAMMESA daily", "BEN annual")
S_PLANTS, S_EXPORTS, S_SA, S_VALID = "Full plants", "Exports", "Not counted by ONS-CAMMESA", "Validation"


def num(s):
    """Spanish-formatted number ('19.276', '1.549.645,8', '2,5') -> float."""
    s = str(s).strip().replace(" ", "")
    if "," in s:
        s = s.replace(".", "").replace(",", ".")
    elif s.count(".") >= 1 and all(len(p) == 3 for p in s.split(".")[1:]):
        s = s.replace(".", "")
    return float(s)


def load(path, sheet, index_col=0):
    try:
        df = pd.read_excel(path, sheet_name=sheet, index_col=index_col)
    except (FileNotFoundError, ValueError, KeyError, OSError):
        return pd.DataFrame()
    if index_col is not None:
        df.index = pd.to_datetime(df.index, errors="coerce")
        df = df[df.index.notna()]
        df.index.name = "date"
    return df


def merge(new, old):
    if old is None or old.empty:
        return new.sort_index()
    if new is None or new.empty:
        return old.sort_index()
    out = pd.concat([old[~old.index.isin(new.index)], new])
    return out[~out.index.duplicated(keep="last")].sort_index()


# ------------------------------------------------------------------ 1. Itaipu monthly reports (itaipu.gov.py)
PERIOD_WORDS = [(r"primer mes|mes de enero|durante (?:el mes de )?enero", 1),
                (r"primeros dos meses|primer bimestre", 2), (r"primer trimestre", 3),
                (r"primer cuatrimestre", 4), (r"primeros cinco meses", 5), (r"primer semestre", 6)]


def itaipu_period(title, text, posted):
    """(year, month) a report covers: 'de enero a agosto' -> Aug; 'primer semestre' -> Jun; 'durante el 2025' ->
    Dec 2025. The year is the post's year unless the month is not before the post month (January posts on the
    previous December/year)."""
    t = html.unescape(f"{title} {text}").lower()
    month = None
    m = re.search(r"(?:enero\s+(?:a|al|hasta)\s+(?:el\s+mes\s+de\s+)?|desde enero hasta\s+|hasta\s+)([a-z]+)", t)
    if m and m.group(1) in MONTHS:
        month = MONTHS[m.group(1)]
    if month is None:
        for pat, mo in PERIOD_WORDS:
            if re.search(pat, t):
                month = mo
                break
    if month is None:
        m = re.search(r"(?:durante el|en el|del)\s+(?:año\s+)?(20\d\d)\b", html.unescape(title).lower())
        if m:
            return int(m.group(1)), 12
    if month is None:
        m = re.search(r"mes de ([a-z]+)", t)
        if m and m.group(1) in MONTHS:
            month = MONTHS[m.group(1)]
    if month is None:
        prev = (posted.replace(day=1) - pd.Timedelta(days=1))
        return prev.year, prev.month
    year = posted.year if month < posted.month else posted.year - 1
    return year, month


def parse_itaipu(post):
    title = html.unescape(re.sub(r"<[^>]+>", " ", post["title"]["rendered"]))
    text = html.unescape(re.sub(r"<[^>]+>", " ", post["content"]["rendered"]))
    text = re.sub(r"\s+", " ", text)
    if not re.search(r"(?i)(suministr|inyect)\w*[^.]*?[\d.]+\s*GWh", title):   # a monthly report's headline
        return None
    posted = pd.Timestamp(post["date"][:10])
    year, month = itaipu_period(title, text[:400], posted)
    ytd = re.search(r"([\d.]+(?:,\d+)?)\s*GWh", title)
    gen = re.search(r"generaci[óo]n\s+(?:total\s+)?del?\s+[^.]{0,40}?mes[^.]{0,30}?fue de\s+([\d.]+(?:,\d+)?)\s*GWh",
                    text, re.I)
    g50 = re.search(r"([\d.]+(?:,\d+)?)\s*GWh fueron generados por el sistema de 50\s*Hz", text, re.I)
    ande = re.search(r"de los cuales\s+([\d.]+(?:,\d+)?)\s*GWh fueron suministrados a la ANDE", text, re.I)
    gen_ytd = re.search(r"cantidad total (?:de energ[íi]a )?generada[^.]{0,80}?fue de\s+([\d.]+(?:,\d+)?)\s*GWh", text, re.I)
    return {"date": pd.Timestamp(year, month, 1), "ANDE_YTD_GWh": num(ytd.group(1)) if ytd else None,
            "ANDE_month_GWh_reported": num(ande.group(1)) if ande else None,
            "Generation_GWh_reported": num(gen.group(1)) if gen else None,
            "Generation_50Hz_GWh": num(g50.group(1)) if g50 else None,
            "Generation_YTD_GWh": num(gen_ytd.group(1)) if gen_ytd else None,
            "posted": posted, "link": post["link"]}


def fetch_itaipu(s, after=None):
    rows, seen = [], set()
    for q in ITAIPU_QUERIES:
        for page in range(1, 20):
            params = {"search": q, "per_page": 100, "page": page, "_fields": "date,link,title,content"}
            if after is not None:
                params["after"] = after.strftime("%Y-%m-%dT00:00:00")
            r = s.get(ITAIPU_WP, params=params, timeout=60)
            if r.status_code == 400:   # past the last page
                break
            r.raise_for_status()
            items = r.json()
            for it in items:
                if it["link"] in seen:
                    continue
                seen.add(it["link"])
                p = parse_itaipu(it)
                if p and p["date"] >= START - pd.DateOffset(months=1) and p["ANDE_YTD_GWh"]:
                    rows.append(p)
            if len(items) < 100:
                break
    if not rows:
        return pd.DataFrame()
    df = pd.DataFrame(rows).sort_values("posted").drop_duplicates("date", keep="last").set_index("date").sort_index()
    print(f"  Itaipu reports: {len(df)} months parsed ({df.index.min():%Y-%m} to {df.index.max():%Y-%m})", flush=True)
    return df


def itaipu_monthly(rep):
    """Monthly ANDE take and total generation (GWh) from the reports: the month's stated values, else the
    difference of consecutive year-to-date totals."""
    rep = rep.sort_index().copy()
    ande, gen = {}, {}
    for d, r in rep.iterrows():
        prev = d - pd.DateOffset(months=1)
        ytd_diff = (r["ANDE_YTD_GWh"] - rep.loc[prev, "ANDE_YTD_GWh"]) if d.month > 1 and prev in rep.index else (
            r["ANDE_YTD_GWh"] if d.month == 1 else None)
        ande[d] = r["ANDE_month_GWh_reported"] if pd.notna(r["ANDE_month_GWh_reported"]) else ytd_diff
        g = r["Generation_GWh_reported"]
        if pd.isna(g) and pd.notna(r.get("Generation_YTD_GWh")):
            if d.month == 1:
                g = r["Generation_YTD_GWh"]
            elif prev in rep.index and pd.notna(rep.loc[prev, "Generation_YTD_GWh"]):
                g = r["Generation_YTD_GWh"] - rep.loc[prev, "Generation_YTD_GWh"]
        gen[d] = g
    return pd.DataFrame({"Itaipu_to_ANDE_GWh": pd.Series(ande), "Itaipu_generation_GWh": pd.Series(gen)})


# ------------------------------------------------------------------ 2. ONS: Itaipu delivered to Brazil (daily)
def ons_itaipu_rows(s, url, cache):
    """Itaipu's hourly rows of one ONS parquet (cached per URL: 2021 exists only as one yearly file)."""
    if url not in cache:
        r = s.get(url, timeout=900)
        if r.status_code != 200:
            cache[url] = None
        else:
            d = pd.read_parquet(io.BytesIO(r.content), columns=["din_instante", "nom_usina", "val_geracao"])
            cache[url] = d[d["nom_usina"].astype(str).str.upper().str.contains("ITAIPU")].copy()
    return cache[url]


def ons_month(s, y, m, cache):
    """Daily MWh of Itaipu 60 Hz and 50 Hz (delivered to Brazil) for one month, or None if ONS hasn't got it."""
    for url in (f"{ONS_BASE}GERACAO_USINA-2_{y}_{m:02d}.parquet", f"{ONS_BASE}GERACAO_USINA-2_{y}.parquet"):
        d = ons_itaipu_rows(s, url, cache)
        if d is None:
            continue
        d = d.copy()
        d["t"] = pd.to_datetime(d["din_instante"])
        d = d[(d["t"].dt.year == y) & (d["t"].dt.month == m)]
        if d.empty:
            continue
        d["mw"] = pd.to_numeric(d["val_geracao"], errors="coerce")
        d["hz"] = d["nom_usina"].str.upper().str.extract(r"(50|60)\s*HZ")[0]
        h = d.pivot_table(index="t", columns="hz", values="mw", aggfunc="sum")
        n = h.groupby(h.index.normalize()).size()
        day = h.groupby(h.index.normalize()).mean() * 24   # hourly MWmed: mean MW x 24 h = MWh
        day = day[n.reindex(day.index) >= 23]
        out = pd.DataFrame({"Itaipu_60Hz_MWh": day.get("60"), "Itaipu_50Hz_to_Brazil_MWh": day.get("50")})
        out["Itaipu_to_Brazil_ONS_MWh"] = out.sum(axis=1, min_count=1)
        out.index.name = "date"
        return out.round(1)
    return None


def update_ons(s, old, today):
    have = set() if old.empty else {(d.year, d.month) for d in old.index}
    months = pd.period_range(START, today.to_period("M"), freq="M")
    todo = [p for p in months if (p.year, p.month) not in have or p >= months[-1] - 2]
    cache = {}
    new = []
    for p in todo:
        try:
            out = ons_month(s, p.year, p.month, cache)
        except Exception as e:  # noqa: BLE001
            print(f"  ONS {p}: FAILED ({type(e).__name__}: {e})", flush=True)
            continue
        if out is not None and not out.empty:
            new.append(out)
            print(f"  ONS {p}: {len(out)} days, {out['Itaipu_to_Brazil_ONS_MWh'].sum() / 1000:,.0f} GWh", flush=True)
    return merge(pd.concat(new) if new else pd.DataFrame(), old)


# ------------------------------------------------------------------ 3. EBY monthly posts (SADI / SINP)
def eby_period(title, url, posted):
    t = html.unescape(f"{title} {url}").lower().replace("-", " ")
    month = next((MONTHS[w] for w in re.findall(r"[a-z]+", t) if w in MONTHS), None)
    y = re.search(r"\b(20\d\d)\b", t)
    if month is None:
        prev = posted.replace(day=1) - pd.Timedelta(days=1)
        return prev.year, prev.month
    year = int(y.group(1)) if y else (posted.year if month < posted.month else posted.year - 1)
    return year, month


def fetch_eby(s, known_links):
    links = {}
    for q in EBY_QUERIES:
        for page in range(1, 8):
            url = EBY_SEARCH.format(page="" if page == 1 else f"page/{page}/", q=q)
            r = s.get(url, timeout=60)
            if r.status_code != 200:
                break
            found = re.findall(r'<time class="entry-date published[^"]*" datetime="([^"]+)".{0,800}?entry-title"><a href="'
                               r'(https://www.eby.gov.py/[^"]+)" rel="bookmark">([^<]+)', r.text, re.S)
            if not found:
                break
            for when, link, title in found:
                if re.search(r"(?i)generaci", title) and "cota" not in link:
                    links[link] = (pd.Timestamp(when[:10]), html.unescape(title))
    rows = []
    for link, (posted, title) in links.items():
        if link in known_links:
            continue
        r = s.get(link, timeout=60)
        if r.status_code != 200:
            continue
        body = r.text.split("Últimas Publicaciones")[0]
        txt = re.sub(r"\s+", " ", html.unescape(re.sub(r"<[^>]+>", " ", body)))
        sadi = re.search(r"(?:SADI|Sistema Argentino de Interconexi[óo]n)[^.]{0,120}?fue de\s+([\d.]+(?:,\d+)?)\s*MWh", txt)
        sinp = re.search(r"(?:SINP|Sistema Interconectado Nacional Paraguayo)[^.]{0,120}?fue de\s+([\d.]+(?:,\d+)?)\s*MWh", txt)
        if not (sadi and sinp):
            print(f"  EBY {link}: no SADI/SINP figures", flush=True)
            continue
        y, m = eby_period(title, link, posted)
        rows.append({"date": pd.Timestamp(y, m, 1), "SADI_MWh": num(sadi.group(1)), "SINP_MWh": num(sinp.group(1)),
                     "posted": posted, "link": link})
    if not rows:
        return pd.DataFrame()
    df = pd.DataFrame(rows).sort_values("posted").drop_duplicates("date", keep="last").set_index("date").sort_index()
    print(f"  EBY reports: {len(df)} new months ({', '.join(f'{d:%Y-%m}' for d in df.index)})", flush=True)
    return df


# ------------------------------------------------------------------ 4. CAMMESA: Yacyreta to the SADI (daily)
def chain_bundle():
    """certifi + the Sectigo OV R36 intermediate that www.eby.org.ar does not send. Verification stays on."""
    import certifi
    der = requests.get(SECTIGO_OV_R36, timeout=30).content
    path = os.path.join(tempfile.gettempdir(), "certifi_plus_sectigo_r36.pem")
    with open(certifi.where()) as f, open(path, "w") as g:
        g.write(f.read() + "\n" + ssl.DER_cert_to_PEM_cert(der))
    return path


def mdb_rows(path, table):
    out = subprocess.run(["mdb-export", path, table], capture_output=True, text=True, timeout=300)
    if out.returncode != 0:
        raise RuntimeError(f"mdb-export {table}: {out.stderr[:200]}")
    return list(csv.DictReader(io.StringIO(out.stdout)))


def cammesa_month_docs(s, A, first):
    nxt = (pd.Timestamp(first) + pd.offsets.MonthBegin(1)).date()
    r = s.get(A.LOOKUP_URL, params={"fechadesde": first.strftime(A.TIME_FMT), "fechahasta": nxt.strftime(A.TIME_FMT),
                                    "nemo": "PARTE_POST_OPERATIVO"}, timeout=60)
    r.raise_for_status()
    out = {}
    for doc in r.json() if isinstance(r.json(), list) else []:
        for att in doc.get("adjuntos", []):
            aid = att.get("id", "")
            if aid.startswith("PO") and aid.endswith(".zip"):
                out[aid] = (doc, att)
    return out


def cammesa_day(s, A, hit, tries=3):
    doc, att = hit
    for attempt in range(tries):
        try:
            r = s.get(A.ATTACHMENT_URL, params={"attachmentId": att["id"], "docId": doc["id"],
                                                "nemo": doc.get("nemo") or "PARTE_POST_OPERATIVO"}, timeout=180)
            r.raise_for_status()
            zf = zipfile.ZipFile(io.BytesIO(r.content))
            name = next(n for n in zf.namelist() if n.lower().endswith(".mdb"))
            with tempfile.NamedTemporaryFile(suffix=".mdb", delete=False) as f:
                f.write(zf.read(name))
                path = f.name
            try:
                gens = {g["GRUPO"]: g for g in mdb_rows(path, "GENERADORES")}
                out, hours = defaultdict(float), set()
                for v in mdb_rows(path, "VALORES_GENERADORES"):
                    hours.add(v["HORA"])
                    g = v["GRUPO"]
                    meta = gens.get(g, {})
                    try:
                        e = float(v["ENERGIA"] or 0)
                    except ValueError:
                        continue
                    if g == "YACYHI":
                        out["YACYHI_MWh"] += e
                    elif g.startswith("YACY"):
                        out["YACYHIPY_MWh"] += e     # Paraguay's Yacyreta energy delivered to the SADI (from 2025)
                    elif str(meta.get("AGENTE", "")).startswith("ANDE"):
                        key = "ANDE_nodes_counted_MWh" if meta.get("INTERCAMBIO") != "S" else "ANDE_nodes_import_MWh"
                        out[key] += e
            finally:
                os.unlink(path)
            if len(hours) < 23:
                raise RuntimeError(f"only {len(hours)} hours")
            return dict(out)
        except Exception:  # noqa: BLE001
            if attempt == tries - 1:
                raise
            time.sleep(5 * (attempt + 1))


def update_cammesa(old, today, budget_min, workers):
    import argentina_generation_mix as A
    s = A.make_session()
    t0 = time.time()
    have = set() if old.empty else {d.date() for d in old.index[old["YACYHI_MWh"].notna()]}
    end = (today - pd.Timedelta(days=1)).date()
    days = [d.date() for d in pd.date_range(START, end)]
    last = max(have) if have else None
    todo = sorted([d for d in days if d not in have or (last and (last - d).days < 3)], reverse=True)
    print(f"  CAMMESA: {len(have):,} days saved, {len(todo):,} to fetch (budget {budget_min:.0f} min)", flush=True)
    listings, rows = {}, {}
    with ThreadPoolExecutor(max_workers=workers) as pool:
        for i in range(0, len(todo), 31):
            if time.time() - t0 > budget_min * 60:
                print(f"  CAMMESA: time budget reached; {len(todo) - i} days left for later runs", flush=True)
                break
            jobs = {}
            for day in todo[i:i + 31]:
                first = day.replace(day=1)
                if first not in listings:
                    try:
                        listings[first] = cammesa_month_docs(s, A, first)
                    except Exception as e:  # noqa: BLE001
                        print(f"  CAMMESA listing {first:%Y-%m} failed: {e}", flush=True)
                        listings[first] = {}
                hit = listings[first].get(day.strftime("PO%y%m%d.zip"))
                if hit:
                    jobs[pool.submit(cammesa_day, s, A, hit)] = day
            for fut in as_completed(jobs):
                day = jobs[fut]
                try:
                    rows[pd.Timestamp(day)] = fut.result()
                except Exception as e:  # noqa: BLE001
                    print(f"  CAMMESA {day}: FAILED ({type(e).__name__}: {str(e)[:120]}) - retried next run", flush=True)
            print(f"  CAMMESA: {len(rows)} days fetched ({time.time() - t0:.0f}s)", flush=True)
    if not rows:
        return old
    new = pd.DataFrame.from_dict(rows, orient="index").reindex(
        columns=["YACYHI_MWh", "YACYHIPY_MWh", "ANDE_nodes_counted_MWh", "ANDE_nodes_import_MWh"]).fillna(0.0)
    new.index.name = "date"
    return merge(new.round(1), old)


# ------------------------------------------------------------------ 5. eby.org.ar latest-month total
def fetch_ebyar(s, today):
    try:
        r = s.get(EBYAR_HOME, timeout=60, verify=chain_bundle())
        r.raise_for_status()
    except Exception as e:  # noqa: BLE001
        print(f"  eby.org.ar: FAILED ({type(e).__name__}: {e})", flush=True)
        return pd.DataFrame()
    m = re.search(r'table-cell">\s*([A-Za-zé]+)\s*</td>\s*</tr>\s*<tr class="row">\s*<td class="cell" '
                  r'data-auto="table-cell">\s*([\d.]+)\s*MWh', r.text)
    if not m or html.unescape(m.group(1)).lower() not in MONTHS:
        print("  eby.org.ar: monthly generation widget not found", flush=True)
        return pd.DataFrame()
    mo = MONTHS[html.unescape(m.group(1)).lower()]
    year = today.year if mo < today.month else today.year - 1
    print(f"  eby.org.ar: {m.group(1)} {year}: {m.group(2)} MWh (net)", flush=True)
    return pd.DataFrame({"Yacyreta_net_MWh": [num(m.group(2))], "read_on": [today.normalize()]},
                        index=pd.DatetimeIndex([pd.Timestamp(year, mo, 1)], name="date"))


# ------------------------------------------------------------------ 6. VMME Balance Energetico Nacional (annual)
def fetch_ben(s, old):
    import pdfplumber
    rows = {}
    for edition, url in sorted(BEN_URLS.items()):
        if not old.empty and edition in old.index.year and (edition - 1) in old.index.year:
            continue
        try:
            r = s.get(url, timeout=120)
            if r.status_code != 200 or not r.content.startswith(b"%PDF"):
                continue
            with pdfplumber.open(io.BytesIO(r.content)) as pdf:
                txt = "\n".join((p.extract_text() or "") for p in pdf.pages[15:30])
        except Exception as e:  # noqa: BLE001
            print(f"  BEN {edition}: FAILED ({e})", flush=True)
            continue
        arg = re.search(r"\nArgentina\s+([\d.]+,\d)\s+([\d.]+,\d)", txt)
        bra = re.search(r"Brasil \(Excedente ITAIPU\)\s+([\d.]+,\d)\s+([\d.]+,\d)", txt)
        tot = re.search(r"Exportaci[óo]n Total\s+([\d.]+,\d)\s+([\d.]+,\d)", txt)
        ande = re.search(r"ANDE \(Energ[íi]a facturada para exportaci[óo]n\)\s+([\d.]+,\d)\s+([\d.]+,\d)", txt)
        gen = re.search(r"generaci[óo]n bruta de energ[íi]a el[ée]ctrica \(([\d.]+,\d) GWh", txt)
        if not (arg and bra):
            print(f"  BEN {edition}: export table not found", flush=True)
            continue
        for k, y in ((1, edition - 1), (2, edition)):
            rows[pd.Timestamp(y, 1, 1)] = {"Export_to_Argentina_GWh": num(arg.group(k)),
                                           "Export_to_Brazil_Itaipu_GWh": num(bra.group(k)),
                                           "Export_total_GWh": num(tot.group(k)) if tot else None,
                                           "ANDE_sales_export_GWh": num(ande.group(k)) if ande else 0.0,
                                           "Gross_generation_GWh": num(gen.group(1)) if (gen and y == edition) else None,
                                           "edition": edition}
        print(f"  BEN {edition}: exports to Argentina {arg.group(1)} / {arg.group(2)} GWh", flush=True)
    if not rows:
        return old
    new = pd.DataFrame.from_dict(rows, orient="index")
    new.index.name = "date"
    if not old.empty:   # keep a gross-generation figure an older edition had
        for d in new.index:
            if d in old.index and pd.isna(new.loc[d, "Gross_generation_GWh"]):
                new.loc[d, "Gross_generation_GWh"] = old.loc[d, "Gross_generation_GWh"]
    return merge(new, old)


# ------------------------------------------------------------------ build
def monthly_sum(daily, cols, min_days=None):
    """Month sums of complete months only (every day of the month present)."""
    if daily.empty:
        return pd.DataFrame(columns=cols)
    d = daily[cols]
    g = d.resample("MS")
    n = g.size()
    out = g.sum(min_count=1)
    full = n == pd.Series(out.index.days_in_month, index=out.index)
    return out[full]


def build(rep, ons, eby, ebyar, cam, ben):
    it = itaipu_monthly(rep) * 1000.0   # GWh -> MWh
    it.columns = ["Itaipu_to_ANDE_MWh", "Itaipu_generation_MWh"]
    enb = monthly_sum(ons, ["Itaipu_to_Brazil_ONS_MWh"])
    cm = monthly_sum(cam, ["YACYHI_MWh", "YACYHIPY_MWh", "ANDE_nodes_counted_MWh"]) if not cam.empty else pd.DataFrame()
    idx = pd.date_range(START, max([x.index.max() for x in (it, enb) if not x.empty]), freq="MS")
    p = pd.DataFrame(index=idx)
    p.index.name = "date"
    p["Itaipu_to_ANDE_MWh"] = it["Itaipu_to_ANDE_MWh"]
    p["Itaipu_to_Brazil_MWh"] = enb["Itaipu_to_Brazil_ONS_MWh"]
    p["Itaipu_generation_MWh"] = it["Itaipu_generation_MWh"]
    p["Itaipu_generation_source"] = "Itaipu monthly report"
    miss = p["Itaipu_generation_MWh"].isna() & p["Itaipu_to_ANDE_MWh"].notna() & p["Itaipu_to_Brazil_MWh"].notna()
    p.loc[miss, "Itaipu_generation_MWh"] = p.loc[miss, ["Itaipu_to_ANDE_MWh", "Itaipu_to_Brazil_MWh"]].sum(axis=1)
    p.loc[miss, "Itaipu_generation_source"] = "supply: ANDE (report) + Brazil (ONS); excludes ~0.6% own use"
    p.loc[p["Itaipu_generation_MWh"].isna(), "Itaipu_generation_source"] = None

    # Yacyreta: SADI from EBY, else CAMMESA; total from EBY (SADI + SINP), else eby.org.ar
    sadi_cam = (cm["YACYHI_MWh"] + cm["YACYHIPY_MWh"]) if not cm.empty else pd.Series(dtype=float)
    p["Yacyreta_to_SADI_MWh"] = eby["SADI_MWh"] if not eby.empty else None
    p["Yacyreta_to_SADI_MWh"] = p["Yacyreta_to_SADI_MWh"].astype(float).fillna(sadi_cam)
    p["Yacyreta_to_SINP_MWh"] = eby["SINP_MWh"] if not eby.empty else None
    p["Yacyreta_to_SINP_MWh"] = p["Yacyreta_to_SINP_MWh"].astype(float)
    p["Yacyreta_SINP_source"] = None
    p.loc[p["Yacyreta_to_SINP_MWh"].notna(), "Yacyreta_SINP_source"] = "EBY monthly report"
    if not ebyar.empty:
        tot = ebyar["Yacyreta_net_MWh"].reindex(p.index)
        ok = p["Yacyreta_to_SINP_MWh"].isna() & tot.notna() & p["Yacyreta_to_SADI_MWh"].notna()
        p.loc[ok, "Yacyreta_to_SINP_MWh"] = tot[ok] - p.loc[ok, "Yacyreta_to_SADI_MWh"]
        p.loc[ok, "Yacyreta_SINP_source"] = "eby.org.ar monthly total - SADI (CAMMESA)"
    estimate_sinp(p, ben)
    p["Yacyreta_total_MWh"] = p["Yacyreta_to_SADI_MWh"] + p["Yacyreta_to_SINP_MWh"]
    p["ANDE_sales_counted_by_CAMMESA_MWh"] = cm["ANDE_nodes_counted_MWh"] if not cm.empty else 0.0
    return p


def estimate_sinp(p, ben):
    """Months with no published SINP figure: ANDE's Yacyreta take is estimated in proportion to its Itaipu take
    that month (both follow Paraguay's demand). The year's ratio comes from (a) VMME's annual export to
    Argentina where the BEN covers the year and every month's SADI is known (annual SINP = SADI - 2 x export,
    less the months already known), else (b) the months with EBY figures that year, else (c) the nearest year."""
    est = p["Yacyreta_to_SINP_MWh"].isna() & p["Yacyreta_to_SADI_MWh"].notna() & p["Itaipu_to_ANDE_MWh"].notna()
    if not est.any():
        return
    ratios, how = {}, {}
    for y in sorted(set(p.index.year)):
        py = p[p.index.year == y]
        known = py["Yacyreta_to_SINP_MWh"].notna()
        b = ben[ben.index.year == y] if not ben.empty else pd.DataFrame()
        if len(b) and py["Yacyreta_to_SADI_MWh"].notna().sum() == 12 and len(py) == 12:
            # BEN's 'Argentina' export row includes ANDE's own market sales; the Yacyreta cession is the rest
            cession = float(b["Export_to_Argentina_GWh"].iloc[0]) - float(b.get("ANDE_sales_export_GWh", pd.Series([0])).fillna(0).iloc[0])
            annual = py["Yacyreta_to_SADI_MWh"].sum() - 2 * cession * 1000
            rest = annual - py.loc[known, "Yacyreta_to_SINP_MWh"].sum()
            base = py.loc[~known, "Itaipu_to_ANDE_MWh"].sum()
            if base > 0 and rest > 0:
                ratios[y], how[y] = rest / base, f"BEN {y} annual export to Argentina"
                continue
        if known.sum() >= 3:
            ratios[y] = py.loc[known, "Yacyreta_to_SINP_MWh"].sum() / py.loc[known, "Itaipu_to_ANDE_MWh"].sum()
            how[y] = f"{known.sum()} EBY-reported months of {y}"
    for d in p.index[est]:
        y = d.year
        if y not in ratios:
            near = min(ratios, key=lambda k: abs(k - y)) if ratios else None
            if near is None:
                continue
            ratios[y], how[y] = ratios[near], f"ratio of {near} ({how[near]})"
        p.loc[d, "Yacyreta_to_SINP_MWh"] = p.loc[d, "Itaipu_to_ANDE_MWh"] * ratios[y]
        p.loc[d, "Yacyreta_SINP_source"] = (f"ESTIMATE: ANDE's Itaipu take x {ratios[y]:.3f} (SINP/ANDE-Itaipu ratio "
                                            f"from {how[y]})")


def outputs(p):
    ok = p["Itaipu_generation_MWh"].notna() & p["Yacyreta_total_MWh"].notna()
    q = p[ok]
    daily = pd.DataFrame(index=q.index)
    daily["Hydro_MWh"] = 0.5 * q["Itaipu_generation_MWh"] + 0.5 * q["Yacyreta_total_MWh"]
    for f in ["Gas", "Wind", "Solar", "Coal", "Nuclear", "Oil", "Bioenergy", "Other"]:
        daily[f"{f}_MWh"] = 0.0
    daily["Total_MWh"] = daily["Hydro_MWh"]
    daily = daily.round(1)
    daily.index.name = "date"

    itaipu_supply = q["Itaipu_to_ANDE_MWh"] + q["Itaipu_to_Brazil_MWh"]
    ex = pd.DataFrame(index=q.index)
    ex["Itaipu_to_Brazil_GWh"] = (0.5 * itaipu_supply - q["Itaipu_to_ANDE_MWh"]) / 1000
    ex["Yacyreta_to_Argentina_GWh"] = (0.5 * q["Yacyreta_total_MWh"] - q["Yacyreta_to_SINP_MWh"]) / 1000
    ex["Paraguay_consumption_GWh"] = (q["Itaipu_to_ANDE_MWh"] + q["Yacyreta_to_SINP_MWh"]) / 1000
    ex["Total_generation_share_GWh"] = daily["Total_MWh"] / 1000
    ex["Yacyreta_SINP_estimated"] = q["Yacyreta_SINP_source"].astype(str).str.startswith("ESTIMATE")
    ex = ex.round(1)
    ex.index.name = "date"

    # What ONS (Brazil) and CAMMESA (Argentina) do NOT already count: ANDE's own take from both plants, less
    # ANDE's market sales CAMMESA books as Argentine generation (INTERCAMBIO='N' ANDE nodes)
    sa = pd.DataFrame(index=q.index)
    sa["Hydro_MWh"] = (q["Itaipu_to_ANDE_MWh"] + q["Yacyreta_to_SINP_MWh"]
                       - q["ANDE_sales_counted_by_CAMMESA_MWh"].fillna(0)).round(1)
    sa["Total_MWh"] = sa["Hydro_MWh"]
    sa.index.name = "date"
    return daily, ex, sa


def validation(daily, ex, ben):
    """Annual totals vs Ember (yearly, Paraguay) and VMME's BEN."""
    yr = daily["Total_MWh"].groupby(daily.index.year).agg(["sum", "count"])
    exy = ex.groupby(ex.index.year)[["Itaipu_to_Brazil_GWh", "Yacyreta_to_Argentina_GWh", "Paraguay_consumption_GWh"]].sum()
    v = pd.DataFrame({"months": yr["count"], "This_pull_share_TWh": (yr["sum"] / 1e6).round(2),
                      "Paraguay_consumption_TWh": (exy["Paraguay_consumption_GWh"] / 1000).round(2),
                      "Ceded_to_Brazil_TWh": (exy["Itaipu_to_Brazil_GWh"] / 1000).round(2),
                      "Ceded_to_Argentina_TWh": (exy["Yacyreta_to_Argentina_GWh"] / 1000).round(2)})
    try:
        e = pd.read_csv("https://storage.googleapis.com/emb-prod-bkt-publicdata/public-downloads/"
                        "yearly_full_release_long_format.csv", low_memory=False)
        e = e[(e["Area"] == "Paraguay") & (e["Unit"] == "TWh")]
        g = e[(e["Category"] == "Electricity generation") & (e["Variable"] == "Total Generation")].set_index("Year")["Value"]
        h = e[(e["Category"] == "Electricity generation") & (e["Variable"] == "Hydro")].set_index("Year")["Value"]
        dm = e[(e["Category"] == "Electricity demand") & (e["Variable"] == "Demand")].set_index("Year")["Value"]
        ni = e[(e["Category"] == "Electricity imports") & (e["Variable"] == "Net Imports")].set_index("Year")["Value"]
        v["Ember_generation_TWh"], v["Ember_hydro_TWh"] = g, h
        v["Ember_demand_TWh"], v["Ember_net_imports_TWh"] = dm, ni
        v["Share_vs_Ember_pct"] = (100 * (v["This_pull_share_TWh"] / v["Ember_generation_TWh"] - 1)).round(1)
    except Exception as ex_:  # noqa: BLE001
        print(f"  Ember check skipped: {ex_}", flush=True)
    if not ben.empty:
        b = ben.copy()
        b.index = b.index.year
        v["BEN_gross_generation_TWh"] = (b["Gross_generation_GWh"] / 1000).round(2)
        v["BEN_export_Brazil_TWh"] = (b["Export_to_Brazil_Itaipu_GWh"] / 1000).round(2)
        v["BEN_export_Argentina_TWh"] = (b["Export_to_Argentina_GWh"] / 1000).round(2)
    v.index.name = "year"
    return v


NOTES = [
    "UNITS",
    "Sheet 'Daily': ONE ROW PER MONTH, dated the 1st, MWh generated in that month (no source publishes Paraguay's "
    "daily split, so the standard daily layout carries monthly values). Hydro_MWh = Paraguay's 50% of Itaipu's "
    "generation + Paraguay's 50% of Yacyreta's output. Other fuels are 0: ANDE's Acaray (210 MW hydro) and its small "
    "thermal/solar units are not included (no reachable source: ande.gov.py is behind a bot captcha) - about 1-2% "
    "of Paraguay's share. Total_MWh = Hydro_MWh.",
    "Sheet 'Full plants': whole-plant monthly figures (MWh) - Itaipu generation (Itaipu's monthly report), energy "
    "supplied to ANDE (Paraguay) and to Brazil (ONS 'ITAIPU 60 HZ' + 'ITAIPU 50 HZ'), Yacyreta's supply to the SADI "
    "(Argentina) and SINP (Paraguay) and their sum, with the source of each SINP figure.",
    "Sheet 'Exports' (GWh per month): Itaipu_to_Brazil_GWh = 50% of Itaipu's supply - ANDE's take (Paraguay's "
    "unused half, ceded to Brazil); Yacyreta_to_Argentina_GWh = 50% of Yacyreta's supply - SINP; "
    "Paraguay_consumption_GWh = ANDE's Itaipu take + Yacyreta's SINP supply (Paraguay's use of its binational "
    "plants); Total_generation_share_GWh = sheet Daily's total. Share ~= consumption + exports (Itaipu's ~0.6% own "
    "use and losses aside). Yacyreta_SINP_estimated flags months whose SINP figure is an estimate.",
    "Sheet 'Not counted by ONS-CAMMESA' (standard layout, MWh per month): the part of Paraguay's generation that "
    "neither Brazil's ONS nor Argentina's CAMMESA series already counts - ANDE's Itaipu take + Yacyreta's SINP "
    "supply, less ANDE's market sales that CAMMESA books as Argentine generation. The South America total uses "
    "this, not the 50% share, so nothing is counted twice.",
    "",
    "DOUBLE COUNTING (why the South America total uses the 'Not counted' sheet)",
    "ONS counts Itaipu's 60 Hz output and the 50 Hz output sent to Brazil ('ITAIPU 50 HZ', id PYIT50): together "
    "they equal Itaipu's supply to ENBPar (2025: 46.68 TWh ONS vs 46.68 TWh Itaipu report), i.e. all of Itaipu except "
    "ANDE's take - so Brazil's series already holds Paraguay's ceded Itaipu energy. CAMMESA counts Yacyreta as "
    "YACYHI (+ YACYHIPY from 2025, 'Paraguayan share delivered to Argentina'), equal to Yacyreta's supply to the SADI "
    "(Nov-2023: CAMMESA ~1,546 GWh vs EBY 1,549.6 GWh) - all of Yacyreta except the SINP. Paraguay's own 50% share "
    "(sheet Daily) therefore overlaps both; only ANDE's take is Paraguay-only.",
    "",
    "ESTIMATES",
    "EBY published SADI/SINP monthly only to Nov-2023 (and not every month). For other months the SINP supply is "
    "estimated as ANDE's Itaipu take that month x a yearly ratio (SINP / ANDE's Itaipu take) from VMME's annual "
    "export to Argentina (annual SINP = SADI - 2 x export), else from that year's EBY months, else the nearest "
    "year; such months are flagged. From the run that first reads it, eby.org.ar's latest-month total replaces the "
    "estimate (SINP = total - CAMMESA SADI).",
    "",
    "SOURCES",
    "Itaipu monthly reports: https://www.itaipu.gov.py/noticias/energia/ (WordPress REST API); "
    "ONS: https://dados.ons.org.br/dataset/geracao-usina-2; EBY: https://www.eby.gov.py/ and https://www.eby.org.ar/; "
    "CAMMESA: https://cammesaweb.cammesa.com/ (PARTE_POST_OPERATIVO); VMME Balance Energetico Nacional: "
    "https://minasyenergia.mopc.gov.py/. Validation against Ember yearly data: sheet 'Validation'.",
]


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", default=OUT)
    ap.add_argument("--budget-min", type=float, default=40, help="CAMMESA fetching budget, minutes")
    ap.add_argument("--workers", type=int, default=4)
    ap.add_argument("--skip-cammesa", action="store_true")
    args = ap.parse_args()
    today = pd.Timestamp.now(tz="America/Asuncion").tz_localize(None).normalize()
    s = requests.Session()
    s.headers.update(UA)

    rep_old, ons_old = load(args.out, S_ITAIPU), load(args.out, S_ONS)
    eby_old, ebyar_old = load(args.out, S_EBY), load(args.out, S_EBYAR)
    cam_old, ben_old = load(args.out, S_CAM), load(args.out, S_BEN)

    print("Itaipu monthly reports ...", flush=True)
    after = None
    if not rep_old.empty:
        full = pd.date_range(START, rep_old.index.max(), freq="MS")
        if full.isin(rep_old.index).all():
            after = rep_old.index.max() - pd.DateOffset(days=10)
    try:
        rep = merge(fetch_itaipu(s, after), rep_old)
    except Exception as e:  # noqa: BLE001
        print(f"  Itaipu reports FAILED ({type(e).__name__}: {e}) - keeping saved months", flush=True)
        rep = rep_old
    print("ONS Itaipu ...", flush=True)
    ons = update_ons(s, ons_old, today)
    print("EBY monthly posts ...", flush=True)
    try:
        eby = merge(fetch_eby(s, set(eby_old["link"]) if not eby_old.empty else set()), eby_old)
    except Exception as e:  # noqa: BLE001
        print(f"  EBY FAILED ({type(e).__name__}: {e})", flush=True)
        eby = eby_old
    ebyar = merge(fetch_ebyar(s, today), ebyar_old)
    print("BEN ...", flush=True)
    ben = fetch_ben(s, ben_old)
    if not args.skip_cammesa:
        print("CAMMESA Yacyreta ...", flush=True)
        cam = update_cammesa(cam_old, today, args.budget_min, args.workers)
    else:
        cam = cam_old
    if rep.empty or ons.empty:
        print("Itaipu reports or ONS missing - nothing written.", flush=True)
        sys.exit(1)

    p = build(rep, ons, eby, ebyar, cam, ben)
    daily, ex, sa = outputs(p)
    val = validation(daily, ex, ben)
    sheets = {"Daily": daily, S_PLANTS: p.round(1), S_EXPORTS: ex, S_SA: sa, S_VALID: val,
              S_ITAIPU: rep, S_ONS: ons, S_EBY: eby, S_EBYAR: ebyar, S_CAM: cam, S_BEN: ben}
    sheets = {k: v for k, v in sheets.items() if v is not None and not v.empty}
    os.makedirs(os.path.dirname(os.path.abspath(args.out)), exist_ok=True)
    xlsx_notes.write_workbook(args.out, sheets, NOTES, {ln for ln in NOTES if ln and ln.isupper()})
    print(f"Saved {args.out}: {len(daily)} months ({daily.index.min():%b-%Y} to {daily.index.max():%b-%Y})", flush=True)
    print(ex.tail(6).to_string(), flush=True)
    print(val.to_string(), flush=True)
    est = ex["Yacyreta_SINP_estimated"].sum()
    print(f"Yacyreta SINP estimated in {est} of {len(ex)} months", flush=True)


if __name__ == "__main__":
    main()
