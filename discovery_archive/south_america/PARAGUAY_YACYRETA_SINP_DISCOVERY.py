"""
Paraguay's share of Yacyreta (energy delivered to ANDE / the SINP) is official for Jan-2021..Nov-2023 (EBY monthly
posts) and from Dec-2024 (CAMMESA's YACYHIPY group); PARAGUAY_POWER.py estimates 16 months in between and leaves
Sep-2023 and Jan-2024 blank (CAMMESA days 3, 5, 6 Sep-2023 and 14, 15 Jan-2024 unreadable). This probe looks for
official figures for those months:
  1. EBY Paraguay WordPress REST API (posts search, all dates): every post whose text gives an ANDE / SINP figure,
     with the month it covers - more posts than the site search found?
  2. EBY media library: PDFs / spreadsheets named memoria, informe, generacion, estadistic.
  3. CAMMESA PARTE_POST_OPERATIVO for the five missing days: is a file listed, and what fails when reading it?

Usage: python3 PARAGUAY_YACYRETA_SINP_DISCOVERY.py
"""
import html
import os
import re
import sys
import time

import pandas as pd
import requests

HERE = os.path.dirname(os.path.abspath(__file__))
SA = os.path.join(HERE, "..", "..", "south_america")
sys.path.insert(0, SA)
sys.argv = [sys.argv[0]]
import PARAGUAY_POWER as P  # noqa: E402  (main() is guarded)

S = requests.Session()
S.headers.update(P.UA)
WP = "https://www.eby.gov.py/wp-json/wp/v2/"
WANT = [("2022-05", "est"), ("2023-03", "est"), ("2023-04", "est"), ("2023-05", "est"), ("2023-06", "est"),
        ("2023-07", "est"), ("2023-09", "blank"), ("2023-10", "est"), ("2023-12", "est"), ("2024-01", "blank")] + \
       [(f"2024-{m:02d}", "est") for m in range(2, 12)]


def wp(kind, **params):
    out, page = [], 1
    while True:
        r = S.get(WP + kind, params={**params, "per_page": 100, "page": page}, timeout=60)
        if r.status_code != 200:
            if page == 1:
                print(f"  {kind} {params}: HTTP {r.status_code} {r.text[:120]!r}", flush=True)
            break
        batch = r.json()
        out += batch
        if len(batch) < 100:
            break
        page += 1
    return out


print("== 1. EBY posts via the WordPress REST API", flush=True)
posts = {}
for q in ["generación", "generacion", "energía entregada", "ANDE", "SINP", "datos oficiales", "MWh", "Yacyretá generó"]:
    for p in wp("posts", search=q, after="2021-12-01T00:00:00", _fields="id,date,link,title,content"):
        posts[p["id"]] = p
    print(f"  search '{q}': {len(posts)} posts so far", flush=True)
found = {}
for p in sorted(posts.values(), key=lambda x: x["date"]):
    title = html.unescape(re.sub(r"<[^>]+>", "", p["title"]["rendered"]))
    txt = re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", html.unescape(p["content"]["rendered"])))
    sadi, sinp = P.parse_eby(txt)
    if "MWh" not in txt and "GWh" not in txt:
        continue
    y, m = P.eby_period(title, p["link"], pd.Timestamp(p["date"][:10]))
    key = f"{y}-{m:02d}"
    print(f"  {p['date'][:10]} [{key}] SADI={sadi} SINP={sinp} | {title[:90]} | {p['link']}", flush=True)
    if sinp is not None:
        found.setdefault(key, (sadi, sinp, p["link"]))
    elif re.search(r"ANDE|SINP|Paraguay", txt):
        i = max(txt.find("ANDE"), txt.find("SINP"))
        print(f"      text: ...{txt[max(0, i - 200):i + 300]}...", flush=True)
print("\n  months wanted -> official figure found?", flush=True)
for key, why in WANT:
    print(f"   {key} ({why}): {found.get(key, 'NOT FOUND')}", flush=True)

print("\n== 2. EBY media library", flush=True)
for q in ["memoria", "informe", "generaci", "estad", "energ"]:
    for m in wp("media", search=q, _fields="id,date,source_url,title,mime_type"):
        if re.search(r"pdf|sheet|excel|csv", m.get("mime_type", ""), re.I):
            print(f"  [{q}] {m['date'][:10]} {m['mime_type']} {m['source_url']}", flush=True)

print("\n== 3. CAMMESA days missing in the archive", flush=True)
import argentina_generation_mix as A  # noqa: E402
s = A.make_session()
for day in ["2023-09-03", "2023-09-05", "2023-09-06", "2024-01-14", "2024-01-15"]:
    d = pd.Timestamp(day).date()
    try:
        docs = P.cammesa_month_docs(s, A, d.replace(day=1))
    except Exception as e:  # noqa: BLE001
        print(f"  {day}: month listing FAILED {type(e).__name__}: {e}", flush=True)
        continue
    ids = sorted(docs)
    want = f"PO{d:%y%m%d}.zip"
    hits = [i for i in ids if i.upper().startswith(f"PO{d:%y%m%d}")]
    print(f"  {day}: {len(ids)} files listed for the month; this day: {hits or 'NONE'} "
          f"(neighbours {[i for i in ids if i[2:8] in (f'{(d - pd.Timedelta(days=1)):%y%m%d}', f'{(d + pd.Timedelta(days=1)):%y%m%d}')]})",
          flush=True)
    for h in hits:
        try:
            out = P.cammesa_day(s, A, docs[h], tries=2)
            print(f"     read OK: {out}", flush=True)
        except Exception as e:  # noqa: BLE001
            print(f"     read FAILED: {type(e).__name__}: {str(e)[:300]}", flush=True)
    time.sleep(1)
