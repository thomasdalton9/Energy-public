"""QA-fix probe (one-off): PAGASA dam table layout, PUCSL reservoir API aggregation options, CEA capacity Feb/Mar 2025
all-India rows."""
import html as htmllib
import io
import re

import pandas as pd
import requests

H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0 Safari/537.36"}


def sec(t):
    print(f"\n######## {t}", flush=True)


sec("PAGASA flood table")
try:
    r = requests.get("https://www.pagasa.dost.gov.ph/flood", headers=H, timeout=(20, 90))
    tab = next((t for t in re.findall(r"(?is)<table.*?</table>", r.text) if "dam name" in t.lower()), None)
    if tab:
        for tr in re.findall(r"(?is)<tr.*?</tr>", tab):
            ms = re.findall(r"(?is)<t[dh]([^>]*)>(.*?)</t[dh]>", tr)
            cells = [re.sub(r"\s+", " ", htmllib.unescape(re.sub(r"<[^>]+>", " ", c))).strip() for _, c in ms]
            attrs = [a.strip() for a, _ in ms]
            print(" | ".join(f"{c!r}{'{' + a + '}' if 'span' in a else ''}" for c, a in zip(cells, attrs)))
    else:
        print("no table", r.status_code, len(r.text))
except Exception as e:  # noqa: BLE001
    print("PAGASA failed", e)

sec("PUCSL reservoir API")
API = "https://gendata.pucsl.gov.lk/api/reservoir/storage-rainfall"
HP = dict(H, Accept="application/json", Referer="https://gendata.pucsl.gov.lk/")
for agg in ("day", "15min", "hour", "raw", "none", None):
    for d0, d1 in (("2014-12-03", "2014-12-06"), ("2015-11-06", "2015-11-10")):
        p = {"from": f"{d0}T00:00:00.000Z", "to": f"{d1}T00:00:00.000Z"}
        if agg:
            p["dateAggregation"] = agg
        try:
            r = requests.get(API, params=p, headers=HP, timeout=(15, 120))
            print(f"agg={agg} {d0}: HTTP {r.status_code} {len(r.content)} bytes")
            if r.ok:
                d = pd.DataFrame(r.json().get("data") or [])
                if len(d):
                    print("  cols", list(d.columns), "rows", len(d))
                    cols = [c for c in ("reportDate", "reservoirName", "storageInGwh", "rainfallInMm") if c in d]
                    print(d[cols][d["reservoirName"].astype(str).str.contains("Victoria|Randenigala", case=False)]
                          .to_string()[:3000])
            else:
                print("  ", r.text[:300])
        except Exception as e:  # noqa: BLE001
            print(f"agg={agg}: {e}")
# other endpoints that may carry per-reading data
for path in ("reservoir/storage-rainfall/raw", "reservoir/storage", "reservoir/levels", "reservoir"):
    try:
        r = requests.get(f"https://gendata.pucsl.gov.lk/api/{path}", params={"from": "2014-12-03T00:00:00.000Z",
                         "to": "2014-12-05T00:00:00.000Z"}, headers=HP, timeout=(15, 60))
        print(path, r.status_code, r.text[:400])
    except Exception as e:  # noqa: BLE001
        print(path, e)

sec("CEA capacity Feb / Mar 2025")
for ym, mon in (("2025-02", "FEB"), ("2025-03", "MAR")):
    u = f"https://npp.gov.in/public-reports/cea/monthly/installcap/2025/{mon}/capacity1-{ym}.xls"
    try:
        r = requests.get(u, headers=H, timeout=(20, 120))
        df = pd.read_excel(io.BytesIO(r.content), header=None)
        col = df.astype(str).apply(lambda c: c.str.strip().str.lower())
        head = [i for i in range(len(df)) if col.iloc[i].str.contains("^region$").any()]
        tot = [i for i in range(len(df)) if col.iloc[i].str.contains("^total of all").any()]
        print(u)
        for i in list(range(head[0], head[0] + 4)) + tot[:1]:
            print(i, [str(x)[:18] for x in df.iloc[i].tolist()])
    except Exception as e:  # noqa: BLE001
        print(u, e)
