"""Probe 2: Texas border crossings / LNG ports in move/poe2 (ports named in series-description), which Texas
consumption processes EIA publishes (full history), and Texas storage/production ranges."""
import os, requests, pandas as pd
KEY = os.environ["EIA_API_KEY"]
B = "https://api.eia.gov/v2/natural-gas/"


def data(route, extra, start="2000-01"):
    out, off = [], 0
    while True:
        p = {"api_key": KEY, "frequency": "monthly", "data[0]": "value", "start": start, "length": 5000, "offset": off,
             "sort[0][column]": "period", "sort[0][direction]": "desc", **extra}
        j = requests.get(B + route, params=p, timeout=120).json()["response"]
        out += j["data"]
        off += 5000
        if off >= int(j["total"]):
            return out


for proc in ("ENP", "ENG"):
    d = pd.DataFrame(data("move/poe2/data/", {"facets[process][]": proc}, "2025-01"))
    d["value"] = pd.to_numeric(d["value"], errors="coerce")
    g = d.groupby(["duoarea", "series-description"]).agg(last=("period", "max"), first=("period", "min"), n=("value", "count"),
                                                         tot=("value", "sum")).reset_index()
    sel = g[g["series-description"].str.contains(r", TX|Texas|Corpus|Freeport|Sabine|Cameron|Golden|Plaquemines|Calcasieu|Port Arthur", case=False, regex=True)
            | g["duoarea"].str.contains("NMX")]
    print("=====", proc, len(g))
    print(sel.to_string(max_colwidth=110))
    # 'to all countries' LNG port totals
    if proc == "ENG":
        print(g[g["duoarea"].str.endswith("-Z00")].to_string(max_colwidth=110))
# consumption, Texas, all history
d = pd.DataFrame(data("cons/sum/data/", {"facets[duoarea][]": "STX"}, "1990-01"))
d["value"] = pd.to_numeric(d["value"], errors="coerce")
print(d.groupby(["process", "process-name", "series", "units"]).agg(first=("period", "min"), last=("period", "max"), n=("value", "count"),
                                                                   nnull=("value", lambda s: s.isna().sum())).to_string())
for sid in ("N9160TX2", "N9170TX2", "N9140TX2"):
    r = requests.get(f"https://api.eia.gov/v2/seriesid/NG.{sid}.M", params={"api_key": KEY}, timeout=60)
    print(sid, r.status_code, str(r.text)[:300])
print(pd.DataFrame(data("cons/sum/data/", {"facets[duoarea][]": "STX", "facets[process][]": "VC0"}, "2015-01")).head(3))
# marketed production history
d = pd.DataFrame(data("prod/sum/data/", {"facets[duoarea][]": "STX", "facets[process][]": "VGM"}, "2015-01"))
print(d[["period", "value", "units"]].head(4), len(d))
# check duoarea facet names for ports
r = requests.get(B + "move/poe2/facet/duoarea", params={"api_key": KEY}, timeout=60).json()["response"]["facets"]
print([(x["id"], x["name"]) for x in r if x["id"].endswith("-NMX") or x["id"].startswith(("YFLF", "YCCPL", "YSPL", "YPLAQ", "YCRP", "YCAM", "YGPT", "YELP"))])
