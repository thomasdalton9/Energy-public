"""MaStR diagnostics (12 Oct 2026): why MaStR hydro / pumped storage / oil differ from ENTSO-E's Germany figures.
Streams EinheitenWasser, EinheitenStromSpeicher_*, EinheitenVerbrennung and prints: hydro by type and status, the largest storage and
hydro units, oil units by size band and by name pattern."""
import codecs, collections, concurrent.futures as cf, importlib.util, re, sys, time
spec = importlib.util.spec_from_file_location("g", "europe/capacity/GERMANY_MASTR_CAPACITY_MONTHLY.py")
g = importlib.util.module_from_spec(spec); sys.modules["g"] = g; spec.loader.exec_module(g)
EXTRA = re.compile(r"<(NameStromerzeugungseinheit|NameKraftwerk|Technologie|ArtDerWasserkraftanlage|EinheitBetriebsstatus|Nettonennleistung|Energietraeger|Land|Bundesland|Einspeisungsart|Inbetriebnahmedatum|DatumEndgueltigeStilllegung)>([^<]*)<")

def work(a):
    url, name, tag = a
    z = g.open_zip(url); dec = codecs.getincrementaldecoder("utf-16")(); carry = ""; out = []
    with z.open(name) as f:
        while True:
            b = f.read(1 << 24)
            if not b: break
            parts = (carry + dec.decode(b)).split("</%s>" % tag); carry = parts.pop()
            for r in parts:
                d = dict(EXTRA.findall(r)); net = g.num(d.get("Nettonennleistung")) / 1000
                if tag == "EinheitVerbrennung" and d.get("Energietraeger") != "2409": continue
                if net >= 0.5 or tag == "EinheitVerbrennung":
                    out.append((net, d))
    return name, out

url = g.export_url(); z = g.open_zip(url); vals = g.read_catalogue(z)
nm = lambda i: vals.get(i, ("?", "?"))[0]
tasks = [(url, "EinheitenWasser.xml", "EinheitWasser"), (url, "EinheitenVerbrennung.xml", "EinheitVerbrennung")] + \
        [(url, i.filename, "EinheitStromSpeicher") for i in z.infolist() if i.filename.startswith("EinheitenStromSpeicher_")]
res = {}
with cf.ProcessPoolExecutor(max_workers=14) as p:
    for name, out in p.map(work, tasks):
        res.setdefault(name.split("_")[0].replace(".xml", ""), []).extend(out)
print("### HYDRO (EinheitenWasser) MW by Art / status")
c = collections.Counter()
for net, d in res["EinheitenWasser"]: c[(nm(d.get("ArtDerWasserkraftanlage")), d.get("EinheitBetriebsstatus"))] += net
for k, v in sorted(c.items(), key=lambda x: -x[1]): print("  ", k, round(v))
print("### largest hydro units")
for net, d in sorted(res["EinheitenWasser"], key=lambda x: -x[0])[:15]: print("  ", round(net), d.get("NameStromerzeugungseinheit"), d.get("NameKraftwerk"), nm(d.get("ArtDerWasserkraftanlage")), d.get("EinheitBetriebsstatus"), d.get("Land"))
print("### STORAGE MW by Technologie / status")
c = collections.Counter()
for net, d in res["EinheitenStromSpeicher"]: c[(nm(d.get("Technologie")), d.get("EinheitBetriebsstatus"))] += net
for k, v in sorted(c.items(), key=lambda x: -x[1]): print("  ", k, round(v))
print("### largest storage units")
for net, d in sorted(res["EinheitenStromSpeicher"], key=lambda x: -x[0])[:25]: print("  ", round(net), d.get("NameStromerzeugungseinheit"), nm(d.get("Technologie")), d.get("EinheitBetriebsstatus"), d.get("Land"), d.get("Inbetriebnahmedatum"))
print("### OIL (Energietraeger 2409) by size band, status 35/37 only")
bands = [(0, 0.1), (0.1, 1), (1, 10), (10, 1e9)]
for lo, hi in bands:
    sel = [(n, d) for n, d in res["EinheitenVerbrennung"] if lo <= n < hi and d.get("EinheitBetriebsstatus") in ("35", "37")]
    print("  band", lo, hi, "units", len(sel), "MW", round(sum(n for n, _ in sel)))
pat = re.compile(r"notstrom|netzersatz|\bNEA\b|notfall|ersatzstrom|diesel|aggregat|usv|bhkw", re.I)
c = collections.Counter(); cn = collections.Counter()
for n, d in res["EinheitenVerbrennung"]:
    if d.get("EinheitBetriebsstatus") in ("35", "37"):
        m = pat.search((d.get("NameStromerzeugungseinheit") or "") + " " + (d.get("NameKraftwerk") or ""))
        k = m.group(0).lower() if m else "(no pattern)"
        c[k] += n; cn[k] += 1
for k, v in c.most_common(): print("  name pattern", k, "units", cn[k], "MW", round(v))
c = collections.Counter()
for n, d in res["EinheitenVerbrennung"]:
    if d.get("EinheitBetriebsstatus") in ("35", "37"): c[(nm(d.get("Einspeisungsart")), nm(d.get("Technologie")))] += n
for k, v in c.most_common(10): print("  einspeisung/technologie", k, round(v))
print("### largest oil units")
for net, d in sorted([x for x in res["EinheitenVerbrennung"] if x[1].get("EinheitBetriebsstatus") in ("35", "37")], key=lambda x: -x[0])[:15]: print("  ", round(net), d.get("NameStromerzeugungseinheit"), "|", d.get("NameKraftwerk"))
