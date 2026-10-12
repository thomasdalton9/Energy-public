"""
Probe 2: MaStR 'In Planung' (EinheitBetriebsstatus 31) and shutdown-planned unit records - which fields exist (planned commissioning date,
name, planned shutdown) in EinheitenWind / Verbrennung / Kernkraft / StromSpeicher / Solar. Streams only the first ~96 MB of each member.
"""
import codecs, collections, importlib.util, os, re, sys
HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(os.path.dirname(HERE))
sys.path.insert(0, ROOT)
spec = importlib.util.spec_from_file_location("g", os.path.join(ROOT, "europe/capacity/GERMANY_MASTR_CAPACITY_MONTHLY.py"))
g = importlib.util.module_from_spec(spec); spec.loader.exec_module(g)
url = g.export_url(); print(url)
z = g.open_zip(url)
names = [(i.filename, i.file_size, i.compress_size) for i in z.infolist()]
for n in names:
    if n[0].startswith("Einheiten"):
        print(n)
LIM = 96 << 20
for name in ["EinheitenWind.xml", "EinheitenVerbrennung.xml", "EinheitenKernkraft.xml", "EinheitenStromSpeicher_1.xml", "EinheitenSolar_1.xml", "EinheitenBiomasse.xml", "EinheitenWasser.xml"]:
    if not any(n[0] == name for n in names):
        print("missing", name); continue
    tag = "Einheit" + name[len("Einheiten"):].split("_")[0].split(".")[0]
    close = "</%s>" % tag
    dec = codecs.getincrementaldecoder("utf-16")()
    carry = ""; got = 0; stat = collections.Counter(); samples = {}
    with z.open(name) as f:
        while got < LIM:
            b = f.read(1 << 22)
            if not b: break
            got += len(b)
            parts = (carry + dec.decode(b)).split(close); carry = parts.pop()
            for rec in parts:
                d = dict(re.findall(r"<(\w+)>([^<]*)<", rec))
                s = d.get("EinheitBetriebsstatus"); stat[s] += 1
                if s == "31" and "plan" not in samples and float(d.get("Nettonennleistung", 0) or 0) > 1000:
                    samples["plan"] = d
                if s != "31" and d.get("DatumEndgueltigeStilllegung") is None and d.get("GeplanteStilllegung") and "stop" not in samples:
                    samples["stop"] = d
                for k in d:
                    if re.search(r"(?i)geplant|plan", k): samples.setdefault("k_" + k, d[k])
    print("=====", name, dict(stat))
    for k, v in samples.items():
        print(k, v if not isinstance(v, dict) else v)
