"""One-off probe 4: COES new-project document folders (Peru) via the portal file browser."""
import html
import re

import requests

PORTAL = "https://www.coes.org.pe/Portal/"
S = requests.Session()
S.headers.update({"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0"})


def browse(path):
    r = S.post(PORTAL + "browser/vistadatos", data={"baseDirectory": path, "url": path, "indicador": "",
                                                    "initialLink": "", "orderFolder": ""}, timeout=(10, 120))
    out = []
    for m in re.finditer(r"openBlob\('([^']+)',\s*'(\w)'", r.text):
        it = (html.unescape(m.group(1)), m.group(2))
        if it not in out:
            out.append(it)
    return out, r.text


for page in ("Planificacion/NuevosProyectos/EstudiosPO", "Planificacion/NuevosProyectos/OperacionComercialUnidades",
             "Planificacion/NuevosProyectos/EstudiosO", "Planificacion/NuevosProyectos/ConclusionOperacion"):
    t = S.get(PORTAL + page, timeout=(10, 120)).text
    vals = {k: html.unescape(v) for k, v in re.findall(r'id="(hf\w+)"[^>]*value="([^"]*)"', t)}
    vals.update({k: html.unescape(v) for k, v in re.findall(r'value="([^"]*)"[^>]*id="(hf\w+)"', t)[::-1]} if False else {})
    print(f"=== {page}: {vals}", flush=True)
    base = vals.get("hfBaseDirectory") or vals.get("hfRelativeDirectory")
    if not base:
        continue
    items, raw = browse(base)
    print(f"  {base}: {len(items)} items; first: {items[:25]}", flush=True)
    if not items:
        print("  raw:", re.sub(r"\s+", " ", raw)[:1500])
    folders = [p for p, k in items if k.upper() == "F"] or [p for p, k in items][:3]
    for f in folders[-3:]:
        sub, _ = browse(f)
        print(f"    {f}: {len(sub)} items: {sub[:30]}", flush=True)
        for g in [p for p, k in sub if k.upper() == "F"][:2]:
            sub2, _ = browse(g)
            print(f"      {g}: {len(sub2)} items: {sub2[:20]}", flush=True)
