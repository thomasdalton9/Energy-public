"""
Discovery probe, round 2, for south_america/SA_POWER_PRICES_DAILY.py (compact output).

Round 1 (SA_POWER_PRICES_DISCOVERY.py) found:
  - Colombia: XM PrecBolsNaci (hourly, COP/kWh, MaxDays 31) and PrecEsca (daily) work; TRM on datos.gov.co
    (32sa-8pi3); banrep.gov.co is behind a bot-manager captcha.
  - Argentina: CAMMESA PARTE_POST_OPERATIVO mdb has PRECIOS_AREA (hourly CMO marginal cost and PRECIO_AREA spot,
    ARS/MWh); BCRA estadisticas v4.0 variable 5 (A 3500) works, v3.0 is retired.
  - Peru: BCRP series work (PD04640PD SBS venta, PD04638PD interbank venta); COES costosmarginales URLs 404.
  - Brazil: dadosabertos.ccee.org.br returns 403 "Acesso bloqueado" to GitHub runners; BCB SGS 1 / PTAX ok.
  - Uruguay: ADME menu has /mme_admin/sancionado.php; Bolivia CNDC API has /cm-diario, /historico/monomicos.
  - Ecuador: cenace.gob.ec fails TLS verification (missing intermediate) - not used (no spot market anyway).
"""
import json
import re
import sys
import traceback

import requests

UA = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) "
                    "Chrome/124.0 Safari/537.36"}
ONLY = set(sys.argv[1:])


def flat(s, n):
    return re.sub(r"\s+", " ", s)[:n]


def text(html):
    return re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", re.sub(r"(?is)<(script|style).*?</\1>", " ", html)))


def req(method, url, n=300, **kw):
    kw.setdefault("timeout", 60)
    kw.setdefault("headers", UA)
    try:
        r = requests.request(method, url, **kw)
        body = r.text if "html" not in str(r.headers.get("content-type")) else "[html] " + text(r.text)
        print(f"  {r.status_code} {str(r.headers.get('content-type'))[:30]} {len(r.content)}B {r.url[:160]}\n"
              f"     {flat(body, n)}")
        return r
    except Exception as e:  # noqa: BLE001
        print(f"  !! {url[:160]}: {type(e).__name__}: {str(e)[:200]}")
        return None


def hrefs(html, pat, n=60):
    out = sorted(set(re.findall(r"""(?:href|src|action)\s*=\s*["']([^"']+)["']""", html, re.I)))
    out = [u for u in out if re.search(pat, u, re.I)]
    for u in out[:n]:
        print("     link", u)
    return out


def run(name, fn):
    if ONLY and name not in ONLY:
        return
    print("\n==== " + name, flush=True)
    try:
        fn()
    except Exception:  # noqa: BLE001
        traceback.print_exc(limit=2)


def brazil():
    for ua in (UA, {"User-Agent": "python-requests/2.32"}, {"User-Agent": "curl/8.5.0"}, {}):
        req("GET", "https://dadosabertos.ccee.org.br/api/3/action/package_show?id=pld_horario", n=500, headers=ua)
    for u in ("https://dadosabertos.ccee.org.br/", "https://www.ccee.org.br/", "https://www.ccee.org.br/precos",
              "https://pda-download.ccee.org.br/", "https://www.ccee.org.br/web/guest/precos/painel-precos"):
        req("GET", u, n=300)
    # ONS open data (S3) - CMO, as a fallback reference
    for u in ("https://dados.ons.org.br/api/3/action/package_search?q=cmo&rows=20",
              "https://ons-aws-prod-opendata.s3.amazonaws.com/?prefix=dataset/cmo&max-keys=50"):
        r = req("GET", u, n=200)
        if r is not None and r.ok:
            if "json" in str(r.headers.get("content-type")):
                for p in r.json()["result"]["results"]:
                    print("     PKG", p["name"], "|", p.get("title"))
                    for res in p.get("resources", [])[:6]:
                        print("        ", res.get("name"), res.get("format"), str(res.get("url"))[:150])
            else:
                for k in re.findall(r"<Key>([^<]+)</Key>", r.text)[:50]:
                    print("     key", k)


def peru():
    r = req("GET", "https://www.coes.org.pe/Portal/", n=100)
    if r is not None:
        hrefs(r.text, r"marginal|cmg|costo|precio|tarif|mercado")
    r = req("GET", "https://www.coes.org.pe/Portal/portalinformacion", n=100)
    if r is not None:
        hrefs(r.text, r"portalinformacion|marginal|cmg")
        for m in sorted(set(re.findall(r"""['"](/Portal/portalinformacion/[A-Za-z]+)['"]""", r.text))):
            print("     pi", m)
        i = r.text.lower().find("marginal")
        print("     ctx:", flat(r.text[max(0, i - 400): i + 400], 800) if i >= 0 else "none")


def chile():
    media = "https://www.cne.cl/wp-json/wp/v2/media"
    seen = set()
    for q in ("marginal", "costo marginal", "CMg", "Costos Marginales", "precio medio", "precio de mercado", "PMM",
              "nudo", "Reporte Mensual", "Precio"):
        try:
            r = requests.get(media, params={"search": q, "per_page": 100, "_fields": "date,source_url"},
                             headers=UA, timeout=60)
            items = r.json() if r.ok else []
            print(f"  search {q!r}: {r.status_code} {len(items)}")
            for i in items:
                u = i.get("source_url", "")
                if u not in seen and not re.search(r"\.(png|jpe?g|gif|svg)$", u, re.I):
                    seen.add(u)
                    print(f"     {i.get('date', '')[:10]} {u}")
        except Exception as e:  # noqa: BLE001
            print("  !!", e)
    for u in ("https://www.cne.cl/nuestros-servicios/reportes/informacion-y-estadisticas/",
              "https://www.cne.cl/estadisticas/electricidad/"):
        r = req("GET", u, n=100)
        if r is not None:
            hrefs(r.text, r"margin|cmg|precio|pmm|xls|csv")


def uruguay():
    for u in ("https://adme.com.uy/mme_admin/sancionado.php", "https://adme.com.uy/mmee/infmensual.php",
              "https://adme.com.uy/mmee/detallehorarioejecutado.php"):
        r = req("GET", u, n=600)
        if r is not None and r.ok:
            hrefs(r.text, r".")
            for f in re.findall(r"(?is)<form.*?</form>", r.text)[:3]:
                print("     form", flat(f, 600))
            for m in re.findall(r"""(?:url|src)\s*[:=]\s*['"]([^'"]+\.(?:php|cgi|csv|xls|xlsx|ods|json)[^'"]*)['"]""",
                                r.text)[:30]:
                print("     ref", m)


def bolivia():
    api = "https://www.cndc.bo/wp-json/cndc/v1/"
    idx = requests.get(api, headers=UA, timeout=60).json()["routes"]
    for k in ("/cndc/v1/cm-diario", "/cndc/v1/historico/monomicos", "/cndc/v1/historico/monomicos/detalle",
              "/cndc/v1/dashboard/precios", "/cndc/v1/precmp", "/cndc/v1/ga/precios-monomonicos"):
        args = {a: {kk: vv for kk, vv in v.items() if kk in ("required", "type", "default", "description")}
                for e in idx.get(k, {}).get("endpoints", []) for a, v in e.get("args", {}).items()}
        print(f"  route {k} args={json.dumps(args)[:400]}")
    for q in ("cm-diario", "cm-diario?fecha=2026-09-01", "cm-diario?desde=2026-09-01&hasta=2026-09-03",
              "historico/monomicos?desde=2021&hasta=2026", "dashboard/precios", "precmp",
              "ga/precios-monomonicos"):
        req("GET", api + q, n=700)


run("brazil", brazil)
run("peru", peru)
run("chile", chile)
run("uruguay", uruguay)
run("bolivia", bolivia)
print("\nDONE")
