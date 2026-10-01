"""
Discovery probe, round 3, for south_america/SA_POWER_PRICES_DAILY.py (compact output).

Round 2 (SA_POWER_PRICES_DISCOVERY2.py) found:
  - Brazil: every CCEE host (www, dadosabertos, pda-download) answers 403 "Acesso bloqueado ... nao atender as
    politicas de seguranca da CCEE" to GitHub runners, whatever the user agent. ONS open data (S3, no key) has
    'cmo-semi-horario' (CMO_SEMIHORARIO_YYYY.csv, from 2020) and 'cmo-semanal'.
  - Peru: COES menu links /Portal/mercadomayorista/costosmarginales/index (and /revisados).
  - Chile: cne.cl media has monthly Precio_Medio_de_Mercado.xlsx (2024-2026) and old BarCMg zips (node-price
    studies); no marginal-cost statistics files.
  - Uruguay: ADME menu 'Precio Spot sancionado' (/mme_admin/sancionado.php), /detalleejecucionhoraria/,
    /datosabiertos.html.
  - Bolivia: CNDC API /cm-diario (arg dias, default 30), /precmp (arg fecha), /dashboard/precios (modo, anio, mes),
    /historico/monomicos (desde, hasta).
"""
import io
import re
import sys
import traceback

import requests

UA = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) "
                    "Chrome/124.0 Safari/537.36"}
ONLY = set(sys.argv[1:])


def flat(s, n):
    return re.sub(r"\s+", " ", str(s))[:n]


def get(url, n=300, quiet=False, **kw):
    kw.setdefault("timeout", 90)
    kw.setdefault("headers", UA)
    try:
        r = requests.get(url, **kw)
        if not quiet:
            body = r.text if "html" not in str(r.headers.get("content-type")) else "[html]"
            print(f"  {r.status_code} {str(r.headers.get('content-type'))[:25]} {len(r.content)}B {r.url[:150]}"
                  f"\n     {flat(body, n)}")
        return r
    except Exception as e:  # noqa: BLE001
        print(f"  !! {url[:150]}: {type(e).__name__}: {str(e)[:200]}")
        return None


def hrefs(html, pat, n=25):
    out = sorted(set(re.findall(r"""(?:href|src|action)\s*=\s*["']([^"']+)["']""", html, re.I)))
    out = [u for u in out if re.search(pat, u, re.I)]
    for u in out[:n]:
        print("     link", u[:200])
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
    r = get("https://ons-aws-prod-opendata.s3.amazonaws.com/dataset/cmo_tm/CMO_SEMIHORARIO_2026.csv", quiet=True)
    print(f"  2026 csv {r.status_code} {len(r.content)}B")
    lines = r.content.decode("utf-8", "replace").splitlines()
    for line in lines[:6] + ["..."] + lines[-6:]:
        print("    ", line[:200])
    r = get("https://ons-aws-prod-opendata.s3.amazonaws.com/dataset/cmo_tm/CMO_SEMIHORARIO_2021.csv", quiet=True)
    lines = r.content.decode("utf-8", "replace").splitlines()
    print(f"  2021 csv {r.status_code} {len(r.content)}B rows={len(lines)}")
    for line in lines[:3]:
        print("    ", line[:200])


def peru():
    base = "https://www.coes.org.pe"
    for path in ("/Portal/mercadomayorista/costosmarginales/index", "/Portal/mercadomayorista/costosmarginales/revisados"):
        r = get(base + path, quiet=True)
        print(f"  {path}: {r.status_code} {r.url[-60:]}")
        html = r.text
        js = [u for u in hrefs(html, r"\.js", n=0) if "costo" in u.lower() or "marginal" in u.lower()]
        for m in sorted(set(re.findall(r"""(?:url|action)\s*[:=]\s*["']([^"']+)["']""", html)))[:20]:
            print("     url", m)
        for m in re.findall(r"""<(?:input|select)[^>]*id=["']([^"']+)["'][^>]*>""", html)[:25]:
            print("     field", m)
        for u in js:
            print("     js", u)
            j = get(base + u if u.startswith("/") else u, quiet=True)
            if j is not None:
                for m in sorted(set(re.findall(r"""controlador\s*\+\s*["']([^"']+)["']|url\s*:\s*([^,\n]+)""", j.text)))[:30]:
                    print("        ", m)
                for m in sorted(set(re.findall(r"""var\s+(controlador|urlBase|url\w*)\s*=\s*([^;]+);""", j.text)))[:10]:
                    print("        var", m)
                for m in sorted(set(re.findall(r"""data\s*:\s*\{([^}]{0,200})\}""", j.text)))[:8]:
                    print("        data", flat(m, 200))


def chile():
    r = get("https://www.cne.cl/wp-content/uploads/2026/09/Precio_Medio_de_Mercado.xlsx", quiet=True)
    print(f"  PMM xlsx {r.status_code} {len(r.content)}B")
    import openpyxl
    wb = openpyxl.load_workbook(io.BytesIO(r.content), read_only=True, data_only=True)
    for ws in wb.worksheets:
        print(f"   sheet {ws.title!r} dims={ws.dimensions}")
        for i, row in enumerate(ws.iter_rows(values_only=True)):
            if i >= 14:
                break
            vals = [v for v in row if v is not None]
            if vals:
                print("      ", flat(vals, 220))
    for u in ("http://reportediario.cne.cl", "http://server.reportediario.cne.cl/server", "https://www.energiaabierta.cl/",
              "https://energiaabierta.cl/visualizaciones/?t=Electricidad"):
        r = get(u, n=120)
        if r is not None and r.ok:
            hrefs(r.text, r"\.js|api|marginal|cmg|costo", n=15)


def uruguay():
    r = get("https://adme.com.uy/mme_admin/sancionado.php", quiet=True)
    print("  sancionado:", r.status_code)
    hrefs(r.text, r"spot|sanc|\.xls|\.ods|\.csv|\.pdf|\.zip|db-docs", n=30)
    i = r.text.lower().find("sancionado")
    print("     ctx:", flat(re.sub(r"<[^>]+>", " ", r.text[i:i + 3000]), 600))
    for u in ("https://adme.com.uy/detalleejecucionhoraria/", "https://adme.com.uy/datosabiertos.html"):
        r = get(u, quiet=True)
        print(f"  {u}: {r.status_code}")
        if r is not None and r.ok:
            hrefs(r.text, r"spot|precio|sanc|\.xls|\.ods|\.csv|\.json|php|cgi|api", n=30)
            for m in sorted(set(re.findall(r"""(?:url|fetch)\s*[:(=]\s*["'`]([^"'`]+)["'`]""", r.text)))[:20]:
                print("     url", m)


def bolivia():
    api = "https://www.cndc.bo/wp-json/cndc/v1/"
    for q in ("cm-diario?dias=3", "cm-diario?dias=2000", "precmp?fecha=2026-09-01",
              "dashboard/precios?modo=mes&anio=2026&mes=8", "dashboard/precios?modo=anio&anio=2021",
              "historico/monomicos?desde=2021&hasta=2026", "historico/monomicos/detalle?anio=2026"):
        r = get(api + q, n=600)
        if r is not None and r.ok and q.startswith("cm-diario?dias=2000"):
            try:
                j = r.json()
                print("     n=", len(j) if isinstance(j, list) else list(j)[:10], flat(j[-1] if isinstance(j, list) else "", 300))
            except Exception as e:  # noqa: BLE001
                print("     !!", e)


run("brazil", brazil)
run("peru", peru)
run("chile", chile)
run("uruguay", uruguay)
run("bolivia", bolivia)
print("\nDONE")
