"""Probe 4: FR register stats, ES alternatives, PT DGEG PDF/xls, IT Terna pages.
Probe 3 findings: ODRE register 'registre-national-installation-production-stockage-electricite-agrege' (current, 'au 31/07/2026') has
dateraccordement, datederaccordement, datemiseenservice, filiere, puismaxinstallee, nbinstallations and <36 kW aggregates by commune with
their dates; REData potencia-instalada keeps returning HTTP 500 (estructura-generacion works); DGEG has monthly production xls and a
monthly 'Estatisticas rapidas das renovaveis' PDF; Energy-Charts has yearly only for FR/ES/PT/IT."""
import json, re, subprocess, sys
from html import unescape
UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0"

def get(u, n=800, show=True, hdr=(), out=None):
    cmd = ["curl", "-sL", "-m", "60", "-A", UA] + [x for h in hdr for x in ("-H", h)]
    if out:
        cmd += ["-o", out]
    cmd += ["-w", "\n#HTTP %{http_code} %{content_type} %{size_download}", u]
    o = subprocess.run(cmd, capture_output=True, text=True, timeout=90)
    t = o.stdout
    head = t[t.rfind("#HTTP"):].strip()
    body = t[:t.rfind("#HTTP")]
    if show:
        print("\n=====", u, "|", head, flush=True)
        print(body[:n].replace("\n", " "), flush=True)
    return body

def anchors(html, pat, lim=80):
    k = 0
    for m in re.finditer(r'<a [^>]*href="([^"]+)"[^>]*>(.*?)</a>', html, re.S | re.I):
        h, tx = m.group(1), re.sub(r"<[^>]+>", " ", m.group(2))
        tx = " ".join(unescape(tx).split())
        if re.search(pat, h + " " + tx, re.I):
            print("   ", h[:170], "|", tx[:100]); k += 1
            if k >= lim: break

def text(h):
    h = re.sub(r"<(script|style)[^>]*>.*?</\1>", " ", h, flags=re.S | re.I)
    return " ".join(unescape(re.sub(r"<[^>]+>", " ", h)).split())

B = "https://odre.opendatasoft.com/api/explore/v2.1/catalog/datasets/"
R = "registre-national-installation-production-stockage-electricite-agrege"
print("#### A. France register stats")
def q(ds, params, n=2500):
    return get(f"{B}{ds}/records?{params}", n)
q(R, "select=filiere,count(*) as n,sum(puismaxinstallee) as p,sum(nbinstallations) as ni&group_by=filiere&limit=50")
q(R, "select=regime,count(*) as n&group_by=regime&limit=20")
q(R, "select=count(*) as n&where=datederaccordement is not null&limit=5")
q(R, "select=filiere,combustible,sum(puismaxinstallee) as p&group_by=filiere,combustible&limit=100", 3500)
q(R, "select=filiere,sum(puismaxinstallee) as p&where=datemiseenservice_date<=date'2024-12-31'&group_by=filiere&limit=50")
q(R, "select=filiere,sum(puismaxinstallee) as p&where=datemiseenservice_date<=date'2020-12-31'&group_by=filiere&limit=50")
q(R, "select=count(*) as n&where=datemiseenservice_date is null&limit=5")
q(R, "select=filiere,count(*) as n,sum(puismaxinstallee) as p&where=datemiseenservice_date is null&group_by=filiere&limit=20")
q(R, "select=min(datemiseenservice_date) as a,max(datemiseenservice_date) as b&limit=2")
for ds in ["registre-national-installation-production-stockage-electricite-agrege-311224", "registre-national-installation-production-stockage-electricite-agrege-311220"]:
    print("SNAPSHOT", ds)
    q(ds, "select=filiere,count(*) as n,sum(puismaxinstallee) as p&group_by=filiere&limit=50")
q(R, "select=filiere,sum(puismaxinstallee) as p&where=filiere='Nucléaire'&group_by=filiere&limit=5")
q(R, "select=nominstallation,puismaxinstallee,datemiseenservice,filiere,codecombustible&where=puismaxinstallee>1000&order_by=puismaxinstallee desc&limit=8", 2500)

print("\n#### B. Spain")
get("https://apidatos.ree.es/en/datos", 600, hdr=("Accept: application/json",))
get("https://apidatos.ree.es/en/datos/generacion", 600, hdr=("Accept: application/json",))
for w in ["potencia-instalada-renovable", "potencia-instalada-tecnologia", "potencia", "capacidad-instalada"]:
    get(f"https://apidatos.ree.es/en/datos/generacion/{w}?start_date=2025-01-01T00:00&end_date=2025-12-31T23:59&time_trunc=month", 200, hdr=("Accept: application/json",))
for u in ["https://www.sistemaelectrico-ree.es/en/spanish-electrical-system/installed-capacity",
          "https://www.sistemaelectrico-ree.es/en/generation/installed-capacity",
          "https://www.sistemaelectrico-ree.es/en/spanish-electrical-system",
          "https://www.miteco.gob.es/es/energia/energia-electrica/registro-administrativo-instalaciones-produccion-energia-electrica.html",
          "https://energia.gob.es/electricidad/Paginas/estadisticas-sistema-electrico.aspx",
          "https://www.cnmc.es/ambitos-de-actuacion/energia/registro-instalaciones-energia-electrica"]:
    h = get(u, show=False)
    print(u, len(h), text(h)[:200])
    anchors(h, r"\.xls|\.csv|\.pdf|potencia|installed|capacity|raipee|registro", 30)

print("\n#### C. Portugal")
get("https://www.dgeg.gov.pt/media/cnkfg5w5/dgeg-arr-2026-08.pdf", show=False, out="/tmp/arr.pdf")
get("https://www.dgeg.gov.pt/media/uejfuum3/dgeg-epm-2026-08.xls", show=False, out="/tmp/epm.xls")
get("https://www.dgeg.gov.pt/media/qlkzsdqk/dummy", show=False)
try:
    import fitz
except Exception:
    subprocess.run([sys.executable, "-m", "pip", "install", "-q", "pymupdf", "xlrd"])
    import fitz
d = fitz.open("/tmp/arr.pdf")
print("PDF pages", len(d))
for i, p in enumerate(d):
    t = p.get_text()
    if re.search(r"pot[eê]ncia", t, re.I):
        print("--- page", i + 1)
        print(t[:2500])
import xlrd
try:
    wb = xlrd.open_workbook("/tmp/epm.xls")
    for sh in wb.sheets():
        print("SHEET", sh.name, sh.nrows, sh.ncols)
        for r in range(min(sh.nrows, 12)):
            print("  ", [str(c.value)[:18] for c in sh.row(r)[:10]])
except Exception as e:
    print("xls", e)

print("\n#### D. Italy Terna")
for u in ["https://dati.terna.it/en/generation/installed-capacity", "https://dati.terna.it/en/download-center",
          "https://www.terna.it/en/electric-system/statistics/statistical-publications",
          "https://www.terna.it/it/sistema-elettrico/statistiche/rapporto-mensile"]:
    h = get(u, show=False)
    print("\n", u, len(h))
    print("TEXT:", text(h)[:1800])
    print("SCRIPTS:", re.findall(r'<script[^>]+src="([^"]+)"', h)[:15])
    print("API-ish:", sorted(set(re.findall(r'["\'](/?(?:api|dati-api|DesktopModules)[^"\']{3,120})["\']', h)))[:25])
    anchors(h, r"AdactoBackend|\.xls|\.csv|\.pdf|download", 40)
