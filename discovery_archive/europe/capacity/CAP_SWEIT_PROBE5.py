"""Probe 5: France register stats (URL-encoded this time), Spain REE site data feeds, Terna download centre, DGEG monthly PDF archive.
Probe 4 findings: PT DGEG 'Estatisticas Rapidas das Renovaveis' PDF (monthly, e.g. media/cnkfg5w5/dgeg-arr-2026-08.pdf) has a 'Potencia instalada'
MW table: renewable total, hydro (of which pumping, by size), wind, biomass (+RSU, biogas), geothermal, PV (UPAC/UPP/micro), columns = year-end values
of 2017-2025 and the report month of 2026; REData potencia-instalada still 500; sistemaelectrico-ree.es has installed-capacity pages;
Terna dati.terna.it/en/generation/installed-capacity is a 404 page, the download centre is server-rendered."""
import json, re, subprocess, sys
from html import unescape
from urllib.parse import quote
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
        print("\n=====", u[:230], "|", head, flush=True)
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
def q(ds, **kw):
    qs = "&".join(f"{k}={quote(v, safe='(),*=<>\\'')}" for k, v in kw.items())
    return get(f"{B}{ds}/records?{qs}", 2600)

print("#### A. France register stats")
q(R, select="filiere,count(*) as n,sum(puismaxinstallee) as p,sum(nbinstallations) as ni", group_by="filiere", limit="50")
q(R, select="regime,count(*) as n", group_by="regime", limit="20")
q(R, select="count(*) as n", where="datederaccordement is not null", limit="5")
q(R, select="filiere,combustible,sum(puismaxinstallee) as p", group_by="filiere,combustible", limit="100")
q(R, select="filiere,sum(puismaxinstallee) as p", where="datemiseenservice_date<=date'2024-12-31'", group_by="filiere", limit="50")
q(R, select="filiere,sum(puismaxinstallee) as p", where="datemiseenservice_date<=date'2020-12-31'", group_by="filiere", limit="50")
q(R, select="filiere,count(*) as n,sum(puismaxinstallee) as p", where="datemiseenservice_date is null", group_by="filiere", limit="20")
q(R, select="min(datemiseenservice_date) as a,max(datemiseenservice_date) as b", limit="2")
for ds in [R + "-311224", R + "-311220"]:
    print("SNAPSHOT", ds)
    q(ds, select="filiere,count(*) as n,sum(puismaxinstallee) as p", group_by="filiere", limit="50")
q(R, select="nominstallation,puismaxinstallee,datemiseenservice,filiere,codecombustible,regime", where="puismaxinstallee>1000", order_by="puismaxinstallee desc", limit="8")
q(R, select="nominstallation,puismaxinstallee,datemiseenservice,datederaccordement,filiere,codecombustible,regime", where="datederaccordement is not null", limit="6")

print("\n#### B. Spain REE site")
for u in ["https://www.sistemaelectrico-ree.es/en/spanish-electricity-system/installed-capacity",
          "https://www.sistemaelectrico-ree.es/en/renewable-energies-report/wind/installed-capacity-wind"]:
    h = get(u, show=False)
    print("\n", u, len(h))
    print("TEXT:", text(h)[:1500])
    anchors(h, r"\.xls|\.csv|download|descarg|apidatos|/api/", 20)
    print("API refs:", sorted(set(re.findall(r'(https?://[^"\'\s<>]*(?:apidatos|ree\.es/[a-z-]*api)[^"\'\s<>]*)', h)))[:10])
    print("data-attrs:", sorted(set(re.findall(r'data-(?:url|src|endpoint|api|widget)[a-z-]*="([^"]+)"', h)))[:15])
    print("snips:", [h[m.start()-80:m.start()+200].replace("\n", " ") for m in re.finditer(r"potencia-instalada|apidatos", h)][:4])

print("\n#### C. Terna download centre")
h = get("https://dati.terna.it/en/download-center", show=False)
t = text(h)
i = t.find("Download Center", 600)
print("TEXT after intro:", t[i:i + 3000])
for m in list(re.finditer(r"(?i)installed|capacity|efficient|potenza", h))[:12]:
    print("   CTX:", h[max(0, m.start() - 150):m.start() + 250].replace("\n", " "))
print("onclick/data-download:", sorted(set(re.findall(r'(?:data-(?:download|file|url|href)[a-z-]*|onclick)="([^"]{5,200})"', h)))[:20])
anchors(h, r"download|\.xls|\.csv|\.zip|api", 40)

print("\n#### D. DGEG monthly PDF archive (slug pattern)")
months = ["janeiro", "fevereiro", "marco", "abril", "maio", "junho", "julho", "agosto", "setembro", "outubro", "novembro", "dezembro"]
for (y, m) in [(2026, 9), (2026, 8), (2026, 1), (2025, 12), (2025, 6), (2024, 12), (2023, 6), (2022, 3), (2021, 12), (2021, 1)]:
    slug = f"https://www.dgeg.gov.pt/pt/destaques/estatisticas-rapidas-das-renovaveis-{months[m-1]}-{y}/"
    h = get(slug, show=False)
    pdfs = sorted(set(re.findall(r'(/media/[a-z0-9]+/[^"\']+\.pdf)', h)))
    print(y, m, slug.split("/")[-2], len(h), pdfs[:3])
