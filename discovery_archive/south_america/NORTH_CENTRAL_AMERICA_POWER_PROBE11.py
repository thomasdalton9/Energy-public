"""
Round 11 (Oct-2026): El Salvador and Belize alternatives, and El Salvador LNG.
UT (ut.com.sv, estadistico.ut.com.sv) is unreachable from GitHub runners
(its name servers and web server time out outside El Salvador - PROBE3..5),
so look at:
  siget  SIGET (regulator) statistics / boletines           www.siget.gob.sv
  dgehm  Direccion General de Energia, Hidrocarburos y Minas (took over the
         CNE's statistics; cne.gob.sv no longer resolves)
  bcr    Banco Central de Reserva - LNG imports
  bel    BEL annual reports (all years) and PUC Belize
  ut     UT again via other host names / ports, in case the block is partial
Prints data-file links (xls/xlsx/csv/pdf/zip) and pages mentioning
generacion / estadistica / gas natural.

Usage: python3 NORTH_CENTRAL_AMERICA_POWER_PROBE11.py siget|dgehm|bcr|bel|ut
"""

print("STARTING", flush=True)

import io
import re
import socket
import sys
from urllib.parse import urljoin, urlparse

import requests

H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) "
                   "Chrome/124.0 Safari/537.36", "Accept-Language": "es-ES,es;q=0.9,en;q=0.8"}
S = requests.Session()
S.headers.update(H)
DATA = re.compile(r"\.(xlsx?|csv|zip|pdf|ods|json)(\?|$)", re.I)
SKIP = re.compile(r"facebook|twitter|linkedin|youtube|instagram|whatsapp|mailto:|tel:|javascript:|wp-json|xmlrpc|"
                  r"\.(jpg|jpeg|png|gif|svg|webp|css|ico|mp4)(\?|$)", re.I)


def get(url, **kw):
    try:
        r = S.get(url, timeout=45, **kw)
        print(f"[{r.status_code}] {r.url[:180]} {r.headers.get('content-type')} {len(r.content):,} B", flush=True)
        return r
    except requests.RequestException as e:
        print(f"ERR {url}: {type(e).__name__}: {str(e)[:200]}", flush=True)
        return None


def crawl(starts, follow, budget=45, hosts=None):
    hosts = hosts or {urlparse(u).netloc.replace("www.", "") for u in starts}
    queue, seen, data = list(starts), set(), {}
    n = 0
    while queue and n < budget:
        u = queue.pop(0)
        if u in seen:
            continue
        seen.add(u)
        r = get(u)
        n += 1
        if r is None or "html" not in (r.headers.get("content-type") or ""):
            continue
        t = r.text
        title = re.search(r"<title[^>]*>(.*?)</title>", t, re.S | re.I)
        print("   title:", re.sub(r"\s+", " ", title.group(1)).strip()[:100] if title else "", flush=True)
        for m in re.finditer(r"<iframe[^>]*src=[\"']([^\"']+)", t, re.I):
            print("   IFRAME", urljoin(r.url, m.group(1))[:200], flush=True)
        for m in re.finditer(r"<a\b[^>]*href=[\"']([^\"']+)[\"'][^>]*>(.*?)</a>", t, re.I | re.S):
            href, text = m.group(1).strip(), re.sub(r"<[^>]+>|\s+", " ", m.group(2)).strip()
            if SKIP.search(href):
                continue
            link = urljoin(r.url, href)
            if DATA.search(link) or "powerbi" in link or "tableau" in link:
                data.setdefault(link, text[:90])
                continue
            host = urlparse(link).netloc.replace("www.", "")
            if any(host.endswith(h) for h in hosts) and re.search(follow, link + " " + text, re.I) and link not in seen:
                queue.append(link)
    print(f"-- {len(data)} data links", flush=True)
    for link, text in data.items():
        if re.search(r"estad|generac|energ|boletin|gas|anual|annual|report|informe|mensual|import|hidrocarb|mercado|19|20",
                     link + text, re.I):
            print(f"   DATA {link[:220]}  [{text}]", flush=True)
    print(f"-- queue left: {queue[:30]}", flush=True)


def siget():
    crawl(["https://www.siget.gob.sv/", "https://www.siget.gob.sv/electricidad/",
           "https://www.siget.gob.sv/estadisticas/", "https://www.siget.gob.sv/category/estadisticas/"],
          r"estad|electric|boletin|mercado|generac|public|document|informe|transparen", 60)


def dgehm():
    for h in ["dgehm.gob.sv", "www.dgehm.gob.sv", "energia.gob.sv", "www.energia.gob.sv", "cne.gob.sv",
              "estadisticas.cne.gob.sv", "www.minec.gob.sv", "www.mh.gob.sv", "www.bcr.gob.sv",
              "www.enfoque.gob.sv", "infoenergia.gob.sv", "www.dgehm.gob.sv"]:
        try:
            print(h, "->", socket.gethostbyname(h), flush=True)
        except OSError as e:
            print(h, "-> no DNS", e, flush=True)
    crawl(["https://www.dgehm.gob.sv/", "https://dgehm.gob.sv/", "https://www.energia.gob.sv/"],
          r"estad|electric|boletin|mercado|generac|hidrocarb|gas|import|public|document|informe|energ", 60)


def bcr():
    crawl(["https://www.bcr.gob.sv/", "https://www.bcr.gob.sv/estadisticas/"],
          r"estad|comercio|import|hidrocarb|petrol|gas|serie|base", 40)


def bel():
    for y in range(2018, 2027):
        for name in [f"Annual Report {y}.pdf", f"Annual_Report_{y}.pdf", f"BEL Annual Report {y}.pdf"]:
            r = get(f"https://bel.com.bz/annual_reports/{name}")
            if r is not None and r.ok and r.content[:4] == b"%PDF":
                import pdfplumber
                with pdfplumber.open(io.BytesIO(r.content)) as pdf:
                    print(f"  {name}: {len(pdf.pages)} pages", flush=True)
                    for i, p in enumerate(pdf.pages):
                        t = p.extract_text() or ""
                        if re.search(r"(BECOL|BELCOGEN|CFE|Hydro|Solar|Biomass|Bagasse).*\n.*(MWh|GWh|kWh)|"
                                     r"(MWh|GWh).*(BECOL|BELCOGEN|CFE)|sources of (energy|supply)|energy (purchases|sources)",
                                     t, re.I | re.S):
                            print(f"  --- {name} page {i + 1}\n{t[:1800]}", flush=True)
                            break
                break
    crawl(["https://www.bel.com.bz/", "https://www.bel.com.bz/Annual_Reports.aspx", "https://www.bel.com.bz/About_Us.aspx"],
          r"report|annual|energy|source|statistic|about|investor|financial", 25)
    crawl(["https://www.puc.bz/", "https://puc.bz/", "https://www.puc.bz/electricity/"],
          r"electric|statistic|report|annual|tariff|review|data|bel|generat", 40)


def ut():
    for h in ["www.ut.com.sv", "ut.com.sv", "estadistico.ut.com.sv", "www.ut.sv", "sim.ut.com.sv", "transparencia.ut.com.sv"]:
        try:
            print(h, "->", socket.getaddrinfo(h, 443)[:2], flush=True)
        except OSError as e:
            print(h, "-> no DNS", e, flush=True)
    for ip, port in [("190.120.15.116", 443), ("190.120.15.116", 8080), ("190.120.15.116", 8443)]:
        s = socket.socket()
        s.settimeout(10)
        try:
            s.connect((ip, port))
            print(f"  TCP {ip}:{port} open", flush=True)
        except OSError as e:
            print(f"  TCP {ip}:{port} {e}", flush=True)
        finally:
            s.close()


if __name__ == "__main__":
    {"siget": siget, "dgehm": dgehm, "bcr": bcr, "bel": bel, "ut": ut}[sys.argv[1]]()
    print("DONE", flush=True)
