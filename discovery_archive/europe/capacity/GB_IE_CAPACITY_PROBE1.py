"""Probe 1 (12 Oct 2026): which raw sources give monthly/quarterly installed capacity for GB / NI / Ireland.

FINDINGS (12 Oct 2026, Actions): gov.uk Energy Trends 6 page lists ET 6.1 (quarterly renewable capacity, UK/NI, Q1 2010 on) and ET 6.3
(monthly solar PV, GB/NI/UK, Jan 2010 on) via the content API; REPD csv/xlsx on the REPD page (Operational date per project);
DUKES 5.12 / 5.16 annual. Elexon IGCA / IGCPU return no rows; the BMU registry (reference/bmunits/all) has generationCapacity and
fuelType but no dates or history (3.3 GW 'COAL' still registered). NESO TEC / embedded registers = live agreements only.
Energy-Charts monthly installed_power exists for DE only (UK/IE 404/400). EirGrid System-and-Renewable-Data-Summary-Report xlsx
has monthly wind (IE/NI), connected-generator registers with energisation dates (wind, solar, other, storage) and thermal MEC.
SEAI and ESB Networks pages block or have no files (ESBN data is copied into the EirGrid workbook). Probes 2-5 dump the layouts."""
import re, sys, requests
S = requests.Session()
S.headers["User-Agent"] = "Mozilla/5.0 (X11; Linux x86_64) energy-research"
def get(u, **k):
    try:
        r = S.get(u, timeout=60, **k)
        return r
    except Exception as e:
        print("  ERR", type(e).__name__, str(e)[:150]); return None
def show(u, n=400, **k):
    print("\n=== GET", u)
    r = get(u, **k)
    if r is None: return None
    print("  status", r.status_code, r.headers.get("content-type"), len(r.content), r.headers.get("last-modified"))
    print("  ", r.text[:n].replace("\n", " ") if "text" in (r.headers.get("content-type") or "") or "json" in (r.headers.get("content-type") or "") else "")
    return r
def links(u, pat=r"\.(xlsx|xls|csv|ods|zip)"):
    r = show(u, 100)
    if r is not None and r.ok:
        for m in sorted(set(re.findall(r'href="([^"]+)"', r.text))):
            if re.search(pat, m, re.I): print("   LINK", m)
for u in [
 "https://www.gov.uk/government/statistics/energy-trends-section-6-renewables",
 "https://www.gov.uk/government/publications/renewable-energy-planning-database-monthly-extract",
 "https://www.gov.uk/government/statistics/solar-photovoltaics-deployment",
 "https://www.gov.uk/government/statistics/electricity-chapter-5-digest-of-united-kingdom-energy-statistics-dukes",
 "https://www.gov.uk/government/statistics/energy-trends-section-5-electricity",
]: links(u)
# gov.uk content API (lists attachments)
for slug in ["energy-trends-section-6-renewables","renewable-energy-planning-database-monthly-extract","solar-photovoltaics-deployment","electricity-chapter-5-digest-of-united-kingdom-energy-statistics-dukes"]:
    r = show("https://www.gov.uk/api/content/government/statistics/"+slug if slug!="renewable-energy-planning-database-monthly-extract" else "https://www.gov.uk/api/content/government/publications/"+slug, 50)
    if r is not None and r.ok:
        try:
            j = r.json()
            def walk(o):
                if isinstance(o, dict):
                    if "url" in o and re.search(r"\.(xlsx|csv|ods|xls)$", str(o["url"]), re.I): print("   ATT", o.get("title"), o["url"])
                    for v in o.values(): walk(v)
                elif isinstance(o, list):
                    for v in o: walk(v)
            walk(j)
            print("   updated", j.get("public_updated_at"))
        except Exception as e: print("  json err", e)
# Elexon Insights
B = "https://data.elexon.co.uk/bmrs/api/v1"
for p in ["/datasets/IGCA?publishDateTimeFrom=2026-01-01T00:00Z&publishDateTimeTo=2026-10-01T00:00Z&format=json",
          "/datasets/IGCPU?publishDateTimeFrom=2026-09-01T00:00Z&publishDateTimeTo=2026-10-01T00:00Z&format=json",
          "/reference/bmunits/all?format=json",
          "/generation/outturn/summary?format=json"]:
    show(B+p, 700)
# NESO CKAN
N = "https://api.neso.energy/api/3/action/"
for q in ["TEC register","embedded register","installed capacity","generation capacity","connection"]:
    r = show(N+"package_search?rows=10&q="+q.replace(" ","+"), 50)
    if r is not None and r.ok:
        try:
            for x in r.json()["result"]["results"]:
                print("   PKG", x["name"], x.get("metadata_modified"), [(rr["name"], rr["format"]) for rr in x["resources"]][:5])
        except Exception as e: print("  err", e)
# Energy-Charts fallback
for c in ["uk","ie","gb"]:
    show(f"https://api.energy-charts.info/installed_power?country={c}&time_step=monthly&installation_decommission=false", 500)
# Ireland
for u in ["https://www.eirgrid.ie/grid/system-and-renewable-data-reports",
          "https://www.eirgrid.ie/industry/data-and-reports",
          "https://www.seai.ie/data-and-insights/seai-statistics/key-publications/energy-in-ireland/",
          "https://www.esbnetworks.ie/who-we-are/our-publications",
          "https://www.cru.ie/",
          "https://www.semcommittee.com/",
          "https://www.nienetworks.co.uk/"]:
    links(u, r"\.(xlsx|xls|csv|pdf)")
