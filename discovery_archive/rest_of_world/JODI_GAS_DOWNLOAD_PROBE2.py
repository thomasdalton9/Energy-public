"""One-off probe 2: JODI gas current data access. Reads the JODI REST API spec (PDF) and prints it, dumps the
data-downloads form, tries the jodidb.org table viewer over http and an OECD SDMX call for industrial production."""
import io, re
import pandas as pd, requests

H = {"User-Agent": "Mozilla/5.0 (X11; Linux x86_64) energy-data/1.0"}
# 1. REST API spec
try:
    r = requests.get("https://www.jodidata.org/_resources/files/downloads/jodi-rest-api-v1.2.pdf", headers=H, timeout=120)
    print("API pdf", r.status_code, len(r.content))
    try:
        from pypdf import PdfReader
    except ImportError:
        import subprocess, sys; subprocess.run([sys.executable, "-m", "pip", "install", "-q", "pypdf"]); from pypdf import PdfReader
    txt = "\n".join((p.extract_text() or "") for p in PdfReader(io.BytesIO(r.content)).pages)
    print("=====API SPEC=====\n" + txt[:12000] + "\n=====END SPEC=====")
except Exception as e:  # noqa: BLE001
    print("API pdf failed", type(e).__name__, str(e)[:200])
# 2. downloads form
try:
    t = requests.get("https://www.jodidata.org/gas/database/data-downloads.aspx", headers=H, timeout=60).text
    i = t.lower().find("beta version")
    print("=====FORM HTML=====\n" + re.sub(r"\s+", " ", t[max(0, i - 3000): i + 3000]) + "\n=====END=====")
    for m in re.findall(r"<(?:input|button|a|form)[^>]*>", t, flags=re.I):
        if any(k in m.lower() for k in ["csv", "download", "submit", "form", "button", "ivt"]):
            print("  tag:", m[:300])
except Exception as e:  # noqa: BLE001
    print("form failed", type(e).__name__, str(e)[:200])
# 3. jodidb table viewer over http
for u in ["http://www.jodidb.org/TableViewer/tableView.aspx?ReportId=38673", "https://www.jodidb.org/TableViewer/tableView.aspx?ReportId=38673",
          "http://www.jodidb.org/TableViewer/summary.aspx?ReportId=38673"]:
    try:
        r = requests.get(u, headers=H, timeout=60)
        print("jodidb", u, r.status_code, len(r.text))
        for l in sorted(set(re.findall(r'(?:href|src|action)="([^"]+)"', r.text))):
            if any(k in l.lower() for k in ["csv", "download", "export", "report", "ivt"]):
                print("   ", l[:200])
        print("   title:", re.findall(r"<title>(.*?)</title>", r.text, flags=re.S)[:1])
    except Exception as e:  # noqa: BLE001
        print("jodidb", u, "failed", type(e).__name__, str(e)[:120])
# 4. OECD industrial production (SDMX)
for u in ["https://sdmx.oecd.org/public/rest/data/OECD.SDD.STES,DSD_STES@DF_INDSERV,4.0/JPN+KOR+IND+CHN.M.PRVM.IX.C.SA?startPeriod=2010-01&format=csvfilewithlabels",
          "https://sdmx.oecd.org/public/rest/data/OECD.SDD.STES,DSD_STES@DF_INDSERV,4.0/JPN+KOR+IND+CHN.M.PRVM.IX..?startPeriod=2010-01&format=csvfilewithlabels",
          "https://sdmx.oecd.org/public/rest/dataflow/OECD.SDD.STES/DSD_STES@DF_INDSERV/latest"]:
    try:
        r = requests.get(u, headers={**H, "Accept": "application/vnd.sdmx.data+csv;labels=both, text/csv, */*"}, timeout=120)
        print("OECD", r.status_code, len(r.content), u[:120]); print("   ", r.text[:600].replace("\n", " || "))
    except Exception as e:  # noqa: BLE001
        print("OECD failed", type(e).__name__, str(e)[:120])
