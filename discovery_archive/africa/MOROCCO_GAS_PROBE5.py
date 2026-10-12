"""
Morocco gas probe, pass 5.

Targets:
  mem      Ministry of Energy Transition 'Chiffres cles' monthly PDFs (www.mem.gov.ma/Pages/rapports.html): full text of the gas
           supply and electricity pages (is there a MONTHLY gas or per-plant series?), plus the weekly 'TDB veille hebdomadaire'
  hcpold   HCP yearbooks 2012-2022 (.rar, extracted with 7z): table 5-10 thermal production by plant and 5-11 fuel use by plant
  onee     ONEE publications page (certificate not verified: it fails from Actions)

Usage: python3 MOROCCO_GAS_PROBE5.py <target>
"""
import glob
import io
import os
import re
import subprocess
import sys
import time

import pandas as pd
import requests
import urllib3

urllib3.disable_warnings()
UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36"
H = {"User-Agent": UA}
OUT = "probe_files"
os.makedirs(OUT, exist_ok=True)


def get(url, tries=2, timeout=(20, 240), **kw):
    kw.setdefault("headers", H)
    last = None
    for _ in range(tries):
        try:
            return requests.get(url, timeout=timeout, **kw)
        except requests.RequestException as e:
            last = e
            time.sleep(3)
    print(f"  ERR {url[:140]} {type(last).__name__}: {str(last)[:160]}")
    return None


def mem():
    import pdfplumber
    r = get("https://www.mem.gov.ma/Pages/rapports.html")
    if r is None or not r.ok:
        print("rapports page failed")
        return
    pdfs = []
    html = r.content.decode("utf-8", errors="replace")      # the static mirror is UTF-8 but sends no charset: r.text gives mojibake links (404)
    for h in re.findall(r'href="([^"]+\.pdf)"', html, re.I):
        u = requests.compat.urljoin(r.url, h)
        name = requests.utils.unquote(u.split("/")[-1])
        if re.search(r"chiffres|conjoncture", name, re.I):
            pdfs.append((u, name))
    print(len(pdfs), "candidate PDFs")
    for u, name in pdfs:
        print("  ", name)
    os.makedirs(f"{OUT}/mem_text", exist_ok=True)
    for u, name in pdfs[:14]:
        rr = get(u)
        if rr is None or not rr.ok or rr.content[:4] != b"%PDF":
            print("  failed", name, rr.status_code if rr is not None else None)
            continue
        with pdfplumber.open(io.BytesIO(rr.content)) as pdf:
            pages = [(pg.extract_text() or "") for pg in pdf.pages[:40]]
        open(f"{OUT}/mem_text/{re.sub(r'[^0-9A-Za-z]+', '_', name)[:80]}.txt", "w").write(
            u + "\n" + "\n".join(f"##### PAGE {i + 1}\n{t}" for i, t in enumerate(pages)))
        print(f"\n######## {name} pages={len(pages)}")
        for i, t in enumerate(pages):
            if re.search(r"gaz naturel|Tahaddart|Mathar|cycle combin|EET|production.{0,30}(électri|electri)|Electricit", t, re.I):
                print(f"--- page {i + 1}\n{t[:3500]}")


def hcpold():
    for fid in ("248148", "248147", "248146", "248145", "248144", "247157", "247156", "247155", "247154", "247153", "247152"):
        rr = get(f"https://www.hcp.ma/file/{fid}/")
        if rr is None or not rr.ok:
            continue
        d = f"/tmp/hcp_{fid}"
        os.makedirs(d, exist_ok=True)
        open(f"{d}/a.rar", "wb").write(rr.content)
        p = subprocess.run(["7z", "x", "-y", f"-o{d}/x", f"{d}/a.rar"], capture_output=True, text=True)
        files = glob.glob(f"{d}/x/**/*", recursive=True)
        cd = requests.utils.unquote(rr.headers.get("content-disposition", ""))
        print(f"\n##### {fid} {cd[-70:]} 7z rc={p.returncode} files={len(files)}")
        if p.returncode:
            print(p.stderr[-400:])
        books = [f for f in files if re.search(r"nergie", os.path.basename(f), re.I) and f.lower().endswith((".xls", ".xlsx"))]
        if not books:
            print("   no energy workbook among", [os.path.basename(f) for f in files][:30])
        for b in books:
            try:
                x = pd.read_excel(b, sheet_name=None, header=None)
            except Exception as e:  # noqa: BLE001
                print("   read error", b, type(e).__name__, str(e)[:80])
                continue
            for s, df in x.items():
                txt = df.astype(str)
                if not txt.apply(lambda c: c.str.contains(r"Tahaddart|Mathar|gaz naturel", case=False, regex=True)).any().any():
                    continue
                if re.search(r"SOMMAIRE|^pg$", str(s), re.I):
                    continue
                df = df.dropna(how="all").dropna(axis=1, how="all")
                df.to_csv(f"{OUT}/hcp_{fid}_{re.sub(r'[^0-9A-Za-z]+', '_', str(s))}.csv", index=False, header=False)
                print(f"--- {os.path.basename(b)} sheet {s!r}")
                for _, row in df.iterrows():
                    vals = [str(v).strip() for v in row.tolist() if str(v) != "nan" and not re.search(r"[؀-ۿ]", str(v))]
                    line = " | ".join(v[:50] for v in vals)
                    if re.search(r"Tahaddart|Mathar|Gaz|Centrales thermiques|Industrie|Production|Consommation|En milli|20\d\d|Redevance|Importation", line, re.I):
                        print("    ", line[:260])


def onee():
    for u in ["https://www.one.org.ma/fr/pages/interne.asp?esp=2&id1=10&t1=1", "https://www.one.org.ma/FR/pages/interne.asp?esp=2&id1=10&t1=1"]:
        rr = get(u, verify=False, timeout=(25, 60))
        print(f"\n=== {u} -> {rr.status_code if rr is not None else None}")
        if rr is None or not rr.ok:
            continue
        for h, t in re.findall(r'href=["\']([^"\'#]+)["\'][^>]*>(.{0,200}?)</a', rr.text, re.S | re.I):
            t = re.sub(r"<[^>]+>|\s+", " ", t).strip()
            if re.search(r"pdf|rapport|activit|chiffre|statist|publication", h + t, re.I):
                print("   ", requests.compat.urljoin(rr.url, h)[:170], "|", t[:90])
        break
    for u in ["https://www.one.org.ma/fr/doc/publications/rapport_activite_2023.pdf", "https://www.one.org.ma/fr/doc/publications/rapport_activite_2022.pdf"]:
        rr = get(u, verify=False, timeout=(25, 60))
        print(u, rr.status_code if rr is not None else None, len(rr.content) if rr is not None else "")


if __name__ == "__main__":
    {"mem": mem, "hcpold": hcpold, "onee": onee}[sys.argv[1]]()
