"""Probe 3: render BEG/OGJ Haynesville PDF pages (Figs 6-8 are charts), EIA TiE 56361 and its links, company presentations via curl_cffi."""
import os, re, io, time, requests
OUT = "discovery_archive/results/haynesville"
UA = {"User-Agent": "Mozilla/5.0 (research; thomas.dalton@cantab.net)"}
r = requests.get("https://www.beg.utexas.edu/files/content/beg/research/shale/Haynesville%20Shale%20Gas%20Play.pdf", headers=UA, timeout=60)
print("BEG pdf", r.status_code, len(r.content), flush=True)
import fitz
d = fitz.open(stream=r.content, filetype="pdf")
os.makedirs(f"{OUT}/beg_pages", exist_ok=True)
for i, p in enumerate(d):
    p.get_pixmap(dpi=110).save(f"{OUT}/beg_pages/p{i+1}.png")
print("pages", len(d), flush=True)
def save(name, u, imp=None):
    try:
        if imp:
            from curl_cffi import requests as cr
            x = cr.get(u, impersonate="chrome", timeout=60)
        else:
            x = requests.get(u, headers=UA, timeout=60)
    except Exception as e:
        print(f"[{name}] ERR {e}", flush=True); return None
    print(f"[{name}] {x.status_code} {len(x.content)} {x.headers.get('content-type')}", flush=True)
    if x.status_code == 200:
        return x
x = save("eia_tie_56361", "https://www.eia.gov/todayinenergy/detail.php?id=56361")
if x:
    t = re.sub(r"<[^>]+>", " ", x.text); t = re.sub(r"\s+", " ", t)
    open(f"{OUT}/eia_tie_56361.txt", "w").write("URL: https://www.eia.gov/todayinenergy/detail.php?id=56361\n" + t)
    for l in sorted(set(re.findall(r'href="([^"]+)"', x.text))):
        if re.search(r"studies|analysis|pdf|haynes", l, re.I): print("   link", l, flush=True)
for nm, u in {"crk_pres_2026_10": "https://investors.comstockresources.com/static-files/5a596a22-02f6-4b49-a9cc-93ecfe0179c0",
              "crk_pres_2023_12": "https://investors.comstockresources.com/static-files/0f5f7e1b-b8bd-41f0-a34c-3846b8301eb8",
              "exe_3q25_pres": "https://investors.expandenergy.com/static-files/b269b415-dfe3-4eca-b66c-250be02bcbae",
              "crk_ir": "https://investors.comstockresources.com/", "exe_ir": "https://investors.expandenergy.com/"}.items():
    for imp in (False, True):
        x = save(f"{nm} imp={imp}", u, imp)
        if x:
            try:
                if x.content[:5] == b"%PDF-":
                    import pdfplumber
                    with pdfplumber.open(io.BytesIO(x.content)) as pdf:
                        t = "\n".join((p.extract_text() or "") + f"\n[[page {i+1}]]" for i, p in enumerate(pdf.pages))
                else:
                    t = re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", x.text))
                open(f"{OUT}/{nm}.txt", "w").write(f"URL: {u}\n" + t[:1500000]); print("   saved", len(t), flush=True)
            except Exception as e:
                print("   parse err", e, flush=True)
            break
