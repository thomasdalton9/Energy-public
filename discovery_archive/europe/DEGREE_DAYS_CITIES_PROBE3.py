"""Probe 3: Kosovo ASK PxWeb API tree, Montenegro MONSTAT census links, TurkStat ADNKS, Bosnia agency."""
import json, os, re, requests
OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "results", "degree_days")
UA = {"User-Agent": "Mozilla/5.0"}
log = []
def get(url, **kw):
    try:
        r = requests.get(url, headers=UA, timeout=120, **kw)
        log.append(f"{r.status_code} {len(r.content)} {r.url[:160]}")
        return r
    except Exception as e:
        log.append(f"ERR {url[:160]} {e}")
def walk(url, depth, out):
    r = get(url)
    if r is None or not r.ok:
        return
    try:
        items = r.json()
    except Exception:
        return
    for it in items:
        out.append((url, it))
        if it.get("type") == "l" and depth > 0:
            walk(url.rstrip("/") + "/" + requests.utils.quote(it["id"]), depth - 1, out)
base = "https://askdata.rks-gov.net/api/v1/en/ASKdata"
out = []
for sub in ("Census population", "Population"):
    walk(base + "/" + requests.utils.quote(sub), 2, out)
json.dump(out, open(os.path.join(OUT, "kosovo_tree.json"), "w"), ensure_ascii=False, indent=0)
for u in ["https://monstat.org/eng/page.php?id=1042&pageid=2", "https://monstat.org/eng/index.php", "https://monstat.org/eng/page.php?id=1801&pageid=1801"]:
    r = get(u)
    if r is not None and r.ok:
        t = r.content.decode("utf-8", "ignore")
        for h, txt in re.findall(r'<a[^>]+href="([^"]+)"[^>]*>(.*?)</a>', t, flags=re.S):
            txt = re.sub(r"<[^>]+>", "", txt).strip()
            if re.search(r"census|popis|population|stanovni", txt + h, re.I):
                log.append(f"  LINK {h} :: {txt[:120]}")
for u in ["https://data.tuik.gov.tr/Kategori/GetKategori?p=nufus-ve-demografi-109&dil=2",
          "https://data.tuik.gov.tr/Bulten/Index?p=Address-Based-Population-Registration-System-Results-2024-54019&dil=2",
          "https://data.tuik.gov.tr/Bulten/Index?p=Adrese-Dayali-Nufus-Kayit-Sistemi-Sonuclari-2024-54019",
          "https://biruni.tuik.gov.tr/medas/?locale=en&kn=95", "https://biruni.tuik.gov.tr/medas/?kn=95&locale=tr"]:
    r = get(u)
    if r is not None and r.ok:
        t = r.content.decode("utf-8", "ignore")
        open(os.path.join(OUT, "probe3_tuik_%d.html" % abs(hash(u) % 10000)), "w").write(t[:300000])
        for h, txt in re.findall(r'<a[^>]+href="([^"]+)"[^>]*>(.*?)</a>', t, flags=re.S):
            txt = re.sub(r"<[^>]+>", "", txt).strip()
            if re.search(r"nufus|population|ADNKS|il ve|province", txt + h, re.I):
                log.append(f"  LINK {h} :: {txt[:120]}")
for u in ["https://bhas.gov.ba/?lang=en", "https://www.popis2013.ba/", "https://bhas.gov.ba/Calendar/Category?id=24&page=1&statGroup=24&tip=1&lang=en"]:
    r = get(u)
    if r is not None and r.ok:
        t = r.content.decode("utf-8", "ignore")
        for h, txt in re.findall(r'<a[^>]+href="([^"]+)"[^>]*>(.*?)</a>', t, flags=re.S):
            txt = re.sub(r"<[^>]+>", "", txt).strip()
            if re.search(r"census|popis|population|stanovni", txt + h, re.I):
                log.append(f"  LINK {h} :: {txt[:120]}")
open(os.path.join(OUT, "probe3_log.txt"), "w").write("\n".join(log))
print("\n".join(log))
