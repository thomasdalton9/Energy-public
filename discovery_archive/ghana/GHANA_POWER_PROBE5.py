"""Ghana probe 5: Energy Commission weekly WEM PDFs - listing size params, and word-position dumps of sample PDFs
(pymupdf) so the table parser can be written offline. Results go to discovery_archive/results/ghana/."""
import json, os, re, subprocess, sys, requests
H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0 Safari/537.36"}
EC = "https://www.energycom.gov.gh"
BASE = EC + "/index.php/planning/weekly-wholesale-electricity-market-wem-statistics/category/"
CATS = {2019: "53-2019", 2020: "52-2020", 2021: "51-2021", 2022: "47-2022", 2023: "67-2023", 2024: "69-2024", 2025: "72-2025", 2026: "76-2026"}
OUTD = "discovery_archive/results/ghana"
os.makedirs(OUTD, exist_ok=True)
subprocess.run([sys.executable, "-m", "pip", "install", "-q", "pymupdf"], check=False)
import fitz

def links(html):
    out = []
    for m in re.finditer(r'href="([^"]*\?download=(\d+):[^"]*)"[^>]*>', html):
        out.append((int(m.group(2)), m.group(1).replace("&amp;", "&")))
    seen, res = set(), []
    for i, u in out:
        if i not in seen: seen.add(i); res.append((i, u))
    return res

def names(html):
    return re.findall(r"([A-Za-z0-9 _\-\.]+\.pdf)", re.sub(r"<[^>]+>", " ", html))

# 1) listing params
for q in ["", "?limit=0", "?limit=100", "?limit=50", "?start=20", "?limitstart=20"]:
    try:
        r = requests.get(BASE + CATS[2026] + q, headers=H, timeout=30)
        l = links(r.text)
        print("LIST 2026", repr(q), r.status_code, len(l), l[:1], l[-1:], flush=True)
    except Exception as e:
        print("LIST ERR", q, e, flush=True)
tot = {}
for y, c in CATS.items():
    try:
        r = requests.get(BASE + c + "?limit=0", headers=H, timeout=30)
        l = links(r.text); tot[y] = len(l)
        print("YEAR", y, len(l), "ids", (l[0][0], l[-1][0]) if l else None, "names", names(r.text)[:1], names(r.text)[-1:], flush=True)
    except Exception as e:
        print("YEAR ERR", y, e, flush=True)

# 2) sample dumps
samples = []
for y, c in CATS.items():
    r = requests.get(BASE + c + "?limit=0", headers=H, timeout=30)
    l = links(r.text)
    if not l: continue
    pick = [l[0], l[len(l) // 2]] if y in (2026, 2025) else [l[len(l) // 2]]
    for i, u in pick:
        samples.append((y, i, u))
for y, i, u in samples:
    url = EC + u if u.startswith("/") else u
    try:
        r = requests.get(url, headers=H, timeout=(10, 90))
        fn = f"/tmp/s_{i}.pdf"; open(fn, "wb").write(r.content)
        doc = fitz.open(fn)
        d = {"year": y, "id": i, "bytes": len(r.content), "disp": r.headers.get("content-disposition"), "pages": []}
        for n, p in enumerate(doc):
            ws = p.get_text("words")
            txt = p.get_text("text")
            keep = bool(re.search(r"(?i)kpong|akosombo", txt))
            d["pages"].append({"n": n, "rot": p.rotation, "w": round(p.rect.width), "h": round(p.rect.height), "nwords": len(ws), "kept": keep,
                               "head": txt[:120].replace("\n", " | "),
                               "words": [[round(w[0]), round(w[1]), round(w[2]), round(w[3]), w[4]] for w in ws] if keep else []})
        json.dump(d, open(f"{OUTD}/sample_{y}_{i}.json", "w"), separators=(",", ":"))
        print("SAMPLE", y, i, len(r.content), [(p["n"], p["rot"], p["nwords"], p["kept"]) for p in d["pages"]], flush=True)
    except Exception as e:
        print("SAMPLE ERR", y, i, type(e).__name__, str(e)[:120], flush=True)
