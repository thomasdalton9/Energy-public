"""Ghana probe 6: word dumps of the week-4 PDFs that GHANA_WEM_WEEKLY_POWER.py could not parse."""
import json, os, re, subprocess, sys, requests
H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0 Safari/537.36"}
EC = "https://www.energycom.gov.gh"
subprocess.run([sys.executable, "-m", "pip", "install", "-q", "pymupdf"], check=False)
import fitz
OUTD = "discovery_archive/results/ghana"
IDS = {2022: (47, 423), 2023: (67, 647), 2024: (69, 744), 2025: (72, 787), 2026: (76, 862)}
for y, (cat, i) in IDS.items():
    slug = {2022: "47-2022", 2023: "67-2023", 2024: "69-2024", 2025: "72-2025", 2026: "76-2026"}[y]
    tail = {2022: "weekly-electricity-market-statistics-2022", 2023: "weekly-wholesale-electricity-market-wem-statistics-2023",
            2024: "weekly-wholesale-electricity-market-wem-statistics-2024", 2025: "weekly-wholesale-electricity-market-wem-statistics-2025",
            2026: "weekly-wholesale-electricity-market-wem-statistics-2026"}[y]
    # find the real link on the listing page
    html = requests.get(f"{EC}/index.php/planning/weekly-wholesale-electricity-market-wem-statistics/category/{slug}?limit=100", headers=H, timeout=30).text
    m = re.search(rf'href="([^"]*\?download={i}:[^"]*)"', html)
    if not m:
        print("no link", y, i, flush=True); continue
    u = EC + m.group(1).replace("&amp;", "&")
    r = requests.get(u, headers=H, timeout=(10, 90))
    doc = fitz.open(stream=r.content, filetype="pdf")
    d = {"year": y, "id": i, "disp": r.headers.get("content-disposition"), "pages": []}
    for n, p in enumerate(doc):
        ws = p.get_text("words")
        d["pages"].append({"n": n, "rot": p.rotation, "w": round(p.rect.width), "h": round(p.rect.height), "nwords": len(ws),
                           "head": p.get_text("text")[:100].replace("\n", " | "),
                           "words": [[round(w[0]), round(w[1]), round(w[2]), round(w[3]), w[4]] for w in ws] if any(x[4] == "Total" for x in ws) else []})
    json.dump(d, open(f"{OUTD}/fail_{y}_{i}.json", "w"), separators=(",", ":"))
    print("FAIL-DUMP", y, i, d["disp"], [(p["n"], p["nwords"], bool(p["words"])) for p in d["pages"]], flush=True)
