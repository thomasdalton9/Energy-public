"""
Probe 7 (compact): nuclear / closure fact pages - Mochovce 4 (SE), Hinkley Point C and Sizewell C, Paks II, Beznau (Axpo), Doel 4 /
Tihange 3 (Engie), Spanish plants (CSN, 7th PGRR, BOE), EDF UK AGR closure dates. Prints MW/GW and year contexts only.
"""
import io, re
import requests
try:
    import pdfplumber
except Exception:
    pdfplumber = None
UA = {"User-Agent": "Mozilla/5.0 (research; energy data)"}


def text(u):
    r = requests.get(u, headers=UA, timeout=(20, 150))
    ct = r.headers.get("content-type", "")
    print("\n#####", r.status_code, len(r.content), ct[:25], u[:150])
    if r.status_code != 200: return "", r
    if "pdf" in ct or u.lower().split("?")[0].endswith(".pdf") or r.content[:4] == b"%PDF":
        try:
            with pdfplumber.open(io.BytesIO(r.content)) as p:
                return "\n".join((pg.extract_text() or "") for pg in p.pages[:120]), r
        except Exception as e:
            print("pdf err", repr(e)[:80]); return "", r
    return re.sub(r"\s+", " ", re.sub(r"<script.*?</script>|<style.*?</style>|<[^>]+>", " ", r.text, flags=re.S)), r


def ctx(t, pats, w=110, n=3):
    for p in pats:
        ms = list(re.finditer(p, t, re.I))
        print(f"  [{p}] {len(ms)}")
        for m in ms[:n]: print("     ..", t[max(0, m.start() - w): m.end() + w].replace("\n", " "))


MW = r"\d[\d,. ]*\s?(?:MW|GW|MWe|megawatt|gigawatt)"
JOBS = [
 ("https://www.seas.sk/o-nas/nase-elektrarne/atomove-elektrarne/mochovce-34-vo-vystavbe/", [MW, r"Mochovce 4[^.]{0,120}20[23]\d"]),
 ("https://www.edfenergy.com/energy/nuclear-new-build-projects/hinkley-point-c/about/reactor", [MW]),
 ("https://www.edfenergy.com/energy/nuclear-new-build-projects/hinkley-point-c/about/news-and-views", [r"203\d", r"2029|2030|2031"]),
 ("https://www.sizewellc.com/about-sizewell-c/", [MW, r"203\d|2034|2035"]),
 ("https://www.sizewellc.com/about-sizewell-c/faqs/", [MW, r"203\d"]),
 ("https://www.axpo.com/group/en/news-and-stories/media-releases.detail.html/news-and-stories/media-releases/2024/axpo-will-operate-the-beznau-nuclear-power-plant-until-2033--inv.html", [r"2032|2033", MW]),
 ("https://www.axpo.com/ch/en/energy/generation-and-distribution/nuclear-power/beznau.html", [MW, r"2032|2033"]),
 ("https://en.newsroom.engie.com/assets/engie-group-pr-closing-belgium-20250314-pdf-64ff8-314df.html?dl=1", [MW, r"2035"]),
 ("https://en.newsroom.engie.com/news/closing-of-the-agreement-between-engie-and-the-belgian-government-6b37d-314df.html", [MW, r"2035"]),
 ("https://www.csn.es/central-nuclear/almaraz/informacion-general", [MW, r"potencia[^.]{0,100}"]),
 ("https://www.miteco.gob.es/content/dam/miteco/es/energia/files-1/nuclear/Residuos/Documents/20231227%20PGRR%207_Version%2027%20diciembre%202023.pdf", [r"Almaraz I\b[^\n]{0,80}", r"Asc[oó] II[^\n]{0,80}", r"Trillo[^\n]{0,60}20\d\d"]),
 ("https://boe.es/boe/dias/2026/08/14/pdfs/BOE-A-2026-17756.pdf", [r"Almaraz[^.]{0,200}20\d\d", r"hasta el[^.]{0,80}20\d\d"]),
 ("https://www.edfenergy.com/energy/nuclear-power-stations", [r"Heysham[^.]{0,200}20\d\d", r"Torness[^.]{0,200}20\d\d", r"Hartlepool[^.]{0,200}20\d\d"]),
]
for u, pats in JOBS:
    try:
        t, r = text(u); print("  chars", len(t)); ctx(t, pats)
    except Exception as e: print("ERR", u[:80], repr(e)[:100])
# Paks II English pages
try:
    h = requests.get("https://www.paks2.hu/", headers=UA, timeout=60).text
    ls = sorted(set(re.findall(r'href="(/en/[^"#]+)"', h)))
    print("paks2 /en links", ls[:30])
    for l in [x for x in ls if re.search(r"project|about|nuclear|new-build|unit", x, re.I)][:4]:
        t, _ = text("https://www.paks2.hu" + l); ctx(t, [MW, r"20[23]\d"], n=3)
except Exception as e: print("paks2 ERR", repr(e)[:100])
