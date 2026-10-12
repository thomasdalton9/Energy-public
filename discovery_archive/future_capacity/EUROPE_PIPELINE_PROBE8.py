"""
Probe 8 (compact): verify nuclear fact documents - EDF 2025 results (Hinkley Point C), gov.uk (Sizewell C), SE (Mochovce 4), Paks II,
CEZ/MPO (Dukovany), gov.pl (PEJ), Nuclearelectrica (Cernavoda 3-4), Spanish 7th PGRR power figures. Prints contexts only.
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
    print("\n#####", r.status_code, len(r.content), ct[:22], u[:140])
    if r.status_code != 200: return ""
    if "pdf" in ct or r.content[:4] == b"%PDF":
        with pdfplumber.open(io.BytesIO(r.content)) as p:
            return "\n".join((pg.extract_text() or "") for pg in p.pages[:150])
    return re.sub(r"\s+", " ", re.sub(r"<script.*?</script>|<style.*?</style>|<[^>]+>", " ", r.text, flags=re.S))


def ctx(t, pats, w=100, n=2):
    for p in pats:
        ms = list(re.finditer(p, t, re.I))
        print(f"  [{p[:50]}] {len(ms)}")
        for m in ms[:n]: print("     ..", t[max(0, m.start() - w): m.end() + w].replace("\n", " "))


def run(u, pats, **k):
    try: ctx(text(u), pats, **k)
    except Exception as e: print("ERR", u[:80], repr(e)[:100])


run("https://www.edf.fr/en/2025-annual-results", [r"Unit 1[^.]{0,160}20\d\d", r"Hinkley[^.]{0,200}20\d\d"])
run("https://www.edfenergy.com/energy/nuclear-new-build-projects/hinkley-point-c", [r"3,?260\s?MW"])
# gov.uk Sizewell C
try:
    j = requests.get("https://www.gov.uk/api/search.json?q=Sizewell+C+final+investment+decision&count=8&fields=title&fields=link", headers=UA, timeout=60).json()
    for x in j["results"]: print("GOVUK", x["link"], "|", x["title"][:80])
    for x in j["results"][:6]:
        if re.search(r"(?i)sizewell", x["title"]):
            c = requests.get("https://www.gov.uk/api/content" + x["link"], headers=UA, timeout=60).json()
            body = re.sub(r"<[^>]+>", " ", json_dump := str(c.get("details", {}))) if False else re.sub(r"<[^>]+>", " ", str(c.get("details", {})))
            print("--", x["link"]); ctx(body, [r"3[.,]\d{1,3}\s?(?:GW|gigawatt|MW)", r"early 2030s|203\d"], n=2)
except Exception as e: print("govuk ERR", repr(e)[:150])
# SE Mochovce
try:
    h = requests.get("https://www.seas.sk/", headers=UA, timeout=60).text
    ls = sorted({l for l in re.findall(r'href="([^"#]+)"', h) if re.search(r"mochovce", l, re.I)})
    print("SE links", ls[:15])
    for l in ls[:5]:
        run(l if l.startswith("http") else "https://www.seas.sk" + l, [r"471|440\s?MW", r"MW", r"20[23]\d"], n=2)
except Exception as e: print("SE ERR", repr(e)[:100])
run("https://www.paks2.hu/en/web/paks-2-en/", [r"\d[\d,. ]*\s?(?:MW|GW)", r"20[23]\d"])
run("https://mpo.gov.cz/en/guidepost/for-the-media/press-releases/the-government-has-decided-on-a-preferred-supplier-for-the-new-nuclear-power-source-at-dukovany--282181/", [r"MW", r"20[23]\d", r"two units|two new|1,?000"])
run("https://www.gov.pl/web/paa-en/resolution-on-the-construction-of-first-nuclear-power-plant-approved-by-the-government", [r"MW", r"20[23]\d"])
run("https://nuclearelectrica.ro/snn/en/investment-projects/reactors-3-and-4/", [r"7\d\d\s?MW", r"20[23]\d"])
run("https://energie.gov.ro/pas-istoric-pentru-unitatile-3-si-4-de-la-cernavoda/", [r"\d{3}\s?MW", r"2031|2032"])
run("https://www.miteco.gob.es/content/dam/miteco/es/energia/files-1/nuclear/Residuos/Documents/20231227%20PGRR%207_Version%2027%20diciembre%202023.pdf",
    [r"MWe", r"potencia (?:el[eé]ctrica|instalada|neta)[^.]{0,120}", r"Almaraz I[^.\n]{0,40}\d{3}\s?MW"])
run("https://www.axpo.com/group/en/news-and-stories/media-releases.detail.html/news-and-stories/media-releases/2024/axpo-will-operate-the-beznau-nuclear-power-plant-until-2033--inv.html",
    [r"Unit 2[^.]{0,140}2032", r"Unit 1[^.]{0,140}2033"])
