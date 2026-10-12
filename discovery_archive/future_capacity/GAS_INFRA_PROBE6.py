"""Probe 6: operator / authority pages for the curated Europe gas-infrastructure facts (keyword context)."""
import re, requests
UA = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36"}
DOCS = [
 ("https://www.tap-ag.com/news", [r"expansion", r"1\.2"]),
 ("https://www.tap-ag.com/news/news-stories/tap-triggers-the-first-level-of-capacity-expansion", [r"1\.2", r"first level", r"bcm"]),
 ("https://www.omvpetrom.com/en/media/latest-news/2026/neptun-deep-advances-towards-first-gas-expected-in-the-first-half-of-2027", [r"first gas", r"bcm", r"plateau"]),
 ("https://www.gaz-system.pl/en/for-media/press-releases/2026/august/11-08-2026-the-offshore-pipeline-microtunnel-for-the-fsru-has-reached-its-target-length.html", [r"2028", r"bcm|billion", r"commercial"]),
 ("https://www.gaz-system.pl/en/for-media/press-releases/2025/december/23-12-2025-market-consultations-for-the-binding-fsru-2-open-season-procedure.html", [r"4\.5", r"FSRU 2", r"50%"]),
 ("https://www.gaz-system.pl/en/for-media/press-releases/2025/december/11-12-2025-strategic-project-for-real-independence-10-years-of-the-lng-terminal-in-swinoujscie.html", [r"regasification capacity", r"bcm|billion", r"2028|2027|2030"]),
 ("https://totalenergies.com/news/press-releases/france-totalenergies-demobilizes-its-floating-lng-terminal-le-havre", [r"demobili[sz]", r"January", r"Cape Ann", r"bcm|billion"]),
 ("https://energy-terminal.de/en/terminals", [r"bcm|billion cubic", r"capacity"]),
 ("https://energy-terminal.de/en/terminals/stade", [r"bcm|billion cubic", r"2026|2027", r"Energos"]),
 ("https://energy-terminal.de/en/news/det-steigert-beitrag-zur-versorgungssicherheit-um-ein-gutes-drittel", [r"bcm|billion cubic", r"Stade", r"Brunsb"]),
 ("https://www.gasunie.nl/en/news/important-new-phase-for-the-future-of-eemsenergyterminal-at-the-port-of-eemshaven", [r"2027", r"2036", r"bcm"]),
 ("https://www.gasunie.nl/en/newsletter/gasunie-update-2-2026/seven-questions-about-lng-and-the-continued-use-of-eemsenergyterminal", [r"2036", r"cFID|conditional", r"bcm"]),
 ("https://lng.hr/en/about-terminal/", [r"6\.1", r"nominal", r"3\.9"]),
 ("https://gasgrid.fi/en/bulletins/gasgrid-finland-oy-and-excelerate-energy-inc-have-agreed-to-lease-an-lng-terminal-vessel-the-lng-terminal-vessel-exemplar-will-ensure-finlands-security-of-supply-while-also-serving-the-needs-of/", [r"10 years|ten years|ten-year", r"2032|2033|2022"]),
 ("https://www.snam.it/content/dam/snam/pages-attachments/it/i-nostri-business/rigassificazione/documents/codice-tariffe-area-comitato/codice-di-rigassificazione-fsru-italia/eng-regas-code-piombino/02_DES~1.PDF", [r"Piombino", r"36", r"twelve|12 months", r"October"]),
 ("https://www.snam.it/en/our-businesses/energy-morning/lng,-technical-meetings-kick-off-for-vado-ligure.html", [r"Vado", r"2026|2027"]),
 ("https://www.gastrade.gr/further-increase-in-maximum-regasification-capacity/", [r"regasification", r"MMscf|bcm|billion"]),
 ("https://www.gastrade.gr/the-project/", [r"Thrace", r"bcm|billion|Billion", r"5,5|5\.5"]),
 ("https://www.fluxys.com/en/products-services/supplying-europe/belgium/lng-belgium", [r"capacity", r"bcm|billion"]),
 ("https://www.gasunie.nl/en/projects/eemsenergyterminal", [r"bcm"]),
 ("https://he-hub.de/en/", [r"2027|2028|2029", r"bcm|billion"]),
 ("https://www.centrica.com/media-centre/news/", [r"Rough"]),
 ("https://www.norskpetroleum.no/en/facts/field/irpa/", [r"[Ss]tart|production", r"gas"]),
 ("https://www.norskpetroleum.no/en/facts/field/hugin/", [r"[Ss]tart|production"]),
 ("https://www.norskpetroleum.no/en/facts/field/fenris/", [r"[Ss]tart|production"]),
 ("https://www.norskpetroleum.no/en/facts/field/berling/", [r"[Ss]tart|production"]),
 ("https://www.norskpetroleum.no/en/facts/field/symra/", [r"[Ss]tart|production"]),
 ("https://www.norskpetroleum.no/en/facts/field/snohvit/", [r"[Ss]tart|production|Future"]),
 ("https://www.norskpetroleum.no/en/production-and-exports/exports-of-oil-and-gas/", [r"capacity"]),
]
def text(r):
    t = re.sub(r"<(script|style)[^>]*>.*?</\1>", " ", r.text, flags=re.S)
    return re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", t))
for u, pats in DOCS:
    try:
        r = requests.get(u, headers=UA, timeout=(10, 90))
    except Exception as e:
        print("FAIL", u[:100], type(e).__name__); continue
    print(r.status_code, len(r.content), r.headers.get("last-modified", ""), u[:110])
    if r.status_code != 200: continue
    if r.content[:4] == b"%PDF":
        import subprocess
        open("/tmp/x.pdf", "wb").write(r.content); subprocess.run(["pdftotext", "-layout", "/tmp/x.pdf", "/tmp/x.txt"]); t = re.sub(r"\s+", " ", open("/tmp/x.txt").read())
    else: t = text(r)
    for p in pats:
        for m in list(re.finditer(p, t))[:3]:
            print("   >", p, "::", t[max(0, m.start() - 130): m.end() + 130])
