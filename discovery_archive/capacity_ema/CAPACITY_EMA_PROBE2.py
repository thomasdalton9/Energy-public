"""Probe 2: dump the pages / sheets that hold installed-capacity tables (Ghana EC 2025 booklet, Morocco HCP yearbook sheet 4,
Israel Electricity Authority report 2024/25, Cyprus CERA 2023 + 2022, Eskom integrated report 2025, NERC 2026 Q2)."""
import io, os, re, sys, tempfile, traceback, warnings
warnings.filterwarnings("ignore")
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import CAPACITY_EMA_PROBE1 as P1   # noqa  (its main() runs on import guard below)
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "..", "africa"))
get, pdf_text, OUT = P1.get, P1.pdf_text, P1.OUT


def dump(name, pages, which):
    with open(os.path.join(OUT, f"probe2_{name}.txt"), "w") as f:
        for i in which:
            if 1 <= i <= len(pages):
                f.write(f"\n===== PAGE {i} =====\n{pages[i-1]}\n")
    print(name, "dumped pages", which)


def pdf(name, url, which=None, rx=None, cffi=False, maxp=14):
    r = get(url, cffi)
    if r is None or r.status_code != 200:
        print(name, "HTTP", None if r is None else r.status_code); return
    pages = pdf_text(r.content)
    sel = list(which or [])
    if rx:
        sel += [i + 1 for i, t in enumerate(pages) if re.search(rx, t or "")][:maxp]
    dump(name, pages, sorted(set(sel)))


def main():
    r = get("https://www.energycom.gov.gh/index.php/planning/energy-statistics?download=774:2025-energy-statistics")
    pdf("ghana", "https://www.energycom.gov.gh/index.php/planning/energy-statistics?download=774:2025-energy-statistics", [20, 21, 22, 23])
    pdf("israel", "https://www.gov.il/BlobFolder/generalpage/dochmeshek/he/General_2024_25_Report.pdf", [11, 18, 25, 26, 105], cffi=True)
    pdf("cyprus2023", "https://www.cera.org.cy/Templates/00001/data/ektheseis/2023_en.pdf", [62, 63, 64, 66, 67, 68])
    pdf("cyprus2022", "https://www.cera.org.cy/Templates/00001/data/ektheseis/2022_en.pdf", rx=r"total installed capacity of (conventional|wind|photovoltaic)", maxp=4)
    pdf("eskom", "https://www.eskom.co.za/wp-content/uploads/2025/10/Eskom-integrated-report-2025.pdf",
        rx=r"(?i)(installed capacity|nominal capacity|net maximum capacity)", maxp=12)
    pdf("nerc", "https://nerc.gov.ng/wp-content/uploads/2026/09/2026_Q2_Report.pdf", [22, 23, 24])
    try:
        import pandas as pd
        import MOROCCO_GAS as mg
        rr = get("https://www.hcp.ma/file/248151/")
        with tempfile.TemporaryDirectory() as td:
            books = mg._extract(rr.content, td)
            x = pd.read_excel(books[0], sheet_name="4", header=None)
            with open(os.path.join(OUT, "probe2_morocco_sheet4.txt"), "w") as f:
                for ri, row in x.iterrows():
                    f.write(f"r{ri}: " + " | ".join(str(v) for v in row.values if str(v) != "nan") + "\n")
            print("morocco sheet 4", x.shape)
    except Exception:
        traceback.print_exc()
    for u in ("https://www.eehc.gov.eg", "http://www.eehc.gov.eg/eehcportal/"):
        r = get(u)
        if r is not None: print(u, re.sub(r"\s+", " ", r.text)[:600])


if __name__ == "__main__":
    main()
