"""Probe 3: why AFRICA_CAPACITY_ANNUAL.py reads no Ghana Table 3.1 and no Moroccan capacity table from the 2020-2023 yearbooks."""
import io, os, re, sys, tempfile, warnings, traceback
warnings.filterwarnings("ignore")
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "..", "africa"))
import pandas as pd
import CAPACITY_EMA_PROBE1 as P1
import MOROCCO_GAS as mg
OUT = P1.OUT


def main():
    r = P1.get("https://www.energycom.gov.gh/index.php/planning/energy-statistics")
    found = sorted({(int(y), int(i)) for i, y in re.findall(r"energy-statistics\?download=(\d+):(\d{4})-energy-statistics", r.text)}, reverse=True)
    print("ghana found", found[:4])
    ed, did = found[0]
    g = P1.get(f"https://www.energycom.gov.gh/index.php/planning/energy-statistics?download={did}:{ed}-energy-statistics")
    pages = P1.pdf_text(g.content)
    for i, t in enumerate(pages):
        if "Grid Installed and Dependable" in t:
            print("page", i + 1, repr(t[:300]))
    for fid in ("248149", "248148", "248147", "248146"):
        rr = P1.get(f"https://www.hcp.ma/file/{fid}/")
        print("== yearbook file", fid, len(rr.content))
        try:
            with tempfile.TemporaryDirectory() as td:
                books = mg._extract(rr.content, td)
                print(" books", [os.path.basename(b) for b in books])
                lines = []
                for b in books:
                    x = pd.read_excel(b, sheet_name=None, header=None)
                    print("  sheets", list(x)[:30])
                    for s, df in x.items():
                        for ri, row in df.iterrows():
                            txt = " | ".join(str(v) for v in row.values if str(v) != "nan")
                            if re.search(r"(?i)puissance|install|Hydraulique|Charbon", txt):
                                lines.append(f"{os.path.basename(b)}/{s}/r{ri}: {txt[:200]}")
                open(os.path.join(OUT, f"probe3_morocco_{fid}.txt"), "w").write("\n".join(lines))
                for l in lines[:30]: print("   ", l)
        except Exception:
            traceback.print_exc()


if __name__ == "__main__":
    main()
