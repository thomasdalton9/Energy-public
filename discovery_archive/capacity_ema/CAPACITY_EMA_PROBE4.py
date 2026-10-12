"""Probe 4: Ghana Energy Commission booklets - editions listed, and where the newest one prints installed capacity."""
import os, re, sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import CAPACITY_EMA_PROBE1 as P1


def main():
    r = P1.get("https://www.energycom.gov.gh/index.php/planning/energy-statistics")
    found = sorted({(int(y), int(i)) for i, y in re.findall(r"energy-statistics\?download=(\d+):(\d{4})-energy-statistics", r.text)}, reverse=True)
    print("editions", found)
    ed, did = found[0]
    g = P1.get(f"https://www.energycom.gov.gh/index.php/planning/energy-statistics?download={did}:{ed}-energy-statistics")
    pages = P1.pdf_text(g.content)
    print("pages", len(pages), len(g.content))
    sel = [i + 1 for i, t in enumerate(pages) if re.search(r"(?i)dependable", t or "")][:6]
    print("pages with 'dependable'", sel)
    with open(os.path.join(P1.OUT, "probe4_ghana_newest.txt"), "w") as f:
        for i in sel:
            f.write(f"\n===== PAGE {i} =====\n{pages[i-1]}\n")


if __name__ == "__main__":
    main()
