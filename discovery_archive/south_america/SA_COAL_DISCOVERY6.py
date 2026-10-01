"""
South America coal - discovery round 6 (basis of the Brazil gap vs the Energy Institute):
  - EPE BEN chapter 2: full tables 2.4 (steam coal) / 2.5 (metallurgical coal) with notes, table 2.1 coal rows (ktoe)
  - EPE Anexo VIII conversion factors (coal calorific values)
  - EI all-data xlsx (OWID snapshot): 'Coal Production - mt' / '- EJ' footnotes and Brazil / Colombia rows
  - ANM Brazil Anuario Mineral Brasileiro 2023 PDF: coal (carvao) lines - ROM vs beneficiated
"""
import html
import io
import re
from urllib.parse import quote

import pandas as pd
import requests

H = {"User-Agent": "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36"}
pd.set_option("display.width", 250)
pd.set_option("display.max_columns", 20)


def get(url):
    try:
        r = requests.get(url, headers=H, timeout=(15, 180))
        print(f"  GET {url[:200]} -> {r.status_code} {len(r.content)} bytes", flush=True)
        return r if r.status_code == 200 else None
    except Exception as e:
        print(f"  GET {url[:200]} -> {type(e).__name__}: {str(e)[:200]}")
        return None


def epe():
    print("\n==== EPE")
    r = get("https://www.epe.gov.br/pt/publicacoes-dados-abertos/publicacoes/BEN-Series-Historicas-Completas")
    links = [html.unescape(h) for h in re.findall(r'href="([^"]+\.xlsx)"', r.text)]
    for pat in (r"Cap.tulo 2 ", r"Anexo VIII"):
        u = next(h for h in links if re.search(pat, h))
        x = get(quote(u if u.startswith("http") else "https://www.epe.gov.br" + u, safe=":/"))
        df = pd.read_excel(io.BytesIO(x.content), header=None)
        if pat.startswith("Cap"):
            lab = df.iloc[:, 0].astype(str)
            for lo, hi in ((0, 38), (86, 134)):
                sub = df.iloc[lo:hi]
                keep = [0] + list(range(df.shape[1] - 7, df.shape[1]))
                print(sub.iloc[:, keep].to_string(max_colwidth=60))
        else:
            for i, row in df.iterrows():
                t = " | ".join(str(v) for v in row.values if str(v) != "nan")
                if re.search(r"carv|coal|kcal", t, re.I):
                    print(f"  r{i}: {t[:300]}")


def ei():
    print("\n==== EI")
    r = get("https://raw.githubusercontent.com/owid/etl/master/snapshots/energy_institute/2026-06-30/"
            "statistical_review_of_world_energy.xlsx.dvc")
    md5 = re.search(r"md5:\s*([0-9a-f]{32})", r.text).group(1)
    x = get(f"https://snapshots.owid.io/{md5[:2]}/{md5[2:]}")
    for s in ("Coal Production - mt", "Coal Production - EJ", "Coal Consumption - EJ"):
        df = pd.read_excel(io.BytesIO(x.content), sheet_name=s, header=None)
        print(f"--- {s}")
        print(df.iloc[2:3, [0] + list(range(df.shape[1] - 9, df.shape[1]))].to_string())
        for i, row in df.iterrows():
            name = str(row.iloc[0])
            if name.strip() in ("Brazil", "Colombia") or (i > 60 and name != "nan"):
                print(f"  r{i}: {name[:400]} {[v for v in row.values[-9:]] if name.strip() in ('Brazil', 'Colombia') else ''}")


def anm_br():
    print("\n==== ANM Brazil AMB")
    import pdfplumber
    for u in ["https://www.gov.br/anm/pt-br/assuntos/economia-mineral/publicacoes/anuario-mineral/anuario-mineral-brasileiro/amb_2023.pdf",
              "https://www.gov.br/anm/pt-br/assuntos/economia-mineral/publicacoes/anuario-mineral/anuario-mineral-brasileiro/PreviaAMB2022.pdf"]:
        r = get(u)
        if r is None or r.content[:4] != b"%PDF":
            continue
        n = 0
        with pdfplumber.open(io.BytesIO(r.content)) as pdf:
            for i, p in enumerate(pdf.pages):
                for ln in (p.extract_text() or "").splitlines():
                    if re.search(r"carv[aã]o|ROM|beneficiad", ln, re.I):
                        print(f"  p{i + 1}: {ln[:220]}")
                        n += 1
                if n > 60:
                    break


if __name__ == "__main__":
    for f in (epe, ei, anm_br):
        try:
            f()
        except Exception as e:
            import traceback
            traceback.print_exc()
