"""Probe 1: official installed-capacity tables for East Med / Africa countries (Ghana EC booklet, NERC quarterly,
Morocco HCP yearbook energy workbook, Israel Electricity Authority report, Cyprus CERA annual reports, Egypt EEHC, Eskom).
Prints status and the lines mentioning capacity; writes hits to discovery_archive/results/capacity_ema/probe1_<name>.txt"""
import io, os, re, sys, tempfile, traceback
import requests
import urllib3
urllib3.disable_warnings()
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "..", "africa"))
OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "results", "capacity_ema")
os.makedirs(OUT, exist_ok=True)
H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36"}
KW = re.compile(r"(?i)install|capacity|puissance|capacit|הספק|כושר ייצור|القدرة")


def get(url, cffi=False, **kw):
    try:
        if cffi:
            from curl_cffi import requests as c
            return c.get(url, impersonate="chrome", timeout=90, **kw)
        return requests.get(url, headers=H, timeout=(15, 120), verify=False, **kw)
    except Exception as e:
        print("  ERR", url, type(e).__name__, str(e)[:100]); return None


def pdf_text(content):
    import pdfplumber
    out = []
    with pdfplumber.open(io.BytesIO(content)) as pdf:
        for i, p in enumerate(pdf.pages):
            out.append(p.extract_text() or "")
    return out


def hits(name, pages, maxlines=120):
    lines = []
    for i, t in enumerate(pages):
        ls = t.split("\n")
        for j, l in enumerate(ls):
            if KW.search(l) and re.search(r"\d", l):
                lines.append(f"p{i+1}: {l.strip()[:200]}")
    open(os.path.join(OUT, f"probe1_{name}.txt"), "w").write("\n".join(lines))
    print(f"  {name}: {len(pages)} pages, {len(lines)} capacity lines")
    for l in lines[:maxlines]:
        print("   ", l)


def try_pdf(name, url, cffi=False):
    print("==", name, url)
    r = get(url, cffi)
    if r is None: return
    print("  HTTP", r.status_code, len(r.content), r.headers.get("content-type"))
    if r.status_code == 200 and r.content[:4] == b"%PDF":
        try: hits(name, pdf_text(r.content))
        except Exception as e: print("  pdf err", e)


def links(url, pat, cffi=False):
    r = get(url, cffi)
    if r is None or r.status_code != 200:
        print("  list HTTP", None if r is None else r.status_code); return []
    return re.findall(pat, r.text)


def main():
    # Ghana
    print("#### GHANA")
    r = get("https://www.energycom.gov.gh/index.php/planning/energy-statistics")
    if r is not None:
        print(r.status_code)
        ls = re.findall(r'href="([^"]*download=[^"]*)"[^>]*>([^<]*)', r.text)
        print(ls[:12])
        for u, t in ls[:2]:
            u = u if u.startswith("http") else "https://www.energycom.gov.gh" + u.replace("&amp;", "&")
            try_pdf("ghana_" + re.sub(r"\W+", "_", t)[:30], u)
    # Nigeria NERC
    print("#### NIGERIA")
    r = get("https://nerc.gov.ng/resource-category/nerc-reports/")
    if r is not None:
        pdfs = re.findall(r'href="([^"]+\.pdf)"', r.text)
        print(len(pdfs), pdfs[:8])
        q = [p for p in pdfs if re.search(r"(?i)quarter|q[1-4]", p)][:2]
        for p in q:
            try_pdf("nerc_" + re.sub(r"\W+", "_", p[-30:]), p)
    # Morocco HCP
    print("#### MOROCCO")
    try:
        import MOROCCO_GAS as mg
        r = get(mg.HCP_LIST)
        items = []
        for m in re.finditer(r'href="(/file/(\d+)/)"', r.text):
            ctx = re.sub(r"<[^>]+>|\s+", " ", r.text[m.end():m.end() + 600])
            t = re.search(r"Annuaire[^|]{0,80}?ann[ée]e\s*(\d{4})", ctx)
            if t: items.append((m.group(2), int(t.group(1))))
        items = sorted(set(items), key=lambda x: -x[1])
        print(items[:6])
        import pandas as pd
        for fid, yb in items[:2]:
            rr = get(f"https://www.hcp.ma/file/{fid}/")
            with tempfile.TemporaryDirectory() as td:
                books = mg._extract(rr.content, td)
                print(" yearbook", yb, [os.path.basename(b) for b in books])
                lines = []
                for b in books:
                    x = pd.read_excel(b, sheet_name=None, header=None)
                    for s, df in x.items():
                        for ri, row in df.iterrows():
                            txt = " | ".join(str(v) for v in row.values if str(v) != "nan")
                            if re.search(r"(?i)puissance|install|capacit", txt):
                                lines.append(f"{os.path.basename(b)}/{s}/r{ri}: {txt[:220]}")
                open(os.path.join(OUT, f"probe1_morocco_{yb}.txt"), "w").write("\n".join(lines))
                print("  lines", len(lines))
                for l in lines[:60]: print("   ", l)
    except Exception:
        traceback.print_exc()
    # Israel
    print("#### ISRAEL")
    for y in ("2024_25", "2023_24", "2022_23"):
        try_pdf("israel_" + y, f"https://www.gov.il/BlobFolder/generalpage/dochmeshek/he/General_{y}_Report.pdf", True)
    # Cyprus CERA
    print("#### CYPRUS")
    for y in (2024, 2023, 2022):
        try_pdf(f"cyprus_cera_{y}", f"https://www.cera.org.cy/Templates/00001/data/ektheseis/{y}_en.pdf")
    # Egypt
    print("#### EGYPT")
    for u in ("https://www.eehc.gov.eg", "http://www.eehc.gov.eg/eehcportal/", "https://www.moee.gov.eg"):
        r = get(u)
        if r is not None:
            print(u, r.status_code, len(r.text))
            print("  ", re.findall(r'href="([^"]+\.pdf)"', r.text)[:15])
    # South Africa Eskom
    print("#### SOUTH AFRICA")
    for u in ("https://www.eskom.co.za/wp-content/uploads/2025/10/Eskom-integrated-report-2025.pdf",
              "https://www.eskom.co.za/eskom-divisions/gx/generation-capacity/"):
        r = get(u)
        if r is not None:
            print(u, r.status_code, len(r.content), r.headers.get("content-type"))
            if r.status_code == 200 and r.content[:4] != b"%PDF":
                t = re.sub(r"<[^>]+>", " ", r.text); t = re.sub(r"\s+", " ", t)
                for m in re.finditer(r"(?i)installed|nominal capacity", t):
                    print("   ", t[max(0, m.start()-80):m.start()+160])
                    break
                print("  ", re.findall(r'href="([^"]+\.pdf)"', r.text)[:15])
main()
