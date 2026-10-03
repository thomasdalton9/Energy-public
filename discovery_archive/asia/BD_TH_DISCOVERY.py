"""
Bangladesh + Thailand discovery, round 1: sources beyond PGCB (Bangladesh) and EPPO electricity / RID (Thailand).

Bangladesh
  - Petrobangla: daily gas production & supply report (fields, RLNG, supply by sector)
  - BPDB: daily generation report (plant-wise), installed capacity by fuel
Thailand
  - EGAT: daily generation / peak (system control centre statistics)
  - EPPO: natural gas tables (supply by source, LNG, use by sector) and capacity tables
  - DMF (Department of Mineral Fuels): petroleum / gas production statistics

For each site: fetch the landing pages, list links that look like data (report / statistics / daily / pdf / xls),
and fetch the first few candidates to show what they hold.
"""
import io
import re
from urllib.parse import urljoin

import requests
import urllib3

urllib3.disable_warnings()
H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0 Safari/537.36",
     "Accept": "*/*", "Accept-Language": "en-US,en;q=0.9"}
T = (20, 60)
FILE = re.compile(r"\.(pdf|xlsx?|csv|zip)(\?|$)", re.I)


def out(*a):
    print(*a, flush=True)


def get(u, **kw):
    try:
        r = requests.get(u, headers=H, timeout=T, verify=False, **kw)
        out(f"GET {r.url} -> {r.status_code} {r.headers.get('content-type')} {len(r.content)}")
        return r
    except Exception as e:  # noqa: BLE001
        out(f"GET {u} ERROR {type(e).__name__}: {str(e)[:150]}")
        return None


def links(r):
    if r is None or r.status_code != 200 or "html" not in (r.headers.get("content-type") or ""):
        return []
    res = []
    for href, text in re.findall(r'<a[^>]+href=["\']([^"\'#]+)["\'][^>]*>(.*?)</a>', r.text, re.S | re.I):
        text = re.sub(r"<[^>]+>|\s+", " ", text).strip()
        res.append((urljoin(r.url, href.strip()), text[:100]))
    return res


def scan(name, pages, key, deep=8):
    out(f"\n==================== {name}")
    seen, cands = set(), []
    for p in pages:
        for u, t in links(get(p)):
            if u in seen:
                continue
            seen.add(u)
            if re.search(key, u + " " + t, re.I):
                out(f"  [{t}] -> {u}")
                cands.append(u)
    # follow the non-file candidates one level to find files
    files = [u for u in cands if FILE.search(u)]
    for u in [c for c in cands if not FILE.search(c)][:deep]:
        for v, t in links(get(u)):
            if FILE.search(v) and v not in files:
                files.append(v)
                out(f"    file [{t}] -> {v}")
    for f in files[:6]:
        r = get(f)
        if r is not None and r.status_code == 200:
            head = r.content[:8]
            out(f"    {f}: starts {head!r}")
            if r.content[:4] == b"%PDF":
                try:
                    import pdfplumber
                    with pdfplumber.open(io.BytesIO(r.content)) as pdf:
                        t = (pdf.pages[0].extract_text() or "")[:900]
                        out("    page1: " + t.replace("\n", " | "))
                except Exception as e:  # noqa: BLE001
                    out(f"    pdf: {e!r}")
            elif f.lower().split("?")[0].endswith((".xls", ".xlsx")):
                try:
                    import pandas as pd
                    x = pd.ExcelFile(io.BytesIO(r.content))
                    out(f"    sheets {x.sheet_names[:10]}")
                    out(x.parse(x.sheet_names[0], header=None).head(15).to_string()[:1500])
                except Exception as e:  # noqa: BLE001
                    out(f"    xls: {e!r}")


def main():
    # ---------------- Bangladesh
    scan("Petrobangla", ["https://petrobangla.org.bd/", "https://www.petrobangla.org.bd/",
                         "https://petrobangla.org.bd/site/view/daily_report",
                         "https://petrobangla.org.bd/site/page/daily-gas-production"],
         r"daily|production|report|mmcfd|gas.?supply|rlng|lng|statistic|\.pdf|\.xls")
    scan("BPDB", ["https://bpdb.gov.bd/", "https://www.bpdb.gov.bd/", "https://misc.bpdb.gov.bd/",
                  "https://bpdb.gov.bd/site/page/installed-capacity"],
         r"daily|generation|capacity|report|load|statistic|annual|\.pdf|\.xls")
    scan("Bangladesh Power Division / SREDA", ["https://powerdivision.gov.bd/", "https://ndre.sreda.gov.bd/"],
         r"generation|capacity|statistic|daily|report|installed")
    # ---------------- Thailand
    scan("EGAT", ["https://www.egat.co.th/home/", "https://www.egat.co.th/home/en/",
                  "https://www.egat.co.th/home/statistics-all-latest/", "https://www.egat.co.th/home/en/statistics-all-latest/"],
         r"statistic|สถิติ|peak|daily|generation|การผลิต|load|\.xls|\.pdf")
    scan("EPPO natural gas / capacity",
         ["https://www.eppo.go.th/epposite/info/stat/natural-gas", "https://www.eppo.go.th/epposite/info/stat/electricity",
          "https://www.eppo.go.th/index.php/en/en-energystatistics/ngv-statistic",
          "https://www.eppo.go.th/index.php/en/en-energystatistics/electricity-statistic",
          "https://www.eppo.go.th/index.php/en/en-energystatistics"],
         r"T0[1-9]_|gas|ก๊าซ|capacity|กำลังผลิต|lng|\.xls")
    scan("DMF", ["https://dmf.go.th/", "https://www.dmf.go.th/", "https://www.dmf.go.th/public/list/data/index/menu/603"],
         r"statistic|สถิติ|production|การผลิต|gas|ก๊าซ|\.xls|\.pdf")
    scan("ERC Thailand", ["https://www.erc.or.th/", "https://www.erc.or.th/en"],
         r"statistic|capacity|generation|data|\.xls")


if __name__ == "__main__":
    main()
