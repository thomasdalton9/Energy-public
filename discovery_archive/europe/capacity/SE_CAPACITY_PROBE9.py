"""
Probe 9 (SE Europe capacity). Output: results/se_probe9.txt
EPDK list: distinct report months present and their link counts; what the months that are neither PDF nor docx are
(2020-05, 2020-10, 2022-02, 2022-12, 2023-04, 2023-08, 2024-01), and the title/links around missing months 2018-03..2019-01, 2019-08, 2023-12.
"""
import os, re, sys, io, html
import requests
from urllib.parse import urljoin
H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36", "Accept-Language": "tr,en;q=0.8"}
OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "results"); os.makedirs(OUT, exist_ok=True)
LOG = open(os.path.join(OUT, "se_probe9.txt"), "w")
def p(*a):
    s = " ".join(str(x) for x in a); print(s, flush=True); LOG.write(s + "\n"); LOG.flush()
MONTHS = {"Ocak": 1, "Şubat": 2, "Mart": 3, "Nisan": 4, "Mayıs": 5, "Haziran": 6, "Temmuz": 7, "Ağustos": 8, "Eylül": 9, "Ekim": 10, "Kasım": 11, "Aralık": 12}
t = requests.get("https://www.epdk.gov.tr/Detay/Icerik/3-0-23-3/elektrikaylik-sektor-raporlar", headers=H, timeout=60).text
titles = re.findall(r"(\d{4}) Yılı Elektrik Piyasası (\w+) Ayı Sekt(?:&#246;|ö)r Raporu", t)
p("distinct titles", len(set(titles)), "title occurrences", len(titles))
have = sorted(set((int(y), MONTHS[m]) for y, m in titles if m in MONTHS))
allm = [(y, m) for y in range(2016, 2027) for m in range(1, 13) if (y, m) <= (2026, 7)]
p("missing titles:", [f"{y}-{m:02d}" for y, m in allm if (y, m) not in have])
# titles in other wording
for m in re.finditer(r"([^<>]{0,60}Elektrik Piyasası[^<>]{0,60}Raporu[^<>]{0,20})", t):
    x = html.unescape(m.group(1))
    if not re.search(r"\d{4} Yılı Elektrik Piyasası \w+ Ayı Sektör Raporu", x): p("other title:", x.strip()[:120])
named = []
for m in re.finditer(r'href="([^"]*DownloadDocument\?id=[^"]+)"', t):
    pre = re.sub(r"\s+", " ", html.unescape(re.sub(r"<[^>]+>", " ", t[max(0, m.start() - 700):m.start()])))
    mm = re.findall(r"(\d{4}) Yılı Elektrik Piyasası (\w+) Ayı Sektör Raporu", pre)
    named.append(((int(mm[-1][0]), MONTHS.get(mm[-1][1])) if mm else None, html.unescape(m.group(1)), html.unescape(m.group(0))[:0]))
for key in [(2020, 5), (2020, 10), (2022, 2), (2022, 12), (2023, 4), (2023, 8), (2024, 1)]:
    for nm, u, _ in named:
        if nm == key:
            r = requests.get(urljoin("https://www.epdk.gov.tr/", u), headers=H, timeout=120)
            p(key, r.status_code, len(r.content), r.headers.get("content-disposition"), r.content[:8])
for key in [(2018, 3), (2018, 12), (2019, 1), (2019, 8), (2019, 12), (2023, 12)]:
    i = t.find(f"{key[0]} Yılı Elektrik Piyasası {[k for k, v in MONTHS.items() if v == key[1]][0]} Ayı")
    p("title lookup", key, i)
# unattributed links (named None) and their context
un = [x for x in named if x[0] is None]
p("unattributed links", len(un))
for nm, u, _ in un[:10]:
    j = t.find(u.replace("&", "&amp;")) if u not in t else t.find(u)
    p("  ", u, re.sub(r"\s+", " ", html.unescape(re.sub(r"<[^>]+>", " ", t[max(0, j - 400):j])))[-300:])
