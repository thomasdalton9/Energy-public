"""Probe 16: (a) every daily-report item on all 64 listing pages (odd names, '-r1' revisions and their PDF file
names); (b) the developer copy of the main site (neasite.dryicesolutions.net): its categories, and whether the
'Energy Details' panel is served from an endpoint with dates/history; (c) NEA Annual Report 2024/25 and the
'Details of major activities ... up to <month> 2082' posts: do they hold monthly energy tables for 2025?"""
import io
import re
import requests
import urllib3
import pdfplumber
urllib3.disable_warnings()
UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36"
S = requests.Session()
S.headers["User-Agent"] = UA
TD = "https://td.neasite.dryicesolutions.net"
MAIN = "https://neasite.dryicesolutions.net"


def get(u, **kw):
    try:
        return S.get(u, timeout=90, verify=False, **kw)
    except Exception as e:  # noqa
        print("ERR", u, type(e).__name__)
        return None


items = []
for page in range(1, 70):
    r = get(f"{TD}/en/category/daily-operational-reports-1?page={page}")
    if r is None:
        continue
    d = [x for x in dict.fromkeys(re.findall(r'href="[^"]+/detail/([^"]+)"', r.text)) if "operational" in x]
    new = [x for x in d if x not in items]
    if not new:
        break
    items += new
print("daily items:", len(items))
odd = [x for x in items if not re.fullmatch(r"nepal-daily-operational-report-\d{4}-\d\d-\d\d", x)]
print("odd names:", odd)
for slug in odd[:40]:
    r = get(f"{TD}/en/detail/{slug}")
    pdfs = sorted(set(re.findall(r'href="([^"]+\.pdf)"', r.text))) if r is not None else []
    print("  ", slug, "->", [p.split("/")[-1] for p in pdfs])
ds = sorted(re.findall(r"(\d{4})-(\d\d)-(\d\d)", " ".join(items)))
print("daily items span", ds[:1], ds[-1:])

# main site developer copy
r = get(MAIN + "/en/category")
if r is not None:
    t = re.sub(r"<script.*?</script>|<style.*?</style>|<svg.*?</svg>|<[^>]+>", " ", r.text, flags=re.S)
    t = re.sub(r"\s+", " ", t)
    i = t.find("Category Home")
    print("MAIN CATEGORIES:", t[i:i + 1500])
r = get(MAIN + "/en")
if r is not None:
    i = r.text.find("Energy Details")
    print("MAIN COPY PANEL:", re.sub(r"<[^>]+>|\s+", " ", r.text[i:i + 3000])[:600])
    for m in re.findall(r'["\'](/?[a-z0-9_\-/]*(?:energy|api|ajax)[a-z0-9_\-/]*)["\']', r.text, re.I)[:40]:
        print("   endpoint?", m)

# annual report 2024/25 and 'major activities' posts
for slug in ("annual-report-20242025", "details-of-major-activities-self-published-carried-out-by-nepal-electricity-"
             "authority-up-to-chaitra-2082", "details-of-major-activities-self-published-carried-out-by-nepal-"
             "electricity-authority-up-to-poush2082"):
    r = get(f"{MAIN}/en/detail/{slug}")
    pdfs = sorted(set(re.findall(r'href="([^"]+\.pdf)"', r.text))) if r is not None else []
    print("\nDETAIL", slug, r.status_code if r is not None else None, pdfs[:3])
    for p in pdfs[:1]:
        b = get(p)
        if b is None or b.content[:4] != b"%PDF":
            print("  not a pdf", p)
            continue
        with pdfplumber.open(io.BytesIO(b.content)) as pdf:
            print("  pages", len(pdf.pages))
            hits = 0
            for k, pg in enumerate(pdf.pages):
                t = pg.extract_text() or ""
                if re.search(r"(Shrawan|Bhadra|Baisakh|Magh).{0,200}(Shrawan|Bhadra|Baisakh|Magh)", t, re.S) and \
                        re.search(r"GWh|MWh", t) or re.search(r"Load Dispatch|System Operation", t):
                    print(f"  ---- page {k + 1}\n", t[:2500])
                    hits += 1
                    if hits >= 4:
                        break
