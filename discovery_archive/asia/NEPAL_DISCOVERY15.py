"""Probe 15: NEA monthly operational reports (NMOR) and anything newer than Jan 2025. Developer copy (plain
requests): every page of the monthly / yearly / daily categories, all categories on transd, a monthly detail page
and its PDF text; constructed NMOR and NDOR URLs for 2081-2083 BS under several folder names."""
import io
import re
import requests
import urllib3
import pdfplumber
urllib3.disable_warnings()
UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36"
S = requests.Session()
S.headers["User-Agent"] = UA
M = "https://td.neasite.dryicesolutions.net"


def get(u):
    try:
        r = S.get(u, timeout=60, verify=False)
        return r
    except Exception as e:  # noqa
        print("ERR", u, type(e).__name__)
        return None


def details(html):
    return list(dict.fromkeys(re.findall(r'href="[^"]+/detail/([^"]+)"', html)))


# all categories on the transd site copy
r = get(M + "/en/category")
if r is not None:
    t = re.sub(r"<script.*?</script>|<style.*?</style>|<svg.*?</svg>", " ", r.text, flags=re.S)
    cats = sorted(set(re.findall(r'href="([^"]+/category/[^"?#]+)"', r.text)))
    print("CATEGORY LIST TEXT:", re.sub(r"<[^>]+>|\s+", " ", t)[t.find("Category") if "Category" in t else 0:][:2500])
    for c in cats:
        print("  CAT", c)

for slug in ("monthly-operational-reports", "yearly-operational-reports-1", "daily-operational-reports-1"):
    allp = []
    for page in range(1, 80):
        r = get(f"{M}/en/category/{slug}?page={page}")
        if r is None:
            break
        d = details(r.text)
        d = [x for x in d if "operational" in x or "summary" in x or "yearly" in x or "monthly" in x]
        if not d or (allp and d[0] in allp):
            break
        allp += d
        if slug.startswith("daily") and page > 1:
            break   # daily already known: page 1 only
    print(f"\n{slug}: {len(allp)} items; first {allp[:6]}; last {allp[-4:]}")
    if slug.startswith("monthly"):
        months = allp

# a monthly detail page -> PDF -> text
for slug in months[:1] + months[-1:]:
    r = get(f"{M}/en/detail/{slug}")
    pdfs = re.findall(r'href="([^"]+\.pdf)"', r.text) if r is not None else []
    print("\nDETAIL", slug, pdfs[:2])
    if pdfs:
        b = get(pdfs[0])
        if b is not None and b.content[:4] == b"%PDF":
            with pdfplumber.open(io.BytesIO(b.content)) as pdf:
                print("pages", len(pdf.pages))
                for p in pdf.pages[:3]:
                    print("-----"); print(p.extract_text())

# constructed URLs for newer files
folders = ["Daily_op_reports", "Monthly_op_reports", "monthly_op_reports", "Monthly_Op_Reports", "Yearly_op_reports"]
for y, mth in [(2081, 9), (2081, 10), (2081, 12), (2082, 1), (2082, 6), (2082, 12), (2083, 3), (2083, 5)]:
    for f in folders:
        for name in (f"NMOR%20{y}_{mth:02d}.pdf", f"NMOR%20{y}_{mth:02d}%20rev1.pdf", f"NDOR%20{y}_{mth:02d}_15.pdf"):
            if name.startswith("NDOR") and f != "Daily_op_reports":
                continue
            r = get(f"{M}/uploads/shares/{f}/{name}")
            if r is not None and r.content[:4] == b"%PDF":
                print("FOUND", f, name, len(r.content))
print("constructed-URL scan done")
# main-site style categories on the developer copy of nea.org.np, if any
for host in ("https://neasite.dryicesolutions.net", "https://nea.neasite.dryicesolutions.net",
             "https://genrd.neasite.dryicesolutions.net", "https://gd.neasite.dryicesolutions.net"):
    r = get(host + "/en")
    print("HOST", host, r.status_code if r is not None else None, len(r.text) if r is not None else 0)
