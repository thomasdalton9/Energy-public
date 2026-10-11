"""
Probe (Oct 2026): does curl_cffi (browser TLS impersonation, which got past Cloudflare on Pakistan's OGRA site) get
past the Cloudflare challenge on gov.il, where the Ministry of Energy publishes Israel's natural gas statistics?
For each page / API: status, title, size, whether it is a challenge page, GUIDs (office ids) and file links
(BlobFolder, .xlsx/.xls/.pdf/.csv) found. EGYPT_ISRAEL_GAS_POWER_PROBE.py found plain requests and Playwright
blocked; files under gov.il/BlobFolder download without a challenge.

RUN 1 (2026-10-11): BLOCKED. Every page and API (ministry/topic landing pages, natural_gas_economy, collectors,
PublicationApi) answers 403 'Just a moment...' / 'Attention Required! | Cloudflare' with impersonate chrome,
chrome124, safari17_0 and edge101 (safari got one 200, itself a challenge stub). BlobFolder files still download
(control PDF 200, 2.2 MB). So gov.il listings need a real browser outside GitHub Actions; a pull can only fetch
BlobFolder file URLs supplied to it.

Workflow (manual): discovery_archive/workflows/israel_govil_curl_cffi_probe.yml.
"""
import json
import re

from curl_cffi import requests

PAGES = [
    "https://www.gov.il/he/departments/ministry_of_energy/govil-landing-page",
    "https://www.gov.il/en/departments/ministry_of_energy/govil-landing-page",
    "https://www.gov.il/he/departments/topics/natural_gas_sector/govil-landing-page",
    "https://www.gov.il/en/pages/natural_gas_economy",
    "https://www.gov.il/he/pages/natural_gas_economy",
    "https://www.gov.il/he/departments/natural_gas_authority/govil-landing-page",
    "https://www.gov.il/he/collectors/publications?officeId=ministry_of_energy",
    "https://www.gov.il/he/api/PublicationApi/Index?limit=20&OfficeId=ministry_of_energy",
    "https://www.gov.il/he/api/PublicationApi/Index?limit=20&skip=0&Topic=natural_gas_sector",
]
FILE_RE = re.compile(r'(?:https?://www\.gov\.il)?/BlobFolder/[^"\'\s<>]+|https?://[^"\'\s<>]+\.(?:xlsx?|csv|pdf)', re.I)
GUID_RE = re.compile(r"[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}", re.I)


def show(url, imp):
    try:
        r = requests.get(url, impersonate=imp, timeout=60)
    except Exception as e:  # noqa: BLE001
        print(f"  [{imp}] {url}: {type(e).__name__} {str(e)[:120]}")
        return None
    t = r.text
    title = re.search(r"<title>(.*?)</title>", t, re.S)
    challenge = any(k in t for k in ("cf-chl", "Just a moment", "challenge-platform", "you have been blocked", "Attention Required"))
    print(f"  [{imp}] {url}: HTTP {r.status_code}, {len(t):,} chars, {r.headers.get('content-type')}, "
          f"title {title.group(1).strip()[:80] if title else None!r}, challenge={challenge}")
    if r.status_code == 200 and not challenge:
        files = sorted(set(FILE_RE.findall(t)))
        guids = sorted(set(GUID_RE.findall(t)))
        print(f"      files ({len(files)}): {files[:25]}")
        print(f"      guids ({len(guids)}): {guids[:10]}")
        if "json" in str(r.headers.get("content-type")):
            print("      json head:", t[:1500])
        return r
    return None


ok = None
for imp in ["chrome", "chrome124", "safari17_0", "edge101"]:
    print(f"\n##### impersonate={imp}")
    for u in PAGES:
        got = show(u, imp)
        ok = ok or got
    if ok:
        break

if ok:
    print("\n##### follow-up: CollectorsWebApi with office ids seen on the pages")
    guids = sorted(set(GUID_RE.findall(ok.text)))
    for g in guids[:6]:
        for ct in ("reports", "publications", "news"):
            show(f"https://www.gov.il/CollectorsWebApi/api/DataCollector/GetResults?CollectorType={ct}&officeId={g}&skip=0&limit=20",
                 "chrome")
print("\n##### BlobFolder control (known file)")
show("https://www.gov.il/BlobFolder/generalpage/freedom_of_information_energy/he/dov_196_2025.pdf", "chrome")
