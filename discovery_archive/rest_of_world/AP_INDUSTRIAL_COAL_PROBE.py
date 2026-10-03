"""
One-off probe (prints only) for an Asia-Pacific industrial coal-to-LNG switching study: where can coal use by
INDUSTRY SUBSECTOR and by COAL TYPE (coking vs thermal) be downloaded?
Round 1: IEA 403; India coal.nic.in / coal.gov.in pages 404.
Round 2: EGEDA balance form posts OTYPE (1,2,3,7,8,9) to ../database/rev_newbalance_select_cond2.php; UNdata has an
SDMX web service at data.un.org/legacy/ws/rest (dataflow list).
Round 3: post each OTYPE to the EGEDA condition page and print its form; list UNdata SDMX dataflows (id + name).
"""
import re
import requests

H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0"}
EG = "https://www.egeda.ewg.apec.org/egeda/database/"
s = requests.Session(); s.headers.update(H)


def show_forms(html):
    for f in re.findall(r"<form.*?</form>", html, re.S | re.I):
        print("  FORM", re.findall(r'<form[^>]*>', f, re.I)[0][:200])
        for name, body in re.findall(r'<select[^>]*name="([^"]+)"[^>]*>(.*?)</select>', f, re.S | re.I):
            opts = re.findall(r'<option[^>]*value="([^"]*)"[^>]*>\s*([^<]*)', body, re.I)
            print(f"    select {name}: {len(opts)} options e.g. {[(v, t.strip()[:30]) for v, t in opts[:40]]}")
        for inp in re.findall(r'<input[^>]*>', f, re.I)[:25]:
            print("    input", inp[:180])


print("=" * 20, "EGEDA condition pages", "=" * 20)
for ot in ("1", "2", "3", "7", "8", "9"):
    r = s.post(EG + "rev_newbalance_select_cond2.php", data={"OTYPE": ot}, timeout=60)
    text = re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", r.text))
    print(f"\nOTYPE {ot}: {r.status_code} {len(r.content):,} B :: {text[:300]}")
    show_forms(r.text)

print("\n" + "=" * 20, "UNdata SDMX dataflows", "=" * 20)
r = s.get("https://data.un.org/legacy/ws/rest/dataflow/all/all/latest", timeout=60)
for fid, name in re.findall(r'<structure:Dataflow[^>]*id="([^"]+)".*?<common:Name[^>]*>([^<]+)</common:Name>', r.text, re.S):
    print(" ", fid, "|", name)
