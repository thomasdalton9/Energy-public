"""
One-off probe (prints only) for an Asia-Pacific industrial coal-to-LNG switching study.
Round 1-2: IEA 403; India coal ministry pages 404; EGEDA balance form posts OTYPE to rev_newbalance_select_cond2.php.
Round 3: EGEDA OTYPE 9 page posts to ./php/rev_newbalance2/balance.php with Y1/Y2 (1980-2023), fE[] products (coking
coal, other bituminous, sub-bituminous, ...), fS[] flows (95: transformation, industry subsectors ...), fC[] economies
(001-021), U unit, HEAD=Y. UNdata SDMX has DF_UNData_EnergyBalance and DF_UNDATA_ENERGY.
Round 4: full fE/fS/fC code lists; one trial balance.php query; UNdata energy-balance structure (dimensions, codes).
"""
import re
import requests

H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0"}
EG = "https://www.egeda.ewg.apec.org/egeda/database/"
s = requests.Session(); s.headers.update(H)
r = s.post(EG + "rev_newbalance_select_cond2.php", data={"OTYPE": "9"}, timeout=60)
html = r.text
for name in ("fE[]", "fS[]"):
    body = re.search(r'<select[^>]*name="%s"[^>]*>(.*?)</select>' % re.escape(name), html, re.S | re.I).group(1)
    opts = re.findall(r'<option[^>]*value="([^"]*)"[^>]*>\s*([^<]*)', body, re.I)
    print(f"\n{name} ({len(opts)}):")
    for v, t in opts:
        print(f"  {v} {t.strip()}")
print("\nfC[] economies:")
for v, label in re.findall(r'NAME="fC\[\]" VALUE="([^"]+)"[^>]*>\s*([^<]{2,40})', html, re.I):
    print(f"  {v} {label.strip()}")
print("\nunit radios:", re.findall(r'NAME="U" VALUE="([^"]+)"[^>]*>\s*([^<]{2,30})', html, re.I))

# trial query: all economies? use first economy only, coal + gas totals, all flows, 2019
data = [("Y1", "2019"), ("Y2", "2019"), ("U", "001"), ("HEAD", "Y"), ("fC[]", "001"),
        ("fE[]", "001000000"), ("fE[]", "001001002"), ("fS[]", "allallall")]
r2 = s.post(EG + "php/rev_newbalance2/balance.php", data=data, timeout=120)
txt = re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " | ", r2.text))
print(f"\nTRIAL balance.php: {r2.status_code} {len(r2.content):,} B ctype {r2.headers.get('content-type')}")
print(txt[:3000])
print("links/forms in result:", re.findall(r'href="([^"]+)"', r2.text)[:10], re.findall(r'<form[^>]*>', r2.text, re.I)[:3])

print("\n===== UNdata energy balance structure =====")
r3 = s.get("https://data.un.org/legacy/ws/rest/dataflow/all/DF_UNData_EnergyBalance/latest?references=all", timeout=90)
print(r3.status_code, len(r3.content))
x = r3.text
for dim in re.findall(r'<structure:Dimension [^>]*id="([^"]+)"', x):
    print("  dimension", dim)
for cl_id, body in re.findall(r'<structure:Codelist [^>]*id="([^"]+)"[^>]*>(.*?)</structure:Codelist>', x, re.S):
    codes = re.findall(r'<structure:Code id="([^"]+)"[^>]*>\s*<common:Name[^>]*>([^<]+)', body)
    print(f"  codelist {cl_id}: {len(codes)} e.g. {codes[:60] if 'COMMOD' in cl_id.upper() or 'TRANS' in cl_id.upper() else codes[:12]}")
