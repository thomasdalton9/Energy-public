"""
One-off probe (prints only) for an Asia-Pacific industrial coal-to-LNG switching study.
Round 1-3: IEA 403; India coal ministry 404; EGEDA balance form (OTYPE 9) -> ./php/rev_newbalance2/balance.php.
Round 4: EGEDA codes: fE coking 001001002, other bituminous 001001001, sub-bit 001001003, anthracite 001002000,
lignite 001003000, natural gas 005022000; fS industry 000020000 + subsectors 000020003 (iron & steel) ...
000020122, autoproducers 000024000; fC China 005, Japan 008, Korea 009, Taipei 016, Thailand 017, Vietnam 021,
Malaysia 010, Philippines 014, Indonesia 007, Singapore 015, HK 006. A plain POST to balance.php returned HTTP 500.
UNdata SDMX DF_UNData_EnergyBalance: dims REF_AREA.COMMODITY.TRANSACTION.UNIT; coal only as B00_CL (no coking split);
industry subsectors B27_1211 ... B39_1214o.
Round 5: EGEDA page JavaScript (how OkSubmit builds the request); UNdata data query for India; DF_UNDATA_ENERGY codes.
"""
import re
import requests

H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0"}
EG = "https://www.egeda.ewg.apec.org/egeda/database/"
s = requests.Session(); s.headers.update(H)
r = s.post(EG + "rev_newbalance_select_cond2.php", data={"OTYPE": "9"}, timeout=60)
for sc in re.findall(r"<script[^>]*>(.*?)</script>", r.text, re.S | re.I):
    print("----- script -----\n" + sc[:3500])
print("----- form tag + hidden inputs -----")
print(re.findall(r"<FORM[^>]*>", r.text, re.I))
print([i for i in re.findall(r"<INPUT[^>]*>", r.text, re.I) if "HIDDEN" in i.upper() or "SUBMIT" in i.upper()])

print("\n===== UNdata energy balance query (India 356) =====")
B = "https://data.un.org/legacy/ws/rest/"
for url, acc in [(B + "data/DF_UNData_EnergyBalance/356.B00_CL+B04_NG../ALL/?startPeriod=2018&endPeriod=2023", "application/vnd.sdmx.data+csv"),
                 (B + "data/DF_UNData_EnergyBalance/356...?startPeriod=2021&endPeriod=2021", "application/vnd.sdmx.data+csv")]:
    rr = s.get(url, headers={"Accept": acc}, timeout=120)
    print(rr.status_code, len(rr.content), rr.headers.get("content-type"), url)
    print(rr.text[:1500])

print("\n===== DF_UNDATA_ENERGY structure =====")
rr = s.get(B + "dataflow/all/DF_UNDATA_ENERGY/latest?references=all", timeout=120)
x = rr.text
print(rr.status_code, len(rr.content))
for dim in re.findall(r'<structure:Dimension [^>]*id="([^"]+)"', x):
    print("  dimension", dim)
for cl_id, body in re.findall(r'<structure:Codelist [^>]*id="([^"]+)"[^>]*>(.*?)</structure:Codelist>', x, re.S):
    codes = re.findall(r'<structure:Code id="([^"]+)"[^>]*>\s*<common:Name[^>]*>([^<]+)', body)
    if "AREA" in cl_id:
        print(f"  codelist {cl_id}: {len(codes)}"); continue
    flt = [c for c in codes if re.search(r"coal|coking|bitum|lignite|natural gas|iron|industr|chemical|mineral|paper|food|textile|autoprod|manufact", c[1], re.I)]
    print(f"  codelist {cl_id}: {len(codes)} codes; relevant: {flt[:90]}")
