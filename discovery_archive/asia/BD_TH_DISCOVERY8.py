"""
Sabah / Sarawak discovery, round 3: the MEIH statistics views fill their tables by JavaScript. Print the script
snippets that build the request (resource / action URLs, parameter names), then try that request for Sabah and
Sarawak installed capacity and for electricity generation, and print the result.
"""
import re

import requests
import urllib3

urllib3.disable_warnings()
H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0 Safari/537.36"}
T = (20, 90)


def out(*a):
    print(*a, flush=True)


def main():
    s = requests.Session()
    s.headers.update(H)
    r = s.get("https://meih.st.gov.my/statistics", timeout=T, verify=False)
    links = {re.sub(r"<[^>]+>|\s+", " ", t).strip(): h.replace("&amp;", "&")
             for h, t in re.findall(r'<a[^>]+href="([^"]+)"[^>]*>(.*?)</a>', r.text, re.S)}
    for name in ("Installed Capacity", "Electricity Generation"):
        p = s.get(links[name], timeout=T, verify=False)
        out(f"\n######## {name}: {p.status_code} {len(p.text)}")
        for m in re.finditer(r"(?:resourceURL|actionURL|renderURL|p_p_resource_id|p_p_lifecycle=2|\$\.ajax|ajax\(|"
                             r"\.submit\(|_eventId|url\s*:|window\.location|viewReport|generate)", p.text):
            out("  ctx: " + re.sub(r"\s+", " ", p.text[max(0, m.start() - 200): m.start() + 400]))
        urls = sorted(set(u.replace("&amp;", "&") for u in re.findall(r'["\'](https?://meih\.st\.gov\.my[^"\']+)["\']', p.text)))
        out(f"  urls ({len(urls)}):")
        for u in urls:
            if "Statistic" in u or "resource" in u or "lifecycle" in u:
                out("    " + u[:300])


if __name__ == "__main__":
    main()
