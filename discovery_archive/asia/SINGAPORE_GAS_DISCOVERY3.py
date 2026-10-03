"""
Singapore gas discovery, round 3: the Gas Network Code (SP Group 'Gas Works' resources, a zip of PDFs). Lists the
zip and prints every line on what the Gas Transporter (PowerGas) publishes or makes available, to settle whether
any metered flow data (receipt points, offtake / meter points, city gas) is public or shipper-only.
"""
import io
import re
import zipfile

import pdfplumber
import requests

URL = ("https://www.spgroup.com.sg/dam/jcr:2572ad30-3703-4a48-ae67-e5931a13d66c/"
       "%5BInfo%5D%20Gas%20Network%20Code%20-%20Jul%2026.zip")
H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0 Safari/537.36"}
PAT = re.compile(r"publish|website|bulletin|information exchange|made available to|make available|"
                 r"daily (gas|flow|quantit)|linepack|metered quantit|meter(ing)? (point|station)|receipt point|"
                 r"offtake point|confidential", re.I)


def main():
    r = requests.get(URL, headers=H, timeout=(20, 180))
    print(f"GET -> {r.status_code} {len(r.content)}", flush=True)
    z = zipfile.ZipFile(io.BytesIO(r.content))
    for n in z.namelist():
        print("zip:", n, flush=True)
    for n in z.namelist():
        if not n.lower().endswith(".pdf"):
            continue
        with pdfplumber.open(io.BytesIO(z.read(n))) as pdf:
            print(f"\n##### {n}: {len(pdf.pages)} pages", flush=True)
            for i, page in enumerate(pdf.pages):
                for line in (page.extract_text() or "").split("\n"):
                    if PAT.search(line):
                        print(f"  p{i + 1}: {line[:200]}", flush=True)


if __name__ == "__main__":
    main()
