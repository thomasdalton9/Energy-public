"""
Follow-up to GIIGNL_DISCOVERY.py: downloads GIIGNL's latest annual
report PDF and its "major flows" PDF, and checks whether they contain a
genuine contracted (long-term) vs spot LNG volume breakdown - and
whether the PDF has real extractable text (vs. being scanned images,
which would need OCR instead of a plain text/table extractor).
"""

import sys

import requests

URLS = {
    "annual_report_2026": "https://giignl-documents.s3.fr-par.scw.cloud/public/ar-2026-annual-report.pdf",
    "major_flows_2026": "https://giignl-documents.s3.fr-par.scw.cloud/public/ar-2026-major-flows.pdf",
}
HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                   "(KHTML, like Gecko) Chrome/124.0 Safari/537.36",
}
TIMEOUT = (10, 60)


def download(url, out_path):
    r = requests.get(url, headers=HEADERS, timeout=TIMEOUT)
    r.raise_for_status()
    with open(out_path, "wb") as f:
        f.write(r.content)
    return len(r.content)


def inspect_pdf(path, label):
    import pdfplumber

    print(f"\n{'=' * 70}\n{label}: {path}\n{'=' * 70}")
    with pdfplumber.open(path) as pdf:
        print(f"  {len(pdf.pages)} pages")
        total_chars = 0
        keyword_pages = []
        for i, page in enumerate(pdf.pages):
            text = page.extract_text() or ""
            total_chars += len(text)
            lower = text.lower()
            if any(kw in lower for kw in ["spot", "short-term", "long-term", "sma contract", "shipper"]):
                keyword_pages.append(i)
        print(f"  total extracted text: {total_chars:,} chars "
              f"({'looks like real text' if total_chars > 500 * len(pdf.pages) else 'suspiciously little - may be scanned/image-based'})")
        print(f"  pages mentioning spot/long-term/contract keywords: {keyword_pages[:20]}")

        for i in keyword_pages[:5]:
            page = pdf.pages[i]
            text = page.extract_text() or ""
            print(f"\n  --- page {i} text preview (first 800 chars) ---")
            print("  " + text[:800].replace("\n", "\n  "))

            tables = page.extract_tables()
            print(f"  --- page {i} has {len(tables)} table(s) ---")
            for t in tables[:2]:
                for row in t[:6]:
                    print(f"    {row}")


def main():
    for label, url in URLS.items():
        out_path = f"/tmp/{label}.pdf"
        print(f"Downloading {url} ...", file=sys.stderr)
        size = download(url, out_path)
        print(f"  {size:,} bytes -> {out_path}", file=sys.stderr)
        inspect_pdf(out_path, label)


if __name__ == "__main__":
    main()
