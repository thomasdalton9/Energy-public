"""
Follow-up to RHINE_HISTORICAL_ARCHIVE_INSPECT.py, which confirmed the
historische-zeitreihen endpoint returns a real ZIP - but its content is
compressed binary, never actually decoded/printed. Extracting the CSV
inside and printing its real header/first rows before building the
production script's column-rename mapping on a guess.
"""
import io
import sys
import zipfile

import requests

PREPARE_URL = "https://www.pegelonline.wsv.de/gast/historische-zeitreihen/prepare-download"
HOST = "https://www.pegelonline.wsv.de"
KAUB_UUID = "1d26e504-7f9e-480a-b52c-5932be6549ab"
HEADERS = {"User-Agent": "Mozilla/5.0"}
TIMEOUT = (10, 60)

params = {
    "uuid": KAUB_UUID,
    "parameter": "W",
    "start": "2024-01-01T00:00:00.000Z",
    "end": "2024-01-10T23:59:59.000Z",
    "format": "csv",
}

r = requests.post(PREPARE_URL, headers=HEADERS, data=params, timeout=TIMEOUT, allow_redirects=False)
print(f"prepare status={r.status_code}", file=sys.stderr)
redirect_url = r.headers["Location"]
if redirect_url.startswith("/"):
    redirect_url = HOST + redirect_url
r2 = requests.get(redirect_url, headers=HEADERS, timeout=TIMEOUT)
print(f"download status={r2.status_code} bytes={len(r2.content)}", file=sys.stderr)

with zipfile.ZipFile(io.BytesIO(r2.content)) as zf:
    names = zf.namelist()
    print(f"zip names: {names}", file=sys.stderr)
    with zf.open(names[0]) as f:
        raw = f.read()
    for encoding in ["utf-8", "latin-1", "utf-16"]:
        try:
            text = raw.decode(encoding)
            print(f"\n=== decoded as {encoding} ===", file=sys.stderr)
            print(text[:1500], file=sys.stderr)
            break
        except UnicodeDecodeError as e:
            print(f"  {encoding} failed: {e}", file=sys.stderr)
