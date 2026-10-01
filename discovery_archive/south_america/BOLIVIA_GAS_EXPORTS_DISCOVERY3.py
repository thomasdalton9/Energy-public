"""
Bolivia gas production and exports - round 3.

Round 2 found:
  - YPFB's 'Gas natural' page (https://www.ypfb.gob.bo/Gas_natural) carries
    'PRODUCCION FISCALIZADA DE GAS NATURAL - PROMEDIOS', 'VOLUMEN
    COMERCIALIZADO GAS NATURAL MERCADO INTERNO' and 'EXPORTACION GAS NATURAL
    A 60F POR PAIS (COMPARATIVA)' - find how the numbers are served.
  - MHE quarterly bulletin: exports by country, quarterly only (MMm3 60F).
  - INE export micro-data, one file per year (2021 ... Ene-Ago 2026p):
    look for natural gas (NANDINA 2711.21) rows and any physical quantity.
Not reachable from the editing sandbox; runs in GitHub Actions.
"""
import io
import re
import zipfile
from urllib.parse import urljoin

import pandas as pd
import requests

H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0 Safari/537.36"}
T = (10, 180)


def out(*a):
    print(*a, flush=True)


def get(url, **kw):
    try:
        r = requests.get(url, headers=H, timeout=T, **kw)
        out(f"GET {url} -> {r.status_code} {r.headers.get('content-type', '')[:50]} {len(r.content)}B")
        return r
    except Exception as e:
        out(f"GET {url} -> ERR {type(e).__name__}: {str(e)[:160]}")
        return None


out("=================== 1. YPFB Gas natural page")
r = get("https://www.ypfb.gob.bo/Gas_natural")
if r is not None and r.status_code == 200:
    html = r.text
    for tag in ("iframe", "img", "script", "embed", "object", "canvas"):
        for m in re.finditer(rf"<{tag}\b[^>]*>", html, re.I):
            s = m.group(0)
            if tag == "script" and "src=" not in s:
                continue
            out(f"   <{tag}> {s[:300]}")
    for m in re.finditer(r'(href|src|data-[a-z-]+)="([^"]+\.(?:pdf|xlsx?|csv|json|png|jpe?g|svg)[^"]*)"', html, re.I):
        out(f"   FILE {m.group(2)}")
    i = html.find("PRODUCCI")
    body = html[max(0, i - 3000): i + 20000]
    out("   --- raw HTML around the headings ---")
    for ln in re.sub(r"\n\s*\n+", "\n", body).splitlines():
        if ln.strip():
            out("   |", ln.strip()[:300])
    for m in re.finditer(r"<script\b[^>]*>(.*?)</script>", html, re.S | re.I):
        js = m.group(1)
        if re.search(r"chart|highcharts|data\s*:|series|labels", js, re.I):
            out("   --- inline script with chart data ---")
            out("   ", js[:6000])

out("\n=================== 2. INE export micro-data")
r = get("https://nube.ine.gob.bo/index.php/s/CKdL20Qa170R1Sg/download")
if r is not None and r.status_code == 200:
    try:
        for name, df in pd.read_excel(io.BytesIO(r.content), sheet_name=None, header=None).items():
            out(f"   dict sheet {name!r} {df.shape}")
            for row in df.head(40).itertuples(index=False):
                vals = [str(x)[:70] for x in row if str(x) != "nan"]
                if vals:
                    out("      ", vals)
    except Exception as e:
        out("   dictionary not xlsx:", e, r.content[:200])


def read_any(content):
    """Return a list of (name, DataFrame) from xlsx / csv / zip content."""
    if content[:2] == b"PK":
        try:
            return list(pd.read_excel(io.BytesIO(content), sheet_name=None).items())
        except Exception:
            z = zipfile.ZipFile(io.BytesIO(content))
            res = []
            for n in z.namelist():
                out(f"   zip member {n}")
                b = z.read(n)
                if n.lower().endswith((".xlsx", ".xls")):
                    res += list(pd.read_excel(io.BytesIO(b), sheet_name=None).items())
                elif n.lower().endswith((".csv", ".txt")):
                    for sep in (";", ",", "|", "\t"):
                        try:
                            df = pd.read_csv(io.BytesIO(b), sep=sep, encoding="latin-1", low_memory=False)
                            if df.shape[1] > 3:
                                res.append((n, df))
                                break
                        except Exception:
                            continue
            return res
    for sep in (";", ",", "|", "\t"):
        try:
            df = pd.read_csv(io.BytesIO(content), sep=sep, encoding="latin-1", low_memory=False)
            if df.shape[1] > 3:
                return [("csv", df)]
        except Exception:
            continue
    return []


for label, u in (("2026 Ene-Ago", "https://nube.ine.gob.bo/index.php/s/a8BfyyBxkysiXNG/download"),
                 ("2021", "https://nube.ine.gob.bo/index.php/s/82Noqtb9OyGFuqg/download")):
    r = get(u)
    if r is None or r.status_code != 200:
        continue
    out(f"   {label}: first bytes {r.content[:8]!r}")
    for name, df in read_any(r.content):
        out(f"   sheet {name!r} {df.shape}; columns {list(df.columns)}")
        out(df.head(3).to_string()[:1500])
        code_col = next((c for c in df.columns if re.search(r"nandina|arancel|partida|subpartida|codigo", str(c), re.I)), None)
        desc_col = next((c for c in df.columns if re.search(r"desc", str(c), re.I)), None)
        mask = pd.Series(False, index=df.index)
        if code_col is not None:
            mask |= df[code_col].astype(str).str.replace(r"\D", "", regex=True).str.startswith("271121")
        if desc_col is not None:
            mask |= df[desc_col].astype(str).str.contains(r"gas natural", case=False, regex=True)
        g = df[mask]
        out(f"   gas rows: {len(g)} (code col {code_col!r}, desc col {desc_col!r})")
        if len(g):
            out(g.head(15).to_string()[:5000])
            num = [c for c in g.columns if pd.api.types.is_numeric_dtype(g[c])]
            keys = [c for c in g.columns if re.search(r"mes|pais|país|destino", str(c), re.I)]
            out(f"   group keys {keys}; numeric {num}")
            if keys:
                out(g.groupby(keys)[num].sum().to_string()[:6000])
