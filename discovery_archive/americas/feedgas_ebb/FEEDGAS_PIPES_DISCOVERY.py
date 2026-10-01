"""
Discovery script for the pipeline EBBs (Electronic Bulletin Boards)
feeding US LNG export terminals - started out scoped to Permian/
Haynesville-fed Gulf Coast terminals only, now widened (per request)
to also cover Cove Point and Elba Island, the two operating terminals
that DON'T draw from those two basins (Appalachian Marcellus/Utica gas
instead) - NOT a real data pull, and not part of any pipeline.

TERMINAL -> RECEIVING PIPELINE(S), confirmed via current (2025-2026)
industry reporting, not guessed - the reference map this script's
scope is built from:
  - Sabine Pass LNG: Trunkline (Creole Trail Interconnect at Gillis
    Hub - the most direct Haynesville route), Transco (Gillis West
    interconnect).
  - Corpus Christi LNG: Gulf Coast Express/GCX (Kinder Morgan, Waha->
    Agua Dulce) and Whistler (Enbridge/WhiteWater, same route), then
    ADCC (Cheniere/WhiteWater JV) for the final Agua Dulce->terminal
    leg - Whistler/ADCC's own EBB URLs not found yet, see below.
  - Cameron LNG: Cameron Interstate Pipeline (Sempra) - a DIFFERENT
    EBB platform (gasnom.com) - sourced from Gillis Hub, so also
    Haynesville; Cameron's own system map additionally names Gulf
    South, LEAP, Acadian, and Williams' LEG (Louisiana Energy Gateway)
    as interconnecting pipelines there.
  - Freeport LNG: THREE interstate meters - Gulf South Stratton Ridge,
    Texas Eastern (TETCO) BIG Pipeline, and Texas Eastern (TETCO)
    Stratton Ridge - so Texas Eastern covers 2 of 3; Gulf South
    (Boardwalk Pipelines) is a separate, not-yet-found platform.
  - Calcasieu Pass LNG: Venture Global's own TransCameron Pipeline
    (likely intrastate - see caveat below), interconnecting with
    Texas Eastern's East Lateral, plus Trunkline via Gillis Hub.
  - Plaquemines LNG: Venture Global's own Gator Express Pipeline
    (likely intrastate), interconnecting with Tennessee Gas Pipeline
    and Texas Eastern Transmission.
  - Golden Pass LNG: BOTH basins directly - up to 1 Bcf/d of Permian
    gas via Kinder Morgan's Trident pipeline from the Katy hub, and
    Haynesville/Louisiana gas via a 20-year NGPL firm transport deal
    (340,000 Dth/d) - NGPL's system was itself expanded to carry more
    Eagle Ford/Haynesville/Permian gas for this.
  - Cove Point LNG (Dominion/Berkshire, Lusby, MD) - NOT Permian/
    Haynesville: an 88-mile bi-directional pipeline connects it to
    Transco (Pleasant Valley Interconnect, Fairfax County VA),
    Columbia Gas Transmission/TCO, and DETI (both in Loudoun County
    VA) - feedgas is Marcellus/Utica (Appalachian) sourced.
  - Elba Island LNG (Kinder Morgan/Southern LNG, near Savannah, GA) -
    also NOT Permian/Haynesville: the ~200-mile bidirectional Elba
    Express Pipeline (Kinder Morgan, code EEC) connects it to Transco
    in Hart County GA / Anderson County SC.

NOT YET IN SCOPE - real, Permian/Haynesville-relevant, but newer
terminals still starting up (2026-2028) rather than established flow
histories, so lower priority for a first pass:
  - Rio Grande LNG (NextDecade, Brownsville - targeting H2 2026
    startup): fed by the Rio Bravo Pipeline (138 mi, 4.5 Bcf/d), itself
    fed by Bay Runner Twin (Agua Dulce -> Rio Grande, 2.6 Bcf/d) - so
    on the SAME Agua Dulce/Permian supply chain as Corpus Christi LNG.
  - Port Arthur LNG (Sempra, Phase 1 - commercial startup targeted
    Dec 2027): fed by a new dedicated Port Arthur Pipeline system
    (2.0 Bcf/d, in service June 2026) plus a Louisiana Connector
    (2.0 Bcf/d, H2 2026). EBB CONFIRMED - Port Arthur Pipeline, LLC
    is on gasnom.com, the same platform as Cameron Interstate - now
    in TARGETS below.
  - CP2 LNG (Venture Global, next to Calcasieu Pass - FID March 2026,
    first production targeted late 2028): fed by CP Express Pipeline
    (2.2 Bcf/d, Venture Global's own lateral - likely intrastate, see
    caveat below), itself fed from the Blackfin Pipeline (a 190-mile,
    48-inch INTRASTATE pipeline, confirmed - a Venture Global/
    WhiteWater 50/50 JV) carrying Permian gas off the Matterhorn
    Express pipeline. Both legs are intrastate, so likely have no
    mandatory public EBB at all - not a research gap, a structural one.
  - Delfin LNG (floating/offshore - FID'd): takes capacity on Kinder
    Morgan's Texas-Louisiana Expansion (the same NGPL system expansion
    that also serves Golden Pass) - so likely covered by the NGPL
    target already in this script.

Two more new (2026-2027) Permian/Haynesville trunk pipelines - checked
for an EBB, genuinely don't have one to find yet because neither is in
commercial service:
  - Blackcomb Pipeline (WhiteWater 50.6% / MPLX 30.4% / Enbridge 19.0%
    JV, WhiteWater-operated - same operator as Whistler): Waha -> Agua
    Dulce, 2.5 Bcf/d, commissioning started late July 2026, ramping
    through 2H 2026 - a THIRD major Permian-to-Agua-Dulce artery
    alongside GCX and Whistler. Once it's fully in commercial service
    it may share a platform with Whistler (same operator) - worth
    rechecking once Whistler's own EBB is found.
  - Pelican Pipeline (WhiteWater + FIC/Stonepeak/Trace Capital):
    northern Louisiana Haynesville -> Gillis Hub, 1.75 Bcf/d (not
    2.5 - corrected from an earlier pass), 170 miles - targeting
    service in 1H 2027, so no EBB posting exists to find yet.
  - Rio Bravo Pipeline (Enbridge, NextDecade retains capacity rights):
    feeds Rio Grande LNG from the Agua Dulce area, 4.5 Bcf/d - still
    under construction (Enbridge just took over the build after
    Rio Grande LNG's FID); likely to end up on Enbridge's LINK
    platform once operational, given Enbridge already runs Texas
    Eastern there, but nothing to find yet.

Intrastate-pipeline caveat: TransCameron and Gator Express are Venture
Global's own laterals, and FERC's NAESB/EBB open-access posting mandate
applies to INTERSTATE pipelines - a purely intrastate line may not have
a public EBB at all. Confirming this either way is exactly what running
this script (or checking those two specifically) would tell us.

Platforms actually checked here (six vendor platforms cover most of
the pipelines above, once the platform's structure is understood the
same approach likely extends to other pipelines on it via a different
`code=`/path parameter):
  - Williams "1Line" (1line.williams.com) - Transco.
  - Energy Transfer "Messenger+" ({code}messenger.energytransfer.com)
    - Trunkline/TGC.
  - Kinder Morgan (pipeline2.kindermorgan.com) - GCX (Permian side),
    NGPL (Golden Pass's Haynesville/Louisiana side), and Elba Express
    (code EEC, feeds Elba Island LNG) - same platform, different
    `code=` each time.
  - Enbridge "LINK" (infopost.enbridge.com / rtba.enbridge.com) -
    Texas Eastern/TETCO (Freeport, Calcasieu Pass, Plaquemines all
    touch this one).
  - gasnom.com - Cameron Interstate Pipeline (Cameron LNG) - a fifth,
    different platform.
  - TC Energy "eConnects" (ebb.tceconnects.com) - Columbia Gas
    Transmission/TCO, one of Cove Point LNG's three interconnects - a
    sixth platform.

Also checked: piperiv.com/ip/transco - a THIRD-PARTY aggregator site
(PipeRiv) that appears to mirror/link informational postings per
pipeline in a possibly cleaner format than going straight to each
operator's own site - worth checking whether that's actually a better
single source than six separate vendor platforms.

Not yet found: EBB URLs for Whistler Pipeline, ADCC, Gulf South
Pipeline (Boardwalk), Tennessee Gas Pipeline's own code on the Kinder
Morgan platform, TransCameron, Gator Express, and DETI (Dominion
Energy Transmission, Cove Point's third interconnect) - natural next
discovery passes.

Each pipeline operator runs its EBB on different platform software,
so this checks each one separately and reports what it actually finds
- HTTP status, whether the page looks like a real server-rendered page
or a JS-only single-page-app shell (which would need Playwright, like
india_daily_generation.py's Grid-India CDN, rather than plain
requests), and any visible links/keywords pointing at actual flow/
scheduled-quantity data (as opposed to just a portal homepage).

Not runnable from the environment that wrote this - every one of these
domains is unreachable here (confirmed via web-fetch tool blocks and a
direct proxy check across every pipeline site tried this session,
including api.eia.gov - this sandbox blocks essentially all outbound
domains except a small allowlist). Two ways to actually run it:
  1. By hand, from the repo root: python3 americas/feedgas_ebb/FEEDGAS_PIPES_DISCOVERY.py
     - then share the console output (or the saved HTML files under
     americas/feedgas_ebb/feedgas_discovery_html/) back so real
     per-pipeline pull scripts can be written against the actual page
     structure.
  2. Via .github/workflows/feedgas_ebb_discovery.yml (workflow_dispatch,
     manually triggered from the Actions tab or the API) - runs on
     GitHub's own runners, which aren't behind this sandbox's network
     block, and commits the saved HTML + a run log back into this same
     americas/feedgas_ebb/ subfolder.

Same discovery-first pattern as COLOMBIA_XM_GENERATION_DISCOVERY.py,
PAKISTAN_OPENDATA_DISCOVERY.py, and TRANSCO_Z6_NY_DISCOVERY.py.
"""

print("STARTING", flush=True)

import os
import re

import requests

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
    )
}

# Anchored to this script's own folder (not the caller's cwd) so output
# always lands inside americas/feedgas_ebb/, whether this is run by hand
# from the repo root or by the GitHub Actions workflow in
# .github/workflows/feedgas_ebb_discovery.yml - this subfolder is meant
# to stay self-contained and not spill files into the rest of the repo.
OUTPUT_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "feedgas_discovery_html")

# (label, url) - the starting point for each operator's EBB, not
# necessarily the final flow-data page itself.
TARGETS = [
    ("Transco (Williams) - site map", "https://www.1line.williams.com/Transco/site-map.html"),
    ("Trunkline Gas / TGC (Energy Transfer) - info postings", "https://tgcmessenger.energytransfer.com/ipost/TGC/customer-activities/information"),
    ("Gulf Coast Express / GCX (Kinder Morgan) - notices", "https://pipeline2.kindermorgan.com/Notices/Notices.aspx?type=P&code=GCX"),
    ("Gulf Coast Express / GCX (Kinder Morgan) - point catalog", "https://pipeline2.kindermorgan.com/PortalWeb/PointCatalog.aspx?code=GCX"),
    ("NGPL (Kinder Morgan) - notices, feeds Golden Pass", "https://pipeline2.kindermorgan.com/Notices/Notices.aspx?type=P&code=NGPL"),
    ("Texas Eastern / TE (Enbridge) - LINK infopost home", "https://infopost.enbridge.com/infopost/TEHome.asp?Pipe=TE"),
    ("Texas Eastern / TE (Enbridge) - operationally available", "https://rtba.enbridge.com/InformationalPosting/Default.aspx?bu=TE&Type=OA"),
    ("Cameron Interstate Pipeline (gasnom.com) - info postings, feeds Cameron LNG", "https://www.gasnom.com/ip/CAMERON/"),
    ("Cameron Interstate Pipeline - system map", "https://www.gasnom.com/ip/cameron/map/"),
    ("Elba Express / EEC (Kinder Morgan) - feeds Elba Island LNG", "https://pipeline2.kindermorgan.com/default.aspx?code=EEC"),
    ("Columbia Gas Transmission / TCO (TC Energy) - infopost home, one of Cove Point's interconnects", "https://ebb.tceconnects.com/infopost/"),
    ("Port Arthur Pipeline (gasnom.com) - feeds Port Arthur LNG", "https://www.gasnom.com/ip/portarthurpipeline/"),
    ("Transco (PipeRiv third-party aggregator)", "https://www.piperiv.com/ip/transco"),
]

DATA_KEYWORDS = [
    "scheduled quantity", "scheduled quantities", "nomination", "flow", "receipt", "delivery",
    "download", ".csv", ".xls", ".xlsx", "operationally available", "capacity",
]

JS_APP_MARKERS = [
    "you need to enable javascript", "please enable javascript", "id=\"root\"", "id=\"app\"",
    "ng-app", "data-reactroot", "__next",
]


def fetch(url):
    try:
        r = requests.get(url, headers=HEADERS, timeout=30, verify=False)
        return r
    except requests.RequestException as e:
        print(f"  FAILED: {type(e).__name__}: {e}", flush=True)
        return None


def analyse(label, url, html):
    lower = html.lower()

    js_app_signals = [m for m in JS_APP_MARKERS if m in lower]
    if js_app_signals:
        print(f"  Looks like a JS-rendered single-page app (markers: {js_app_signals}) - plain requests may not see the real data; Playwright may be needed.", flush=True)
    else:
        print("  Looks like a server-rendered page (no obvious SPA-shell markers).", flush=True)

    found_keywords = sorted({k for k in DATA_KEYWORDS if k in lower})
    print(f"  Data-related keywords present: {found_keywords or 'none'}", flush=True)

    links = re.findall(r'href=[\'"]([^\'" >]+)', html, flags=re.IGNORECASE)
    interesting_links = [
        l for l in links
        if any(k.replace(" ", "") in l.lower().replace("-", "").replace("_", "") for k in
               ["schedule", "nomination", "flow", "download", "csv", "xls", "capacity", "operational"])
    ]
    if interesting_links:
        print(f"  Interesting link(s) found ({len(interesting_links)}, showing up to 15):", flush=True)
        for l in interesting_links[:15]:
            print(f"    {l}", flush=True)


def main():
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    for label, url in TARGETS:
        print(f"\n{'=' * 70}\n{label}\n{url}\n{'=' * 70}", flush=True)
        r = fetch(url)
        if r is None:
            continue

        print(f"  status: {r.status_code}, content-length: {len(r.content):,} bytes", flush=True)

        if r.status_code != 200:
            print(f"  Non-200 response - first 500 chars: {r.text[:500]!r}", flush=True)
            continue

        safe_name = re.sub(r"[^a-zA-Z0-9]+", "_", label)[:60] + ".html"
        out_path = os.path.join(OUTPUT_DIR, safe_name)
        with open(out_path, "w", encoding="utf-8") as f:
            f.write(r.text)
        print(f"  Saved -> {out_path}", flush=True)

        analyse(label, url, r.text)

    print(
        "\nDONE. Please share the console output above (or the saved HTML files in "
        f"{OUTPUT_DIR}/) so real per-pipeline pull scripts can be written against the "
        "actual page structure - especially wherever it found a real scheduled-quantity/"
        "flow data link rather than just a portal homepage.",
        flush=True,
    )


if __name__ == "__main__":
    main()
