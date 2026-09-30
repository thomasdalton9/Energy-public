import subprocess
import sys
from pathlib import Path
import time

BASE_DIR = Path(__file__).resolve().parent

# Every South America pull, then south_america_dashboard.py, which puts
# all their charts into one workbook: South_America_Dashboard.xlsx.
#   Brazil     - "ONS Brazil.py" (daily + monthly generation, cached),
#                BRAZIL_ONS_HYDRO_BY_REGION.py (reservoir storage)
#   Colombia   - COLOMBIA_XM_HYDRO.py (reservoir storage),
#                COLOMBIA_XM_GENERATION.py (generation by fuel, cached)
#   Argentina  - argentina_generation_mix.py (only fetches new days)
#   Uruguay    - URUGUAY_ADME.py (generation by source, cached)
#   Ecuador    - ECUADOR_CENACE.py (generation by source; one day per run,
#                CENACE keeps no archive - run daily to build history)
#   Chile      - CHILE_CEN_HYDRO.py (needs CEN_USER_KEY in api_keys.py -
#                free at https://sipub.coordinador.cl/)
#
# A pull that fails doesn't stop the rest: the dashboard is still built
# from whatever is on disk (that chart is skipped or shows the last
# successful pull), and the failures are listed at the end.
#
# COLOMBIA_XM_GENERATION_DISCOVERY.py is deliberately NOT in this list
# - it's a one-off, run-it-yourself script for capturing real API
# responses, not a regular pipeline step.
SCRIPTS = [
    BASE_DIR / "ONS Brazil.py",
    BASE_DIR / "BRAZIL_ONS_HYDRO_BY_REGION.py",
    BASE_DIR / "COLOMBIA_XM_HYDRO.py",
    BASE_DIR / "COLOMBIA_XM_GENERATION.py",
    BASE_DIR / "argentina_generation_mix.py",
    BASE_DIR / "URUGUAY_ADME.py",
    BASE_DIR / "ECUADOR_CENACE.py",
    BASE_DIR / "CHILE_CEN_HYDRO.py",
    BASE_DIR / "south_america_dashboard.py",
]


def run_script(script_path):
    script_name = script_path.name
    print()
    print("=" * 80)
    print(f"RUNNING: {script_name}")
    print("=" * 80)

    if not script_path.is_file():
        raise RuntimeError(
            f"{script_path} does not exist - "
            f"is your local checkout up to date? (git pull)"
        )

    start = time.perf_counter()

    result = subprocess.run(
        [sys.executable, str(script_path)],
        cwd=str(script_path.parent),
    )

    if result.returncode != 0:
        raise RuntimeError(
            f"{script_name} failed "
            f"(exit code {result.returncode})"
        )

    elapsed = time.perf_counter() - start

    print(
        f"SUCCESS: {script_name} "
        f"({elapsed:.1f} seconds)"
    )


def main():
    print()
    print("=" * 80)
    print("SOUTH AMERICA DASHBOARD")
    print("=" * 80)

    failed = []
    for script in SCRIPTS:
        try:
            run_script(script)
        except RuntimeError as e:
            print(f"FAILED: {e}")
            failed.append(script.name)

    print()
    print("=" * 80)
    print("PIPELINE COMPLETE" + (f" - {len(failed)} failed: {', '.join(failed)}" if failed else ""))
    print("=" * 80)
    if "south_america_dashboard.py" in failed:
        sys.exit(1)


if __name__ == "__main__":
    main()
