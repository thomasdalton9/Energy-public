"""Dispatch every scheduled workflow once, by hand (run_all_once.yml).

Order: the pulls in the order of their first scheduled time of day, in waves; then, once the pulls have
finished (or MAX_WAIT_MIN has passed), the master rebuilds (*_master.yml); then the private-repo sync.
Each pull already retries its own push, so waves only keep the push queue short.
"""
import glob, json, os, subprocess, sys, time
import yaml

WF_DIR = ".github/workflows"
SELF = "run_all_once.yml"
SYNC = "sync_to_private.yml"
WAVE = int(os.environ.get("WAVE", "12"))
GAP_S = int(os.environ.get("GAP_S", "240"))
MAX_WAIT_MIN = int(os.environ.get("MAX_WAIT_MIN", "200"))
RUN_SYNC = os.environ.get("RUN_SYNC", "true") == "true"


def scheduled():
    rows = []
    for path in glob.glob(f"{WF_DIR}/*.yml"):
        name = os.path.basename(path)
        if name == SELF:
            continue
        d = yaml.safe_load(open(path))
        on = d.get(True) or d.get("on") or {}
        if not isinstance(on, dict) or "schedule" not in on or "workflow_dispatch" not in on:
            continue
        minute, hour = on["schedule"][0]["cron"].split()[:2]
        try:
            t = int(hour) * 60 + int(minute)
        except ValueError:  # '*/n' style: run early
            t = 0
        rows.append((t, name))
    return [n for _, n in sorted(rows)]


def gh(*args):
    return subprocess.run(["gh", *args], capture_output=True, text=True)


def dispatch(names):
    for n in names:
        r = gh("workflow", "run", n, "--ref", "main")
        print(("started  " if r.returncode == 0 else "FAILED   ") + n + ("" if r.returncode == 0 else ": " + r.stderr.strip()), flush=True)


def title(name):
    d = yaml.safe_load(open(f"{WF_DIR}/{name}"))
    return d.get("name") or name


def active(names, since):
    r = gh("run", "list", "--event", "workflow_dispatch", "--created", f">={since}", "--limit", "1000",
           "--json", "status,workflowName")
    if r.returncode != 0:
        print("gh run list failed: " + r.stderr.strip(), flush=True)
        return [None]  # unknown: keep waiting
    titles = {title(n) for n in names}
    return [x for x in json.loads(r.stdout or "[]") if x["status"] != "completed" and x["workflowName"] in titles]


def wait_for(names, since, limit_min):
    end = time.time() + limit_min * 60
    while time.time() < end:
        left = active(names, since)
        print(f"{time.strftime('%H:%M')} UTC: {len(left)} still running or queued", flush=True)
        if not left:
            return
        time.sleep(120)
    print(f"Stopped waiting after {limit_min} min; continuing.", flush=True)


def main():
    since = time.strftime("%Y-%m-%dT%H:%M:%SZ", time.gmtime())
    names = scheduled()
    masters = [n for n in names if n.endswith("_master.yml")]
    pulls = [n for n in names if n not in masters and n != SYNC]
    print(f"{len(pulls)} pulls, {len(masters)} masters, sync={RUN_SYNC}", flush=True)
    for i in range(0, len(pulls), WAVE):
        dispatch(pulls[i:i + WAVE])
        if i + WAVE < len(pulls):
            time.sleep(GAP_S)
    wait_for(pulls, since, MAX_WAIT_MIN)
    dispatch(masters)
    wait_for(masters, since, 60)
    if RUN_SYNC:
        dispatch([SYNC])


if __name__ == "__main__":
    sys.exit(main())
