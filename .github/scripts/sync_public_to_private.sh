#!/usr/bin/env bash
# Copy every change made on Energy-public's main into one branch of the private Energy repo.
# Energy-public is the source of truth for whatever it changes: each file added or modified on public main since
# the last sync is copied over as public has it now, each file it deleted is deleted. Files public has not
# touched since the last sync (e.g. private-only files) are left alone. On the private development branch (anything
# but main) generated outputs stay out: a file that branch's .gitignore excludes and that it doesn't already track
# is skipped, as the dev branch keeps only source.
# The last synced public commit is kept in .public_sync_sha on the private branch.
# usage: sync_public_to_private.sh <public checkout> <private checkout> <private branch>
set -euo pipefail
PUB=$1; PRIV=$2; BR=$3
BASELINE=2a9da3d6ffb99159fce963a35ab655ef8616c807   # Energy-public's first commit = private main 952568d
cd "$PRIV"
git checkout -q "$BR"
last=$(cat .public_sync_sha 2>/dev/null || echo "$BASELINE")
head=$(git -C "$PUB" rev-parse HEAD)
if [ "$last" = "$head" ]; then echo "$BR: up to date at ${head:0:7}"; exit 0; fi
copied=0; removed=0
while IFS= read -r -d '' path; do
  [ "$path" = ".public_sync_sha" ] && continue
  if [ "$BR" != "main" ] && ! git ls-files --error-unmatch -- "$path" >/dev/null 2>&1 && git check-ignore -q -- "$path"; then
    continue   # generated output: not kept on the development branch
  fi
  if git -C "$PUB" cat-file -e "$head:$path" 2>/dev/null; then
    mkdir -p "$(dirname "$path")"
    git -C "$PUB" show "$head:$path" > "$path"
    [ "$(git -C "$PUB" ls-tree "$head" -- "$path" | cut -c1-6)" = "100755" ] && chmod +x "$path"
    git add -f -- "$path"; copied=$((copied+1))
  elif [ -e "$path" ]; then
    git rm -q -f -- "$path"; removed=$((removed+1))
  fi
done < <(git -C "$PUB" diff -z --name-only --no-renames "$last" "$head")
echo "$head" > .public_sync_sha
git add .public_sync_sha
if git diff --cached --quiet -- . ':(exclude).public_sync_sha'; then
  note="no file changes"
else
  note="$copied files updated, $removed removed"
fi
git commit -q -m "Sync from Energy-public ${last:0:7}..${head:0:7}: $note [skip ci]"
echo "$BR: synced ${last:0:7}..${head:0:7}: $note"
