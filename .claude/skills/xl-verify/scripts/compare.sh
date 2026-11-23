#!/usr/bin/env bash
# Run one flow with two xl launchers and compare everything they produce.
# Usage: compare.sh <flow.sh> <base-launcher> <branch-launcher> <out-dir> [runs]
#
# A flow is a bash script that reads $XL (the launcher), $IN (a read-only fixture dir it may copy
# from) and $OUT (an empty dir it writes into), and drives xl the way a user would: CLI commands
# only, every stdout it cares about saved under $OUT. compare.sh runs it once per launcher (plus
# runs-1 extra timed repeats), then:
#   - unpacks every .xlsx under $OUT into <name>.xlsx.d/ with each XML part pretty-printed,
#   - diffs base vs branch (out-dir/diff.txt) and lists changed files (out-dir/changed.txt),
#   - records exit codes and wall times per run (out-dir/times.tsv).
# Text outputs have the run's own $OUT path replaced by <OUT>, so only real differences remain.
# It never decides pass or fail: the checker reads the diff against what the change claims.
set -euo pipefail
[ $# -ge 4 ] || { echo "usage: $0 <flow.sh> <base-launcher> <branch-launcher> <out-dir> [runs]" >&2; exit 2; }
FLOW=$(cd "$(dirname "$1")" && pwd)/$(basename "$1"); BASE=$2; BRANCH=$3; RUNS=${5:-1}
mkdir -p "$4"; DIR=$(cd "$4" && pwd)
IN=${IN:-$(dirname "$FLOW")/fixtures}
: > "$DIR/times.tsv"
pretty() { if command -v xmllint >/dev/null; then xmllint --format "$1" 2>/dev/null || cat "$1"; else cat "$1"; fi; }
for side in base branch; do
  L=$BASE; [ "$side" = branch ] && L=$BRANCH
  for i in $(seq 1 "$RUNS"); do
    O="$DIR/$side"; [ "$i" -gt 1 ] && O="$DIR/$side.run$i"
    rm -rf "$O"; mkdir -p "$O"
    start=$(python3 -c 'import time;print(time.time())')
    set +e; XL="$L" IN="$IN" OUT="$O" bash "$FLOW" >"$O/_flow.stdout" 2>"$O/_flow.stderr"; rc=$?; set -e
    end=$(python3 -c 'import time;print(time.time())')
    printf '%s\t%s\t%s\t%.3f\n' "$side" "$i" "$rc" "$(python3 -c "print($end-$start)")" >> "$DIR/times.tsv"
    # paths differ per side by construction; replace them so the diff shows only real differences
    find "$O" -type f ! -name '*.xlsx' -exec perl -pi -e "s{\Q$O\E}{<OUT>}g" {} +
  done
  find "$DIR/$side" -name '*.xlsx' -type f | while read -r x; do
    d="$x.d"; rm -rf "$d"; mkdir -p "$d"; (cd "$d" && unzip -qo "$x")
    find "$d" -type f \( -name '*.xml' -o -name '*.rels' \) | while read -r f; do pretty "$f" > "$f.tmp" && mv "$f.tmp" "$f"; done
  done
done
(cd "$DIR" && diff -ru base branch > diff.txt || true)
(cd "$DIR" && diff -rq base branch 2>/dev/null | sed 's/^/  /' > changed.txt || true)
echo "times (side run exit seconds):"; cat "$DIR/times.tsv"
echo "changed files: $(wc -l < "$DIR/changed.txt" | tr -d ' ')   diff lines: $(wc -l < "$DIR/diff.txt" | tr -d ' ')"
