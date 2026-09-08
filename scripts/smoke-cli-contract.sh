#!/usr/bin/env bash
# Smoke the agent-visible contract of a PACKAGED xl — the assembly JAR in ci.yml, every native
# binary in release.yml, or a local install — the way an agent drives it: `--version`, `--help`,
# `sheets`, `view --json`, `schema --json`, `lint`, and two failures whose envelopes must carry the
# error code (GH-592). The golden corpus (GoldenSpec) proves the in-process harness; this proves the
# artifact an agent actually runs: assembly resource merging, native-image reflection, the version
# baked in at build time.
#
# Usage: scripts/smoke-cli-contract.sh <expected-version|-> <xl command...>
#   scripts/smoke-cli-contract.sh - java -jar out/xl-cli/assembly.dest/out.jar
#   scripts/smoke-cli-contract.sh 0.21.0 out/xl-cli/nativeImage.dest/native-executable
#   scripts/smoke-cli-contract.sh - xl        # whatever `make install` put on PATH
#
# `-` takes the version the binary itself prints — it must look like a release, X.Y.Z with an
# optional pre-release suffix — and holds every other place xl prints a version to it. release.yml
# passes the tag instead, so the binary must also agree with the release it ships in. Output is read
# CRLF-tolerant: the Windows binary ends its lines with \r\n. `::error::` lines become annotations
# on GitHub and are plain lines anywhere else.
set -euo pipefail

if [ $# -lt 2 ]; then
  echo "usage: $0 <expected-version|-> <xl command...>" >&2
  exit 2
fi
EXPECTED=$1; shift
XL=("$@")

RELEASE_VERSION='^[0-9]+\.[0-9]+\.[0-9]+([-+][0-9A-Za-z.+-]+)?$'
if [ "$EXPECTED" != "-" ] && ! [[ "$EXPECTED" =~ $RELEASE_VERSION ]]; then
  echo "::error::expected version '$EXPECTED' is not a release version (X.Y.Z[-suffix])"; exit 1
fi

ROOT=$(cd "$(dirname "$0")/.." && pwd)
FIXTURE="$ROOT/xl-ooxml/test/resources/fixtures/small-values.xlsx"
[ -f "$FIXTURE" ] || { echo "::error::fixture missing: $FIXTURE"; exit 1; }

ERR=$(mktemp)
trap 'rm -f "$ERR"' EXIT

# run <expected exit> <args...>: stdout in $OUT (CRs stripped), stderr in the $ERR file;
# any other exit code fails the smoke with both channels shown.
run() {
  local want=$1; shift
  set +e
  OUT=$("${XL[@]}" "$@" 2>"$ERR"); local status=$?
  set -e
  OUT=${OUT//$'\r'/}
  if [ "$status" -ne "$want" ]; then
    echo "::error::xl $* exited $status, expected $want"
    echo "--- stdout"; echo "$OUT"; echo "--- stderr"; cat "$ERR"
    exit 1
  fi
}
# check <what> <jq filter>: the filter must hold of $OUT; `$version` inside it is the expected version.
check() {
  local what=$1 filter=$2
  if ! echo "$OUT" | jq -e --arg version "$VERSION" "$filter" >/dev/null; then
    echo "::error::$what: jq '$filter' does not hold of:"; echo "$OUT"; exit 1
  fi
}

run 0 --version
if [ "$EXPECTED" = "-" ]; then
  VERSION=$OUT
  [[ "$VERSION" =~ $RELEASE_VERSION ]] || { echo "::error::--version printed '$VERSION', not a release version"; exit 1; }
else
  VERSION=$EXPECTED
  [ "$OUT" = "$VERSION" ] || { echo "::error::--version printed '$OUT'; expected $VERSION"; exit 1; }
fi

run 0 --help
[ -z "$OUT" ] || { echo "::error::--help wrote to stdout; the contract puts usage on stderr"; echo "$OUT"; exit 1; }
grep -q '^Usage:' "$ERR" || { echo "::error::--help did not print usage on stderr"; cat "$ERR"; exit 1; }

run 0 -f "$FIXTURE" sheets
grep -q 'Values' <<<"$OUT" || { echo "::error::sheets did not list the Values sheet"; echo "$OUT"; exit 1; }

run 0 -f "$FIXTURE" --json sheets
check "sheets --json" '.ok == true and .verb == "sheets" and .version == $version and .data[0].name == "Values" and .error == null'

run 0 -f "$FIXTURE" --json view A1:C4
check "view --json" '.ok == true and .verb == "view" and .data.sheet == "Values" and (.data.rows | length) == 4 and .data.rows[0].cells[1].value == 42'

run 0 --json schema
check "schema --json" '.ok == true and .verb == "schema" and .data.version == $version and (.data.verbs | length) > 0 and ([.data.errorCodes[].code] | index("SHEET_REQUIRED")) != null and (.data.warningCodes | index("TRUNCATED")) != null and (.data.exitCodes | map(.code)) == [0, 1, 2, 3]'

run 0 -f "$FIXTURE" lint
grep -q 'clean' <<<"$OUT" || { echo "::error::lint did not report the fixture clean"; echo "$OUT"; exit 1; }

run 0 -f "$FIXTURE" --json lint
check "lint --json" '.ok == true and .verb == "lint" and .data.clean == true'

# Failures are envelopes too: the code, never the message, is the contract.
run 3 -f does-not-exist.xlsx --json sheets
check "missing-file envelope" '.ok == false and .exitCode == 3 and .error.code == "IO_READ"'

run 2 -f "$FIXTURE" --json nonsense-verb
check "unknown-verb envelope" '.ok == false and .exitCode == 2 and .error.code == "UNKNOWN_VERB"'

echo "smoke: xl $VERSION (${XL[*]}) answered every verb as the contract says"
