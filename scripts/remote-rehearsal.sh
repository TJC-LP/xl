#!/usr/bin/env bash
# Rehearse a Claude Code cloud sandbox locally, in Docker.
#
# Reproduces what Anthropic's hosted environment gives a session: Ubuntu 24.04 with OpenJDK 21 on
# PATH, no scala-cli, no coursier, no cached toolchain, and the committed tree of this checkout.
# Then proves, in order:
#   1. `./mill` self-provisions the JDK named in .mill-jvm-version (no setup needed for the build);
#   2. scripts/remote-setup.sh, run the way the SessionStart hook runs it, puts java 25 and
#      scala-cli on PATH through CLAUDE_ENV_FILE;
#   3. a module compiles and one suite runs on that toolchain (or everything, with --full).
#
# Only committed files take part (git archive HEAD), so commit before rehearsing.
# The real sandbox is x86_64; pass --platform linux/amd64 to emulate it on Apple silicon (slower).
#
# Usage: scripts/remote-rehearsal.sh [--platform linux/amd64] [--full]
set -euo pipefail
cd "$(dirname "$0")/.."

PLATFORM=()
FULL=false
while [[ $# -gt 0 ]]; do
  case "$1" in
    --platform) PLATFORM=(--platform "$2"); shift 2 ;;
    --full) FULL=true; shift ;;
    *) echo "unknown argument: $1" >&2; exit 2 ;;
  esac
done

command -v docker > /dev/null || { echo "docker is required" >&2; exit 2; }

WORK="$(mktemp -d)"
trap 'rm -rf "$WORK"' EXIT
git archive --format=tar HEAD > "$WORK/src.tar"
if [[ -n "$(git status --porcelain -- .mill-jvm-version .mill-version build.mill scripts/remote-setup.sh .claude/settings.json)" ]]; then
  echo "note: uncommitted changes to toolchain files are NOT part of the rehearsal (git archive HEAD)" >&2
fi

docker run --rm "${PLATFORM[@]}" -v "$WORK":/in:ro -e FULL="$FULL" ubuntu:24.04 bash -s <<'IN'
set -euo pipefail
export DEBIAN_FRONTEND=noninteractive
apt-get update -qq > /dev/null
apt-get install -y -qq --no-install-recommends openjdk-21-jdk-headless curl ca-certificates git gzip > /dev/null
mkdir -p /work && tar -xf /in/src.tar -C /work && cd /work

echo "== sandbox baseline: $(java -version 2>&1 | head -1)"

echo "== 1. ./mill self-provisions its JDK (.mill-jvm-version = $(cat .mill-jvm-version))"
./mill --no-daemon --version 2>&1 | grep -E 'Mill|java.home|java.version' || ./mill --no-daemon --version

echo "== 2. SessionStart hook: scripts/remote-setup.sh with CLAUDE_CODE_REMOTE=true"
export CLAUDE_CODE_REMOTE=true CLAUDE_PROJECT_DIR=/work CLAUDE_ENV_FILE=/tmp/claude-env
time bash scripts/remote-setup.sh
echo "-- CLAUDE_ENV_FILE --"; cat /tmp/claude-env
# shellcheck disable=SC1091
source /tmp/claude-env
echo "-- after hook: $(java -version 2>&1 | head -1)"
echo "-- scala-cli: $(scala-cli version 2>/dev/null | head -1 || echo MISSING)"

echo "== 3. build on that toolchain"
if [[ "$FULL" == "true" ]]; then
  time ./mill --no-daemon __.compile
  time ./mill --no-daemon -i __.test
else
  time ./mill --no-daemon xl-core.compile
  ./mill --no-daemon xl-core.test.testOnly com.tjclp.xl.addressing.ColumnSpec
fi
echo "== rehearsal OK"
IN
