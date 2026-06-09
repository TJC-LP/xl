#!/usr/bin/env bash
# Compile-verify every example script against the LOCAL build, and run a curated fast subset.
# This is the anti-rot guarantee for examples/*.sc and for the xl-scripting skill snippets
# derived from them (plugin/skills/xl-scripting/).
#
# Usage: scripts/test-examples.sh [--compile-only]
set -euo pipefail

cd "$(dirname "$0")/.."

COMPILE_ONLY=false
[[ "${1:-}" == "--compile-only" ]] && COMPILE_ONLY=true

# ---------- 1. Version drift guard ----------
# examples/project.scala must pin the same version build.mill falls back to, so publishLocal
# artifacts resolve. Both bump together at release-prep time.
PROJECT_VERSION=$(grep -oE 'com\.tjclp::xl:[0-9A-Za-z.+-]+' examples/project.scala | cut -d: -f4)
BUILD_VERSION=$(grep -oE '"PUBLISH_VERSION", "[^"]+"' build.mill | sed -E 's/.*"PUBLISH_VERSION", "([^"]+)".*/\1/')

if [[ -z "$PROJECT_VERSION" || -z "$BUILD_VERSION" ]]; then
  echo "✗ Could not extract versions (project.scala: '$PROJECT_VERSION', build.mill: '$BUILD_VERSION')" >&2
  exit 1
fi
if [[ "$PROJECT_VERSION" != "$BUILD_VERSION" ]]; then
  echo "✗ Version drift: examples/project.scala pins $PROJECT_VERSION but build.mill falls back to $BUILD_VERSION" >&2
  echo "  Bump both together (see .claude/commands/release-prep.md)" >&2
  exit 1
fi
echo "✓ Version drift guard: $PROJECT_VERSION"

# ---------- 2. Publish the local build ----------
echo "→ Publishing local build ($BUILD_VERSION) to ivy2Local..."
./mill __.publishLocal > /dev/null

# Resolve ivy2Local first so the just-published artifacts win over Maven Central.
export COURSIER_REPOSITORIES="ivy2Local|central"

# ---------- 3. Compile every example ----------
# Perf/large-data scripts are compile-verified but never run.
RUN_SET=(
  quick_start.sc
  demo.sc
  easy_mode_demo.sc
  scripting_tour.sc
  financial_model.sc
  data_validation.sc
)

FAILED=()
for sc in examples/*.sc; do
  name=$(basename "$sc")
  if scala-cli compile "$sc" > /dev/null 2>&1; then
    echo "✓ compile $name"
  else
    echo "✗ compile $name" >&2
    FAILED+=("compile:$name")
  fi
done

# ---------- 4. Run the curated fast subset ----------
# Per-example timeout: a hung example must fail loudly, never block CI. (`timeout` is coreutils;
# present on ubuntu runners and via brew coreutils locally — fall back to no timeout if absent.)
RUN_TIMEOUT=()
if command -v timeout > /dev/null 2>&1; then
  RUN_TIMEOUT=(timeout 300)
fi
if ! $COMPILE_ONLY; then
  for name in "${RUN_SET[@]}"; do
    if "${RUN_TIMEOUT[@]}" scala-cli run "examples/$name" > /dev/null 2>&1; then
      echo "✓ run     $name"
    else
      echo "✗ run     $name" >&2
      FAILED+=("run:$name")
    fi
  done
fi

# ---------- 5. Report ----------
if (( ${#FAILED[@]} > 0 )); then
  echo "" >&2
  echo "✗ ${#FAILED[@]} example check(s) failed:" >&2
  printf '  %s\n' "${FAILED[@]}" >&2
  exit 1
fi
echo ""
echo "✓ All example checks passed"
