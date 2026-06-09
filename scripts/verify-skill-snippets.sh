#!/usr/bin/env bash
# Extract every fenced ```scala block that begins with `//> using` from the xl-scripting skill
# and compile it with scala-cli. Blocks without the directive header are fragments and skipped.
#
# Usage:
#   scripts/verify-skill-snippets.sh           # compile against the pinned Maven version
#   scripts/verify-skill-snippets.sh --local   # publishLocal first, substitute the local version
set -euo pipefail
cd "$(dirname "$0")/.."

LOCAL=false
[[ "${1:-}" == "--local" ]] && LOCAL=true

SKILL_DIR="plugin/skills/xl-scripting"
TMP=$(mktemp -d)
trap 'rm -rf "$TMP"' EXIT

for md in "$SKILL_DIR/SKILL.md" "$SKILL_DIR/reference/RECIPES.md" "$SKILL_DIR/reference/API.md"; do
  [[ -f "$md" ]] || continue
  awk -v outdir="$TMP" -v src="$(basename "$md" .md)" '
    BEGIN { inblock = 0; n = 0 }
    /^```scala$/ { inblock = 1; buf = ""; next }
    /^```$/ {
      if (inblock) {
        inblock = 0
        if (buf ~ /^\/\/> using/) {
          n++
          file = outdir "/" src "-" n ".sc"
          printf "%s", buf > file
          close(file)
        }
      }
      next
    }
    { if (inblock) buf = buf $0 "\n" }
  ' "$md"
done

COUNT=$(ls "$TMP"/*.sc 2>/dev/null | wc -l | tr -d ' ')
if [[ "$COUNT" == "0" ]]; then
  echo "✗ No compilable snippets found — extraction is broken" >&2
  exit 1
fi
echo "→ Extracted $COUNT complete-script snippet(s)"

if $LOCAL; then
  BUILD_VERSION=$(grep -oE '"PUBLISH_VERSION", "[^"]+"' build.mill | sed -E 's/.*"PUBLISH_VERSION", "([^"]+)".*/\1/')
  echo "→ Publishing local build ($BUILD_VERSION) and substituting the pin for verification..."
  ./mill __.publishLocal > /dev/null
  export COURSIER_REPOSITORIES="ivy2Local|central"
  sed -i.bak -E "s|(//> using dep com.tjclp::xl:)[0-9A-Za-z.+-]+|\\1$BUILD_VERSION|" "$TMP"/*.sc
  rm -f "$TMP"/*.bak
fi

FAILED=()
for sc in "$TMP"/*.sc; do
  name=$(basename "$sc")
  if scala-cli compile "$sc" > "$TMP/$name.log" 2>&1; then
    echo "✓ $name"
  else
    echo "✗ $name" >&2
    sed 's/^/    /' "$TMP/$name.log" | grep -iE "error" -A4 | head -20 >&2
    FAILED+=("$name")
  fi
done

if (( ${#FAILED[@]} > 0 )); then
  echo "" >&2
  echo "✗ ${#FAILED[@]} skill snippet(s) failed to compile" >&2
  exit 1
fi
echo "✓ All skill snippets compile"
