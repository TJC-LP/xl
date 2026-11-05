# Installing the `xl` binary

Read this page only when `xl --version` fails or prints a version below the one
[SKILL.md](../SKILL.md) requires. A deployment that pre-installs `xl` documents that in
`reference/LOCAL.md` next to the skill; when that file exists, nothing here applies.

Check: `which xl || echo "not installed"`; then `xl --version`.

## macOS / Linux / Windows Bash

The latest published native binary; no JDK required.

```bash
set -euo pipefail
LATEST=$(curl -fLsS https://api.github.com/repos/TJC-LP/xl/releases/latest | grep '"tag_name"' | cut -d'"' -f4)
VERSION=${LATEST#v}
case "$VERSION" in *[!0-9A-Za-z.+-]*|'') echo "Invalid release version" >&2; exit 1 ;; esac
case "$(uname -s)-$(uname -m)" in
  Linux-x86_64) PLATFORM=linux-amd64 ;;
  Linux-aarch64) PLATFORM=linux-arm64 ;;
  Darwin-x86_64) PLATFORM=darwin-amd64 ;;
  Darwin-arm64) PLATFORM=darwin-arm64 ;;
  MINGW*-*|MSYS*-*|CYGWIN*-*) PLATFORM=windows-amd64.exe ;;
  *) echo "Unsupported platform; use the verified JAR distribution" >&2; exit 1 ;;
esac
BINARY="xl-${VERSION}-${PLATFORM}"
INSTALL_DIR="${XL_INSTALL_DIR:-${HOME}/.local/bin}"
mkdir -p "$INSTALL_DIR"
WORK=$(mktemp -d "$INSTALL_DIR/.xl-download.XXXXXX")
trap 'rm -rf "$WORK"' EXIT
BASE="https://github.com/TJC-LP/xl/releases/download/v${VERSION}"
curl -fLsS "$BASE/$BINARY" -o "$WORK/$BINARY"
curl -fLsS "$BASE/SHA256SUMS" -o "$WORK/SHA256SUMS"
EXPECTED=$(awk -v name="$BINARY" '$2 == name { print $1; n++ } END { if (n != 1) exit 1 }' "$WORK/SHA256SUMS")
case "$EXPECTED" in *[!0-9a-f]*|'') echo "Invalid checksum manifest" >&2; exit 1 ;; esac
[[ ${#EXPECTED} == 64 ]] || { echo "Invalid checksum length" >&2; exit 1; }
if command -v sha256sum >/dev/null 2>&1; then
  ACTUAL=$(sha256sum "$WORK/$BINARY")
else
  ACTUAL=$(shasum -a 256 "$WORK/$BINARY")
fi
[[ ${ACTUAL%% *} == "$EXPECTED" ]] || { echo "Release checksum mismatch" >&2; exit 1; }
chmod +x "$WORK/$BINARY"
mv "$WORK/$BINARY" "$INSTALL_DIR/xl"
echo "Installed verified xl $VERSION to $INSTALL_DIR/xl"
```

The release must provide `SHA256SUMS`; older releases without it fail closed. The manifest and
binary are trusted through the same HTTPS release source; this detects changed bytes, not a
compromised publisher. Ensure `~/.local/bin` is on PATH.

## Windows (PowerShell)

Download the matching `.exe` and `SHA256SUMS` from the same release into a temporary directory.
Compare `(Get-FileHash -Algorithm SHA256 path/to/xl.exe).Hash` with the manifest entry for that exact
asset before moving it to your installation directory. Refuse a missing, duplicate, or mismatched
entry; do not execute an unverified file.

## JAR distribution

Every release also attaches `xl-cli-<version>.tar.gz`: the assembly JAR plus a wrapper script,
for platforms without a native binary. Verify the tarball against its exact `SHA256SUMS` entry
before extraction or running its installer. It needs a JRE (25 or later).

## Rasterizer for PNG / JPEG / PDF / WebP export

The native binary renders SVG itself but needs one external tool to rasterize it; `xl rasterizers`
lists what it found.

```bash
# macOS
brew install librsvg

# Linux (Debian/Ubuntu)
apt install librsvg2-bin

# Python alternative
pip install cairosvg
```
