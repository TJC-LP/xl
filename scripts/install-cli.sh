#!/usr/bin/env bash
set -euo pipefail
VERSION="__VERSION__"
case "$VERSION" in *[!0-9A-Za-z.+-]*|'') echo "Invalid release version" >&2; exit 1 ;; esac
case "$(uname -s)-$(uname -m)" in
  Linux-x86_64) PLATFORM=linux-amd64 ;;
  Linux-aarch64) PLATFORM=linux-arm64 ;;
  Darwin-x86_64) PLATFORM=darwin-amd64 ;;
  Darwin-arm64) PLATFORM=darwin-arm64 ;;
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
