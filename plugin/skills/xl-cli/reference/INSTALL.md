# Installing the `xl` binary

Read this page only when `xl --version` fails or prints a version below the one
[SKILL.md](../SKILL.md) requires. A deployment that pre-installs `xl` documents that in
`reference/LOCAL.md` next to the skill; when that file exists, nothing here applies.

Check: `which xl || echo "not installed"`; then `xl --version`.

## macOS / Linux / Windows Bash

The pinned v0.26.0 native binary; no JDK required. The script verifies it against the
exact entry for its asset in that release's `SHA256SUMS` and fails closed when the manifest
or entry is missing. The inline digests below are the legacy v0.25.0 fallback, which
predates release manifests; they come from the official GitHub release asset metadata
(checked November 6, 2026).

```bash
(
set -euo pipefail
VERSION=0.26.0
case "$VERSION" in *[!0-9A-Za-z.+-]*|'') echo "Invalid release version" >&2; exit 1 ;; esac
case "$(uname -s)-$(uname -m)" in
  Linux-x86_64) PLATFORM=linux-amd64; EXPECTED=b213be3184a37c936736a0a25719bf8fa9efa2b3c2539082ce5a87ed71319a53 ;;
  Linux-aarch64) PLATFORM=linux-arm64; EXPECTED=347abb99781d8dea77a14a0e4eee4bcf4d5098c4771fb6360a96736c643c5182 ;;
  Darwin-x86_64) PLATFORM=darwin-amd64; EXPECTED=90a0aeefdd6400eebd2e28d8efa29c873ffec3a8c13e761691f6f945363fbb42 ;;
  Darwin-arm64) PLATFORM=darwin-arm64; EXPECTED=0acdfd68526f2a4b730eb73e478270ffb066dce24663658072e91711973e8480 ;;
  MINGW*-*|MSYS*-*|CYGWIN*-*) PLATFORM=windows-amd64.exe; EXPECTED=3a1dae1db9e779bee7a6137fcd76ed8cec41f1c3b867c19d8014d002e1990e33 ;;
  *) echo "Unsupported platform; use the verified JAR distribution" >&2; exit 1 ;;
esac
BINARY="xl-${VERSION}-${PLATFORM}"
INSTALL_DIR="${XL_INSTALL_DIR:-${HOME}/.local/bin}"
mkdir -p "$INSTALL_DIR"
WORK=$(mktemp -d "$INSTALL_DIR/.xl-download.XXXXXX")
trap 'rm -rf "$WORK"' EXIT
BASE="https://github.com/TJC-LP/xl/releases/download/v${VERSION}"
# v0.25.0 predates release manifests; its exact digests above are the only fallback.
if [[ "$VERSION" != "0.25.0" ]]; then
  curl -fLsS "$BASE/SHA256SUMS" -o "$WORK/SHA256SUMS"
  EXPECTED=$(awk -v asset="$BINARY" '$2 == asset { count++; digest=$1 } END { if (count != 1) exit 1; print digest }' "$WORK/SHA256SUMS")
  [[ "$EXPECTED" =~ ^[0-9a-f]{64}$ ]] || { echo "Invalid release checksum" >&2; exit 1; }
fi
curl -fLsS "$BASE/$BINARY" -o "$WORK/$BINARY"
if command -v sha256sum >/dev/null 2>&1; then
  ACTUAL=$(sha256sum "$WORK/$BINARY")
else
  ACTUAL=$(shasum -a 256 "$WORK/$BINARY")
fi
[[ ${ACTUAL%% *} == "$EXPECTED" ]] || { echo "Release checksum mismatch" >&2; exit 1; }
chmod +x "$WORK/$BINARY"
mv "$WORK/$BINARY" "$INSTALL_DIR/xl"
echo "Installed verified xl $VERSION to $INSTALL_DIR/xl"
)
```

The pinned digests detect changed bytes, not a compromised publisher. Keep version and
digests together for the legacy v0.25.0 fallback. Release preparation updates `VERSION`
to match the skill and plugin; subsequent versions require their own `SHA256SUMS`
with exactly one matching asset entry and fail closed when it is missing. Ensure `~/.local/bin` is on PATH.

## Windows (PowerShell)

For v0.25.0, download `xl-0.25.0-windows-amd64.exe` from the official release.
Compare `(Get-FileHash -Algorithm SHA256 path/to/xl.exe).Hash` with
`3a1dae1db9e779bee7a6137fcd76ed8cec41f1c3b867c19d8014d002e1990e33` before moving or executing it.
For later releases, require the exact asset entry in that release’s `SHA256SUMS`.

## JAR distribution

Every release also attaches `xl-cli-<version>.tar.gz`: the assembly JAR plus a wrapper script,
for platforms without a native binary. For v0.25.0, verify `xl-cli-0.25.0.tar.gz` against
`b94b9e539df29cce397e6acb65014686202fdd6066758dc7251bf48ce205fd10`
before extraction or running its installer. Later releases require the exact `SHA256SUMS`
entry. It needs a JRE (25 or later).

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
