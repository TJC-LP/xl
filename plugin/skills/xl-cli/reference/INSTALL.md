# Installing the `xl` binary

Read this page only when `xl --version` fails or prints a version below the one
[SKILL.md](../SKILL.md) requires. A deployment that pre-installs `xl` documents that in
`reference/LOCAL.md` next to the skill; when that file exists, nothing here applies.

Check: `which xl || echo "not installed"`; then `xl --version`.

## macOS / Linux (recommended)

The latest published native binary; no JDK required.

```bash
# Auto-detect platform and install latest release
REPO="TJC-LP/xl"
LATEST=$(curl -s "https://api.github.com/repos/$REPO/releases/latest" | grep '"tag_name"' | cut -d'"' -f4)
VERSION=${LATEST#v}
case "$(uname -s)-$(uname -m)" in
  Linux-x86_64)  BINARY="xl-$VERSION-linux-amd64" ;;
  Linux-aarch64) BINARY="xl-$VERSION-linux-arm64" ;;
  Darwin-x86_64) BINARY="xl-$VERSION-darwin-amd64" ;;
  Darwin-arm64)  BINARY="xl-$VERSION-darwin-arm64" ;;
  *) echo "Unsupported: $(uname -s)-$(uname -m)" && exit 1 ;;
esac
mkdir -p ~/.local/bin
curl -fsSL "https://github.com/$REPO/releases/download/$LATEST/$BINARY" -o ~/.local/bin/xl || {
  echo "Error: no $BINARY published for $LATEST — use the JAR distribution xl-cli-$VERSION.tar.gz instead" >&2
  exit 1
}
chmod +x ~/.local/bin/xl
echo "Installed xl $VERSION to ~/.local/bin/xl"
xl --version
```

Ensure `~/.local/bin` is in your PATH: `export PATH="$HOME/.local/bin:$PATH"`

## Alternative: GitHub CLI

```bash
# If gh is installed (simpler, handles auth for private repos)
gh release download --repo TJC-LP/xl --pattern "xl-*-$(uname -s | tr A-Z a-z)-$(uname -m | sed 's/x86_64/amd64/;s/aarch64/arm64/')" -D /tmp
mv /tmp/xl-* ~/.local/bin/xl && chmod +x ~/.local/bin/xl
```

## Windows (PowerShell)

```powershell
$repo = "TJC-LP/xl"
$latest = (Invoke-RestMethod "https://api.github.com/repos/$repo/releases/latest").tag_name
$version = $latest -replace '^v', ''
$url = "https://github.com/$repo/releases/download/$latest/xl-$version-windows-amd64.exe"
Invoke-WebRequest -Uri $url -OutFile "$env:LOCALAPPDATA\xl.exe"
Write-Host "Installed xl $version"
```

## JAR distribution

Every release also attaches `xl-cli-<version>.tar.gz`: the assembly JAR plus a wrapper script,
for platforms without a native binary. It needs a JRE (25 or later).

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
