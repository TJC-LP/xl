# Installing scala-cli

Read this page only when `scala-cli --version` fails. A deployment that pre-installs scala-cli
(and pre-caches the xl dependency tree) documents that in `reference/LOCAL.md` next to the
skill; when that file exists, nothing here applies.

Check: `which scala-cli || echo "not installed"`

| Platform | Command |
|----------|---------|
| macOS | `brew install Virtuslab/scala-cli/scala-cli` |
| Linux | Use the verified Linux installation below (no Coursier prerequisite) |
| Windows | `winget install virtuslab.scalacli` |

No JDK prerequisite: scala-cli provisions a JVM through coursier on first use. The first run of
a script downloads the Scala compiler and the `com.tjclp::xl` dependency tree (30-60 s); later
runs are cached.

## Verified Linux installation

The version and SHA-256 values are pinned from the official v1.9.1 release asset
metadata (checked November 6, 2026). No JDK or Coursier is needed to install.

```bash
(
set -euo pipefail
case "$(uname -m)" in
  x86_64|amd64) ARCH=x86_64; EXPECTED=dd69825dd712e365bfed624e75f2de79479c3980f302fe93f32790d526a8a50d ;;
  aarch64|arm64) ARCH=aarch64; EXPECTED=bf78b11d82d1eb31883c8bbe9af4539f7ad1f1503aa65569b25bad51c4f4f8d3 ;;
  *) echo "Unsupported Linux architecture" >&2; exit 1 ;;
esac
INSTALL_DIR="${SCALA_CLI_INSTALL_DIR:-$HOME/.local/bin}"
mkdir -p "$INSTALL_DIR"
WORK=$(mktemp -d "$INSTALL_DIR/.scala-cli-download.XXXXXX")
trap 'rm -rf "$WORK"' EXIT
curl -fLsS "https://github.com/VirtusLab/scala-cli/releases/download/v1.9.1/scala-cli-$ARCH-pc-linux.gz" -o "$WORK/scala-cli.gz"
if command -v sha256sum >/dev/null 2>&1; then
  ACTUAL=$(sha256sum "$WORK/scala-cli.gz")
else
  ACTUAL=$(shasum -a 256 "$WORK/scala-cli.gz")
fi
[[ ${ACTUAL%% *} == "$EXPECTED" ]] || { echo "Scala CLI checksum mismatch" >&2; exit 1; }
gzip -dc "$WORK/scala-cli.gz" > "$WORK/scala-cli"
chmod +x "$WORK/scala-cli"
mv "$WORK/scala-cli" "$INSTALL_DIR/scala-cli"
echo "Installed verified Scala CLI 1.9.1 to $INSTALL_DIR/scala-cli"
)
```

Ensure `~/.local/bin` is on PATH. Update the version and both digests together.

## Offline or egress-restricted hosts

Warm the cache once while the network is available, then run offline:

```bash
printf '//> using scala 3.9.0\n//> using dep com.tjclp::xl:0.26.0\nimport com.tjclp.xl.scripting.{*, given}\nval _ = Sheet("warm").put(ref"A1", 1)\n' > /tmp/warm.sc
scala-cli compile --server=false /tmp/warm.sc          # downloads everything the pins need
scala-cli --power run --server=false --offline script.sc   # no network from here on
```

The pins in the warm script must equal the pins in every script you run; any other pair misses
the cache. `--server=false` skips the Bloop build server, which is the right choice for one-shot
scripts on an ephemeral host.

## Which Scala version

The published xl artifacts are compiled with the Scala version the skeleton in
[SKILL.md](../SKILL.md) pins. An older compiler cannot read newer TASTy, so keep
`//> using scala` at that version or later.
