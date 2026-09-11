# Installing scala-cli

Read this page only when `scala-cli --version` fails. A deployment that pre-installs scala-cli
(and pre-caches the xl dependency tree) documents that in `reference/LOCAL.md` next to the
skill; when that file exists, nothing here applies.

Check: `which scala-cli || echo "not installed"`

| Platform | Command |
|----------|---------|
| macOS | `brew install Virtuslab/scala-cli/scala-cli` |
| Linux | `curl -sSLf https://scala-cli.virtuslab.org/get \| sh` |
| Windows | `winget install virtuslab.scalacli` |

No JDK prerequisite: scala-cli provisions a JVM through coursier on first use. The first run of
a script downloads the Scala compiler and the `com.tjclp::xl` dependency tree (30-60 s); later
runs are cached.

## Offline or egress-restricted hosts

Warm the cache once while the network is available, then run offline:

```bash
printf '//> using scala 3.9.0\n//> using dep com.tjclp::xl:0.22.0\nimport com.tjclp.xl.scripting.{*, given}\nval _ = Sheet("warm").put(ref"A1", 1)\n' > /tmp/warm.sc
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
