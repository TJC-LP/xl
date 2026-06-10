# xl Claude Code Plugin

LLM-friendly Excel tooling for [Claude Code](https://claude.ai/code): read, write, style, and
analyze `.xlsx` files via the `xl` CLI, or script complex transformations against the xl Scala
library. The plugin packages two complementary skills.

## Skills

### xl-cli — stateless CLI operations

Drives the `xl` command-line binary: view ranges, read cells, search, evaluate formulas
(104 supported functions), export to CSV/JSON/PNG/PDF, style cells, edit rows/columns, and apply
atomic batch operations. Every invocation reads the file, applies one change set, and writes the
result — no session state. The skill auto-detects the latest released binary from the GitHub API,
so it never needs a version bump. See [`skills/xl-cli/SKILL.md`](skills/xl-cli/SKILL.md).

### xl-scripting — type-safe Scala scripting

Writes scala-cli scripts against the `com.tjclp::xl` library (one-import prelude,
compile-time-validated cell references, total APIs, whole-workbook recalculation with per-cell
error reporting). Suited to bulk or conditional transformations, multi-file pipelines, typed data
extraction, formula-heavy model building, and streaming 100k+ row files in constant memory. Pins
the library version; every documented snippet is compile-verified in CI. See
[`skills/xl-scripting/SKILL.md`](skills/xl-scripting/SKILL.md).

### Which one?

One-shot operations (quick reads, single edits, search, visual exports) → **xl-cli**. Anything
with loops, intermediate computation, or many dependent edits that would round-trip the file
repeatedly → **xl-scripting**. They compose well: generate with a script, verify visually with
the CLI.

## Installation

In Claude Code:

```
/plugin marketplace add TJC-LP/xl
/plugin install xl@xl-marketplace
```

## Release artifacts

Each [GitHub release](https://github.com/TJC-LP/xl/releases) also attaches the skills as
standalone zips for non-marketplace installs: `xl-skill-<version>.zip` (xl-cli) and
`xl-scripting-skill-<version>.zip` (xl-scripting), alongside the native `xl` binaries.
