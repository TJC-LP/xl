# xl Claude Code Plugin

LLM-friendly Excel tooling for [Claude Code](https://claude.ai/code): read, write, style, and
analyze `.xlsx` files via the `xl` CLI, or script complex transformations against the xl Scala
library. The plugin packages two complementary skills.

## Skills

### xl-cli — stateless CLI operations

Drives the `xl` command-line binary: view ranges, read cells, search, evaluate formulas
(108 supported functions), export to CSV/JSON/PNG/PDF, style cells, edit rows/columns, and apply
atomic batch operations. Every invocation reads the file, applies one change set, and writes the
result — no session state. The skill states the minimum `xl` version it documents; its
[install page](skills/xl-cli/reference/INSTALL.md) fetches the latest release. See
[`skills/xl-cli/SKILL.md`](skills/xl-cli/SKILL.md).

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

## Vendoring the skills into another repository

Each skill directory is a drop-in: copy `plugin/skills/<skill>/` from a release tag byte for byte
and re-copy it on every bump. Nothing in the tree is meant to be edited downstream.

- **Installation lives in `reference/INSTALL.md`**, not in `SKILL.md`, so a host that pre-installs
  `xl` or scala-cli pays no context for instructions it never runs.
- **Local facts go in `reference/LOCAL.md`**, a file the upstream tree never contains. Both
  `SKILL.md` files tell the agent to read it first when it exists: the binary's path, the
  pre-warmed scala-cli cache, output directories, house rules. Keep it and any other local
  reference pages out of the upstream file set, and a re-vendor is a plain replace of the
  upstream files.
- **Pin one version everywhere.** `SKILL.md` names the minimum `xl` version it documents
  (`Requires xl >= X.Y.Z`; CI checks it against `plugin.json`) and the scripting skill pins
  `com.tjclp::xl:X.Y.Z` and the Scala version the artifacts were compiled with. The binary a host
  installs, the library its scripts resolve and the skill text should come from the same tag.
