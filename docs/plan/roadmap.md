# XL Roadmap

> **Track Progress**: [GitHub Issues](https://github.com/TJC-LP/xl/issues)

**Last Updated**: 2026-06-10

> **Completed release records**: [archive/plan/v0.10.0-execution.md](../archive/plan/v0.10.0-execution.md) (0.10.0 tracker) and [archive/plan/v0.10.0-triage.md](../archive/plan/v0.10.0-triage.md) (rationale + per-issue verdicts).

---

## TL;DR

**Current Status**: Production-ready with **104 formula functions** (incl. dynamic arrays SEQUENCE/SORT/UNIQUE/FILTER and OFFSET), **structural editing** (insert/delete rows & columns with formula rewriting), the **scripting prelude** (`com.tjclp.xl.scripting`), whole-workbook `recalculate`, named-range & hyperlink authoring, SAX streaming (36% faster than POI), Excel tables, and full OOXML round-trip. 3005+ tests passing.

**Current Version**: **0.11.0 "Scripting"** (released 2026-06-10)

---

## Release Roadmap

### v0.12.0 candidates (not yet scheduled)

This roadmap is the single source of truth for scheduling; entries below are **candidates, not commitments** — nothing is started until it is pulled into a release plan.

| Candidate | Notes |
|-----------|-------|
| **"Visual" theme**: Chart model ([#222](https://github.com/TJC-LP/xl/issues/222)) on a DrawingML/image layer ([#221](https://github.com/TJC-LP/xl/issues/221)) | Deferred from 0.10.0/0.11.0; charts structurally depend on the drawing layer |
| 0.11.x follow-up: leading `=+` unary plus rejected by parser ([#271](https://github.com/TJC-LP/xl/issues/271)) | Banker idiom (`=+A1`); flagged as the top 0.11.1 candidate |
| 0.11.x follow-up: trailing empty format sections ([#262](https://github.com/TJC-LP/xl/issues/262)) | Number-format section selection edge case |
| 0.11.x follow-up: quoting for cell-ref-shaped sheet names ([#263](https://github.com/TJC-LP/xl/issues/263)) | e.g. a sheet literally named `A1` |
| 0.11.x follow-up: streaming StylePatcher indent + border-merge unification ([#264](https://github.com/TJC-LP/xl/issues/264)) | Parity between streaming and in-memory style edits |
| 0.11.x follow-up: SaxStax DirectSaxEmitter metadata gaps ([#265](https://github.com/TJC-LP/xl/issues/265)) | Streaming writer metadata parity |
| 0.11.x follow-up: even/first-page headers + `fitToPage` flag ([#266](https://github.com/TJC-LP/xl/issues/266)) | Completes the 0.11.0 `PageSetup` print extensions |
| Display strategy: `excel""` interpolator re-evaluates formulas sheet-locally instead of using the `recalculate` cache | Cross-sheet formulas can display stale/erroring values in scripts; unify display on the recalc cache |

### v0.11.0 "Scripting" (Released 2026-06-10)

Make library scripting (scala-cli + `com.tjclp.xl.scripting` prelude) the turbo-charged agent path — goal: the best functional Excel scripting DSL. Tracked in [#252](https://github.com/TJC-LP/xl/issues/252).

| Feature | Status |
|---------|--------|
| `com.tjclp.xl.scripting` prelude (one import; pure base import unchanged) | ✅ Done |
| Opaque-type extension fix for external consumers (toA1/shift/col/row...) + `xlprelude` probes | ✅ Done |
| Total DSL: `range := v` fill (Ctrl+Enter), ARef `down/up/right/left` | ✅ Done |
| `Workbook.upsert`, `readTypedOr`/`readTypedOpt`, `wb.evaluateFormula(formula, onSheet)` | ✅ Done |
| `Workbook.recalculate` — total, per-cell `CellEvalError`s, cycle isolation, cross-sheet fix | ✅ Done |
| `FormattedParsers.detect` promotion (CLI delegates) + prelude `String.toFormatted` | ✅ Done |
| xl-scripting skill (SKILL.md + API.md + 7 recipes) + release packaging & version gate | ✅ Done |
| Anti-rot CI: examples job + skill-verify workflow | ✅ Done |
| Future: typed row/record extraction (RowCodec derivation) | 🔵 Proposed |
| Future: bounds-checked `shift`/navigation variants | 🔵 Proposed |

### v0.10.0 "Trust & Author" (Released)

Build version **0.10.0**. Focus: trust (surgical-edit fidelity) and authoring (write the parts XL previously only read).

| Feature | Status |
|---------|--------|
| Named-range authoring (`DefinedName` serialization + CLI `name add/rm`) | ✅ Done |
| Hyperlink authoring (`Cell.hyperlink` serialization) | ✅ Done |
| Structural editing — insert/delete rows & columns with formula rewriting (cross-sheet, `#REF!` generation) | ✅ Done |
| Formula breadth — registry 88 → **104 functions** (IFS, SWITCH, CHOOSE, LARGE, SMALL, RANK, PERCENTILE, QUARTILE, HLOOKUP, MAXIFS, MINIFS, OFFSET, dynamic arrays SEQUENCE/SORT/UNIQUE/FILTER) | ✅ Done |
| Trust fixes C1–C5 — preserve inline worksheet elements (dataValidations, sheetProtection, autoFilter) through edits | ✅ Done |
| Theme color resolution (`ThemePalette.resolve`, `toResolvedArgb`/`toResolvedHex`) | ✅ Done |
| Configurable file size limits (CLI `--max-size`, `0` = unlimited) | ✅ Done |

### Shipped in prior releases

CLI expansion and evaluator/tooling fixes (v0.6.x):

| Feature | Status |
|---------|--------|
| `csv` import, `comment`, `clear`, `fill`, `sort` commands | ✅ Done |
| Batch `put` smart mode, `--auto-fit` flag | ✅ Done |
| `rasterizers` command + multi-backend rasterization (`--rasterizer <name>`) | ✅ Done |
| Cross-sheet formula fix, eager-recalculation fix | ✅ Done |

### Planned / Future

Authoring and rendering features not yet shipped:

| Feature | Status |
|---------|--------|
| XLSM macro preservation policy + tests (macros never executed) | Planned |
| Data-validation / conditional-formatting **authoring** (currently preserved through edits, no write API) | Planned |
| Drawing Layer (Images, Shapes) | Planned |
| Chart Model | Planned |
| Two-phase streaming (SST + styles in row-stream write path) | Planned |
| Merged Cells in row-stream Write | Backlog |
| Query API | Backlog |
| Pivot Tables | Backlog |

---

## Completed Work

All completed phases are documented in git history. Key milestones:

- **P0-P8**: Foundation, OOXML, streaming, codecs, macros
- **WI-07/08/09**: Formula parser, evaluator (now **104 functions** after the 0.10.0 breadth pass)
- **TJC-1055** (closes GH-116): Text functions — TRIM, MID, FIND, SUBSTITUTE, VALUE, TEXT
- **WI-10**: Excel table support
- **WI-17**: SAX streaming write (36% faster than POI)
- **WI-19**: Row/column property serialization
- **0.10.0 "Trust & Author"**: named-range & hyperlink authoring, structural editing, theme resolution, function breadth (88 → 104), surgical-edit trust fixes C1–C5

For historical details: `git log --oneline docs/plan/`

---

## Related Documentation

| Doc | Purpose |
|-----|---------|
| [STATUS.md](../STATUS.md) | Current capabilities |
| [LIMITATIONS.md](../LIMITATIONS.md) | Known limitations |
| [QUICK-START.md](../QUICK-START.md) | Get started in 5 minutes |
| [reference/cli.md](../reference/cli.md) | CLI command reference |
| [reference/performance-guide.md](../reference/performance-guide.md) | Optimization guide |

---

## Contributing

1. Check [GitHub Issues](https://github.com/TJC-LP/xl/issues) for open tasks
2. See [CONTRIBUTING.md](../CONTRIBUTING.md) for code guidelines
3. Reference issue number in commits: `fix(ooxml): implement feature (#123)`
