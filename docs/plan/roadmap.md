# XL Roadmap

> **Track Progress**: [GitHub Issues](https://github.com/TJC-LP/xl/issues)

**Last Updated**: 2026-06-10

> **Completed release records**: [archive/plan/v0.10.0-execution.md](../archive/plan/v0.10.0-execution.md) (0.10.0 tracker) and [archive/plan/v0.10.0-triage.md](../archive/plan/v0.10.0-triage.md) (rationale + per-issue verdicts).

---

## TL;DR

**Current Status**: Production-ready with **104 formula functions** (incl. dynamic arrays SEQUENCE/SORT/UNIQUE/FILTER and OFFSET), **structural editing** (insert/delete rows & columns with formula rewriting), the **scripting prelude** (`com.tjclp.xl.scripting`), whole-workbook `recalculate`, named-range & hyperlink authoring, SAX streaming (36% faster than POI), Excel tables, and full OOXML round-trip. 3005+ tests passing.

**Current Version**: **0.12.1 "Clean Sweep"** (released 2026-06-11)

---

## Release Roadmap

The full open backlog (triaged 2026-06-10) is scheduled as **six waves → four releases**, each wave
executed as a parallel multi-agent run via `.claude/workflows/issue-wave.js` (baseline gate →
worktree-isolated TDD clusters → adversarial review → integration). This roadmap is the single
source of truth for scheduling.

### v0.11.1 "Totality" — wave 1 (Released 2026-06-10)

All open bugs, one patch release (PR #276). Reviewer-discovered gaps filed as #277–#285.

| Issue | Fix |
|-------|-----|
| [#271](https://github.com/TJC-LP/xl/issues/271) | Leading unary plus (`=+A1`) parses as identity; printer normalizes |
| [#263](https://github.com/TJC-LP/xl/issues/263) | Cell-ref-shaped sheet names quoted via shared `SheetName.needsQuoting` |
| [#262](https://github.com/TJC-LP/xl/issues/262) | Trailing empty format sections preserved (`"0.0;;"` hide-zero idiom) |
| [#264](https://github.com/TJC-LP/xl/issues/264) | Streaming `StylePatcher` totality (malformed attribute hardening) |
| [#266](https://github.com/TJC-LP/xl/issues/266) | Even/first-page headers+footers, `fitToPage` flag |
| [#265](https://github.com/TJC-LP/xl/issues/265) | `DirectSaxEmitter` emits sheet metadata (+ hyperlinks) for fresh sheets |
| [#275](https://github.com/TJC-LP/xl/issues/275) | Evaluating display prefers `recalculate`'s cached values |
| [#48](https://github.com/TJC-LP/xl/issues/48) | `SheetEvaluator` var-free refactor |
| [#17](https://github.com/TJC-LP/xl/issues/17) | SST surgical whitespace: not reproducible; regression spec added |

### v0.11.2 "Laws & Functions" — waves 2 + 3 (Released 2026-06-10)

| Wave | Issues |
|------|--------|
| 2 — test infrastructure (PR #299) | [#240](https://github.com/TJC-LP/xl/issues/240) real-fixture corpus + generative round-trip law + streaming/in-memory parity; [#40](https://github.com/TJC-LP/xl/issues/40) `Sheet.put` benchmark; [#47](https://github.com/TJC-LP/xl/issues/47) renderer edge tests |
| 3 — evaluator breadth + law-found fixes #277/#287-#290 (PR #300) | [#193](https://github.com/TJC-LP/xl/issues/193) LET; [#274](https://github.com/TJC-LP/xl/issues/274) INDIRECT (design-first); [#93](https://github.com/TJC-LP/xl/issues/93) YEARFRAC parity; [#115](https://github.com/TJC-LP/xl/issues/115) RAND/RANDBETWEEN (seeded-RNG capability); [#184](https://github.com/TJC-LP/xl/issues/184) formula numFmt inheritance |

### v0.11.3 "Robustness" — waves 4 + 5 (Released 2026-06-11)

| Wave | Issues |
|------|--------|
| 4 — streaming/OOXML + parity/totality follow-ups #278/#283/#285/#291/#293/#305 (PR #310) | [#223](https://github.com/TJC-LP/xl/issues/223) two-pass streaming SST + style registry; [#242](https://github.com/TJC-LP/xl/issues/242) docProps emission; [#243](https://github.com/TJC-LP/xl/issues/243) 1904 dates + `NumFmt.Fraction` + autofilter authoring |
| 5 — CLI/UX + display/SVG/evaluator follow-ups #279-#282/#296/#298/#301/#302/#306-#308 (PR #311) | [#134](https://github.com/TJC-LP/xl/issues/134) `filter --where` (row predicates); [#137](https://github.com/TJC-LP/xl/issues/137) `diff`; [#159](https://github.com/TJC-LP/xl/issues/159) markdown import; [#156](https://github.com/TJC-LP/xl/issues/156) AWT-metric autofit; [#86](https://github.com/TJC-LP/xl/issues/86) Batik-first rasterization |

### v0.12.0 "Visual" — wave 6 (Released 2026-06-11)

Phased: (a) verbatim chart/drawing preservation proven by the wave-2 fixture corpus; (b) `Drawing`/`Image`/anchor domain model + image authoring ([#221](https://github.com/TJC-LP/xl/issues/221)); (c) typed chart AST (bar/line/pie) + authoring + `xl chart` CLI ([#222](https://github.com/TJC-LP/xl/issues/222)). Re-scoped by its own design panel when reached.

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
