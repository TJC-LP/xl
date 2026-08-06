# XL Roadmap

> **Track Progress**: [GitHub Issues](https://github.com/TJC-LP/xl/issues)

**Last Updated**: 2026-07-29

> **Completed release records**: [archive/plan/v0.10.0-execution.md](../archive/plan/v0.10.0-execution.md) (0.10.0 tracker) and [archive/plan/v0.10.0-triage.md](../archive/plan/v0.10.0-triage.md) (rationale + per-issue verdicts).

---

## TL;DR

**Current Status**: Production-ready with **112 formula functions** (incl. dynamic arrays SEQUENCE/SORT/UNIQUE/FILTER and OFFSET), **structural editing** (insert/delete rows & columns with formula rewriting), the **scripting prelude** (`com.tjclp.xl.scripting`), whole-workbook `recalculate`, named-range & hyperlink authoring, **typed charts + embedded pictures** (0.12.0), **conditional formatting** (0.12.1), SAX streaming (36% faster than POI), Excel tables, and full OOXML round-trip. 5,424 tests passing.

**Current Version**: **0.19.2** "Fixpoint" (recalculation & seeding integrity, released 2026-08-06)

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

### v0.12.1 "Clean Sweep" — wave 7 (Released 2026-06-11)

Every remaining open issue closed in one wave. **Conditional formatting** ([#136](https://github.com/TJC-LP/xl/issues/136)) is the headline — typed cellIs/expression/colorScale/dataBar/top10/text rules + `dxf` differential formats, `sheet.conditionalFormat` authoring with auto-priority, structural-edit range shifting, unmodeled families preserved byte-faithfully — alongside twelve fidelity/writer fixes: openpyxl comment subdirectory dialect (#292), RichText SST keying (#303), exact surgical SST counts (#304), `[Content_Types]` preservation (#314), identity-keyed source mappings (#315), activeTab (#294), fitToPage tri-state (#284), `Cell.comment` deprecated→`Sheet.comments` (#295). Codec `put` paths 2.4x faster (#297).

### v0.19.2 "Fixpoint" — wave 24 (Released 2026-08-06)

The 0.19.1 blind-regression round. Four new silent-wrong-number bugs, **three of them one root
cause**: evaluation reading a stale or unevaluated cache instead of the value. Six worktree-isolated
TDD clusters, one design panel, and an unusually hard adversarial review — three clusters needed a
second round and `cli-cache-safety` needed four, each round refuting the last with an empirical
counter-example.

| Issue | Fix |
|-------|-----|
| [#492](https://github.com/TJC-LP/xl/issues/492) [#491](https://github.com/TJC-LP/xl/issues/491) | **One pass = the global fixpoint** — SCC condensation walk replaces preOrder / flat-Jacobi / postOrder; per-component budgets and verdicts; `RecalcResult.cycles`/`.unconverged`/`.certified` |
| [#469](https://github.com/TJC-LP/xl/issues/469) | Cycles warm-start from numeric caches, zero as fallback |
| [#493](https://github.com/TJC-LP/xl/issues/493) [#494](https://github.com/TJC-LP/xl/issues/494) | Data-table what-if lanes evaluate their precedent cone; `ErrorGuardFired` / `ConeUnresolved` warnings |
| [#495](https://github.com/TJC-LP/xl/issues/495) | Structural edits refuse to tear a data-table interior (was a silent degrade to constants) |
| [#468](https://github.com/TJC-LP/xl/issues/468) [#481](https://github.com/TJC-LP/xl/issues/481) [#496](https://github.com/TJC-LP/xl/issues/496) | Dirty-cone-scoped recalc, `--no-recalc`, `--strict` on write verbs, `batch` honors `calcPr` |
| [#459](https://github.com/TJC-LP/xl/issues/459) | `####` for numbers too wide for their column (was a leading-digit clip) |
| [#474](https://github.com/TJC-LP/xl/issues/474) [#475](https://github.com/TJC-LP/xl/issues/475) [#476](https://github.com/TJC-LP/xl/issues/476) [#488](https://github.com/TJC-LP/xl/issues/488) [#486](https://github.com/TJC-LP/xl/issues/486) [#490](https://github.com/TJC-LP/xl/issues/490) | Field asks: hidden lines in `view`, streaming numFmt parity, 112 functions, lookup date plane, lint docs, skill fix |

**Deferred to 0.20.0**, in priority order: [#499](https://github.com/TJC-LP/xl/issues/499) (range-argument
reads never re-derive uncached formula cells and silently shorten arrays — the general form of #494),
[#507](https://github.com/TJC-LP/xl/issues/507) (`staleCaches` under-approximates through unparseable
defined names, DEFAULT path included), [#503](https://github.com/TJC-LP/xl/issues/503) +
[#509](https://github.com/TJC-LP/xl/issues/509) (`--no-recalc` over-invalidation and the
`fullCalcOnLoad` design that fixes it), [#497](https://github.com/TJC-LP/xl/issues/497) (CF evaluation
in the render path), [#498](https://github.com/TJC-LP/xl/issues/498), #500–#502, #504–#506, #508, plus
the standing tail #460 #462 #463 #465 #477 #479 #480 #482 #485.

### v0.19.0 "Canon" — wave 21 epilogue + wave 22 (Released 2026-08-03)

Tracker-zero release (PRs #447 + #451, three worktree-isolated TDD clusters in the sweep): **col/row default styles** ([#445](https://github.com/TJC-LP/xl/issues/445) — `<col style=>`/`<row s=>` emission on both backends + read-back, `Sheet.withColumnStyle`/`withRowStyle`), **sheet view modes** ([#446](https://github.com/TJC-LP/xl/issues/446)), **Excel-canonical XML forms + the `slot.ordinal` theme-index swap fix** ([#448](https://github.com/TJC-LP/xl/issues/448)), **DateTime arithmetic** ([#449](https://github.com/TJC-LP/xl/issues/449) — date−date day counts via the writer's serial conversion at every numeric boundary), **data-table tear/seed lints + `xl recalc --tables`** ([#442](https://github.com/TJC-LP/xl/issues/442)), and **ca/aca + del1/del2 fidelity** ([#435](https://github.com/TJC-LP/xl/issues/435), source-breaking `FormulaKind.Normal(aca, ca)`).

### v0.18.0 "Quill" — wave 21 (Released 2026-07-29)

Burn-down finale (PR #443, five clusters, one design panel, one rework round): **the tracker hits zero** — every issue from the 2026-07-22 25-issue plan of record is closed across three waves/releases in one week. **Two-variable Data Table authoring** ([#419](https://github.com/TJC-LP/xl/issues/419) — `sheet.dataTable` on the 0.16.0 FormulaKind substrate, Excel-fixture-verified record placement, autoNoTable cache seeding), **AutoFilter CLI** ([#432](https://github.com/TJC-LP/xl/issues/432)) + **outline grouping** ([#421](https://github.com/TJC-LP/xl/issues/421)) — 32 batch ops, **Underline enum** ([#423](https://github.com/TJC-LP/xl/issues/423), breaking w/ bridge) + **Normal font** ([#425](https://github.com/TJC-LP/xl/issues/425)), **CELL()** ([#424](https://github.com/TJC-LP/xl/issues/424), 109 functions) + **range-slot name chains** ([#411](https://github.com/TJC-LP/xl/issues/411)), **pie per-slice dPt** ([#418](https://github.com/TJC-LP/xl/issues/418)). Follow-ups: #435, #442.

### v0.17.0 "Parity" — wave 20 (Released 2026-07-29)

Burn-down wave 2 of 3 (PR #440, six worktree-isolated TDD clusters, zero rework rounds, zero merge conflicts): **numFmt parity across every path** ([#408](https://github.com/TJC-LP/xl/issues/408) StylePatcher delegates to `NumFmt.builtInId` — streamed styles match in-memory ids 2/9/7; [#410](https://github.com/TJC-LP/xl/issues/410) NumFmtFormatter built-in arms render via FormatCodeParser, property-tested identical), **batch values[] op-level format** ([#416](https://github.com/TJC-LP/xl/issues/416)), **usage errors name their flags + lint positional file** ([#422](https://github.com/TJC-LP/xl/issues/422)), **localSheetId remap on sheet order mutations** ([#434](https://github.com/TJC-LP/xl/issues/434)), **removal-orphan part pruning** ([#417](https://github.com/TJC-LP/xl/issues/417)), **comment author-prefix dedup + run-color preservation** ([#433](https://github.com/TJC-LP/xl/issues/433)), **`<sheetFormatPr>` write path** ([#426](https://github.com/TJC-LP/xl/issues/426)), and **`Sheet.named`** ([#420](https://github.com/TJC-LP/xl/issues/420)). Remaining: wave 21 "authoring" (#411 #418 #419 #421 #423 #424 #425 #432) → 0.18.0 = tracker zero (+#435 follow-up).

### v0.16.0 "Bedrock" — wave 19 + #437/#438 (Released 2026-07-29)

Burn-down wave 1 of 3 (25-issue tracker → zero; plan of record 2026-07-22). Structural integrity and preservation — the corruption/degradation class from the field-gotcha audit, five worktree-isolated TDD clusters, adversarially reviewed (one rework round on the structural cluster): **structural edits stop poisoning files** ([#427](https://github.com/TJC-LP/xl/issues/427) — equals-free re-print, stale caches invalidated rather than shipped (recalc re-bakes); openpyxl reads single-`=`), **shift clamping at the sheet edge** ([#428](https://github.com/TJC-LP/xl/issues/428) — full-height CF/DV/merges stay ≤ row 1,048,576/col XFD; Excel-refuses-file class closed), **Excel-authored DVs/print areas/tables/autoFilter shift** ([#429](https://github.com/TJC-LP/xl/issues/429) — design-panel-directed `SqrefShift` engine for preserved payloads; widened DV participation), **formula-record preservation** ([#430](https://github.com/TJC-LP/xl/issues/430), design-panel-directed — `FormulaKind` on `CellValue.Formula`: `t="array"` keeps its CSE marker, `t="dataTable"` interiors no longer bake to constants; the wave-21 #419 authoring substrate), **bytes-read preservation parity** ([#412](https://github.com/TJC-LP/xl/issues/412) — `SourceContent.OnDisk|InMemory`, path≡bytes write law), **lint extensions + `ref-out-of-bounds`** ([#413](https://github.com/TJC-LP/xl/issues/413) — flags the #428 class lint used to pass), and the **suite's only parallel-load flake structurally pinned** ([#414](https://github.com/TJC-LP/xl/issues/414)). Post-wave follow-throughs in the same release: structural commands recalculate before writing (PR #437 — fresh-correct `<v>`, the #352 batch contract) and **positional `put` smart detection** ([#431](https://github.com/TJC-LP/xl/issues/431), PR #438 — batch-parity detection incl. `--stream`, `--no-detect` opt-out, explicit-date-format errors). Follow-ups filed: #435 (ca/aca attrs; delete-band data-table degradation fidelity). Next: wave 20 "CLI & fidelity correctness" (#408 #410 #416 #417 #420 #422 #426 #433 #434) → 0.17.0; wave 21 "authoring" (#411 #418 #419 #421 #423 #424 #425 #432) → 0.18.0.

### v0.15.0 "Fidelity" — wave 18 (2026-07-17)

Field-hardening from the first production QA cycle (FinAgent LBO build) — every open issue closed in one wave (four worktree-isolated TDD clusters, adversarially reviewed): the **numFmt round-trip corruption** ([#404](https://github.com/TJC-LP/xl/issues/404) — built-in-equal formatCodes no longer degrade to General; total `NumFmt.formatCode` inverse; generators un-dodged), the deferred **#385 range-date slice** ([#405](https://github.com/TJC-LP/xl/issues/405) — XIRR/XNPV/holiday ranges coerce raw serials; post-round-trip recalc no longer poisons returns blocks), variadic blank-arg parity ([#395](https://github.com/TJC-LP/xl/issues/395)), decoder coercion edges ([#396](https://github.com/TJC-LP/xl/issues/396)), **names + sheet-qualified refs in range-typed slots** ([#394](https://github.com/TJC-LP/xl/issues/394) — `RangeLocation.Name`/`SheetNameRef`, `cellRange`→`rangeLocation` migration, plus a latent wrong-sheet fix for cross-sheet lookup/financial args), hyperlink `#` normalization ([#406](https://github.com/TJC-LP/xl/issues/406)), **per-series chart spPr + `--series-colors` + `chart` batch op** ([#407](https://github.com/TJC-LP/xl/issues/407) — LibreOffice renders chart-add output out of the box), and **`xl lint`** ([#397](https://github.com/TJC-LP/xl/issues/397) — CT child-order + r:id resolution on raw zip parts, with preserved CT_Workbook children now re-emitted in schema position). Follow-ups filed during the wave: #408 (StylePatcher id table).

### v0.14.0 "Candor" — wave 17 + #400 (Released 2026-07-16)

The replication campaign's finale (PRs #401, #402): the **typed-result refactor** ([#344](https://github.com/TJC-LP/xl/issues/344), design-panel-directed) makes Excel error values first-class evaluation results — `=1/0` → `#DIV/0!` as a catchable, cacheable value; aggregate/logical/comparison propagation per Excel per-function policy; op-level `#NUM!` classification; `#N/A` dimension padding; TRUE/FALSE text conditions; host failures stay loud behind a law-tested boundary — plus xl-agent robustness (pause_turn auto-resume, errored-task accounting, trace-overwrite guard) and **full calcPr authoring** ([#400](https://github.com/TJC-LP/xl/issues/400): calcMode/fullCalcOnLoad/calcId — the last tjc-modeling zip patch retired). Campaign complete: every issue in waves W1–W7 closed across 0.12.7 → 0.13.0 → 0.14.0. Open follow-ups: [#394](https://github.com/TJC-LP/xl/issues/394)–[#397](https://github.com/TJC-LP/xl/issues/397).

### v0.13.0 "Fixpoint" — waves 13–16 (Released 2026-07-16)

The replication-campaign feature train, four waves in one minor (PRs #391, #392, #393, #398 — every cluster worktree-isolated, TDD, adversarially reviewed): parser parity (percent postfix [#355](https://github.com/TJC-LP/xl/issues/355), preserved leading unary plus [#374](https://github.com/TJC-LP/xl/issues/374)), appearance read/write parity (freeze-pane read + scrolled panes + tabSelected [#372](https://github.com/TJC-LP/xl/issues/372)/[#382](https://github.com/TJC-LP/xl/issues/382), tabColor model+CLI [#358](https://github.com/TJC-LP/xl/issues/358)), authoring API (data validation [#375](https://github.com/TJC-LP/xl/issues/375), calcPr [#373](https://github.com/TJC-LP/xl/issues/373), Patch comment/CF [#379](https://github.com/TJC-LP/xl/issues/379), textRotation [#380](https://github.com/TJC-LP/xl/issues/380), runtime columns [#361](https://github.com/TJC-LP/xl/issues/361), writeRecalculated [#360](https://github.com/TJC-LP/xl/issues/360)), the evaluator milestone pair (defined names [#384](https://github.com/TJC-LP/xl/issues/384) — 926/1,571 probe rejections on one real LBO — and opt-in Jacobi iterative recalculation completing [#373](https://github.com/TJC-LP/xl/issues/373); coercion parity [#385](https://github.com/TJC-LP/xl/issues/385), MROUND [#386](https://github.com/TJC-LP/xl/issues/386)), and CLI tooling ([#356](https://github.com/TJC-LP/xl/issues/356), [#357](https://github.com/TJC-LP/xl/issues/357), [#324](https://github.com/TJC-LP/xl/issues/324), [#359](https://github.com/TJC-LP/xl/issues/359)). Milestones M1 (house formulas evaluate end-to-end) and M2 (recalculate() verifies a real LBO incl. circular debt schedules) unlocked. Remaining: #344 error-propagation parity → 0.14.0; evaluator follow-ups #394–#396; workbook lint #397.

### v0.12.7 "Integrity" — wave 12 (Released 2026-07-16)

File-integrity bugs from the tjc-modeling byte-exact replication campaign (PR #389, four worktree-isolated TDD clusters, each adversarially reviewed with revert-and-rerun refutation + openpyxl 3.1.5 cross-checks): property-only rows survive scratch writes ([#381](https://github.com/TJC-LP/xl/issues/381)), formula cells keep cached DateTime as Excel serials on all backends ([#378](https://github.com/TJC-LP/xl/issues/378)), schema-valid `<rFont>` comment/SST/inline rich runs + reader acceptance ([#383](https://github.com/TJC-LP/xl/issues/383)), identity-named modified-sheet parts on reordered sources ([#327](https://github.com/TJC-LP/xl/issues/327)), comment-removal CT/rels/legacyDrawing pruning ([#328](https://github.com/TJC-LP/xl/issues/328)), default theme part for scratch theme-color workbooks ([#387](https://github.com/TJC-LP/xl/issues/387)), and per-cell containment of financial-function divergence with a recalculate() totality backstop ([#388](https://github.com/TJC-LP/xl/issues/388)). Remaining campaign backlog: W2 parser unblockers (#374, #355) → W3 read/write parity (#372, #382, #358) → W4 authoring API (#373, #379, #380, #375, #361, #360) → W5 evaluator (#386, #385, #384) → W6 CLI/tooling (#356, #357, #324, #359) composing 0.13.0; then #344 as 0.14.0.

### v0.12.6 "Fieldwork" (Released 2026-07-15)

Field-reported bugs from production agent use on real deal workbooks ([#349](https://github.com/TJC-LP/xl/issues/349)–[#354](https://github.com/TJC-LP/xl/issues/354), PRs #363–#368, each adversarially reviewed): external-workbook reference parsing with Excel-cache pinning ([#353](https://github.com/TJC-LP/xl/issues/353)), batch recalculation + the `xl recalc` command incl. a latent recalc-cache/surgical-write fix ([#352](https://github.com/TJC-LP/xl/issues/352)), visible `view`/`search` truncation ([#351](https://github.com/TJC-LP/xl/issues/351)), DOCTYPE-tolerant core-part reads with line/column diagnostics ([#350](https://github.com/TJC-LP/xl/issues/350)), native-image Xerces message bundles + per-platform release smoke tests ([#349](https://github.com/TJC-LP/xl/issues/349)), and linux-arm64 native binaries ([#354](https://github.com/TJC-LP/xl/issues/354)). Companion skill-docs refresh (#362). Enhancement backlog from the same field triage: #355–#361.

### v0.12.5 "Memo" (Released 2026-07-13)

Evaluator performance and correctness on recursive models ([#346](https://github.com/TJC-LP/xl/issues/346)): pass-local memoization of recursively evaluated uncached formula references (direct refs, aggregate readers, array materialization — each cell computes once per pass instead of once per dependency path), and `Workbook.recalculate()` rebuilt on a workbook-level qualified graph with global Tarjan cycle isolation + one global Kahn order. An LBO-style debt schedule that previously ran for hours (exponential) recalculates in milliseconds; cross-sheet cycles report as circular/blocked; cross-sheet aggregates over uncached formulas are sheet-order-independent.

### v0.12.4 "Carriage" — wave 11 (Released 2026-07-09)

Elementwise error semantics through array operations ([#337](https://github.com/TJC-LP/xl/issues/337) — errors carry per-element with a property-tested Left-only-on-dimension-mismatch invariant; aggregates fail loudly on carried errors instead of silently mis-summing), unused IF/IFS branch errors no longer poison array formulas ([#339](https://github.com/TJC-LP/xl/issues/339)), and loud benchmark failure diagnostics ([#340](https://github.com/TJC-LP/xl/issues/340) — stop_reason capture, 32K default output cap + `--max-tokens`, partial-usage recovery, red glyphs). Live probe: task 13894 xl 3/3 first time. Deliberate divergences tracked in [#344](https://github.com/TJC-LP/xl/issues/344).

### v0.12.3 "Parity" — waves 10–11a (Released 2026-07-09)

Evaluator correctness gaps found by live SpreadsheetBench dogfooding: Excel comparison total order — text/cross-type/date/empty semantics ([#335](https://github.com/TJC-LP/xl/issues/335)), array-aware IF per CSE + elimination of the aggregator crash family ([#333](https://github.com/TJC-LP/xl/issues/333)), AND/OR array aggregation with NOT/IFS broadcast ([#338](https://github.com/TJC-LP/xl/issues/338)), and benchmark failure diagnostics ([#334](https://github.com/TJC-LP/xl/issues/334)). Plus the xl-agent harness refresh (#332): Claude 5 registry, anthropic-java 2.48.0, prompt caching (−52–59%/task), version-agnostic release-asset resolution, Skills API drift fixes. Follow-ups filed: #337 (elementwise error propagation), #339 (IF eager branches, depends on #337), #340 (diagnostics polish).

### v0.12.2 "Interop" — wave 9 (Released 2026-06-11)

The LibreOffice edit-corruption fix and the writer follow-ups it surfaced: editing LO-produced workbooks no longer corrupts them ([#320](https://github.com/TJC-LP/xl/issues/320) — `workbook.xml.rels` regenerates in the same pass as `workbook.xml`, so sheet rIds stay consistent), comment content-type registration follows actual emitted paths (#321), no dangling `[Content_Types]` overrides for dropped writer-owned parts (#322), and new sheets join surgical SST accounting (#323).

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
| Data-validation **authoring** (currently preserved through edits, no write API; conditional-formatting authoring shipped in 0.12.1) | Planned |
| Drawing Layer — Shapes/connectors authoring (images shipped in 0.12.0) | Planned |
| Two-phase streaming (SST + styles in row-stream write path) | Planned |
| Merged Cells in row-stream Write | Backlog |
| Query API | Backlog |
| Pivot Tables | Backlog |

---

## Completed Work

All completed phases are documented in git history. Key milestones:

- **P0-P8**: Foundation, OOXML, streaming, codecs, macros
- **WI-07/08/09**: Formula parser, evaluator (**108 functions**; the 0.10.0 breadth pass took the registry 88→104, then 0.11.2 added INDIRECT/RAND/RANDBETWEEN/LET)
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
