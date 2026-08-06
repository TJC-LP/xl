# XL Project Status

**Last Updated**: 2026-08-06 (0.19.2)

## Current State

> **For detailed phase completion status and roadmap, see [plan/roadmap.md](plan/roadmap.md)**

### What Works (Production-Ready)

**New in 0.19.2 "Fixpoint"** (2026-08-06) — recalculation & seeding integrity, wave 24:
- ✅ **One pass = the global fixpoint** (#492, #491) — iterative recalculation walks the SCC condensation in dependency-first order instead of splitting into preOrder / one flat Jacobi / postOrder, so no read can fall back to a stale cache and re-recalculating a correct multi-SCC circular book no longer corrupts it. `RecalcResult.cycles` / `.unconverged` / `.certified` report per-component verdicts
- ✅ **Cycles warm-start from numeric caches** (#469) — zero-seeding no longer wipes valid caches of mutually-ISERROR-guarded pairs
- ✅ **Data-table what-if lanes evaluate their precedent cone** (#493, #494) — acyclic grids no longer seed silently FLAT, guarded XIRR corners seed real rates, and a guard resolving to its error arm surfaces as `ErrorGuardFired` instead of banking its text arm
- ✅ **Structural edits refuse to tear a data-table interior** (#495) — was a silent degrade to constants that `data-table-torn` could not see
- ✅ **Cache-safe CLI writes** (#468, #481, #496) — dirty-cone-scoped recalc, `--no-recalc`/`--preserve-caches`, `--strict` exit codes on write verbs, `batch` honoring declared `calcPr`
- ✅ **`####` for overflowing numbers in the raster** (#459) — was a leading-digit clip that rendered a plausible wrong number
- ✅ **112 functions** (#476 — SEARCH, N, HYPERLINK); hidden rows/cols in `view` (#474); streaming numFmt parity (#475); VLOOKUP/HLOOKUP date keys (#488)

**New in 0.19.1** (2026-08-04) — field hardening, waves 23 + 23b (#478/#487):
- ✅ **Circular-book data-table seeding** (#453) — the seeder fixpoints cycle members per axis combination under the book's CalcPr instead of pinning them at loaded caches (grids no longer seed silently FLAT); `seedDataTablesReport()` + IterativeCalc overloads
- ✅ **`RecalcResult.converged` / `iterationsUsed`** (#454) — maxIter exhaustion distinguishable from convergence; CLI recalc honors declared `calcPr iterate="1"` (#461)
- ✅ **Reprint integrity** (#455, #484) — structural edits preserve associativity parentheses and serialize Excel's file form; scripting `<f>` leading-`=` canonicalized + `formula-leading-equals` lint (#456)
- ✅ **Evaluator wrong-value classes dead** (#466, #467) — text-operand criteria (`"<>x"`, `">m"`, non-blank `"<>"`) evaluate with Excel semantics; MATCH/XLOOKUP comparator is total (blanks match nothing, DateTime-by-serial, cell-ref lookup values dereference)
- ✅ **Write-integrity guards** (#470, #471, #472, #473) — `copy(sheets=…)` reductions honored under preservation; fresh-write style plane re-registered (numFmt by code, dxfs carried); structural edits refuse past-bounds shifts and rewrite general defined names
- ✅ **Native image opens name-bloated books** (#457) — JAXP limits lifted (62k–110k definedName banker files); lint accepts `xlPathMissing` external links (#458)

**New in 0.19.0** (2026-08-03):
- ✅ **Column/row default styles** (#445) — `<col style=>` / `<row s= customFormat="1">` emit on both writer backends (StyleIndex-remapped like cell styleIds) AND parse back, so read→modify→write keeps source column styles; `Sheet.withColumnStyle`/`withRowStyle` author the sheet-wide-body-font-without-Normal mechanism
- ✅ **Sheet view modes** (#446) — `SheetView.view` (normal/pageBreakPreview/pageLayout) + `zoomScaleNormal`/`zoomScaleSheetLayoutView`/`topLeftCell`, set-or-remove with foreign values riding preservation
- ✅ **Excel-canonical XML forms** (#448) — integral `sz`, 17-sig-digit plain tints (`tint="0"` omitted), bare gray125, derived `outlineLevelRow/Col` summary attrs; **theme-index swap fixed** — SAX path + comments wrote Dark2 as Light2 via `slot.ordinal`
- ✅ **DateTime arithmetic** (#449) — `=end-start` day counts, `=date+30` offsets, and MIN/MAX/COUNT/SUM over date columns evaluate via `dateTimeToExcelSerial` (the writer's conversion); result is a serial Number, booleans stay skipped in aggregates
- ✅ **Data-table lints + `xl recalc --tables`** (#442) — `data-table-torn` (5 tear classes incl. del-flagged records) + `data-table-unseeded` (autoNoTable doctrine), DOM/SAX finding-identical, O(1) streaming; `recalc --tables` seeds after recalculation, default pinned-cache path byte-identical
- ✅ **ca/aca + del1/del2 fidelity** (#435) — plain-formula calc flags on `FormulaKind.Normal(aca, ca)` survive every path (source-breaking: `FormulaKind.Normal()`); input-deleting structural edits keep the record del-flagged with caches intact

**New in 0.18.0** (2026-07-29):
- ✅ **Two-variable Data Table authoring** (#419) — `sheet.dataTable(range, rowInput, colInput)` authors native `<f t="dataTable">` records with autoNoTable-safe cache seeding (design-panel-verified against Excel fixtures); the house sensitivity engine is authorable
- ✅ **AutoFilter + outline grouping CLI** (#432, #421) — `xl autofilter <range>`/`--clear`, `group-rows`/`group-cols`/`ungroup-*` + 5 batch twins (32 ops); authored filters shift under structural edits
- ✅ **Underline enum** (#423, breaking w/ deprecated bridge) — singleAccounting/doubleAccounting round-trip typed; **configurable Normal font** (#425) — `WorkbookMetadata.defaultFont` on both backends
- ✅ **CELL()** (#424) — filename/address/row/col arms, volatile (109 functions); **name→name chains in range slots** (#411)
- ✅ **Pie per-slice colors** (#418) — `--series-colors` → typed `Series.pointFills` → `<c:dPt>` fills

**New in 0.17.0** (2026-07-29):
- ✅ **numFmt parity everywhere** (#408, #410) — StylePatcher delegates to NumFmt.builtInId (streamed Decimal/Percent/Currency ids 2/9/7 match in-memory); NumFmtFormatter built-in arms render via FormatCodeParser (PercentDecimal shows Excel's 15.60%)
- ✅ **Batch values[] op-level format** (#416) + **usage errors name flags, lint positional file** (#422)
- ✅ **Sheet lifecycle bookkeeping** (#434, #417) — localSheetId remap on remove/insert/reorder; removal-orphan chart/drawing part pruning (shared parts survive)
- ✅ **Comment round-trip fidelity** (#433) — no duplicate author prefix; run colors survive via raw rPr preservation
- ✅ **sheetFormatPr write path** (#426) + **Sheet.named dynamic factory** (#420)

**New in 0.16.0** (2026-07-29):
- ✅ **Array/data-table formula records survive rewrites** (#430) — `<f t="array" ref>` and `<f t="dataTable" ref dt2D dtr r1 r2 del1 del2 ca/>` are modeled per cell (`FormulaKind` on `CellValue.Formula`) and re-emit byte-exactly through dirty sheet regeneration on every writer (DOM, SAX, streaming); two-variable Data Tables no longer bake to static grids on `put`. DataTable caches are pinned (recalc/copy/eval never parse the derived `TABLE(...)` display text); StructuralEditor shifts record payloads, degrades interior-tearing edits to cached constants, and sets `del1`/`del2` when a delete removes an input cell (#435); plain-formula `ca`/`aca` ride on `FormulaKind.Normal` (#435); CLI shows `{=TABLE(A1,A2)}` braces + JSON `formulaKind`, and `putf` rejects top-level `TABLE(` (authoring lands with #419); new `formula-records.xlsx` corpus fixture rides every round-trip/parity law
- ✅ **Structural edits stop poisoning files** (#427/#428/#429) — equals-free `<f>` re-print with caches invalidated then **re-baked by a global recalc before write** (#352 contract); range shifts clamp at row 1,048,576/col XFD (full-axis `A:A`/`1:1` shapes preserved); Excel-authored dataValidations, print areas, tables, and preserved autoFilter shift with the edit via the `SqrefShift` engine
- ✅ **Positional `put` smart detection** (#431) — currency/percent/ISO-date/number/boolean detection at batch parity incl. `--stream` (with style-preserving detected formats); `--no-detect` opt-out; explicit `format:"date"` on unparseable input errors instead of writing text dressed as a date
- ✅ **Bytes-read preservation parity** (#412) — `SourceContent.OnDisk | InMemory`; `write(read(path)) ≡ write(readFromBytes(bytes))` law-tested byte-identically (breaking: `SourceContext.sourcePath: Path` → `content: SourceContent`; `fromFile` unchanged)
- ✅ **`xl lint` extensions** (#413) — chartsheet/dialogsheet order tables, externalLink r:id chain, `[Content_Types]` registration, O(1) SAX mode, and `ref-out-of-bounds` (flags over-max ranges Excel refuses; previously reported clean)
- ✅ **Suite flake removed** (#414) — the wall-clock single-pass race is now structural (100/100 under saturated load)

**New in 0.15.0** (2026-07-17):
- ✅ **numFmt round-trip fidelity** (#404) — `<numFmts>` declarations survive read→write verbatim (built-in-equal codes like `"0.00%"` no longer degrade to `General`); total `NumFmt.formatCode` inverse; whole-code `"General"` renders correctly (incl. `TEXT(n,"General")`)
- ✅ **Range-date serial coercion** (#405) — XIRR/XNPV date ranges and NETWORKDAYS/WORKDAY holidays accept raw serial Numbers; post-round-trip recalc no longer poisons returns blocks
- ✅ **Variadic blank-arg parity** (#395) + **decoder coercion edges** (#396) — direct blank refs ignored by COUNT/COUNTA/AVERAGE (`=AVERAGE(blank,blank)` → `#DIV/0!`); `decodeAsInt`/`decodeAsDate` mirror ScalarCoercion (Empty/Bool/cached-formula arms)
- ✅ **Names + sheet-qualified refs in range-typed slots** (#394) — `=VLOOKUP(x, named_table, 2)`, ad-hoc `=XIRR(S!A1:B1,…)`, sheet-scoped names (`=Model!case`), `.`/`\` in names; new `RangeLocation.Name`/`TExpr.SheetNameRef`; latent cross-sheet wrong-sheet bug fixed for INDEX/MATCH/XLOOKUP/XIRR/NPV
- ✅ **Chart series styling** (#407) — every series emits type-appropriate `<c:spPr>` (bar fills, line strokes, pie per-slice `dPt`) cycling `DefaultTheme.accents`; `Series.fill`, `chart add --series-colors`, `chart` batch op (27 ops), `<c:tx>` always emitted — LibreOffice renders chart output out of the box
- ✅ **`xl lint`** (#397) — raw-zip validation of CT child order + r:id resolution (exit 0/1/2, `--format json`); preserved CT_Workbook children re-emit in schema position; hyperlink `#` normalization (#406)

**New in 0.14.0** (2026-07-16):
- ✅ **Excel error values as first-class results** (#344) — `=1/0` evaluates to `#DIV/0!` (IFERROR/ISERROR-catchable, cached as `t="e"` cells) instead of failing the formula; aggregates/logical folds/comparisons propagate per Excel policy (COUNT-family still skips); op-level `#NUM!`/`#DIV/0!`/`#N/A` codes; `#N/A` dimension padding; `"TRUE"`/`"FALSE"` text coerces in conditions; host failures stay loud — the boundary is law-tested (`RecalcResult.excelErrors` vs `errors`); design record `docs/design/error-propagation.md`
- ✅ **Full calcPr authoring** (#400) — `calcMode`/`fullCalcOnLoad`/`calcId` join the iterate triple; the TJC house `<calcPr>` authors byte-exactly; last tjc-modeling zip patch retired
- ✅ **xl-agent robustness** (#344) — bounded `pause_turn` auto-resume with preserved container id, errored tasks counted in summaries, skill traces never overwritten by the engine fallback

**New in 0.13.0** (2026-07-16):
- ✅ **Defined-name resolution** (#384) — `=IF(case=2,…)`, `=entry_mult*ltm_ebitda` evaluate; sheet-scoped shadowing, name-chains with cycle guard, dependency-graph edges; was 926/1,571 probe rejections on a real LBO
- ✅ **Opt-in iterative recalculation** (#373) — `recalculate(IterativeCalc(maxIter, maxChange))` Jacobi-fixpoints declared cycles (circular debt schedules verify); calcPr authoring for scratch workbooks
- ✅ **Coercion parity** (#385) + **MROUND** (#386) — serial Numbers in date positions, blanks as 0 in scalar numeric contexts (aggregates still skip); 108 registry functions
- ✅ **Parser parity** (#355, #374) — percent postfix operator with Excel precedence and byte-identical round-trip; leading unary plus preserved through print
- ✅ **Appearance round-trip** (#372, #382, #358) — freeze panes read into the model (incl. scrolled panes), `tabSelected`, `Sheet.tabColor`; CLI: `sheet-view`, `tab-color` (theme syntax), `page-setup`, `header-footer`
- ✅ **Authoring API** (#375, #379, #380, #361, #360) — data-validation dropdowns (typed + preserved), Patch comment/CF cases, `Align.textRotation`, `Column.parse` runtime handles, `Excel.writeRecalculated`
- ✅ **CLI tooling** (#356, #357, #324, #359) — batch `putf` format field, JSON `formula` field, `cf add`/`cf list`, actionable rasterizer diagnostics with native-image awareness

**New in 0.12.7** (2026-07-16):
- ✅ **Property-only rows survive scratch writes** (#381) — rows with only `RowProperties` (no cells) emit `<row>` on the default backend; backend parity pinned by tests
- ✅ **Cached DateTime on formula cells** (#378) — `=TODAY()`/`=EDATE()`/`=EOMONTH()` caches serialize as Excel serials (`t="n"` + `<v>`) on all three writer backends; streaming writer gained cached-Text `t="str"`
- ✅ **Schema-valid comment rich runs** (#383) — `<rFont>` per CT_RPrElt (was styles-shape `<name>`); openpyxl opens XL-authored comments; reader accepts both dialects
- ✅ **Identity-named modified-sheet parts** (#327) — no duplicate-zip-entry failure on Excel-reordered sources; physical writes, sheet rels, and workbook rels derive from one path
- ✅ **Comment-removal package pruning** (#328) — dropping a sheet's last comment prunes CT overrides, sheet rels, and legacyDrawing (resolved-path-aware, openpyxl dialect included)
- ✅ **Default theme part for scratch theme colors** (#387) — `Color.Theme` styles/dxfs ship `xl/theme/theme1.xml` + override + rel; RGB-only workbooks unaffected
- ✅ **Total recalculate() under financial divergence** (#388) — IRR/XIRR/XNPV/NPV + TVM overflow class returns per-cell errors; workbook-evaluator NonFatal backstop guarantees totality

**New in 0.12.6** (2026-07-15):
- ✅ **External-workbook references** (#353) — `[2]Book1!A1` forms parse (dedicated AST node, exact printer round-trip, anchor-aware shifting), contribute no dependency edges, and `recalculate()` pins their Excel-written caches verbatim while dependents compute from them; uncached external cells report a clear per-cell error instead of `UnexpectedChar([`
- ✅ **Batch recalculation + `xl recalc`** (#352) — cell-mutating batches end with one global recalculation (cached `<v>` for `putf` cells, errors surfaced in the summary), plus a latent fix: recalculated caches survive surgical writes of disk-read workbooks
- ✅ **Visible truncation** (#351) — `view`/`search` report `showing N of M rows` when `--limit` clips (markdown trailer, stderr for csv, `truncated`/`totalRows` in json); `--limit 0` = unlimited
- ✅ **DOCTYPE-tolerant core-part reads** (#350) — benign `<!DOCTYPE>` prologs stripped by a conservative scanner (parser stays locked down); parse errors carry line/column
- ✅ **Native-image diagnostics + arm64** (#349, #354) — Xerces message bundles registered (real parse errors on the shipped binary, release smoke test per platform); linux-arm64 native binaries join the release matrix

**New in 0.12.5** (2026-07-13):
- ✅ **Evaluator memoization + workbook-level recalculation** (#346) — recursive uncached-reference evaluation is memoized per pass, and `recalculate()` runs one topological order over the qualified (sheet!cell) graph: the recursive debt-schedule shape drops from hours (exponential in dependency paths) to milliseconds; cross-sheet cycles now report circular/blocked like same-sheet ones; cross-sheet aggregates over uncached formula cells compute correctly regardless of sheet order

**New in 0.12.4** (2026-07-09):
- ✅ **Elementwise error carriage** (#337) — array comparisons/arithmetic carry `#DIV/0!`/`#REF!`/`#VALUE!` per element (Left only on dimension mismatch, property-tested); aggregates fail loudly on carried errors (IFERROR-catchable)
- ✅ **Array-IF branch errors** (#339) — unused-branch failures demote to discarded error elements per CSE
- ✅ **Benchmark truncation diagnostics** (#340, internal) — stop_reason capture, 32K output cap + `--max-tokens`; divergence backlog in #344

**New in 0.12.3** (2026-07-09):
- ✅ **Excel comparison semantics** (#335) — ordered comparisons follow Excel's total order (case-insensitive lexicographic text, `number < text < logical`, dates by serial, empty coercion) in scalar and array paths
- ✅ **Array-aware IF + crash-free logical functions** (#333, #338) — CSE broadcast for IF/IFS, AND/OR aggregate over arrays, NOT spills elementwise; the `MIN(IF(...))` ClassCastException family is gone
- ✅ **Benchmark failure diagnostics** (#334, internal) — per-case errors in reports, partial traces survive mid-run failures

**New in 0.12.0–0.12.2** (2026-06-11):
- ✅ **Typed charts** (0.12.0, #222) — bar/line/pie via `Chart.bar`/`line`/`pie` + `Sheet.addChart` (CLI `chart add`); typed-parse-or-Preserved hybrid read, structural-edit + rename reference tracking (see LIMITATIONS §12)
- ✅ **Embedded pictures** (0.12.0, #221) — `Sheet.addImage`/`pictures`/`removeDrawing`, three anchor forms, 7-format sniffing, sha-deduped media; non-picture drawings preserved (see LIMITATIONS §13)
- ✅ **Conditional formatting** (0.12.1, #136) — `Sheet.conditionalFormat` with cellIs/expression/colorScale/dataBar/top10/text rules + `Dxf` differential formats; joins the generative round-trip law (library API; see LIMITATIONS §10)
- ✅ **LibreOffice interop** (0.12.2) — editing LibreOffice-produced workbooks no longer corrupts them; `[Content_Types]`/SST writer accounting hardened (#320–#323)
- ✅ Codec `put` paths ~2.4x faster (0.12.1, #297)

**New in 0.11.0 "Scripting"** (2026-06-10):
- ✅ **Scripting prelude** `com.tjclp.xl.scripting.{*, given}` — ONE import for scripts: core API + DSL + compile-time literals + formula evaluation + sync `Excel` + streaming `ExcelIO` + smart detection (`String.toFormatted`) + `.unsafe` boundary
- ✅ **Range fill**: `range := value` puts the value in every cell (Excel Ctrl+Enter semantics; previously a silent no-op)
- ✅ **Total ARef navigation**: `down`/`up`/`right`/`left` (default 1) for Either-free loops
- ✅ **`Workbook.upsert(name, f)`** — total update-or-create counterpart of `update`
- ✅ **`Workbook.recalculate(clock)`** — total whole-workbook recalculation returning `RecalcResult` (cached workbook + per-sheet values + per-cell `CellEvalError`s; cycles isolated, acyclic remainder still evaluates)
- ✅ **`FormattedParsers.detect`** — total smart value detection (currency/accounting/percent/ISO date/number/boolean/text), promoted from CLI-internal code
- ✅ **Per-side borders + outlines**: `CellStyle.borderTop/borderBottom/borderLeft/borderRight` and `range.outlined(style[, color])` via the new `Patch.MergeBorder`
- ✅ **Alignment indent**: `CellStyle.indent(n)`
- ✅ **Sheet view settings**: `SheetView(showGridLines, zoomScale)` on `Sheet.viewSettings` (SVG renderer respects gridline suppression)
- ✅ **Print setup extensions**: `PageSetup` gains `headerFooter`, `margins`, `printArea`, `repeatRows` (sheet-scoped `_xlnm` defined names; even/first headers + fitToPage tracked in #266)
- ✅ **xl-scripting skill** (`plugin/skills/xl-scripting/`) — SKILL.md + API reference + runnable recipes, compile-verified on CI
- ✅ **Anti-rot CI**: examples job (`scripts/test-examples.sh`) + skill-verify workflow (`scripts/verify-skill-snippets.sh`)

**Core Features**:
- ✅ Type-safe addressing (Column, Row, ARef with 64-bit packing)
- ✅ Compile-time validated literals: `ref"A1"` and `ref"A1:B10"`
- ✅ Immutable domain model (Cell, Sheet, Workbook)
- ✅ Patch Monoid for declarative updates
- ✅ Complete style system (Font, Fill, Border, Color, NumFmt, Align)
- ✅ StylePatch Monoid for style composition
- ✅ StyleRegistry for per-sheet style management
- ✅ **End-to-end XLSX read/write** (creates real Excel files)
- ✅ **Surgical modification** (read → modify → write preserves unknown parts: charts, images, drawings, comments, and inline worksheet elements — dataValidations, sheetProtection, autoFilter — preserved through edits as of 0.10.0 / C1)
- ✅ **Embedded images** (#221): `Sheet.addImage(image, at | range | anchor)` with natural-size sniffing (png/jpeg/gif/bmp), `Sheet.pictures`, `removeDrawing`; one-cell/two-cell/absolute anchors; shapes ride through as `Drawing.Preserved`; media sha-256 dedup; in-memory read/write only (see LIMITATIONS §13)
- ✅ **Typed charts** (#222): bar (clustered/stacked/percent-stacked, column/horizontal), line, pie — `Chart`/`Series`/`DataRef` model, `Sheet.addChart`/`charts`, CLI `chart add` + `add-image`; typed-parse-or-Preserved hybrid read (out-of-fence charts stay byte-preserved); structural edits + rename track chart references; value caches resolved from stored cells on write (see LIMITATIONS §12)
- ✅ **Structural editing** (0.10.0): insert/delete rows & columns shift cells, merges, row/col properties, freeze panes, and rewrite all affected formulas (cross-sheet) with `#REF!` generation
- ✅ **Named ranges & hyperlinks authoring** (0.10.0): `DefinedName` and `Cell.hyperlink` are now serialized (previously read-only)
- ✅ Hybrid write optimization (11x speedup for unmodified workbooks, 2-5x for partial modifications)
- ✅ Shared Strings Table (SST) deduplication
- ✅ Styles.xml with component deduplication
- ✅ Multi-sheet workbooks
- ✅ All cell types: Text, Number, Bool, Formula, Error, DateTime
- ✅ RichText support (multiple formats within one cell)
- ✅ DateTime serialization (Excel serial number conversion)
- ✅ **Excel Tables** (structured data ranges with headers, AutoFilter, and styling)
- ✅ **True streaming I/O** (constant memory, 100k+ rows)

**Ergonomics & Type Safety**:
- ✅ Given conversions: `sheet.put(ref"A1", "Hello")` (no wrapper needed)
- ✅ Batch put via varargs `Sheet.put(ref -> value, ...)`
- ✅ Formatted literals: `money"$1,234.56"`, `percent"45.5%"`, `date"2025-11-10"`
- ✅ **String interpolation**: `ref"$sheet!$cell"`, `money"$$${amount}"` with runtime validation
- ✅ Compile-time optimization for literal interpolations (zero runtime overhead)
- ✅ **CellCodec[A]** for 9 primitive types (String, Int, Long, Double, BigDecimal, Boolean, LocalDate, LocalDateTime, RichText)
- ✅ Batch `Sheet.put` with auto-inferred formatting (former `putMixed` API)
- ✅ `readTyped[A]` for type-safe cell reading
- ✅ **Optics** module (Lens, Optional, focus DSL)
- ✅ RichText DSL: `"Bold".bold.red + " normal " + "Italic".italic.blue`
- ✅ HTML export: `sheet.toHtml(ref"A1:B10")`
- ✅ **Formula Parsing** (WI-07 complete): TExpr GADT, FormulaParser, FormulaPrinter with round-trip verification and scientific notation
- ✅ **Formula Evaluation** (WI-08 complete): Pure functional evaluator with total error handling, short-circuit semantics, and Excel-compatible behavior
- ✅ **Function Library**: **108 built-in functions** (aggregate, conditional, logical, text, date, financial, lookup, math, statistical, dynamic arrays), extensible type class parser, evaluation API. 0.10.0 added IFS, SWITCH, CHOOSE, LARGE, SMALL, RANK, PERCENTILE, QUARTILE, HLOOKUP, MAXIFS, MINIFS, OFFSET, and the spill functions SEQUENCE/SORT/UNIQUE/FILTER (#76, #120, #122); 0.11.2 added INDIRECT (GH-274), RAND/RANDBETWEEN via the seeded Rng capability (GH-115), and LET lexical bindings (GH-193) (dynamic text-to-reference resolution with deferred-bucket recalculation).
- ✅ **Dependency Graph** (WI-09d complete): Circular reference detection (Tarjan's SCC), topological sort (Kahn's algorithm), safe evaluation with cycle detection
- ✅ **Cross-Sheet Formula References** (TJC-351): Single cell refs (`=Sales!A1`), range refs (`=SUM(Sales!A1:A10)`), arithmetic with cross-sheet refs, workbook-level cycle detection (`DependencyGraph.fromWorkbook`)

**Performance** (JMH Benchmarked - WI-15; figures captured on an earlier release, not re-validated for 0.12.x — treat as indicative):
- ✅ **Streaming reads: 35% faster than POI for small files** (0.887ms vs 1.357ms @ 1k rows)
- ✅ **Streaming reads: Competitive with POI for large files** (8.408ms vs 7.773ms @ 10k rows - within 8%)
- ✅ **In-memory reads: 26% faster than POI for small files** (1.225ms vs 1.650ms @ 1k rows)
- ✅ Inline hot paths (SAX parser: 3.8x speedup vs fs2-data-xml)
- ✅ Zero-overhead opaque types
- ✅ Macros compile away (no runtime parsing)
- ⚠️ Writes: POI 49% faster (future optimization work - Phase 3)

**Streaming API**:
- ✅ Excel[F[_]] algebra trait
- ✅ ExcelIO[IO] interpreter
- ✅ `readStream` / `readSheetStream` / `readStreamByIndex` – constant‑memory streaming read (fs2.io.readInputStream + fs2‑data‑xml)
- ✅ `writeStream` / `writeStreamsSeq` – constant‑memory streaming write (fs2‑data‑xml)
- ✅ `writeWorkbookStream` – lower-allocation SAX/StAX write for in-memory workbooks; preserves merges, comments, tables, row/column properties, and freeze panes
- ✅ **`writeFast`** – SAX/StAX streaming write (opt-in via `ExcelIO.writeFast()` or `WriterConfig(backend = XmlBackend.SaxStax)`)
- ✅ Benchmark: 100k rows in ~1.8s read (~10MB constant memory) / ~1.1s write (~10MB constant memory)

**Output Configuration** (P6.7 Complete):
- ✅ WriterConfig with compression and prettyPrint options
- ✅ Compression.Deflated default (5-10x smaller files)
- ✅ WriterConfig.debug for debugging (STORED + prettyPrint)
- ✅ Backward compatible API (writeWith for custom config)
- ✅ 4 compression tests verify behavior

**Infrastructure**:
- ✅ Mill build system
- ✅ Scalafmt 3.10.1 integration
- ✅ GitHub Actions CI pipeline
- ✅ Comprehensive documentation (README.md, CLAUDE.md)

### Test Coverage

**5,430 tests** (verified via `./mill __.test`, 2026-08-06):

| Module | Tests | Covers |
|--------|-------|--------|
| xl-evaluator | 1571 | parser, evaluator, 108-function library, dependency graph, cross-sheet formulas, recalculation, structural editing, Excel comparison total order, array CSE semantics |
| xl-core | 1104 | addressing laws, Patch/StylePatch monoids, codecs, optics, RichText, interpolation, render (HTML/SVG), styles DSL, charts, drawings, conditional formatting |
| xl-ooxml | 684 | round-trips (cells, styles, tables, comments, hyperlinks, charts, drawings, conditional formatting), compression, security (XXE, ZIP bomb), preservation |
| xl-cli | 406 | command parsing, batch ops, view/eval/export, streaming mode |
| xl-cats-effect | 110 | streaming I/O, O(1) memory verification, SAX/StAX write |
| xl-agent | 102 | benchmark engine, skill abstraction, failure-path diagnostics, release-asset resolution |
| xl (prelude) | 19 | external-consumer probes (`xl/test/src/xlprelude/`) |
| xl-testkit | 0 | placeholder (no sources yet) |

See [reference/testing-guide.md](reference/testing-guide.md) for suite structure and testing patterns.

---

## Current Limitations & Known Issues

### XML Serialization

**Formula System** (WI-07, WI-08, WI-09a/b/c/d - Production Ready):
- ✅ **Parsing** (WI-07): Typed AST (TExpr GADT), FormulaParser, FormulaPrinter, round-trip verification, 57 tests
- ✅ **Evaluation** (WI-08): Pure functional evaluator, total error handling, short-circuit semantics, 58 tests
- ✅ **Function Library** (WI-09a-h + TJC-1055 complete): **108 built-in functions**, extensible type class parser, evaluation API
  - **Aggregate** (12): SUM, COUNT, COUNTA, COUNTBLANK, AVERAGE, MEDIAN, MIN, MAX, STDEV, STDEVP, VAR, VARP
  - **Statistical** (5): LARGE, SMALL, RANK, PERCENTILE, QUARTILE
  - **Conditional** (9): SUMIF, COUNTIF, SUMIFS, COUNTIFS, AVERAGEIF, AVERAGEIFS, MAXIFS, MINIFS, SUMPRODUCT
  - **Logical / Selection** (13): IF, IFS, IFERROR, SWITCH, CHOOSE, AND, OR, NOT, ISNUMBER, ISTEXT, ISBLANK, ISERR, ISERROR
  - **Text** (12): CONCATENATE, LEFT, RIGHT, MID, LEN, UPPER, LOWER, TRIM, FIND, SUBSTITUTE, TEXT, VALUE
  - **Date** (12): TODAY, NOW, DATE, YEAR, MONTH, DAY, EOMONTH, EDATE, DATEDIF, NETWORKDAYS, WORKDAY, YEARFRAC
  - **Math** (16): ABS, ROUND, ROUNDUP, ROUNDDOWN, INT, MOD, POWER, SQRT, LOG, LN, EXP, FLOOR, CEILING, TRUNC, SIGN, PI
  - **Financial** (9): NPV, IRR, XNPV, XIRR, PMT, FV, PV, RATE, NPER
  - **Lookup / Reference** (12): VLOOKUP, HLOOKUP, XLOOKUP, INDEX, MATCH, OFFSET, INDIRECT, ROW, COLUMN, ROWS, COLUMNS, ADDRESS
  - **Dynamic Arrays** (5): TRANSPOSE, SEQUENCE, SORT, UNIQUE, FILTER
  - **Random** (2): RAND, RANDBETWEEN
  - FunctionSpec registry: macro-collected specs with extensible registry
  - APIs: sheet.evaluateFormula(), sheet.evaluateCell(), sheet.evaluateAllFormulas()
  - Clock trait for pure date/time functions (deterministic testing)
- ✅ **Dependency Graph** (WI-09d): Circular reference detection + topological sort, 52 tests
  - Tarjan's SCC algorithm: O(V+E) cycle detection with early exit
  - Kahn's algorithm: O(V+E) topological sort for correct evaluation order
  - Precedent/dependent queries: O(1) lookups via adjacency lists
  - Safe evaluation: sheet.evaluateWithDependencyCheck() (production-ready)
  - Performance: Handles 10k formula cells in <10ms
- ⚠️ Merged cells are supported by the in-memory OOXML path and `writeWorkbookStream`. Pure row-stream generation (`writeStream` / `writeStreamsSeq`) has no merge API.
- ✅ Hyperlinks serialized as of 0.10.0 (`Cell.hyperlink` → `<hyperlinks>` + worksheet relationships; populated on read).
- ✅ Column/row properties (width, height, hidden, outlineLevel, collapsed) are fully serialized via DirectSaxEmitter.

### Style System

**Minor Limitations**:
- ✅ Theme colors resolved via `Color.toResolvedArgb(theme)` / `toResolvedHex(theme)` (slot lookup + tint application through `ThemePalette.resolve`)
- ⚠️  StyleRegistry requires explicit initialization per sheet (design choice for purity)

### OOXML Coverage

**Missing Parts** (not critical for MVP):
- ❌ docProps/core.xml, docProps/app.xml (metadata)
- ⚠️ xl/theme/theme1.xml (theme palette) — preserved from source on round-trip, not generated for new workbooks
- ❌ xl/calcChain.xml (formula calculation order)
- ✅ Worksheet relationships (`_rels/sheetN.xml.rels`) — written when a sheet has comments, tables, or hyperlinks
- ⚠️ Print settings, page setup — odd + even/first header/footer, margins, print area, repeat rows (#259, #266), and `fitToPage` tri-state (#284); shipped across 0.11.0–0.12.1
- ✅ Conditional formatting (0.12.1, #136): typed `Sheet.conditionalFormat` rules (cellIs/expression/colorScale/dataBar/top10/text) + `Dxf` differential formats; library API (no CLI yet) — see LIMITATIONS §10
- ❌ Data validation (preserved through edits, but no authoring API yet)
- ✅ Named ranges (authoring shipped in 0.10.0: `DefinedName` serialization + CLI `name add/rm`)

### Streaming I/O Limitations

**Row-stream write path** (✅ Working):
- ✅ True constant-memory row streaming with `writeStream` / `writeStreamsSeq`
- ✅ O(1) memory regardless of file size
- ⚠️  No SST support (inline strings only - larger files)
- ⚠️  Minimal styles (default only - no rich formatting)
- ⚠️  No row-stream API for workbook metadata such as merged ranges, comments, tables, and freeze panes

**In-memory workbook SAX/StAX write path** (✅ Working):
- ✅ `writeWorkbookStream` writes an already-materialized `Workbook` through the SAX/StAX backend
- ✅ Preserves full workbook metadata handled by the OOXML writer, including merges, comments, tables, row/column properties, and freeze panes
- ⚠️  Not a row-input streaming API; the `Workbook` is already in memory

**Read Path** (✅ P6.6 Complete):
- ✅ **True constant-memory streaming** - uses `fs2.io.readInputStream`
- ✅ O(1) memory for worksheet data (unlimited rows supported)
- ✅ Streams worksheet XML incrementally (4KB chunks)
- ⚠️  SharedStrings Table (SST) materialized in memory (~10MB typical, scales with unique strings)
- ✅ Large files (500k+ rows) process without OOM
- ✅ Memory tests verify O(1) behavior

**Result**:
- Both streaming **read and write** achieve constant memory for worksheet data ✅
- 500k rows: ~10-20MB memory (worksheet streaming + SST materialized)
- 1M+ rows supported without memory issues (unless >100k unique strings)
- **Design tradeoff**: SST materialization acceptable for most use cases (text typically <10MB)

### Security & Safety

**Implemented**:
- ✅ ZIP bomb detection
- ✅ XXE (XML External Entity) prevention
- ✅ Formula injection guards in in-memory and streaming writes

**Remaining**:
- ❌ XLSM macro preservation policy and tests (macros are never executed)

**Implemented (continued)**:
- ✅ Configurable file size limits via CLI `--max-size <MB>` (default 100MB; `--max-size 0` = unlimited)

### Advanced Features

**Completed** (P6, P7, P8, P31, WI-07/08/09, WI-10, WI-15, WI-17):
- ✅ P6: CellCodec primitives (9 types with auto-formatting)
- ✅ P7: String interpolation Phase 1 (runtime validation for all macros)
- ✅ P8: String interpolation Phase 2 (compile-time optimization)
- ✅ P31: Optics, RichText, HTML export, enhanced ergonomics
- ✅ **Formula System** (WI-07/08/09): Parser, evaluator, 108 functions, dependency graph, cycle detection
- ✅ **Excel Tables** (WI-10): Structured data with headers, AutoFilter, styling
- ✅ **Benchmarks** (WI-15): JMH performance suite (XL vs POI)
- ✅ **SAX Write** (WI-17): Fast SAX/StAX streaming write path
- ✅ **Security Hardening** (WI-30): ZIP bomb detection, XXE prevention, formula injection guards

**Future (and recently shipped)**:
- ❌ P6b: Full case class codec derivation (Magnolia/Shapeless)
- ❌ P9: Advanced macros (path macro, style literal)
- ✅ P10: Drawings — embedded pictures (#221): `Sheet.addImage`/`pictures`/`removeDrawing`, three anchor forms, 7-format classification, sha-deduped media, hybrid byte-preservation of non-picture drawings (shapes as `Drawing.Preserved`); shape *authoring* still future. In-memory only — see LIMITATIONS §13
- ✅ P11: Charts (0.12.0, #222): typed bar/line/pie via `Chart.bar`/`line`/`pie` + `Sheet.addChart` (CLI `chart add`); out-of-fence/Excel-authored charts stay byte-preserved — see LIMITATIONS §12
- ❌ Pivot Tables (remaining part of P12)

---

## Next Steps

> **For detailed roadmap and future plans, see [plan/roadmap.md](plan/roadmap.md)**

### Priority 1: P6.5 - Performance & Quality Polish

**Focus**: Address PR review feedback
- Optimize style indexOf from O(n²) to O(1)
- Extract whitespace check utilities
- Add error path tests
- Full round-trip integration tests

### Priority 2: P6b - Full Codec Derivation

**Focus**: Automatic case class mapping
- Derive RowCodec[A] for case classes
- Header-based column binding
- Type-safe bulk operations

### Priority 3: P9 - Advanced Macros

**Focus**: Additional compile-time validation
- `path` macro for file path validation
- `style` literal for CellStyle DSLs
- Enhanced diagnostics

### Priority 4: P10-P13 - Advanced Features

**Focus**: Drawings, Charts, Tables, Security
- See [plan/roadmap.md](plan/roadmap.md) for detailed breakdown

---

## File Structure

### Completed Modules
```
xl/src/com/tjclp/xl/
└── scripting.scala        ✅ One-import scripting prelude (com.tjclp.xl.scripting)

xl-core/src/com/tjclp/xl/
├── addressing/            ✅ Opaque types (Column, Row, ARef packing), CellRange, SheetName
├── cells/                 ✅ Cell, CellValue, CellError, Comment
├── sheets/                ✅ Sheet, SheetView, PageSetup
├── workbooks/             ✅ Workbook, WorkbookMetadata, DefinedName
├── patch/                 ✅ Patch Monoid (incl. MergeBorder)
├── styles/                ✅ CellStyle, Font, Fill, Border, Color, NumFmt, StylePatch, style DSL
├── codec/                 ✅ CellCodec (9 primitive types), readTyped/readTypedOr/readTypedOpt
├── macros/                ✅ ref"", fx"", money"" … compile-time literals (lives in xl-core; no separate xl-macros module)
├── optics/                ✅ Lens, Optional, focus DSL
├── richtext/              ✅ TextRun, RichText, DSL extensions
├── formatted/             ✅ Formatted literals + FormattedParsers.detect
├── render/                ✅ HTML/SVG renderers (sheet.toHtml / toSvg)
└── dsl/                   ✅ Ergonomic patch operators (:=, ++, outlined)

xl-ooxml/src/com/tjclp/xl/ooxml/
├── ContentTypes.scala     ✅ [Content_Types].xml
├── Relationships.scala    ✅ .rels files
├── Workbook.scala         ✅ xl/workbook.xml
├── worksheet/             ✅ xl/worksheets/sheet#.xml (RichText, merges, hyperlinks, page setup)
├── SharedStrings.scala    ✅ xl/sharedStrings.xml (SST with RichText)
├── Styles.scala           ✅ xl/styles.xml
├── XlsxWriter.scala       ✅ ZIP assembly + surgical modification
└── XlsxReader.scala       ✅ ZIP parsing + security limits

xl-cats-effect/src/com/tjclp/xl/io/
├── Excel.scala            ✅ Algebra trait
├── ExcelIO.scala          ✅ Interpreter with true streaming
└── Sax/StAX + fs2 writers ✅ Event-based streaming read/write
```

### Completed Modules (Additional)
- `xl-evaluator/` ✅ **Complete** (WI-07/08/09 - formula parsing, evaluation, 108 functions, dependency graph, structural editing, recalculation)
- `xl-cli/` ✅ **Complete** (stateless `xl` CLI: 46 subcommands, 27 batch ops, rendering, streaming mode)
- `xl-agent/` ✅ **Complete** (AI agent benchmark runner)
- `xl-benchmarks/` ✅ **Complete** (WI-15 - JMH performance benchmarks)

### Not Started (Future Phases)
- `xl-testkit/` (law helpers, golden test framework) — still a placeholder

> Drawings and charts shipped in 0.12.0 as `com.tjclp.xl.drawings` / `com.tjclp.xl.charts` **within `xl-core`**, not as separate modules.

---

## Technical Debt

### Completed ✅
1. ~~StreamingXmlWriter compilation~~ - ✅ fs2-data-xml integration complete
2. ~~DateTime serialization~~ - ✅ Excel serial number conversion implemented
3. ~~Cell → CellStyle linkage~~ - ✅ StyleRegistry provides sheet-level style management
4. ~~Comments~~ - ✅ Full OOXML round-trip (xl/commentsN.xml + VML drawings), rich text support, 12+ tests
5. ~~Merged cells~~ - ✅ `mergedRanges` serialized via `<mergeCells>` (in-memory + `writeWorkbookStream` paths)
6. ~~Column/row properties~~ - ✅ width/height/hidden/outline serialized via DirectSaxEmitter
7. ~~Hyperlinks~~ - ✅ serialized with worksheet relationships (0.10.0)

### Remaining
1. **Theme resolution** - Improve Theme color ARGB approximations (currently functional but not perfect)

---

## Performance Results (Actual)

> ⚠️ **Captured on an earlier release (JDK 25, Apple Silicon); not re-validated for 0.12.x — treat as indicative.** 0.12.1 made codec `put` paths ~2.4x faster (#297), not reflected below.

### JMH Benchmark Results (WI-15) - XL vs Apache POI

**XL vs Apache POI** (Apple Silicon M-series, JDK 25):

#### Streaming Reads (SAX Parser - Production Recommendation)
| Rows | POI | XL | Result |
|------|-----|----|--------|
| **1,000** | 1.357 ± 0.076 ms | **0.887 ± 0.060 ms** | ✨ **XL 35% faster** |
| **10,000** | 7.773 ± 0.590 ms | 8.408 ± 0.153 ms | Competitive (XL within 8%) |

#### In-Memory Reads (For Modification Workflows)
| Rows | POI | XL | Result |
|------|-----|----|--------|
| **1,000** | 1.650 ± 0.055 ms | **1.225 ± 0.086 ms** | ✨ **XL 26% faster** |
| **10,000** | 13.784 ± 0.377 ms | 14.115 ± 1.250 ms | Competitive (XL within 2%) |

#### Writes
| Rows | POI | XL | Result |
|------|-----|----|--------|
| **1,000** | 1.280 ± 0.041 ms | 1.906 ± 0.245 ms | POI 49% faster |
| **10,000** | 10.228 ± 0.417 ms | 15.248 ± 1.315 ms | POI 49% faster |

**Key Findings**:
- ✨ **XL is fastest for small-medium files** (< 5k rows): 35% faster streaming, 26% faster in-memory
- ✅ **XL competitive on large files**: Within 8% of POI on 10k row streaming reads
- 🔧 **Write optimization**: Future work (Phase 3) - POI currently 49% faster
- 💾 **Constant memory**: Streaming uses O(1) memory regardless of file size
- ⚡ **SAX parser**: 3.8x speedup vs previous fs2-data-xml implementation

**Recommendation**: Use `ExcelIO.readStream()` for production workloads (fastest for <5k rows, constant memory).

### Streaming Implementation (P5 + P6.6 Complete) ✅

**Memory characteristics**:
- **Write**: **~10MB constant memory** (O(1)) ✅
- **Read**: **~10MB constant memory** (O(1)) ✅ (P6.6 fixed with fs2.io.readInputStream)
- **Scalability**: Can handle 1M+ rows without OOM ✅

**Performance characteristics** (validated with JMH):
- **Streaming vs in-memory**: Streaming 1.7x faster for large files (8.4ms streaming vs 14.1ms in-memory @ 10k rows)
- **XL vs POI**: XL is fastest for small files (35% faster @ 1k rows), competitive for large files (within 8% @ 10k rows)

### Comparison to Apache POI (True Streaming Read + Write)

| Operation | XL | Apache POI SXSSF | Improvement |
|-----------|-----|------------------|-------------|
| **Write 100k** | 1.1s @ 10MB | ~5s @ 800MB | **4.5x faster, 80x less memory** ✅ |
| **Write 1M** | ~11s @ 10MB | ~50s @ 800MB | **4.5x faster, constant memory** ✅ |
| **Read 100k** | 1.8s @ 10MB | ~8s @ 1GB | **4.4x faster, 100x less memory** ✅ |
| **Read 500k** | ~9s @ 10MB | OOM @ 1GB+ | **Constant memory vs OOM** ✅ |

**Note**: The SXSSF comparison figures above are approximate (POI columns marked `~` are estimates); the JMH tables earlier in this section are the measured numbers. True O(1) streaming is verified by memory tests.

**Result for Writes**: Exceeded goal of 3-5x throughput, achieved 80x memory improvement with constant memory.

---

## Commands for Next Session

```bash
# Quick start
./mill __.compile
./mill __.test

# Work on streaming
./mill xl-cats-effect.compile
./mill xl-cats-effect.test

# Format code
./mill __.reformat

# Create sample file (once fixed)
./mill xl-ooxml.test.runMain com.tjclp.xl.ooxml.Demo
```

---

## Critical Success Factors

1. **Purity maintained** - Core is 100% pure, zero side effects
2. **Laws verified** - All Monoids tested with property-based tests
3. **Deterministic output** - Same input = same bytes (stable diffs)
4. **Zero overhead** - Opaque types, inline, compile-time macros
5. **Real files** - Creates valid XLSX that Excel/LibreOffice opens
6. **Type safety** - Opaque types prevent mixing units; codecs enforce type correctness
7. **Performance** - ~35% faster than Apache POI on streaming reads (JMH validated), writes currently ~49% slower, 80x less memory
