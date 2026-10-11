# A `.xl` source format for xl: research and design

**Status**: Research / design proposal (not accepted; no code). Prompted by the 2026-10-11 "Open source
Excel" exchange between Richie Caputo and Weaver (Appendix A).
**Codebase examined**: `origin/main` `6697409` (0.23.1 plus the unreleased Weaver wave).
**Method**: six parallel research lanes (prior art, Excel interop, fit with the code, agent ergonomics,
risk and effort, benchmark reproduction) synthesised into this document. Every claim is tagged where it
matters: **[V]** verified in the repo or against a cited source, **[I]** inferred from verified facts,
**[U]** could not be verified from the sandbox (Appendix B lists them all).

---

## 1. Executive summary and recommendation

**Recommendation: go-with-changes on Weaver's framing.**

Weaver's distinction holds up against everything the research found: *yes* to a `.xl` **source** format
that compiles to xlsx, *no* to a `.xl` **storage** format that replaces it. Seven historical data points
(XLSB, SpreadsheetML 2003, SYLK, DIF, Lotus, ODS, Gnumeric) say that no spreadsheet format has ever won
by being better; the one that displaced its predecessor did so by becoming the dominant app's default
Save As. Even the two from-scratch Rust engines of 2024 to 2026 (gridcraft, IronCalc) chose xlsx as
their native file. The efficiency problem Richie named is real but it is not the file's: the benchmark
lane reproduced the shape of Weaver's numbers (§8), and the 30x gap between a Parquet read and an xl
read is xl's reader and model, which a new file format would not touch.

What *is* worth building is the authoring model: a declarative, name-addressed, git-diffable text source
that compiles through the existing `Edit` algebra to an xlsx with **live formulas**, carries a lock of
computed values, and stamps its source hash into the xlsx. The four products that shipped "names, not
addresses" before (Quantrix, Anaplan, Grist, Causal) all exported *values* to Excel, which is exactly
why they stayed niche in deal work. The gap a `.xl` compiler fills is a name-addressed source whose
render a banker can trace with F2. Nobody has shipped that.

The changes to Weaver's plan, each argued in the body:

| # | Weaver's plan | Change | Why (section) |
|---|---|---|---|
| 1 | Build the compiler on the existing library | **Precondition**: finish the ADR-017 Wave 2 lowering so batch `putf` and the verbs run through `Edit` and the library recalc path. The two bugs Weaver found (bare XLOOKUP with no cached value; FILTER written as a scalar) are symptoms of the verb-layer/library-layer divergence that still exists at `WriteCommands.scala:478-612` [V]. A compiler gate that trusts verb-layer caches would pass both. | §5.4, §9 row 9 |
| 2 | `lock/values.parquet` | **The v1 lock is sorted text** (JSON Lines, gzip), produced in the pure core. Parquet is a later opt-in only after a measured native-image build on all five release platforms; Delta Kernel never enters the binary (Weaver's own position, confirmed: `delta-kernel-defaults` hard-depends on the 30 MB shaded Hadoop client runtime [V]). | §5.5, §9 rows 7 and 15 |
| 3 | Decompiler (step 4) | **v1 is exact-only** (read back the stamped source). The lossy decompiler for inherited banker models is a separate, open-ended track informed by the live-model proof, with a fidelity law per inference pass. Hand-edited xlsx is a **fork**, not a merge, in v1. | §4.2, §7.5, §9 rows 2 and 6 |
| 4 | "Keep Excel's formula language" | Confirmed, plus: compile names to **Excel's own name machinery** where it fits (defined names for scalars and rows, Tables for typed blocks) so the xlsx is self-explanatory without the `.xl`. Emit every formula cell with a cached value from the lock, because most readers of the xlsx (previews, phones, pandas, Google Drive) never recalculate. | §3.4, §4.5 |
| 5 | Stamp the source hash in a custom part | Stamp in **two carriers**: a custom document property and a hidden workbook-scoped defined name. Excel desktop keeps both; Google Sheets keeps names; whether Excel preserves an *unknown* custom part is unverified [U]. | §4.1 |
| 6 | Grammar derived from batch JSON and specs | Define the `.xl` AST as a **pure ADT** with a text grammar as its parser/printer under a round-trip law (the `FormulaParser`/`FormulaPrinter` discipline), compile to `Vector[Edit]`, and export the ADT builder from the scripting prelude so scripts can build models from data. A versioned `xl 1` header from the first line. | §5.2, §7.6 |
| 7 | Dynamic arrays | **Blocking dependency**: xl drops the `cm` marker and cannot author a spill today (issue #714, open; draft PR #721 unverified in Excel 365) [V]. Any model using FILTER/SORT/UNIQUE/XLOOKUP-arrays needs this first. | §4.4 |

**The "or not" case** (no xlsx at all): no as a direction, yes as a mode. Every consumer named in this
repo's own history opened an xlsx; without xlsx there is no LibreOffice/Excel oracle that defines
"correct" for the evaluator. `xl compile --no-xlsx` (source + lock only) captures the efficiency for
agent-internal iterations (§10.3).

**Conditions that would flip the decision to no-go** (§10.5): the proof model needs explicit placement
for a large share of its cells; humans must author `.xl` rather than review diffs; or the evaluator's
function gap (IPMT/PPMT, SUBTOTAL/AGGREGATE, the LAMBDA family) sits on the critical path of the first
live model.

---

## 2. Weaver's claims, scored

| Claim (Weaver, 2026-10-11) | Verdict | Evidence |
|---|---|---|
| "Adoption decides exchange formats, not efficiency" (XLSB) | **Confirmed** | Library of Congress: XLSB "not widely supported except in Microsoft products"; Google Sheets import list omits it; pandas needs a separate engine; gridcraft imports it values-only (§3.1) |
| XLSB is "3-5x faster" | **Qualified** | No primary Microsoft benchmark exists; vendor and forum figures are 2 to 3x faster open/save, up to ~50% smaller [U for the multiplier] |
| "Names, not addresses … impossible inside xlsx" | **Qualified** | Excel Tables with structured references survive row insertion; LET/LAMBDA/AFE give named formulas. What xlsx cannot hold is layout intent and a stable identity for a row outside a Table (§3.4) |
| "Keep Excel's formula language. Do not invent one." | **Confirmed** | OpenFormula took six years to agree a syntax while Excel's became de facto; every names-not-cells product that invented a language exported values (§3.3) |
| 200k x 8: CSV 10.7 MB, sheet XML 55.8 MB, 8.9 MB zipped, Parquet 3.0 MB | **Reproduced within the expected range** (dataset composition differs) | §8 |
| Read times: Parquet 0.04 s, `xl --stream` 7.6 s, in-memory 38-41 s | **Reproduced in shape** (JAR here, native image there) | §8 |
| "The other 30x is our implementation, not the file. A better reader would." | **Confirmed** | A plain SAX pass over the same sheet XML is far below xl's stream time (§8); roadmap #634 (streaming-default read engine) is the relevant item [V] |
| Delta Kernel too heavy for the native binary | **Confirmed** | `delta-kernel-defaults` 4.0.0 depends on `hadoop-client-runtime` (30.4 MB shaded) + `parquet-hadoop` 1.12.3 + Jackson; native image reflect-config today is 1.6 KB, Batik/Xerces only [V] |
| "Write plain Parquet locally and let Databricks commit it" | **Go one step simpler**: plain text | Parquet writers for the JVM drag Hadoop or are beta (Hardwood); text is zero-dependency, git-diffable and a first-class Auto Loader input (§5.5) |
| "xlsx to .xl is exact for files we compiled (stamp the source hash)" | **Confirmed with a carrier caveat** | Custom doc property and defined names survive Excel; unknown custom parts and `extLst` are unverified in Excel 365 (§4.1) |
| The two `putf` bugs (bare XLOOKUP uncached; FILTER scalar) | **Located, not yet filed** | `WriteCommands.scala:562-612` builds `CellValue.Formula(f, None)` and relies on the after-edit recalc; "spill" does not occur in the file [V]. Issue #714 covers the spill half [V]. Neither is in CHANGELOG/roadmap [V] |

---

## 3. Prior art

### 3.1 File formats: efficient ones lost, the default won

| Format | Shape | What happened | Lesson |
|---|---|---|---|
| BIFF `.xls` (1987-2007) | OLE2 binary | Default until 2007; still outnumbered `.xlsx` on the web in 2019 by the LoC's count [U, summary-only] | A legacy default outlives its replacement by decades |
| OOXML `.xlsx` (2007-) | zip of XML | Won by being Office 2007's default Save As, not by compactness | Default + openness beat efficiency |
| **XLSB** (2007-) | same zip, binary records, [MS-XLSB] under the OSP | LoC: "not widely supported except in Microsoft products"; omitted from Google Sheets' import list; pandas needs `pyxlsb`/`calamine`; openpyxl cannot read it; gridcraft imports values only | Same vendor, same model, same package, faster, documented: zero exchange adoption |
| SpreadsheetML 2003 | single text XML file | Microsoft's own *text* spreadsheet format for two releases; not full fidelity; superseded 2007 | Text alone is not a draw |
| SYLK (1986), DIF (1981) | ASCII | Died with their apps | A text interchange format with formulas existed in 1986 |
| Lotus WK* | binary, spec public domain | Dead with the product | A public spec did not save it |
| ODS / OpenFormula | zip of XML; `of:=` formulas | ODF 1.0 shipped without a formula spec; OpenFormula only in 2011; Microsoft wrote OOXML formulas into ODS instead | Six years to agree a syntax; the incumbent's became the standard |
| Apple Numbers | Snappy protobuf, private schema | Apple-only, reverse-engineered per build | A private binary schema is the opposite of what an agent source needs |
| gridcraft (2026, Rust), IronCalc (2024, Rust) | xlsx is the native format | Both from-scratch engines chose xlsx | Even a 2026 engine does not invent a file |

Sources: LoC fdd000510/fdd000512; support.google.com/docs/answer/40608; pandas `ExcelFile` docs;
github.com/storytold/gridcraft; github.com/ironcalc/IronCalc; Wikipedia OpenFormula/SpreadsheetML (the LoC,
Microsoft and Wikipedia pages were summary-only from the sandbox).

### 3.2 Text-diffable and git-friendly spreadsheets

All existing text spreadsheet formats are **cell-addressed**: pyspread `.pysu` (one line per cell, Python
expressions), sc-im `.sc` (`let A1 = …` commands), VisiData `.vdj` (a replayable edit log, i.e. the
batch-JSON idea), flat ODS `.fods`. Flat ODS in git is the cautionary tale: a one-cell edit produced a
297+/297- diff because LibreOffice renumbers automatic styles on every save (libreoffice.org users list
2016/msg00810; `fods_minimizer` exists to strip them). **Generated ids and styles kill diffability**; a
`.xl` must emit styles by role and carry no timestamps or GUIDs in source. Git textconv drivers and Git XL
(xlwings; diffs VBA only in the open tool) show demand for Excel-in-git, and that diffing the *render* is
lossy and never supports merge.

Grist is the best shipping example of names-not-addresses: Python formulas over `$Column`, with rename
propagation into other formulas (support.getgrist.com/formulas). It chose SQLite plus its own action
history over text plus git; its xlsx export's formula handling is undocumented [U].

### 3.3 "Names, not addresses" has shipped four times, and all four exported values

| Product | Model | Excel hand-off | Status |
|---|---|---|---|
| Quantrix Modeler (2004-) | matrices over named categories; one formula per line item; "eclipse" conflicts when two formulas hit a cell | "Formulae are not included in the Export from the Qloud"; formula export only as CSV text | 20 years, niche; "significant learning curve" |
| Anaplan (2006-) | formulas belong to line items; dimensions apply automatically; no textual model language | Excel Add-in pushes data only, <= 1M cells | Platform, not a file |
| Grist (2017-) | named columns, Python formulas | one sheet per table; formula preservation undocumented [U] | Open source, own VCS |
| Causal (2019-2024) | variables x dimensions x time; purest names-not-cells | **values only** ("Raw Values" / "Formatted Values" export) | Acquired by Lucanet 2024; product reportedly folded [U] |

The lesson is the strongest single finding of the prior-art lane: **the compiler must emit live,
traceable Excel formulas**, lowered to Excel's own constructs (defined names, Tables) wherever possible,
or the partner gets a dead model and `.xl` joins this list.

### 3.4 Excel's own partial answers

| Weaver's claim | What Excel already does | What it still cannot do |
|---|---|---|
| Names, not addresses | Defined names (workbook/sheet scope); LET; LAMBDA with named functions (MSR: Turing-complete); Excel Labs Advanced Formula Environment edits named formulas as text modules with gist import | A name's `refersTo` is A1 text; names are untyped, undimensioned, carry no layout intent |
| Inserting a row never breaks a reference | **Tables**: structured references (`Table1[Sales]`, `[@Sales]`) adjust on insert/delete; calculated columns fill down | Only inside a Table. A defined name over a fixed range grows for an insert *inside* it but not for a row appended at its edge (hence the OFFSET/COUNTA folk pattern). Tables cannot hold banker free layout (merged bands, mixed orientation) |
| Dynamic arrays | Native since 2018 | xl cannot author them yet (#714) |
| Formulas as source | AFE, Power Query M (stored in a binary `DataMashup` part), Python in Excel `xl()` | Microsoft built "named formulas as text modules" and left it in Labs; none of it is a source of truth in the file |
| Elastic row declarations | MSR "Elastic Sheet-Defined Functions" (JFP 2020): variable-size inputs with principal generalisation | Exactly the semantics `ebitda[Y1:Y5] = prev * (1 + growth)` needs when Y5 becomes Y7 |

Relevant repo facts [V]: `DefinedName` with all twelve ECMA attributes and a lazy index
(`xl-core/src/com/tjclp/xl/workbooks/DefinedName.scala:51`); `TableSpec` exists and is written
(`xl-core/src/com/tjclp/xl/tables/TableSpec.scala:143`, `xl-ooxml/.../Table.scala`) but **structured
references and LAMBDA do not parse** (`xl-evaluator/src/com/tjclp/xl/formula/parser/UnparseableFormula.scala:14`);
LET is a parser special form (`FormulaParser.scala:929`).

### 3.5 Text-to-xlsx compilers and decompilers

Thin and embedded: Haskell `clerk` (monadic placement eDSL), F# `Kookerella.FsOpenXmlDsl` (bidirectional:
it reads an xlsx back into the DSL, a small decompiler precedent), .NET `ExcelReportLib` (text DSL ->
parser -> workbook AST -> layout engine -> renderer, architecturally the closest to Weaver's pipeline),
Python `pyproforma`/`PyCashflow` (named line items, values out). Decompilers `pycel` and `excel_to_code`
name the hard edge: INDIRECT/OFFSET must be statically resolvable. Nobody has shipped a standalone,
git-native, name-addressed source whose compiler emits a banker-openable xlsx with live formulas and whose
decompiler reads one back.

Columnar grids (Row Zero, Quadratic with Parquet ingest, Gigasheet, FINOS Perspective over Arrow) all keep a
scalar, cell-addressed formula layer over a columnar data store. No shipping product stores formulas
columnarly. Weaver's `data/*.parquet` plus scalar source split matches where the field converged.

---

## 4. Excel interop findings

### 4.1 Stamping provenance into the xlsx

| Carrier | Excel desktop re-save | Excel web | Google Sheets | LibreOffice | xl today [V] |
|---|---|---|---|---|---|
| `docProps/custom.xml` custom property | Kept (first-class feature) | Kept [I] | Dropped [I, U] | Kept | Preserved verbatim, **not authored** (`DocProps.scala` writes core/app only) |
| `customXml/item1.xml` part | Kept; Document Inspector offers to strip it | Kept but `itemProps` loses schema refs (office-js #3160) | Dropped [I] | Dropped (ONLYOFFICE dropped until 7.0) | Preserved verbatim, not authored |
| Private part + custom relationship | **Unknown for Excel** (Word documents preservation; no Excel statement found) [U] | [U] | Dropped | Dropped | Preserved if present |
| `extLst/ext uri="{guid}"` | "Intended" to be retained per an SC34 engineer; a 2013 experiment found inconsistent discarding [U] | [U] | Dropped | Known x14 kept, rest dropped | Preserved |
| Hidden workbook-scoped defined name | Kept, and adjusted through inserts | Kept | Kept (Sheets exports named ranges) | Kept | **Authored today** (`Edit.DefineName`, `name add`, `hidden` attr, #696) |

Recommendation: two carriers, a custom document property (`xl.source = sha256:…; grammar=1; lock=…`,
user-visible, survives Excel and LibreOffice) and a hidden defined name `_xl.src` (survives every engine
that keeps names). The stamp proves lineage, never intactness (§4.2). Authoring `custom.xml` is a small
`xl-ooxml` addition; the defined name needs nothing new.

### 4.2 What Excel rewrites on save: "two truths" detection

Excel never copies a package through; an edit-free open/save rewrites `calcChain` (regenerated), shared
strings (re-deduplicated), `styles.xml` (renumbered), formula text (re-serialised with `_xlfn.`/`SINGLE`/
`ANCHORARRAY`), `xr:uid`/`revisionPtr` GUIDs, `calcId`, `app.xml`, and every view. Byte or hash comparison
of the xlsx is therefore dead on arrival. The usable definition is **model-level**: read the xlsx with xl,
normalise (formula AST via the parser/printer, resolved `CellStyle.canonicalKey`, numeric values), and diff
against the lock's cell-to-symbol map. Signals in increasing cost: stamp present and matching (lineage);
`app.xml`/`revisionPtr` changed (touched, cheap, not proof); formula-AST or style diff on a mapped cell
(edited, authoritative); value diff on input cells (edited inputs); value diff on formula cells only (stale
cache or engine divergence, a calc question). `xl diff -g` and `WorkbookLint` are most of this already [V].

Because Excel rewrites `refersTo` and table `ref` on structural edits, defined names and Tables emitted by
the compiler are the anchors that let a hand-inserted row still map back to a symbol.

### 4.3 Opening a `.xl` directly in Excel: friction, not a path

Windows `ExtensionHardening` prompts on an extension/content mismatch and a text `.xl` falls through to the
text-import wizard; macOS UTIs cannot be added to Excel's own handler list; Power Query custom connectors
(`.mez`) load only in Power BI Desktop; VSTO is Windows-only. Realistic: `xl compile model.xl -o model.xlsx`
in the CLI, a shell association on a helper that compiles and opens, and later an Office.js task pane using
`Excel.createWorkbook(base64)` to open a compiled workbook. Nothing in Excel parses the source.

### 4.4 Fidelity: what the compiler must emit and what xl writes today [V]

| Feature | xl today | Gap for a compiler |
|---|---|---|
| Defined names | Full model, authoring, resolution, rename rewriting | None; policy: one name per symbol, absolute refs |
| Tables (ListObject) | Authored and written; structured refs **not parsed** (LIMITATIONS §15) | Medium diff in xl-evaluator if `.xl` formulas reference Tables |
| **Dynamic arrays / spills** | `cm` dropped on every write; no authoring (#714 open; PR #721 draft, 3,295 additions, "not checked in Excel itself") | **Blocking** for FILTER/SORT/UNIQUE/SEQUENCE/XLOOKUP-array models; also needs spilled cells written as constants for cache-only readers |
| Implicit intersection `@` | Modelled (#604) | Compiler decides legacy vs array per formula |
| Data tables (what-if) | Byte-exact model, seeded by `recalc --tables`; no authoring sugar (GH-419) | Small-medium: an `Edit.AddDataTable` |
| Charts | bar/line/pie typed; no scatter/combo/waterfall/chartex | Large for a finance deck; v1 declares charts and refuses unsupported kinds loudly |
| Conditional formatting | Six families | Small-medium for the rest |
| Data validation | List dropdowns only | Small-medium |
| Comments | Legacy authored; threaded preserved only | Notes now, threaded later |
| Sheet/workbook protection | Preserved, not authored | Small; LP models often lock input sheets |
| Outlines, number formats, views, print, images, hyperlinks, merges | Full | None |
| Doc properties | core/app written (GH-242); `custom.xml` not | Add custom.xml (§4.1) |

### 4.5 "Lossless xlsx export", defined

Byte-identity is meaningless against Excel. Proposed definition: an export is lossless when (a) it opens in
Excel 365 desktop and web and in LibreOffice without a repair prompt; (b) after Excel's full recalculation
every formula cell equals the lock's value within 1e-9 relative (dates, text, booleans and errors exact);
(c) `xl decompile` of the export reproduces the source modulo canonical formatting; (d) `xl decompile` of
the file after an edit-free Excel open/save reports no cell-level diff against the lock. (b) and (d) are CI
tests with LibreOffice headless plus a periodic Excel pass; (c) is a pure property test.

Out of scope for compiled models and carried as opaque asset hashes for inherited ones: VBA (xl strips it
today, LIMITATIONS §25), pivots, external links, OLE/ActiveX, Power Query, slicers, chartex, shapes.

### 4.6 Cached values: why the lock matters

Excel recalculates on open when `fullCalcOnLoad=1` or `calcId` is absent or old (then *all* formulas).
LibreOffice shows whatever `<v>` is present and computes only cells lacking one; its default ignores the
xlsx flags. Google Sheets recalculates with its own engine (`#NAME?` for unsupported functions). Numbers
substitutes cached values. Quick Look, Outlook preview, SharePoint cards, pandas, openpyxl `data_only` and
`xl view` without `--eval` are cache-only [I, the preview renderers are U]. Consequences: every formula
cell carries a `<v>` from the lock; `fullCalcOnLoad` is a fallback for cells the evaluator could not
compute, never a strategy (it also dirties the workbook on open); omit `calcId` (xl already does,
`Workbook.scala:377`) so Excel recalculates once as a free consistency check; spills need both the anchor
metadata and constants in the spilled cells; data tables must be seeded before export.

---

## 5. Proposed architecture

### 5.1 The line to hold

Weaver's: source plus compiler plus lock; no new UI, no new formula language, no new engine beyond what xl
needs anyway. Operationally: **every `.xl` construct compiles to an xlsx construct with no warning**, and
every formula inside a `.xl` is Excel's grammar with names in place of references. The only syntax that is
not Excel's is the reference form (`name`, `name[Y1]`, `name[prev]`, `Sheet.name`), which the compiler
lowers.

### 5.2 Module placement [V for the constraints, I for the placement]

ADR-017 §2.1 (`docs/design/agent-first-architecture.md:130-139`) forbids the evaluator in `xl-core` and
IO/JSON in `xl-evaluator`; the purity charter keeps `xl-core`/`xl-ooxml` pure. Type-checking names and
cycles needs `FormulaParser`, `NameWalk` and `DependencyGraph`, so the front-end sits above the evaluator:

```
xl-core  ←  xl-ooxml  ←  xl-cats-effect
xl-core  ←  xl-evaluator
{xl-core, xl-evaluator}  ←  xl-source   (NEW, published, pure)
{xl-cats-effect, xl-evaluator, xl-source}  ←  xl-cli
{core, ooxml, cats-effect, evaluator, source}  ←  xl   (aggregate + prelude)
```

| Piece | Module | Notes |
|---|---|---|
| `.xl` AST (`enum ModelDecl` and friends) | `xl-source` (or `xl-core` beside `Edit` if the prelude must build it without the evaluator) | Pure ADT; the text grammar is its parser and printer with a `parse . print = id` law |
| Parser / printer / `fmt` | `xl-source` | Source positions on every node (the diagnostics in §7.3 depend on them) |
| Layout engine (names -> cells) | `xl-source` | Pure `SourceModel => Either[LayoutError, Placement]`; only addressing types from core |
| Compiler (placed model -> `Vector[Edit]` -> `Workbook`) | `xl-source` | Targets `Workbook.edit(...)` (`Workbook.scala:222-227`) and `EditInterpreter` unchanged; gives `--print-batch` for free |
| Gate (recalc, checks, refusals) | `xl-source` calling `Workbook.recalculate` and `evaluateFormula` | On the **library** recalc path, never verb-layer caches |
| Lock encoder (`RecalcResult` + placement -> bytes) | `xl-source` (pure); file write in `xl-cli` / prelude facade | Mirrors `XlsxWriter` pure bytes / `Excel.write` IO |
| Stamp (custom property + hidden name) | `xl-ooxml` (custom.xml authoring, small) + `Edit.DefineName` | The one place `xl-ooxml` changes |
| Decompiler | `xl-source` | Exact in v1; lossy later |
| Verbs `compile`, `check`, `decompile`, `fmt` | `xl-cli` (`Command.scala`, `contract/Schema.scala:202`, goldens, `DocsGenSpec`) | Thin forwarders published in `xl schema --json` |
| Prelude | `xl` aggregate: one `export` + a probe | `XlSource.compile(text)` and the ADT builder |

**Untouched for the minimal slice**: `xl-cats-effect`, and `xl-evaluator` unless the grammar needs
structured references or implicit intersection. Candidate core change: `Edit.AddTable(TableSpec)` (one enum
case, interpreter arm, `EditSchema` row) once typed tables enter the grammar; `Edit.AddDataTable` later.

### 5.3 How the algebra already covers the target language [V]

`Edit` (`xl-core/src/com/tjclp/xl/ops/Edit.scala:37-179`, 49 cases) is exactly "a source format with
addresses": `Put/PutValues/PutFormula/PutFormulas/DragFormula`, `Style(StyleOverlay, StyleMode)`,
`DefineName`, `AddSheet`, `AddChart`, `AddConditionalFormat`, `AutoFit`, `Freeze`, `InsertRows`,
`GroupRows`, and the rest. Batch JSON (34 ops in the generated `batch-ops.md`; the brief's 36 overcounts) is
its address-bound JSON encoding. What a declarative source adds sits *above* the algebra: symbols, axes,
placement, roles, checks. What the algebra lacks: `AddTable` and data-table/validation authoring.

Important nuance: the CLI's `BatchParser.applyOne` (`helpers/BatchParser.scala:1239-1253`) still applies
ops with its own appliers and `applyPutFormulaDragging` calls `FormulaShifter` directly; ADR-017 §2.12
("batch JSON becomes the codec of `Vector[Edit]`", #583) is unfinished. The compiler would be the first
full-coverage consumer of the algebra and the strongest argument for finishing W2.2.

### 5.4 How the evaluator type-checks and computes the lock [V]

- `TExpr.NameRef`/`SheetNameRef` parse today (`ast/TExpr.scala:420,433`; `FormulaParser.scala:1143-1149,
  1308-1311`); names never shift on drag. A static "unknown name" check needs only the symbol table.
- `NameWalk` (#704, `eval/NameWalk.scala`) gives stack-safe name resolution and name-to-name cycle detection;
  `DependencyGraph.fromWorkbook/detectCycles/qualifiedSccOrder` give cell cycles before writing.
- `Workbook.recalculate` returns `RecalcResult(workbook, evaluated: Map[SheetName, Map[ARef, CellValue]],
  errors, converged, …)` (`eval/Recalc.scala:315-322`): **this is the lock table**. `evaluateFormula`
  evaluates a `check` without writing a cell.
- Missing: a relative-row token (`prev`), structured references, R1C1, and a verified implicit-intersection
  rule for a row-shaped name used bare inside another row's formula [U]. The compiler resolves all of these
  *after placement*: `prev` becomes the cell one column left at the anchor; a bare row name becomes the
  same-column cell; the row also gets a `DefineName` for humans and the decompiler; then one
  `Edit.DragFormula` per row, and `FormulaShifter` (laws at `printer/FormulaShifter.scala:40-46`) does the
  rest.
- The gate is `recalculate` + `cacheResults` + every `check` + refusal on unsupported functions, errors,
  non-convergence (exit 1, nothing written, the `--strict` rule since #677). Precondition: the single-verb
  `putf` cache/spill divergence retired under `Edit`.

### 5.5 The lock and the data: text now, Parquet maybe, Delta never in the binary

| Option | Native image | "A directory is not an attachment" | Git review | Databricks |
|---|---|---|---|---|
| A. Embedded Delta Kernel | Worst: `delta-kernel-defaults` 4.0.0 -> `hadoop-client-runtime` 30.4 MB + `parquet-hadoop` 1.12.3 + Jackson; reflection/ServiceLoader-driven; no published reachability metadata [V] | Fails | None | Native |
| B. Plain Parquet from xl | `parquet-hadoop` needs Hadoop `provided`; `carpet` depends on `hadoop-common` compile; `parquet-floor` re-implements the Hadoop shims but native-image reachability is unproven; Arrow maintainers: "GraalVM is not a currently supported configuration" (arrow-java #866); Hardwood is zero-dep and native-claimed but its writer is beta (1.1.0.Beta1) [V for the POMs; U for native builds] | One binary file | Blob | Auto Loader |
| **C. Sorted JSON Lines (gzip) from xl** | Trivial: `upickle` is already an xl-cli dependency; bytes produced purely with sorted rows and `BigDecimal.toPlainString` | One file | A changed cell is a changed line | First-class Auto Loader input |

Verdict: **C for v1**. Parquet later, only if a local columnar consumer (DuckDB, notebooks) needs it *and* a
measured five-platform native build is accepted; a hand-rolled single-row-group PLAIN Parquet writer (the
footer encoder is the only real work) fits xl's no-POI posture better than any dependency. Delta stays
where Weaver put it: a Databricks job commits `models.cells` keyed on `(model_id, source_hash, version)`.
Open data-platform question: snapshots vs changed-cells-only, which moves volume by two orders of magnitude
(a 20-sheet deal model is roughly 40k to 200k populated cells; 50 to 500 compiles per lifecycle [I]).

The `data/*.parquet` bulk blocks of Weaver's container are a **reader** question, not a lock question: they
only pay off through the streaming writer (ADR-011) and the deferred #634 engine. Non-goal for v1.

### 5.6 Proposed container

```
deal-falcon/
  model.xl              # source (git)
  data/comps.csv        # bulk blocks, v1 as CSV; Parquet when the reader can stream it
  lock/values.jsonl.gz  # computed values + cell-to-symbol map + source hash (git, regenerated by compile)
  build/model.xlsx      # the render (not in git; stamped with sha256(model.xl))
```

---

## 6. Grammar sketch and worked examples

The grammar is a sketch for discussion, not a specification. Design rules: Excel formula syntax verbatim
inside expressions; names everywhere else; one axis declaration per block; styles by role; checks as
statements; a versioned header; explicit placement only through the escape hatches of §6.3.

```
file      := "xl" version NL decl*
decl      := sheet | styles | table | check | grid | raw
sheet     := "sheet" Name ("at" SheetName)? NL (axis | inputs | rows | place | formatBlock)*
axis      := "years" Label ".." Label                      -- or: axis <name> = a..b
inputs    := "inputs" NL (ident "=" literal (":" role)? ("@" A1)? ("\"" label "\"")? NL)+
rows      := "rows" NL rowDecl+
rowDecl   := ident ("[" period "]")? "=" excelExpr (":" role)? ("\"" label "\"")? NL
period    := Label | Label ".." | ".." Label | "prev" | "first" | "last"
place     := "place" ident "at" A1                         -- pin a row's first cell / a scalar
table     := "table" ident ("at" Ref)? NL columns NL (rowLit | "from" path) ...
check     := "check" ident ":" excelExpr ("\"" message "\"")?
styles    := "styles" NL (ident "=" styleKeys NL)+          -- roles: header, input, total, currency...
grid      := "grid" ident "at" Ref NL (A1 ("=" excelExpr | literal) ("style" ident)? NL)+
raw       := "raw" Ref "from" path                          -- copied verbatim, never recompiled
excelExpr := Excel formula grammar where a reference may be: ident | ident[period] | Sheet.ident[...]
```

Reference forms inside a row formula: `revenue` is the same period (broadcast), `revenue[prev]` the previous
period, `revenue[Y1]` an absolute period, `Model.ebitda` a symbol on another sheet. A `rows` formula may never
contain an A1 address; if an agent needs one there, the thing should be an input, an alias, or a `grid` cell.

### 6.1 Example 1: names, not addresses (the minimal three-statement fragment)

```
xl 1

sheet Model
  years FY2025..FY2029

  inputs
    revenue_fy25 = 1_000_000 : currency  "Revenue FY2025"
    growth       = 5%                    "Revenue growth"
    margin       = 30%                   "EBITDA margin"
    capex_pct    = 4%                    "Capex % revenue"

  rows
    revenue[first] = revenue_fy25                    : currency
    revenue        = revenue[prev] * (1 + growth)    : currency
    ebitda         = revenue * margin                : currency
    capex          = revenue * capex_pct             : currency
    fcf            = ebitda - capex                  : currency  "Free cash flow"

check ebitda_le_revenue: MIN(revenue - ebitda) >= 0  "EBITDA cannot exceed revenue"

styles
  header = bold, fill navy, font white, center
  input  = font blue
```

What the compiler does (default layout: labels in A, periods from B, inputs block then rows block, one blank
row between, `header` applied to the axis row, `input` to input cells):

| Source symbol | Cells | Emitted |
|---|---|---|
| axis row | `Model!A1:F1` | `Put` labels, `Style(header)` |
| `growth` | `Model!B4` | `Edit.put(…, 0.05, Explicit(Percent))`, `Style(input)`, `DefineName("growth", "Model!$B$4")` |
| `revenue` | `Model!B8:F8` | `PutFormula(B8, "=$B$3")`, `DragFormula(C8:F8, "=B8*(1+growth)", anchor C8)`, `DefineName("revenue", "Model!$B$8:$F$8")` |
| `ebitda` | `Model!B9:F9` | `DragFormula(B9:F9, "=B8*margin", anchor B9)`, `DefineName("ebitda", …)` |
| `check` | not a cell by default | evaluated with `evaluateFormula` on the recalculated book; `--checks visible` also writes a flagged cell |

The xlsx a partner opens contains `=C8*(1+growth)` in D8 (live, traceable, name-qualified) and the names
`growth`, `margin`, `revenue`, `ebitda`, `capex`, `fcf` in the Name Manager. The lock records
`Model!D8 -> revenue[FY2027] = 1,102,500`. The three review diffs from §7.2 are each one to three lines.

### 6.2 Example 2: the placement escape hatch (a banker-style cover sheet)

```
xl 1

sheet Summary at "Summary"
  inputs
    irr = XIRR(Model.fcf, Model.dates) @C12     : pct1   "Sponsor IRR"   -- pin ONE cell
    moic = SUM(Model.fcf) / Model.equity_in @C13 : mult   "MOIC"

  rows
    ebitda_summary = Model.ebitda                 : currency   -- row still spans Model's years
  place ebitda_summary at B20                      -- pin the row's first cell

grid Cover at Cover!A1
  A1  "Project Falcon"                       style title
  A2  "Sources & Uses"                       style header
  B4  = Model.ebitda[FY2025]                 style currency
  C4  = XLOOKUP(B4, 'Comps'!A:A, 'Comps'!D:D) style currency   -- verbatim Excel inside a grid
  A6:D6 merge                                style band

raw Cover!H40:K55 from "legacy/falcon_v7.xlsx"   -- the decompiler's residue; never recompiled
```

Rules for mixing: reservation order is `raw` > `grid` > `@`/`place` pins > automatic layout; automatic layout
never enters a reserved cell; overlapping reservations are `XL_PLACEMENT_COLLISION` naming both statements
and the range. `grid` cells may read names (`Model.ebitda[FY2025]`); names may read grid cells only through a
declared alias (`alias cover_total = Cover!B4`), so a grid cannot become an invisible dependency. `raw` is
opaque to checks; `audit` still runs over the compiled book. `xl check --layout` prints every symbol's
resolved address so an agent reads it back instead of guessing.

### 6.3 Example 3: a typed table, a bulk block and a check over it

```
xl 1

table comps at Comps!A1                      -- compiles to an Excel Table "comps"
  columns name: text, ev: currency, ebitda: currency, multiple = ev / ebitda : mult
  from "data/comps.csv"                      -- v1 CSV; Parquet when the streaming reader lands

sheet Valuation
  inputs
    target_ebitda = Model.ebitda[FY2029]     : currency
  rows
    median_multiple = MEDIAN(comps.multiple)                   : mult
    implied_ev      = target_ebitda * median_multiple          : currency

check comps_nonempty: ROWS(comps) >= 5  "Need at least five comps"
check multiples_sane: AND(MIN(comps.multiple) > 0, MAX(comps.multiple) < 40)
```

Today `comps.multiple` must compile to an A1 range (`Comps!$D$2:$D$40`) because the evaluator does not parse
structured references; the xlsx still gets a real Table with a calculated column so Excel users see
`comps[multiple]`. When the evaluator learns structured references (a medium diff in `xl-evaluator`), the
compiler can emit them directly and the table grows safely under hand edits. `ROWS`/`MEDIAN`/`AND` are in the
registry [V]; the check runs on the library recalc path.

---

## 7. Agent ergonomics

### 7.1 What agents use today and what addresses cost them [V]

Two surfaces: the CLI with batch JSON ("batch is the primary write surface", `plugin/skills/xl-cli/SKILL.md`)
and scala-cli scripts with the prelude and the `Edit` algebra (`plugin/skills/xl-scripting/SKILL.md`). The
CHANGELOG and ADR-017 are a ledger of address-model costs paid by agents: three divergent sheet rules and 17
batch sites re-implementing "requires --sheet"; formula-blind `rename-sheet` (#559) and `insertRows`;
uncached writes shipped by `Excel.write` and by `--strict` still writing `-o` (#677); unparsed batch `putf`
reaching files (#663); implicit intersection making array-vs-scalar a property of *where* a formula sits;
the `transparent inline` literal/runtime trap in the Scala DSL (#420, #465); `deps` listing 18,714
precedents for one whole-column SUMIFS. Per-task token figures for SpreadsheetBench are not committed
anywhere in the repo [U]; the only number is "prompt caching -52 to 59%/task" (roadmap line 311).

### 7.2 Authoring and diffs: the same fragment three ways

The ergonomics lane wrote the §6.1 fragment as batch JSON, as a prelude script, and as `.xl`:

| Representation | Size | What the agent reasons about |
|---|---|---|
| Batch JSON | 1,334 chars, ~333 tokens, 20 lines | a row/column plan it invents and keeps in its head; the sheet on every op; `$B$3` vs `B7`; `from` equal to the first cell of `ref`; names declared before use; a check is a `TRUE` cell nothing reads |
| Prelude script | 1,409 chars, ~352 tokens, 31 lines | all of the above plus `writeChecked` not `write`, `Loc(None, …)` under `editIn`, raw strings in `DragFormula` (no `fx` check), the literal/runtime trap |
| `.xl` | 372 chars, ~93 tokens, 17 lines | nothing spatial; `revenue[prev]` is relative in time, not cells; the check is a statement |

Review diffs for three typical changes:

| Change | Batch JSON | Script | `.xl` |
|---|---|---|---|
| growth 5% to 6% | one line | one token | one line |
| add a year | four to six hunks (every `:F` range and the header array) | a run of `.put(ref"G1", …)` plus every `Area` | `years FY2025..FY2029` to `..FY2030` |
| insert `capex` between `ebitda` and `fcf` | no `insert-rows` batch op [V]; renumber every op at row >= 9 by hand, a dozen hunks whose meaning is "one row" | `Edit.InsertRows` exists, but every literal ref below it now lies | two or three added lines |

`xl diff -g` is good for what it is (typed values, resolved styles, structure since the Weaver wave) but
its `SheetDiff` is keyed by `ARef` with no row alignment, so an inserted row reads as every cell below it
changed [I from the data model]. With a source, the xlsx diff becomes a *compiler-output* check in CI.

### 7.3 Diagnostics and the source map

Today's contract has 52 error codes, a `Location(file, sheet, ref, opIndex)`, caret-annotated formula parse
errors with 0-based positions, per-cell "Circular reference" at recalc, and no check concept [V]. A source
format needs positions in the source and errors at the statement level. Proposed `xl check model.xl --json`
diagnostics (exit per the contract's "the code alone decides" rule): `XL_UNKNOWN_NAME` (3, with
`did you mean`), `XL_CYCLE` (3; 1 under a declared `iterative`), `XL_SHAPE_MISMATCH` (3), `XL_PLACEMENT_COLLISION`
(3), `XL_CHECK_FAILED` (1, nothing written), `XL_SYNTAX` (3, caret re-based to file line/col), `XL_UNPLACED`.
JSON shape LSP-aligned (`range: {start: {line, col}, end}`), so an editor plugin and a Claude Code
`PostToolUse` hook consume the same thing. The compiler emits a bidirectional **source map** (statement span
to `{sheet, range}`; cell to `(symbol, period)`) into the lock and optionally a custom part, so `Model!C7:
#NAME?` resolves to `model.xl:14` without grepping.

### 7.4 When agents need addresses

An inherited banker layout read by habit (`Summary!C12` is IRR); an LP template that expects particular cells
(the tjc-modeling byte-exact replication campaign in CHANGELOG [V]); external workbook references `[2]Book!A1`;
geometry-bound constructs (data tables, charts, print areas, freeze panes, CF `sqref`); spill anchors. The
escape hatch of §6.2 covers each.

### 7.5 What a decompiler's output must look like on inherited models

Recognise dragged runs by the shifter test (`shift(parse(f0), k, 0) == parse(fk)`, AST equality, an R1C1
normalisation by construction [V for the shifter laws]); a plug that breaks a run survives visibly as an
`override`; name rows from the nearest label cell (slugified, `TODO rename` where none); recover a header row
as the axis; lift shared single cells to inputs; styles to roles by `canonicalKey` when repeated; everything
unliftable to `grid`, everything irreproducible to `raw`; and **prove the round trip before showing the
source** (`decompile` immediately compiles and `diff -g`s against the input; refuse or annotate `LOSSY`).
Failure mode: a 200k-line cell dump that is worse than the xlsx. The tell: `grid`+`raw` lines exceeding
`rows`+`inputs` lines means the model was not understood. The second failure mode, false lifting (a section
header mistaken for a row label), recompiles to the same cells and is invisible to the diff guard, so the
decompiler must emit a names summary a banker can reject in eighty lines.

### 7.6 Skills and tooling

A third skill, `xl-model` (a language skill, not a tool skill), with its snippets verified by `xl check` in
`scripts/verify-skill-snippets.sh`. Verbs `compile`, `check`, `decompile`, `fmt` join `xl schema --json` and
the generated docs automatically. A Scala internal DSL instead of a text format was considered and rejected
as the *source*: scala-cli latency against a millisecond native `xl check`; diffs in a host language;
scalac positions, not model positions; the literal/runtime trap; a decompiled program a banker will never
read; non-JVM consumers need the grammar. The resolution is the pure ADT plus text grammar plus a prelude
builder (§5.2), so scripts build models from data and compile through the same path.

---

## 8. Benchmark reproduction

_To be filled from the benchmark lane (in progress at the time of writing; see Appendix B if this section is
still a placeholder)._

---

## 9. Risk register

Likelihood and impact on a three-point scale.

| # | Risk | L | I | Early warning | Mitigation | Decides |
|---|---|---|---|---|---|---|
| 1 | Second-format ecosystem risk: zero tooling outside xl | H | M | The first "I'll just fix it in Excel"; a `.xl` feature with no xlsx rendering | `.xl` is agent-facing in v1; the law "every construct compiles with no warning" is a test | Richie |
| 2 | Two truths after a human edits the compiled xlsx | H | H | Stamp mismatch on a model a partner returned | v1: refuse by default; `--fork` re-ingests as a new lineage; decompile-and-merge is a research track | Richie + deal-team lead |
| 3 | Layout expressiveness for banker hand layout | H | M | The proof model needs more than a small share of escape-hatch cells | Escape hatches from day one; publish the escape ratio on the proof model as the go/no-go metric | Richie |
| 4 | Scope creep into a new spreadsheet, formula language or engine | M | H | Any `.xl`-only expression syntax; a preview UI | ADR-018 states the line; every grammar PR cites the xlsx construct it compiles to | Maintainer |
| 5 | Maintenance burden on ~273k lines with one developer of record (`build.mill:73`) | H | H | Wave 2b stays deferred while `.xl` lands; `WriteCommands` keeps its own `putf` beside `Edit` | Finish the `Edit` consolidation first; one wave per step with a dogfood wave after | Richie |
| 6 | Decompiler fidelity on inherited models | H | H | decompile -> compile -> diff shows value changes on a fixture | Exact decompiler first; lossy passes each gated by a fidelity law | Maintainer |
| 7 | Parquet/Arrow/Delta in the native image (72 MB binary today; 1.6 KB reflect-config) | M-H | M | One platform's native build fails; reflect-config grows by hundreds of entries | Text lock in v1; Parquet only after a measured five-platform build; Delta never | Richie + release owner |
| 8 | Determinism and purity | L | H | A compile that is not byte-identical twice; `Clock.system` reached from the compiler | Compiler takes `Clock` like `recalculate`; sorted lock rows; byte-identity goldens; `law-fuzz.yml` extended | Maintainer |
| 9 | Evaluator coverage vs deal models: 124 functions + LET [V]; missing IPMT/PPMT/CUMIPMT/CUMPRINC, SLN/DDB, MIRR, SUBTOTAL, AGGREGATE, PRODUCT, TEXTJOIN/CONCAT, DAYS, LAMBDA/MAP/REDUCE/SCAN/BYROW, VSTACK/HSTACK/TAKE/DROP/SORTBY | H | H | The proof model hits a missing function; Excel and `view --eval` disagree after the partner opens it | Gate on the library recalc path; typed `UNSUPPORTED_FUNCTION` refusal; LibreOffice (today's oracle) or gridcraft as the second evaluator; fix the two `putf` bugs first | Maintainer |
| 10 | Grammar versioning over years of models in git | M | M | An old model fails to parse after a release | `xl 1` header; every released grammar kept; `xl source migrate`; grammar goldens | Maintainer |
| 11 | Security of an LLM-written text format (formula injection, `=cmd|`, external links, `WEBSERVICE`, hyperlinks) | M | H | Any `.xl` from outside the sandbox | Compiler policy: no external refs, no DDE, allowlisted hyperlink schemes, no macros; typed refusals; a lint rule on compiled output | Richie + security |
| 12 | Adoption inside TJC: agents only, or humans too? | M | H | Analysts asking for a `.xl` editor | Decide "agents write, humans review diffs"; re-evaluate after the proof | Richie + deal-team lead |
| 13 | "xlsx as fork" governance: which artefact is the record? | H | M | Two models with the same name and different numbers in Egnyte/UC | Lineage columns in the lock; the stamped xlsx is the record, an unstamped one is a fork | Richie + Finance/Compliance |
| 14 | Performance expectations: `.xl` does not make xl read faster | M | M | Compile time dominated by `Workbook` construction on bulk data | State as a non-goal; bulk blocks through the streaming writer later (#634) | Maintainer |
| 15 | Dynamic-array authoring is a blocking dependency (#714 open; PR #721 draft, unverified in Excel 365) | H | H | The proof model uses FILTER/SORT/UNIQUE | Land and Excel-verify #721 before the compiler claims spill support | Maintainer |

---

## 10. Effort by complexity

Anchors from this repo's own history [V]: `OpRegistry.scala` 1,062 lines; the `Edit` algebra PR #598 =
5,620 source + 993 test lines over 39 files; `FormulaParser.scala` 1,684 lines; `import-md` ~560 lines
plus 280 test; `diff` 1,027 lines plus ~1,150 test. Recent agent waves land as 10k to 28k-line PRs each
followed by a dogfood wave. No calendar time anywhere.

| Step (Weaver) | Diff | Modules | Hidden sub-steps | Order of magnitude |
|---|---|---|---|---|
| 1. Grammar | Medium-Large | new `xl-source`; `XLError` codes in xl-core; `ErrorCode`/`Schema` in xl-cli; prelude export + probe | parser **and** canonical printer with a round-trip law; typed AST; ScalaCheck generators; `xl 1` header and grammar goldens; docs page; CHANGELOG; ADR-018 | 2.5 to 4k src, 3 to 5k tests |
| 2. Compiler | Large | `xl-source` (layout + lowering); `xl-ooxml` (custom.xml stamp); `xl-evaluator` (typed refusal for unsupported functions); `xl-cli` (verbs, `VerbDoc`, goldens, `DocsGenSpec`); prelude; examples | layout engine is the novel part; gate on the library recalc path (precondition: `putf` divergence retired); stamp as a byte-deterministic contract; `xl-model` skill with verified snippets | 4 to 8k src, 4 to 8k tests (one `Edit`-algebra PR is the analogue) |
| 3. Lock | Small-Medium as text; Large as Parquet-in-native | xl-core/`xl-source` (pure model, canonical order, bytes); xl-cats-effect (IO); xl-cli (`compile` emits it; `diff` learns locks) | schema (`sheet, ref, symbol, period, kind, formula, value, type, numFmt, source_hash, grammar_version`); lock-to-workbook equality law | 0.5 to 1.5k src + tests; +1 to 2k and a build-risk tail for Parquet |
| 4a. Decompiler, exact | Small | xl-ooxml reader + `xl-source` printer | hash verification; refuse on mismatch | 0.3 to 0.6k |
| 4b. Decompiler, lossy | Large, open-ended | `xl-source` inference; xl-evaluator (shifter equivalence, `QualifiedGraph`); xl-cli `decompile`; fixtures from the real-file corpus | one fidelity law per inference pass (decompile -> compile -> `diff --cells-only` empty on every fixture); names summary | 3 to 6k src, as many tests, plus fixtures |
| 5. Prove on a live deal model | Small code, large consequence | none new | by this repo's record a dogfood yields 5 to 7 findings and a 10 to 16k-line follow-up PR; needs a human in Excel 365 | 0 new; budget one wave of fixes |

Dependency order: **1 -> 2 -> (3 in parallel with 4a) -> 5 -> 4b**. Preconditions before step 2: the
`putf` cache/spill fixes inside the `Edit` path; a decision on the stamp carriers; #721 landed and verified
in Excel if the proof model spills. Cross-cutting per step, enforced by CI [V]: format all sources, goldens
re-recorded deliberately, `docs/reference/generated/*` regenerated, `ContractSpec` vocabulary, `plugin.json`/
`SKILL.md` version coupling, a prelude probe per visible addition, `examples/*.sc` still compiling,
CHANGELOG `[Unreleased]`, roadmap and STATUS lines.

Overall: 11 to 20k new source lines plus a comparable test tree, one to two wave-sized PRs, +5 to 10% of
the codebase, all of it a new surface with its own contract, goldens, skill and grammar versions. The
compiler and the lossy decompiler are uncertain by a factor of two; the grammar and the text lock are
bounded by existing analogues.

### 10.3 The "or not" case: no xlsx for internal-only models

What xl would be: `xl-evaluator` as the calc engine, `.xl` as source, the lock as state, `view --format
html|svg|png` as the viewer (it exists in embryo). What it loses, from this repo's own record: every
consumer named in the roadmap (the tjc-modeling byte-exact campaign, the FinAgent LBO, the Crestline
circular build "to six decimals") opened an xlsx; Excel as the UI for what-if, mark-ups and print; the
LibreOffice oracle that *defines* evaluator correctness in `xl-evaluator` tests and `law-fuzz.yml`; the
8,636-test safety net, mostly OOXML round-trip laws. What it gains: no OOXML fidelity backlog, no two
truths, no 8 GB heap on inherited books, freedom to change the data model. Only machine-to-machine
workflows (roll-ups into a Delta table, benchmark harnesses, agent scratch models) could live without xlsx,
and those already have a better home in Databricks. **Verdict: no as a direction; yes as a mode**
(`compile --no-xlsx`).

### 10.4 Delta vs plain Parquet vs text

Covered in §5.5. Verdict: text lock from xl; Parquet deferred; Delta committed by Databricks, never embedded.

### 10.5 Decision framing

**Go-with-changes** (§1 table). The three conditions that would flip it to **no-go**:

1. The proof model needs explicit placement for a large share of its cells (judgment: more than roughly a
   third of non-input rows), or per-cell style overrides dominate. Then "names, not addresses" is not the
   authoring win claimed, and batch JSON over the existing CLI is the cheaper agent surface.
2. Humans must author `.xl`. That is a new editor, viewer and training problem, the scope creep Weaver's
   line forbids.
3. The evaluator gap sits on the critical path of the first live model. Then evaluator coverage (or a second
   evaluator as the gate oracle) comes first and `.xl` waits.

Conditions that would strengthen a plain **go**: a committed agent-wave budget for the dogfood tail; a named
internal consumer (Weaver, FinAgent) producing `.xl` in anger on the proof model; a measured compile-and-gate
on the tjc-modeling template that matches Excel after a partner opens it.

---

## 11. Open questions for Richie

1. **Scope of the first proof.** Which live deal model, and is a partner opening the xlsx in Excel 365 the
   acceptance test? The escape-hatch ratio on that model is the metric this document proposes for go/no-go.
2. **Who writes `.xl`?** Agents only (humans review diffs in PRs), or humans too (editor, viewer, training)?
3. **Hand-edit policy.** Is "a hand-edited xlsx is a fork" acceptable for v1, with the stamped xlsx as the
   record and an unstamped one as a fork, or must v1 attempt decompile-and-merge?
4. **Lock format.** Accept sorted JSON Lines from xl with Databricks committing Delta, and defer Parquet
   until a local columnar consumer needs it?
5. **Sequencing against Wave 2b.** Finish the `Edit` lowering (#583 and the `putf` divergence) before the
   compiler, as this document recommends, or in the same wave?
6. **Dynamic arrays.** Land and Excel-verify PR #721 before or inside the `.xl` wave?
7. **Unity Catalog.** A `models` catalog with `models.cells` keyed on `(model_id, source_hash, version)`:
   snapshots or changed cells only?
8. **Evaluator gap.** Grow the registry toward the deal-model set (IPMT/PPMT family, SUBTOTAL/AGGREGATE,
   TEXTJOIN, the array shapers) before the proof, or accept typed refusals and a second evaluator as the gate?
9. **Security policy.** Confirm the compiler's refusals: no external links, no DDE, allowlisted hyperlink
   schemes, no macros.
10. **ADR-018.** Record Weaver's line verbatim as the scope statement?

---

## Appendix A: the email exchange (Sunday 2026-10-11, 2:06 to 3:39 PM ET)

Subject "Open source Excel", eleven messages between Richie Caputo and Weaver (TJC's internal email work
agent; an AI, every reply signed "Weaver / TJC Internal Work Agent"). No other participants.

**Messages 1 to 7.** Richie asked how storytold/gridcraft (six days old, MIT/Apache, Rust, agent-built,
from ArtCraft) compares to xl. Weaver benchmarked it: a real engine, 501 formula functions vs xl's ~120,
20 to 30x faster on a 1.6M-cell book, but alpha, f64 precision, thin tests, one-maintainer risk. Advice:
keep xl as author and ship-gate, add gridcraft as a second evaluator and use it as an oracle to agent-fill
xl's evaluator; "their speed is the method, not the repo." Adopting or forking it: no. Buying ArtCraft: no;
build a TJC software factory with xl as product one. Two xl bugs noted: single-verb `putf` of a bare XLOOKUP
writes no cached value; single-verb `putf` of FILTER writes a scalar instead of a spill.

**Message 8, Richie (3:30 PM).** "The xlsx format is effectively legacy. Could we build our own .xl format
that is far more efficient? Assume we just need to be able to create xlsx eventually perhaps... or not"

**Message 9, Weaver (3:35 PM).** "Yes to a .xl source format. No to a .xl storage format that replaces
xlsx. The distinction is the whole answer." xlsx "is the PDF of finance"; what is legacy is the authoring
model, "a grid of addressed cells holding text formulas with no source of truth." For agents the primitive
is "a declarative source that compiles to xlsx, the way Typst compiles to PDF." "Names, not addresses. Cells
are assigned at compile time … Inserting a row never breaks a reference, and an agent never reasons about
$G$14." "Keep Excel's formula language. Do not invent one." XLSB: "Adoption decides exchange formats, not
efficiency." Measurements on 200k x 8: CSV 10.7 MB, sheet XML 55.8 MB (8.9 MB zipped), Parquet 3.0 MB; reads
Parquet 0.04 s, gridcraft 1.3 s, `xl --stream` 7.6 s, xl in-memory 38 to 41 s; "the other 30x in xl's number
is our implementation … A better reader would." Proposed container: `model.xl` + `data/*.parquet` +
`lock/values.parquet` + the compiled `.xlsx`; the ship-gate moves into the compiler; round trip exact for
files we compiled (source hash in a custom part), lossy-but-useful for inherited banker models. Risks: a
second format with zero ecosystem, layout expressiveness (needs an explicit-placement escape hatch), scope
creep. "The line to hold: source plus compiler plus lock. No new UI, no new formula language, no new engine
beyond what xl needs anyway." Build plan: grammar from batch JSON, Scala specs and TJC templates; compiler on
the existing library with no new xlsx code; lock; decompiler; prove on one live deal model. "Make .xl the
source, xlsx the render, Parquet the cache, and the efficiency problem and the agent-authoring problem are
the same fix."

**Message 10, Richie (3:37 PM).** "What if we build on the delta format which is already Scala?"

**Message 11, Weaver (3:39 PM).** "Yes, for the lock and the data. Not for the source, and not as the thing
we hand to a human." Delta is a table format over a directory with a transaction log, not a document; the
source belongs in git, xlsx is for humans. "The Scala is delta-spark, which needs Spark. The no-Spark JVM
path is Delta Kernel, and I would not embed even that in the xl native binary. Write plain Parquet locally
and let Databricks commit it." Delta buys versions, SQL across every model, lineage in Unity Catalog. Risks:
a directory is not an attachment; two truths after a hand edit; cell-grain volume. "Delta is the right
database for models, not the right file for them."

**Status at the time of this document.** Richie has not replied to the Delta message. Open decisions: proceed
with a `.xl` DSL under Weaver's line; plain Parquet out of xl vs embedded Delta; a Unity Catalog `models`
catalog; which live deal model is the proof.

---

## Appendix B: verification ledger

**Verified in the repo** (path:line in the body): module dependency graph and purity rules; `Edit` (49
cases) and `EditInterpreter`; `BatchParser` applying ops outside `Edit`; batch op count 34; `TExpr.NameRef`,
`NameWalk`, `DependencyGraph`, `RecalcResult.evaluated`; `FormulaShifter` laws; structured references and
LAMBDA unparseable; `DefinedName` and `TableSpec`; `DocProps` writes core/app only; the surgical writer's
verbatim preservation of unknown parts; `WriteCommands.scala:478-612` single-verb `putf`; the CLI contract
(76 verbs, 52 error codes, 4 exit codes, 161 goldens); native-image settings and the 1.6 KB reflect-config;
124 registry functions + LET; the CI gates; the two skills.

**Verified against GitHub** (this session): issue #714 open (2026-10-10); PR #721 draft, 64 files, +3,295,
"Not checked in Excel itself"; release v0.23.1 native binaries 65.8 to 73.8 MB, JAR tarball 46.1 MB.

**Verified on Maven Central** (POM reads): `delta-kernel-defaults` 4.0.0 -> `hadoop-client-runtime` 3.4.0
(30.4 MB) + `parquet-hadoop` 1.12.3; `parquet-hadoop` 1.15.2 Hadoop `provided`; `carpet-record` 0.4.0 ->
`hadoop-common` compile; `parquet-floor` 1.50; `arrow-vector` 18.3.0; `duckdb_jdbc` 1.3.2.0 76.4 MB;
apache/arrow-java#866 "GraalVM is not a currently supported configuration".

**Summary-only** (the sandbox proxy refused the page; claims rest on search excerpts): Library of Congress
format descriptions; Microsoft Learn and support.microsoft.com pages (structured references, LAMBDA, PY,
CalculationProperties, KB 948615); Grist, Causal, Anaplan, Quantrix documentation; Brian Jones's OOXML history;
LibreOffice help and bugzilla; openpyxl heptapod issues.

**Could not verify**: any primary Microsoft XLSB benchmark; whether Excel 365 preserves an unknown custom part
or `extLst` on save; whether an edit-free Excel save changes `revisionPtr`; whether PR #721's output opens as a
spill in Excel 365; Google Sheets' treatment of `custom.xml`; preview renderers (Quick Look, Outlook,
SharePoint) as cache-only; Excel for Mac extension sniffing; native-image behaviour of any Parquet library;
Hardwood jar size; whether xl's evaluator applies implicit intersection to a row-shaped defined name; Grist's
xlsx formula export; Causal's post-acquisition status; Excel Labs AFE retirement; per-task token figures for
SpreadsheetBench; that the two `putf` bugs are not already filed as issues; the birth size of `import-md`
(#159) and `diff` (#137) in this shallow clone; deal-model cell and compile counts (judgment).
