# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

**XL** is a purely functional, mathematically rigorous Excel (OOXML) library for Scala 3 (3.9 LTS). The design prioritizes **purity, totality, determinism, and law-governed semantics** with zero-overhead opaque types and compile-time DSLs.

> **Guiding Principle**: You are working on **the best Excel library in the world**. Before making any decision, ask yourself: **"What would the best Excel library in the world do?"**

**Package**: `com.tjclp.xl` | **Build**: Mill 1.1.x | **Scala**: 3.9.0 (LTS) | **JDK**: Temurin 25, pinned in `.mill-jvm-version` and downloaded by Mill itself

## Core Philosophy (Non-Negotiables)

1. **Purity & Totality**: No `null`, no partial functions, no thrown exceptions, no hidden effects
2. **Strong Typing**: Opaque types for Column, Row, ARef, SheetName; enums with `derives CanEqual`
3. **Deterministic Output**: Canonical ordering for XML/styles; byte-identical output on re-serialization
4. **Law-Governed**: Monoid laws for Patch/StylePatch; round-trip laws for parsers/printers
5. **Effect Isolation**: Core (`xl-core`, `xl-ooxml`) is 100% pure; only `xl-cats-effect` has IO

## Performance

- **Streaming**: O(1) memory for reads and writes via SAX parser
- **Benchmarks**: JMH suite in `xl-benchmarks/` (work in progress)

*Use `ExcelIO.readStream()` for large files (constant memory), `ExcelIO.read()` for random access + modification.*

## Module Architecture

```
xl/              → Aggregate module + scripting prelude (com.tjclp.xl.scripting) + prelude probes
xl-core/         → Pure domain model (Cell, Sheet, Workbook, Patch, Style), macros, DSL
xl-ooxml/        → Pure OOXML mapping (XlsxReader, XlsxWriter, SharedStrings, Styles)
xl-cats-effect/  → IO interpreters and streaming (Excel[F], ExcelIO, SAX-based streaming)
xl-evaluator/    → Formula parser/evaluator (TExpr GADT, function registry, dependency graphs)
xl-cli/          → Stateless `xl` CLI (internal, native-image capable)
xl-agent/        → AI agent benchmark runner (Anthropic API, skill comparison)
xl-benchmarks/   → JMH performance benchmarks
xl-testkit/      → Test laws, generators, helpers [placeholder — no sources yet]
```

Published to Maven Central: `xl` (aggregate), `xl-core`, `xl-ooxml`, `xl-cats-effect`, `xl-evaluator`. Internal: `xl-cli`, `xl-agent`, `xl-benchmarks`, `xl-testkit`.

## Import Patterns

```scala
import com.tjclp.xl.{*, given}     // Everything: core + formula + display + type class instances (pure)
import com.tjclp.xl.unsafe.*       // .unsafe boundary (explicit opt-in)

// Scripts: ONE import for everything above + sync Excel + ExcelIO + unsafe + String.toFormatted.
// Never combine with com.tjclp.xl.{*, given} in the same file (ambiguous forwarders).
import com.tjclp.xl.scripting.{*, given}

// Core API
val sheet = Sheet("Demo").put("A1" -> 100)
Excel.write(Workbook(sheet), "output.xlsx")

// Formula evaluation (when xl-evaluator is a dependency)
sheet.evaluateFormula("=A1*2")     // SheetEvaluator extension method
FormulaParser.parse("=SUM(A1:A10)") // Parser at package level

// Display formatting
given Sheet = sheet
println(excel"Value: ${ref"A1"}")  // excel interpolator
sheet.displayCell(ref"A1")         // explicit display method
```

Note: `{*, given}` is required because Scala 3's `*` doesn't include given instances by default.

For production code with Cats Effect, use `ExcelIO` instead of `Excel`:
```scala
import com.tjclp.xl.io.ExcelIO
val excel = ExcelIO.instance[IO]
excel.read(path).flatMap(wb => excel.write(wb, outPath))
```

## Key Types

**Addressing** (`xl-core/src/com/tjclp/xl/addressing/`):
- `Column`, `Row` → opaque Int (zero-overhead)
- `ARef` → opaque Long (64-bit packed: `(row << 32) | col`)
- `CellRange(start, end)` → normalized, inclusive range

**Domain** (`xl-core/src/com/tjclp/xl/`):
- `Cell(ref, value, styleId, ...)` | `Sheet(name, cells, ...)` | `Workbook(sheets, ...)`
- `Patch` enum → Monoid for Sheet updates (Put, SetStyle, Merge, Remove, Batch)
- `StylePatch` enum → Monoid for CellStyle updates
- `CellStyle(font, fill, border, numFmt, align)` → complete cell formatting

**Codecs** (`xl-core/src/com/tjclp/xl/codec/`):
- `CellCodec[A]` → Bidirectional for 9 types: String, Int, Long, Double, BigDecimal, Boolean, LocalDate, LocalDateTime, RichText
- Auto-inferred formats: LocalDate→Date, LocalDateTime→DateTime, BigDecimal→Decimal

**Formula** (`xl-evaluator/src/com/tjclp/xl/formula/`):
- `TExpr[A]` GADT → typed formula AST
- `FormulaParser.parse()` / `FormulaPrinter.print()` → round-trip verified
- `DependencyGraph` → cycle detection via Tarjan's SCC, topological sort via Kahn's

## Build Commands

```bash
./mill __.compile          # Compile all (main + test sources)
./mill __.test             # Run all tests (6,788)
./mill xl-core.test        # Test one module
./mill xl-core.test.testOnly com.tjclp.xl.addressing.ColumnSpec -- '*parse*'   # One suite, glob-filtered
./mill mill.scalalib.scalafmt.ScalafmtModule/reformatAll __.sources     # Format (what CI checks; __.reformat skips test sources)
./mill mill.scalalib.scalafmt.ScalafmtModule/checkFormatAll __.sources  # CI format check
./mill clean               # Clean artifacts
make install               # Local only: GraalVM native-image CLI → ~/.local/bin/xl
make install-jar           # Portable: assembly JAR + wrapper → ~/.local/bin/xl (cloud sessions, CI)
```

**IMPORTANT**: After modifying CLI code, reinstall before verifying behavior: `make install` locally (GraalVM native-image), `make install-jar` where there is no GraalVM (cloud sessions, CI). Do NOT manually copy jars.

### Toolchain

`.mill-jvm-version` pins Temurin 25 and `.mill-version` pins Mill 1.1.5; the `./mill` launcher downloads both through Coursier. The only prerequisites are `bash`, `curl`, and network access to GitHub and Maven Central. Never install a JDK by hand to satisfy the build, and never edit `javacOptions` to fit the local JDK. A fresh checkout downloads dependencies on its first build: give `./mill` a 600000 ms timeout.

### Remote sessions (Claude Code on the web, routines, `@claude` in GitHub Actions)

The sandbox is Ubuntu 24.04 with OpenJDK 21 and no scala-cli or GraalVM. `.claude/settings.json` runs `scripts/remote-setup.sh` on SessionStart; when `CLAUDE_CODE_REMOTE=true` it provisions JDK 25 and scala-cli through Coursier, exports `JAVA_HOME`/`PATH` for the session, and prints a summary into context. There, `make install-jar` replaces `make install`, worktrees and `gtr` are unnecessary (the sandbox clone is isolated), and personal memory is absent: everything a session must know lives in this file, `.claude/rules/`, and `.claude/skills/`. Details, the environment setup script, and a Docker rehearsal: `docs/reference/remote-sessions.md`.

## xl-agent Benchmark Runner

The `xl-agent` module runs AI agent benchmarks comparing different Excel manipulation skills (xl-cli vs openpyxl).

### Running Benchmarks

```bash
# Basic benchmark run (no -- needed with Mill)
./mill xl-agent.run --benchmark spreadsheetbench --task 2768 --skills xl

# With streaming console output
./mill xl-agent.run --benchmark spreadsheetbench --task 2768 --skills xl --stream

# Parallel execution (default: 4)
./mill xl-agent.run --benchmark spreadsheetbench --skills xl --parallelism 8

# Compare multiple skills
./mill xl-agent.run --benchmark spreadsheetbench --task 2768 --skills xl,xlsx

# List available skills
./mill xl-agent.run --list-skills

# Force re-upload of skill files (bypasses cache)
./mill xl-agent.run --benchmark spreadsheetbench --task 2768 --skills xl --force-upload
```

### Key Options

| Flag | Description |
|------|-------------|
| `--benchmark <name>` | Benchmark suite: `spreadsheetbench`, `tokenbenchmark` |
| `--task <ids>` | Task ID(s), comma-separated (`--task 2768,2769`); a repeated `--task` replaces the earlier one |
| `--skills <list>` | Comma-separated: `xl`, `xlsx`, or `xl,xlsx` |
| `--parallelism <n>` | Number of parallel work units (default: 4) |
| `--max-tokens <n>` | Per-iteration output cap for agent turns (default: 32768; thinking counts against it) |
| `--stream` | Real-time colored console output |
| `--force-upload` | Bypass file cache, re-upload skill |
| `--output <dir>` | Results directory (default: `results/<timestamp>/`) |

### Architecture

- **BenchmarkEngine**: Orchestrates execution with flattened work scheduling
- **Skill**: Abstraction for different approaches (XlSkill, XlsxSkill)
- **WorkUnit**: Single (task, skill, case) combination for parallel execution
- **ConversationTracer**: Captures agent conversation for debugging

### Output

Results are written under the `--output` directory (default `results/<timestamp>/`):
- `outputs/<taskId>/<skill>/` - Output xlsx files
- `tasks/<taskId>/<skill>/case<N>/conversation.json` - Conversation traces (one per case)
- `summary.json` / `summary.md` - Aggregated results (JSON + Markdown)
- `<skill>/summary.json` - Per-skill summary

## CLI Usage

The `xl` CLI is stateless by design. Key patterns:

```bash
# Global flags (accepted anywhere on the command line, before or after the verb)
-f, --file <path>     # Input file (required)
-s, --sheet <name>    # Sheet to operate on
-o, --output <path>   # Output file for mutations
--max-size <MB>       # Override 100MB security limit (0 = unlimited; the heap still bounds what fits)
--stream              # O(1) memory streaming mode for large files

# ONE sheet rule, for every verb, batch op and --stream path:
#   a qualified ref names the sheet > -s > the only sheet of a single-sheet book > SHEET_REQUIRED (exit 3)
xl -f data.xlsx --sheet "Q1 Report" view A1:D20    # Using --sheet flag
xl -f data.xlsx view "Q1 Report"!A1:D20            # Using qualified ref (wins over -s)
xl -f single.xlsx view A1:D20                      # Single-sheet books auto-select for every verb

# Commands that work without sheet (operate on all sheets)
xl -f data.xlsx sheets                              # List all sheets
xl -f data.xlsx search "Revenue"                    # Search all sheets

# Mutations require -o
xl -f in.xlsx -s Data -o out.xlsx put B5 1000      # Write value
xl -f in.xlsx -s Data -o out.xlsx putf C5 "=B5*1.1" # Write formula

# Formula dragging with $ anchoring: one formula over a range drags from the range's first cell
xl -f in.xlsx -s Data -o out.xlsx putf B2:B10 "=SUM(\$A\$1:A2)"

# Sheet names with spaces: use double quotes around the formula argument
xl -f in.xlsx -s Summary -o out.xlsx putf B4 "='Income Statement'!G8"

# View command flags
--eval                # Evaluate formulas (compute live values)
--formulas            # Show formulas instead of values

# Array formula evaluation (evala command)
xl -f data.xlsx -s Sheet1 evala "=TRANSPOSE(A1:C2)"           # Evaluate and display result
xl -f data.xlsx -s Sheet1 evala "=TRANSPOSE(A1:C2)" --at E1   # Spill result starting at E1
xl -f data.xlsx -s Sheet1 evala "=A1:B2*10"                   # Array arithmetic with broadcasting

# Style command flags (styles merge by default, use --replace for full replacement)
--replace             # Replace entire style instead of merging
--border <style>      # Border style for all sides: none, thin, medium, thick
--border-top <style>  # Top border only
--border-right <style>    # Right border only
--border-bottom <style>   # Bottom border only
--border-left <style>     # Left border only
--border-color <color>    # Border color (applies to all specified borders)

# Rasterization (PNG/JPEG export)
--rasterizer <name>   # Force a specific backend: batik, cairosvg, rsvg-convert, resvg, imagemagick

# Large file handling (100k+ rows)
--stream              # O(1) memory streaming: search, stats, bounds, view, cell, describe, sheets; put, putf, style, batch
--max-size 0          # Disable security limits for in-memory load (lifts the limit, not the heap)
--max-size 500        # Set custom limit in MB
```

**Large File Operations** (~10s vs ~80s for 1M rows):
```bash
# Streaming mode - O(1) memory, 7-8x faster (ONE sheet rule: sheet-scoped verbs need -s or a
# qualified ref on a multi-sheet book, else SHEET_REQUIRED exit 3; search/sheets read the whole book)
xl -f huge.xlsx --stream search "pattern" --limit 10
xl -f huge.xlsx -s Sheet1 --stream stats A1:E100000
xl -f huge.xlsx -s Sheet1 --stream bounds
xl -f huge.xlsx -s Sheet1 --stream view A1:D100 --format csv
xl -f huge.xlsx -s Sheet1 -o out.xlsx --stream putf A2 "=B2*1.1"

# In-memory mode - when you need full workbook access
xl -f huge.xlsx --max-size 0 sheets                # Disable limits
xl -f huge.xlsx --max-size 500 -s Sheet1 cell A1   # 500MB limit
xl -Xmx32g -f huge.xlsx --max-size 0 audit         # Raise the native image's 8 GB heap cap (-Xmx before -f to be safe)
```

**Streaming limitations**: `--stream` covers the reads (`search`, `stats`, `bounds`, `view` in markdown/csv/json, `cell`, `describe`, `sheets`) and the writes (`put`, `putf`, `style`, and `batch` for streamable ops); it never recalculates. An in-memory load (`--max-size`) is needed only for `--eval`, `put --csv`, `--strict` on a write, the html/svg/png/jpeg/webp/pdf renders (they need styles), and the whole-book verbs `audit`, `deps`, `filter` and `describe --full`.

**`--max-size` lifts the security limit, not the heap** (#636): the native binary is built with an 8 GB heap ceiling (`-R:MaxHeapSize=8g` in `xl-cli/package.mill`), raised by `-Xmx<size>` (consumed anywhere on the command line; put it before `-f` to be safe: `xl -Xmx64g -f …`; JAR: `java -Xmx64g -jar`); `--max-size` below 0 is a usage error. Measured on 1M-row books, an in-memory load needs 16–20× its worksheet XML for dense cells (346 MB loads in 6 GB, not 5), 33–41× for wide text rows (the dogfood book: 1.09 GB of XML, 36–45 GB of heap) and 2–3× the `sharedStrings.xml` bytes for long text — so stream large files. With `--max-size` lifted, `MemoryGuard` (xl-cli) sizes the load from the zip's central directory before parsing: a lower estimate (`14 × sheet XML + 2 × SST`) above the heap is refused as `RESOURCE_LIMIT` (exit 3); an upper estimate (`30 × sheet + 3 × SST`) above it proceeds under a `MEMORY_PRESSURE` warning; a load, recalculation, `view --eval` evaluation, render or serialisation that then exhausts the heap is `RESOURCE_LIMIT` too, with the `--stream`/`-Xmx` hint — never a raw `OutOfMemoryError` (residuals: an OOM raised first on another fiber, e.g. `recalc --parallel`, can still be fatal; non-heap OOMs carry the heap wording).

See `docs/design/smart-streaming.md` for future enhancements.

**Batch JSON Syntax** (typed values + smart detection):
```bash
# Native JSON types (numbers, booleans, null)
echo '[{"op":"put","ref":"A1","value":123.45}]' | xl -f in.xlsx -o out.xlsx batch -

# Smart detection: currency, percent, dates (opt-out with "detect":false)
echo '[{"op":"put","ref":"A1","value":"$1,234.56"}]' | xl ...  # → Currency
echo '[{"op":"put","ref":"A1","value":"45.5%"}]' | xl ...      # → Percent (stored as 0.455)
echo '[{"op":"put","ref":"A1","value":"2025-01-15"}]' | xl ... # → Date

# Explicit format hints
echo '[{"op":"put","ref":"A1","value":0.455,"format":"percent"}]' | xl ...

# Batch putf: "formula" accepted as alias for "value"
echo '[{"op":"putf","ref":"D14","formula":"=SUM(D5:D12)"}]' | xl -f in.xlsx -o out.xlsx batch -

# Formula dragging (shifts references like Excel fill-down)
echo '[{"op":"putf","ref":"B2:B10","value":"=A2*2","from":"B2"}]' | xl -f in.xlsx -o out.xlsx --stream batch -

# Explicit formula array
echo '[{"op":"putf","ref":"B2:B4","values":["=A2*2","=A3*2","=A4*2"]}]' | xl ...

# Dry-run: validate batch JSON without writing
echo '[{"op":"putf","ref":"A1","formula":"=1+1"}]' | xl batch --dry-run -
# Also works with --file/--output (skips read/write, just validates)
echo '[{"op":"put","ref":"A1","value":"test"}]' | xl -f in.xlsx -o out.xlsx batch --dry-run -

# The document's JSON Schema: every op, field, alias and example (no -f needed)
xl batch --schema

# Every op accepts "sheet" (except add-sheet/rename-sheet); keys in camelCase or kebab-case
echo '[{"op":"put","sheet":"Summary","ref":"A1","value":1}]' | xl -f in.xlsx -o out.xlsx batch -

# Comments, visibility, autofit, sheet management
echo '[{"op":"comment","ref":"A1","text":"Note","author":"User"}]' | xl ...
echo '[{"op":"clear","range":"A1:B10","all":true}]' | xl ...
echo '[{"op":"col-hide","col":"C"}]' | xl ...
echo '[{"op":"autofit","columns":"A:F"}]' | xl ...
echo '[{"op":"add-sheet","name":"Summary","after":"Sheet1"}]' | xl ...
echo '[{"op":"rename-sheet","from":"Old","to":"New"}]' | xl ...
```

**Every batch operation**, with its fields, aliases and example, is generated from the registry that parses the document: run `xl batch --schema` (JSON Schema) or read `docs/reference/generated/batch-ops.md`. Never copy the list by hand — `DocsGenSpec` fails when the page drifts from the code.

**The CLI contract as data**: `xl schema` prints every verb with what it needs and how it exits; `xl schema --json` publishes exit/error/warning codes, globals, verbs, the batch schema, the function registry and the envelope schema in one document; `docs/reference/generated/{cli-verbs,batch-ops,functions,exit-codes,error-codes}.md` are rendered from it (`XL_UPDATE_DOCS=1 ./mill xl-cli.test.testOnly com.tjclp.xl.cli.contract.DocsGenSpec` regenerates; same discipline as `XL_UPDATE_GOLDEN=1` for the goldens).

**Common mistake**: Using an unqualified range on a multi-sheet book without `--sheet`:
```bash
# ❌ Wrong on a multi-sheet book - SHEET_REQUIRED (exit 3) naming the candidates
xl -f data.xlsx view A1:B4

# ✅ Correct options (a single-sheet book needs neither: its only sheet is selected)
xl -f data.xlsx --sheet "Sheet1" view A1:B4
xl -f data.xlsx view "Sheet1"!A1:B4
```

See `docs/reference/cli.md` for full command reference.

**Directory Structure**:
- `.claude/` - Dev-only commands (release-prep, docs-cleanup-xl)
- `plugin/` - User-facing skill distributed via plugin marketplace

The CLI skill (`plugin/skills/xl-cli/SKILL.md`) auto-detects the latest release from GitHub API—no version placeholders to maintain. The scripting skill (`plugin/skills/xl-scripting/SKILL.md`) pins the library version instead (bumped by release-prep; release workflow gates on it) and its snippets are compile-verified by `scripts/verify-skill-snippets.sh`.

## Code Style

All code must pass the CI format check (`./mill mill.scalalib.scalafmt.ScalafmtModule/checkFormatAll __.sources`). See `docs/design/style-guide.md` and the `xl-scala-style` skill.

**Rules**: opaque types for domain quantities | enums with `derives CanEqual` | `final case class` for data | total functions returning Either/Option | extension methods over implicit classes

**WartRemover** (compile-time enforcement):
- ❌ Error: `null`, `.get` on Try/Either
- ⚠️ Warning: `.head/.tail` on collections, `.get` on Option, `var`/`while`/`return` (acceptable in tests/macros)

**Error Handling**: Always use `XLResult[A] = Either[XLError, A]`

## Architecture Patterns

### 1. Opaque Types
```scala
opaque type Column = Int   // Cannot mix Int as Column
opaque type ARef = Long    // Packed: (row << 32) | col
```

### 2. Monoid Composition
```scala
import com.tjclp.xl.dsl.*
val patch = (ref"A1" := "Hello") ++ ref"A1".styled(boldStyle) ++ ref"A1:B2".merge
```
Note: Using Cats `|+|` requires type ascription on enum cases.

### 3. Compile-Time Macros
```scala
val cellRef: ARef = ref"A1"           // Validated at compile time
val formula = fx"=SUM(A1:B10)"        // CellValue.Formula
val price = money"$$1,234.56"         // Formatted(value, NumFmt.Currency)
```

### 4. Deterministic XML
- Attributes sorted by name (`XmlUtil.elem`)
- Elements sorted by natural key
- `XmlUtil.compact` for production, `XmlUtil.prettyPrint` for debug

### 5. Law-Based Testing
```scala
property("parse . print = id") { forAll { (ref: ARef) => ARef.parse(ref.toA1) == Right(ref) }}
```

## Essential APIs

### Sheet Operations
```scala
// Batch put with type inference (string refs validated at compile time)
sheet.put("A1" -> "Revenue", "B1" -> LocalDate.now, "C1" -> BigDecimal("1000.50"))

// Type-safe reading
sheet.readTyped[BigDecimal](ref"C1") // Either[CodecError, Option[A]]

// Styling
sheet.style("B19:D21", CellStyle.default.withNumFmt(NumFmt.Percent))
```

### Formula Evaluation
```scala
// SheetEvaluator extension methods available from com.tjclp.xl.{*, given}
sheet.evaluateFormula("=SUM(A1:A10)")      // XLResult[CellValue]
sheet.evaluateWithDependencyCheck()         // Safe eval with cycle detection
```

**Functions**: the registry (`FunctionRegistry.all`, macro-collected from the `FunctionSpecs*` traits) plus `LET` (lexical bindings; a parser-level special form, not in the registry). The complete list with arity, argument slots and flags is generated from the code: `xl functions --json` or `docs/reference/generated/functions.md`. Do not maintain a count or a list by hand.

### Rich Text
```scala
val text = "Bold".bold.red + " normal " + "Italic".italic.blue
sheet.put("A1" -> text)
```

### Comments & HTML Export
```scala
sheet.comment(ref"A1", Comment.plainText("Note", Some("Author")))
sheet.toHtml(ref"A1:B10")  // HTML table with inline CSS
```

## Important Constraints

### ARef Packing
```scala
// Pack: (row << 32) | (col & 0xFFFFFFFF)
// Valid: Column 0-16383 (A-XFD), Row 0-1048575 (1-1048576)
```

### Style Deduplication
Styles deduplicated by `CellStyle.canonicalKey`. Build style index before emitting cells.

### CellRange Normalization
`CellRange(start, end)` auto-normalizes (start ≤ end).

## Testing

**Framework**: MUnit + ScalaCheck | **Generators**: `xl-core/test/src/com/tjclp/xl/Generators.scala`

**6,788 tests** by module: xl-evaluator (2417), xl-core (1551), xl-ooxml (1118), xl-cli (1339), xl-cats-effect (167), xl-agent (145), xl prelude probes (51). See `docs/reference/testing-guide.md` for suite structure and patterns.

## Documentation

- **Roadmap**: `docs/plan/roadmap.md` (single source of truth for work scheduling)
- **Status**: `docs/STATUS.md` (current capabilities, 6,788 tests)
- **Design**: `docs/design/*.md` (architecture, purity charter, domain model)
- **Reference**: `docs/reference/*.md` (examples, scaffolds, performance guide)
- **Remote sessions**: `docs/reference/remote-sessions.md` (cloud sandbox, SessionStart hook, GitHub Actions, Docker rehearsal)
- **Skills** (`.claude/skills/`, load on demand): `mill-build` (targeted builds/tests, Mill gotchas), `xl-scala-style` (Scala 3 idioms, compiler landmines), `xl-testing` (MUnit/ScalaCheck patterns, release gates). Always-on process rules: `.claude/rules/workflow.md`.

## AI Agent Workflow

**Issue Tracking**: [GitHub Issues](https://github.com/TJC-LP/xl/issues)

1. Check GitHub Issues for available tasks
2. Locally: `gtr list` to check for conflicting worktrees, then `gtr create issue-XXX-description` (`gtr` is a personal shell helper; `git worktree add ../worktrees/xl/<branch> -b <branch>` is the plain-git equivalent). In a cloud session skip this step: the sandbox clone is already isolated.
3. After PR merge: close GitHub issue, update STATUS.md if needed

**Module Conflict Matrix**:
- High risk: `xl-core/Sheet.scala` (serialize work)
- Medium risk: `xl-ooxml/Worksheet.scala` (coordinate)
- Low risk: `xl-evaluator/` (parallelize freely)

## Known Gotchas

**Monoid syntax needs type ascription**:
```scala
val p = (Patch.Put(ref, value): Patch) |+| (Patch.SetStyle(ref, 1): Patch)
```

**Extension methods need @targetName**:
```scala
extension (sheet: Sheet)
  @annotation.targetName("applyPatchExt")
  def applyPatch(patch: Patch): XLResult[Sheet] = ...
```

**Adding a FunctionSpec invalidates the registry macro**: `FunctionRegistry.all` is
`FunctionRegistryMacro.collect[FunctionSpecs.type]`, so editing any `FunctionSpecs*` trait can leave
Zinc's incremental state inconsistent — you get `Cyclic reference involving trait TExprReferenceOps`
(and TExprLookupOps / TExprAggregateOps), pointing at files you never touched. It is not your code:
`./mill clean xl-evaluator.compile` then recompile. New specs auto-register from the trait; there is
no list to update. Scala 3.9's Zinc bridge records macro type arguments as dependencies, so this
should now be rare; the fix is unchanged when it appears.

**Import order**: Java/javax → Scala stdlib → Cats → Project → Tests

**AWT headless default (render)**: the first `toHtml`/`toSvg` call sets `java.awt.headless=true` unless the embedder set it explicitly — non-headless AWT keeps script JVMs alive after main completes (non-daemon AWT-Shutdown thread). GUI embedders: set `-Djava.awt.headless=false` or initialize your toolkit before rendering. A deliberate, guarded global side effect (see `RenderUtils`).

## CI/CD

GitHub Actions: `./mill __.checkFormat` → `./mill __.compile` → `./mill __.test`

Coursier + Mill artifacts cached for 2-5x speedup.
