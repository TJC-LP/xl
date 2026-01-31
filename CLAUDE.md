# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

**XL** is a purely functional, mathematically rigorous Excel (OOXML) library for Scala 3.7. The design prioritizes **purity, totality, determinism, and law-governed semantics** with zero-overhead opaque types and compile-time DSLs.

> **Guiding Principle**: You are working on **the best Excel library in the world**. Before making any decision, ask yourself: **"What would the best Excel library in the world do?"**

**Package**: `com.tjclp.xl` | **Build**: Mill 0.12.x | **Scala**: 3.7.4

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
xl-core/         → Pure domain model (Cell, Sheet, Workbook, Patch, Style), macros, DSL
xl-ooxml/        → Pure OOXML mapping (XlsxReader, XlsxWriter, SharedStrings, Styles)
xl-cats-effect/  → IO interpreters and streaming (Excel[F], ExcelIO, SAX-based streaming)
xl-benchmarks/   → JMH performance benchmarks
xl-evaluator/    → Formula parser/evaluator (TExpr GADT, 81 functions, dependency graphs)
xl-testkit/      → Test laws, generators, helpers [future]
xl-agent/        → AI agent benchmark runner (Anthropic API, skill comparison)
```

## Import Patterns

```scala
import com.tjclp.xl.{*, given}     // Everything: core + formula + display + type class instances
import com.tjclp.xl.unsafe.*       // .unsafe boundary (explicit opt-in)

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
./mill __.compile          # Compile all
./mill __.test             # Run all tests (731+)
./mill xl-core.test        # Test specific module
./mill __.reformat         # Format (Scalafmt 3.10.1)
./mill __.checkFormat      # CI check
./mill clean               # Clean artifacts
make install               # Install xl CLI to ~/.local/bin/xl
```

**IMPORTANT**: After modifying CLI code, always run `make install` to update the installed CLI. Do NOT manually copy jars.

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

# List available tasks
./mill xl-agent.run --benchmark spreadsheetbench --list-tasks

# List available skills
./mill xl-agent.run --list-skills

# Force re-upload of skill files (bypasses cache)
./mill xl-agent.run --benchmark spreadsheetbench --task 2768 --skills xl --force-upload
```

### Key Options

| Flag | Description |
|------|-------------|
| `--benchmark <name>` | Benchmark suite: `spreadsheetbench`, `tokenbenchmark` |
| `--task <id>` | Run specific task ID (can repeat) |
| `--skills <list>` | Comma-separated: `xl`, `xlsx`, or `xl,xlsx` |
| `--parallelism <n>` | Number of parallel work units (default: 4) |
| `--stream` | Real-time colored console output |
| `--force-upload` | Bypass file cache, re-upload skill |
| `--output-dir <path>` | Results directory (default: `results/`) |

### Architecture

- **BenchmarkEngine**: Orchestrates execution with flattened work scheduling
- **Skill**: Abstraction for different approaches (XlSkill, XlsxSkill)
- **WorkUnit**: Single (task, skill, case) combination for parallel execution
- **ConversationTracer**: Captures agent conversation for debugging

### Output

Results are written to `results/` directory:
- `outputs/<taskId>/<skill>/` - Output xlsx files
- `traces/<taskId>/<skill>/` - Conversation traces (JSON)
- `summary.json` - Aggregated results

## CLI Usage

The `xl` CLI is stateless by design. Key patterns:

```bash
# Global flags (used with all commands)
-f, --file <path>     # Input file (required)
-s, --sheet <name>    # Sheet to operate on
-o, --output <path>   # Output file for mutations
--max-size <MB>       # Override 100MB security limit (0 = unlimited)
--stream              # O(1) memory streaming mode for large files

# Sheet selection is REQUIRED for unqualified ranges
xl -f data.xlsx --sheet "Q1 Report" view A1:D20    # Using --sheet flag
xl -f data.xlsx view "Q1 Report"!A1:D20            # Using qualified ref

# Commands that work without sheet (operate on all sheets)
xl -f data.xlsx sheets                              # List all sheets
xl -f data.xlsx search "Revenue"                    # Search all sheets

# Single cell ops auto-detect sheet if unambiguous
xl -f data.xlsx cell A1                             # Works if only one sheet

# Mutations require -o
xl -f in.xlsx -o out.xlsx put B5 1000              # Write value
xl -f in.xlsx -o out.xlsx putf C5 "=B5*1.1"        # Write formula

# Formula dragging with $ anchoring
xl -f in.xlsx -o out.xlsx putf B2:B10 "=SUM(\$A\$1:A2)" --from B2

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
--use-imagemagick     # Use ImageMagick instead of Batik (needed for native image rasterization)

# Large file handling (100k+ rows)
--stream              # Use O(1) memory streaming (search, stats, bounds, view)
--max-size 0          # Disable security limits for in-memory load
--max-size 500        # Set custom limit in MB
```

**Large File Operations** (~10s vs ~80s for 1M rows):
```bash
# Streaming mode - O(1) memory, 7-8x faster
xl -f huge.xlsx --stream search "pattern" --limit 10
xl -f huge.xlsx --stream stats A1:E100000
xl -f huge.xlsx --stream bounds
xl -f huge.xlsx --stream view A1:D100 --format csv

# In-memory mode - when you need full workbook access
xl -f huge.xlsx --max-size 0 sheets      # Disable limits
xl -f huge.xlsx --max-size 500 cell A1   # 500MB limit
```

**Streaming limitations**: HTML/SVG/PDF need styles (use --max-size instead). Cell details, formula eval, and writes require full workbook load.

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

# Formula dragging (shifts references like Excel fill-down)
echo '[{"op":"putf","ref":"B2:B10","value":"=A2*2","from":"B2"}]' | xl -f in.xlsx -o out.xlsx --stream batch -

# Explicit formula array
echo '[{"op":"putf","ref":"B2:B4","values":["=A2*2","=A3*2","=A4*2"]}]' | xl ...
```

**Common mistake**: Using unqualified range without `--sheet`:
```bash
# ❌ Wrong - will error
xl -f data.xlsx view A1:B4

# ✅ Correct options
xl -f data.xlsx --sheet "Sheet1" view A1:B4
xl -f data.xlsx view "Sheet1"!A1:B4
```

See `docs/reference/cli.md` for full command reference.

**Directory Structure**:
- `.claude/` - Dev-only commands (release-prep, docs-cleanup-xl)
- `plugin/` - User-facing skill distributed via plugin marketplace

The CLI skill (`plugin/skills/xl-cli/SKILL.md`) auto-detects the latest release from GitHub API—no version placeholders to maintain.

## Code Style

All code must pass `./mill __.checkFormat`. See `docs/design/style-guide.md`.

**Rules**: opaque types for domain quantities | enums with `derives CanEqual` | `final case class` for data | total functions returning Either/Option | extension methods over implicit classes

**WartRemover** (compile-time enforcement):
- ❌ Error: `null`, `.head/.tail`, `.get` on Try/Either
- ⚠️ Warning: `.get` on Option, `var`/`while`/`return` (acceptable in tests/macros)

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

**82 Functions**: SUM, SUMIF, SUMIFS, SUMPRODUCT, COUNT, COUNTA, COUNTBLANK, COUNTIF, COUNTIFS, AVERAGE, AVERAGEIF, AVERAGEIFS, MEDIAN, STDEV, STDEVP, VAR, VARP, MIN, MAX, IF, AND, OR, NOT, ISNUMBER, ISTEXT, ISBLANK, ISERR, ISERROR, CONCATENATE, LEFT, RIGHT, MID, LEN, UPPER, LOWER, TRIM, SUBSTITUTE, TEXT, VALUE, TODAY, NOW, DATE, YEAR, MONTH, DAY, HOUR, MINUTE, SECOND, EOMONTH, ABS, ROUND, ROUNDUP, ROUNDDOWN, INT, MOD, POWER, SQRT, LOG, LN, EXP, FLOOR, CEILING, TRUNC, SIGN, PMT, FV, PV, RATE, NPER, NPV, IRR, VLOOKUP, XLOOKUP, PI, ROW, COLUMN, ROWS, COLUMNS, ADDRESS, TRANSPOSE

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

**980+ tests**: addressing (17), patch (21), style (60), datetime (8), codec (42), batch (46), syntax (18), optics (34), OOXML (24), streaming (18), RichText (5), formula (51+), v0.3.0 regressions (36), CLI (100+)

## Documentation

- **Roadmap**: `docs/plan/roadmap.md` (single source of truth for work scheduling)
- **Status**: `docs/STATUS.md` (current capabilities, 980+ tests)
- **Design**: `docs/design/*.md` (architecture, purity charter, domain model)
- **Reference**: `docs/reference/*.md` (examples, scaffolds, performance guide)

## AI Agent Workflow

**Issue Tracking**: [GitHub Issues](https://github.com/TJC-LP/xl/issues)

1. Check GitHub Issues for available tasks
2. Run `gtr list` to verify no conflicting worktrees
3. Create worktree: `gtr create issue-XXX-description`
4. After PR merge: close GitHub issue, update STATUS.md if needed

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

**Import order**: Java/javax → Scala stdlib → Cats → Project → Tests

## CI/CD

GitHub Actions: `./mill __.checkFormat` → `./mill __.compile` → `./mill __.test`

Coursier + Mill artifacts cached for 2-5x speedup.
