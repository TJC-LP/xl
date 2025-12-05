# Changelog

All notable changes to the XL project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

*No unreleased changes*

---

## [0.2.1] - 2025-12-05

### Added

- **`xl` aggregate package**: Single artifact bundling all modules for simpler onboarding
  - Depend on `com.tjclp::xl:0.2.1` instead of 4 separate modules
  - Individual modules (`xl-core`, `xl-ooxml`, `xl-cats-effect`, `xl-evaluator`) remain available for minimal footprint
- **`/release-prep` slash command**: Streamlined version bump workflow for releases

### Changed

- Documentation updated with simplified single-dependency examples
- Examples now use aggregate package by default

---

## [0.2.0] - 2025-12-05

Major CLI enhancement release with 38 commits across 8 PRs.

### Added

- **CLI Sheet Management Commands** (PR #63)
  - `rename` - Rename sheets
  - `move` - Reorder sheets within workbook
  - `copy` - Duplicate sheets
  - `merge` - Combine multiple sheets
  - `add-sheet` - Add new sheet with optional positioning (`--before`, `--after`)
  - `remove-sheet` - Delete sheets
  - `stats` - Sheet statistics (cell count, ranges, formulas)
  - `names` - List defined names/named ranges
  - Qualified reference support: `Sheet1!A1:B10` syntax

- **Formula Dragging** (PR #64)
  - Excel-style formula dragging with `$` anchoring support
  - `putf` command supports `--from` flag for formula range application
  - Per-endpoint anchor modes: relative (`A1`), absolute (`$A$1`), mixed (`$A1`, `A$1`)
  - Example: `xl putf B2:B10 "=SUM($A$1:A1)" --from B2`

- **CLI Syntax Improvements** (PR #65)
  - Map syntax for batch `put` and `style` operations with compile-time validation
  - `--eval` flag for formula caching in repeated evaluations
  - Improved batch operation duplicate detection

- **Repeatable `--sheet` Flag** (PR #69)
  - Create multiple sheets in one command: `xl new out.xlsx --sheet Data --sheet Summary`
  - Backward compatible with existing `--sheet-name` option

- **Example Improvements** (PR #66)
  - Shebang support for direct script execution (`#!/usr/bin/env -S scala-cli shebang`)
  - Centralized dependency management via `project.scala`
  - Standardized naming conventions (underscores instead of hyphens)

### Fixed

- **AVERAGE in Nested Formulas** (PR #68)
  - Dedicated `TExpr.Average` case prevents crash in expressions like `=SUM(A1)+AVERAGE(B1:B2)`
  - Single-pass optimization for MIN/MAX/AVERAGE using foldLeft
  - Added iterator consumption regression tests

- **CLI Batch Performance** (PR #66)
  - O(N) batch operations using grouped sheet updates (was O(N²))
  - File descriptor leak prevention in `BatchParser`

- **Excel File Corruption** (PR #62)
  - Fixed corruption during surgical modification operations
  - Improved release body generation from tag messages

### Changed

- **CLI Architecture Refactor** (PR #66)
  - Split command handlers into focused modules for maintainability
  - Extracted helper modules: `SheetResolver`, `ValueParser`, `StyleBuilder`, `BatchParser`
  - Extracted enum types: `CliCommand`, `ViewFormat`
  - Extracted shared renderer utilities: `RendererCommon`

- **CLI Behavior**
  - Requires explicit `--sheet` flag for unqualified ranges (prevents ambiguity)
  - Clearer error messages for empty range evaluations

- **Documentation**
  - Updated xl-cli skill with full CLI capabilities
  - Clarified unsafe import usage in README

---

## [0.1.4] - 2025-12-04

### Fixed

- Multiple native binary build improvements
- Release asset naming consistency

---

## [0.1.3] - 2025-12-04

### Added

- **GraalVM Native Image Support**
  - Zero-dependency native binaries for macOS (Intel/ARM) and Linux
  - `xl` CLI runs without JDK installation
- **Windows Support**
  - Native Windows binary (`xl.exe`)
  - `mill.bat` for Windows build compatibility
- **Release Automation**
  - CLI tarball included in GitHub releases
  - Versioned skill package in releases

### Fixed

- Dynamic version from `PUBLISH_VERSION` environment variable
- Windows native image build shell configuration

---

## [0.1.2] - 2025-12-04

### Added

- **Maven Central Publishing**
  - Automated publishing via GitHub Actions
  - GPG signing with Mill's `SonatypeCentralPublishModule`

### Fixed

- Release workflow environment variable conventions
- GPG key import process

---

## [0.1.1] - 2025-12-04

### Fixed

- **OOXML Compatibility**
  - Bind `r` namespace in sheet elements for openpyxl compatibility
  - Prevent chained write corruption with atomic temp file strategy

---

## [0.1.0] - 2025-12-04

Initial public release of XL - the pure functional Excel library for Scala 3.

### Added

- **Core Domain Model** (`xl-core`)
  - Pure functional Excel types: `Cell`, `Sheet`, `Workbook`, `CellValue`
  - Opaque types with zero overhead: `Column`, `Row`, `ARef` (packed 64-bit), `SheetName`, `CellRange`
  - `Patch` and `StylePatch` monoids for composable modifications
  - `CellStyle` with font, fill, border, number format, alignment
  - Compile-time validated macros: `ref"A1"`, `fx"=SUM(A1:A10)"`, `money"$1,000"`
  - DSL operators: `:=` for assignment, `++` for patch composition

- **OOXML Read/Write** (`xl-ooxml`)
  - `XlsxReader` and `XlsxWriter` for .xlsx files
  - Shared strings table (SST) with intelligent caching
  - Style deduplication via `CellStyle.canonicalKey`
  - Canonical XML output for byte-identical, diffable files
  - Surgical modification with byte-perfect preservation of unmodified parts

- **Cats Effect Integration** (`xl-cats-effect`)
  - `Excel[F]` type class and `ExcelIO` convenience object
  - SAX-based streaming reads for O(1) memory consumption
  - `fs2` streaming write support for large workbooks
  - `ExcelR[F]` for explicit error handling without exceptions

- **Formula Evaluator** (`xl-evaluator`)
  - `TExpr[A]` GADT for typed formula AST
  - 30 Excel functions: SUM, COUNT, AVERAGE, MIN, MAX, IF, AND, OR, NOT, CONCATENATE, LEFT, RIGHT, LEN, UPPER, LOWER, TODAY, NOW, DATE, YEAR, MONTH, DAY, NPV, IRR, VLOOKUP, SUMIF, COUNTIF, SUMIFS, COUNTIFS, SUMPRODUCT, XLOOKUP
  - `DependencyGraph` with cycle detection (Tarjan's SCC algorithm)
  - Round-trip verified `FormulaParser` and `FormulaPrinter`

- **CLI** (`xl-cli`)
  - Commands: `view`, `cell`, `search`, `put`, `putf`, `batch`, `export`, `new`, `sheets`
  - Export formats: CSV, JSON, TSV, HTML, SVG, Markdown
  - LLM-optimized output with `--format` options

- **Additional Features**
  - Cell-level codecs for 9 types: String, Int, Long, Double, BigDecimal, Boolean, LocalDate, LocalDateTime, RichText
  - Auto-format inference: LocalDate → Date, LocalDateTime → DateTime, BigDecimal → Decimal
  - RichText DSL: `"Bold".bold.red + " normal " + "Italic".italic.blue`
  - HTML export with inline CSS preservation
  - Optics library (Lens, Optional) with zero external dependencies
  - Row/column properties: height, width, hidden, outline level

- **High-Fidelity Rendering**
  - SVG renderer with all 14 Excel border styles
  - HTML renderer with content-aware alignment
  - Theme color resolution from OOXML theme XML
  - Text overflow into adjacent empty cells

### Performance

- **Streaming reads**: 35% faster than Apache POI at 1k rows, competitive at 10k rows
- **O(1) memory**: SAX-based streaming maintains constant memory regardless of file size
- **Surgical writes**: 11x speedup for unmodified workbooks (verbatim copy optimization)
- **Partial modification**: 2-5x speedup when only some sheets are modified

### Testing

- 731+ tests across all modules
- Property-based law verification for Monoid, Lens, and round-trip laws
- ScalaCheck generators for all core types
