# Changelog

All notable changes to the XL project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

*No unreleased changes*

---

## [0.5.0-RC2] - 2025-12-27

### Changed

- **36-39% faster streaming writes** via SaxStax backend and DirectSaxEmitter (PR #145)
  - New `DirectSaxEmitter` bypasses intermediate XML construction
  - Optimized attribute handling and string building
  - Benchmarked against Apache POI with consistent improvements

### Documentation

- Cleaned up internal planning docs, consolidated roadmap
- Removed 14 obsolete planning documents

---

## [0.5.0-RC1] - 2025-12-25

### Added

- **34 new formula functions** bringing total to 81:
  - **Financial (TVM)**: PMT, FV, PV, RATE, NPER - time value of money calculations
  - **Statistical**: MEDIAN, STDEV, STDEVP, VAR, VARP - descriptive statistics
  - **Type checking**: ISNUMBER, ISTEXT, ISBLANK, ISERR, ISERROR - value type inspection
  - **Conditional aggregation**: AVERAGEIF, AVERAGEIFS - conditional averaging
  - **Count/Reference**: COUNTBLANK, ROW, COLUMN, ROWS, COLUMNS, ADDRESS - cell reference utilities
  - **Math**: SQRT, MOD, POWER, LOG, LN, EXP, FLOOR, CEILING, TRUNC, SIGN, INT - numeric operations

- **Variadic aggregate functions**: SUM, COUNT, AVERAGE, MIN, MAX, MEDIAN, STDEV, VAR now support Excel-compatible variadic syntax
  - `=SUM(1,2,3)` - individual values
  - `=SUM(A1:A5, B1:B5)` - multiple ranges
  - `=SUM(A1, 5, B1:B3)` - mixed ranges and values

- **Dynamic `xl functions` command**: Now shows all 81 functions from registry instead of hardcoded list

### Changed

- **Formula system refactored**: Reorganized into modular traits for better maintainability
  - Split `FunctionSpecs.scala` into focused modules (Aggregate, Financial, Lookup, etc.)
  - Extracted `TExpr` helpers into dedicated traits
  - Deduplicated range extraction and criteria parsing helpers

### Fixed

- **Formula parsing edge cases**: Fixed scientific notation, date arithmetic, cross-sheet dependencies
- **WartRemover compliance**: Added `@SuppressWarnings` annotations for intentional type casts in formula DSL

---

## [0.4.3] - 2025-12-19

### Added

- **Rasterizer fallback chain**: Multiple SVG-to-raster backends for native image support in environments where Batik/AWT is unavailable (like Claude.ai)
  - Batik (built-in), cairosvg, rsvg-convert, resvg, ImageMagick
  - Automatic fallback tries each rasterizer in order until one succeeds
  - New `--rasterizer <name>` flag to force a specific backend

- **Targeted range evaluation**: `--eval` now only evaluates formulas within the viewed range (plus their dependencies) instead of the entire sheet
  - Significantly faster for small ranges on large sheets
  - Complexity: O(R + D) where R = range cells, D = dependencies vs O(N) for all formulas

### Fixed

- **Cross-sheet formula evaluation**: Fixed `--eval` flag to work correctly with cross-sheet references like `='Sheet2'!A1`
- **CairoSvg JPEG handling**: Now explicitly rejects JPEG format to allow fallback chain to find appropriate rasterizer

### Breaking Changes

- **`--use-imagemagick` flag removed**: Replaced with `--rasterizer <name>` option for specifying rasterizer backend
  - Migration: `--use-imagemagick` → `--rasterizer imagemagick`

---

## [0.4.2] - 2025-12-19

### Added

- **PI() function**: Mathematical constant (3.14159...) now available in formulas
- **Headless formula evaluation**: `xl eval "=PI()*2"` works without a file for constant formulas
- **`xl functions` command**: Lists all 47 supported formula functions with categories and descriptions
- **Cross-sheet formula references**: `=Sheet1!A1` syntax now fully supported in evaluator

### Fixed

- **PolyRef type safety**: Eliminated unsafe `asInstanceOf` casts in formula evaluator by resolving polymorphic references at parse time
- **VLOOKUP text lookup**: Text lookups now correctly return text results
- **Nested formula evaluation**: Formula cells with cached values now properly extract values during evaluation
- **Wildcard pattern matching**: `Widget*` patterns in SUMIF/COUNTIF now match correctly (TJC-353)
- **Addition with cell references**: Fixed regression where `=A1+B1` failed to evaluate

---

## [0.4.0] - 2025-12-17

### Added

- **Pure JVM rasterization** (PR #91): Apache Batik-based SVG to PNG/JPEG conversion
  - No external ImageMagick dependency required
  - Falls back to ImageMagick in GraalVM native images (AWT unavailable)
  - `--use-imagemagick` flag for explicit ImageMagick usage

- **Indexed color support**: Proper handling of legacy Excel indexed colors
  - Maps indices 0-63 to RGB values per ECMA-376 specification
  - Preserves indexed colors during read/write round-trips

- **Per-side border styling**: New CLI flags for individual border control
  - `--border-top`, `--border-right`, `--border-bottom`, `--border-left`
  - Border merging is now per-side (matches Excel behavior)

- **Date calculation functions**: 6 new Excel-compatible date functions
  - `EOMONTH` - End of month N months from start date
  - `EDATE` - Same day N months from start date
  - `DATEDIF` - Difference between dates (Y/M/D/MD/YM/YD units)
  - `NETWORKDAYS` - Working days between dates (excludes weekends/holidays)
  - `WORKDAY` - Date N working days from start (excludes weekends/holidays)
  - `YEARFRAC` - Year fraction between dates (5 day-count conventions)

### Changed

- **Mill upgraded to 1.1.0-RC3**: Improved Scala 3.7.3 support and build performance
  - Tested across all CI platforms (Linux/macOS/Windows)

- **Style merging is now default**: `xl style` command merges with existing styles
  - Use `--replace` flag for previous replacement behavior
  - Preserves existing formatting when adding new properties

- **Formula evaluation warnings**: `--eval` flag now prints warnings to stderr when formulas fail to evaluate instead of silently continuing

### Breaking Changes

- **`xl style` command now merges by default**: Style commands now merge with existing cell styles instead of replacing them entirely. Use `--replace` flag for the previous behavior.
  - Before: `xl style A1 --bold` replaced entire cell style with just bold
  - After: `xl style A1 --bold` adds bold to existing style
  - Migration: Add `--replace` flag to scripts that rely on replacement behavior

### Fixed

- **Batik @SuppressWarnings scope**: Moved annotation from object level to method level for tighter suppression

- **DRY forkArgs**: Extracted shared JVM options to `BuildConfig.lazyValsJvmArgs` constant

- **Partial file cleanup**: Batik rasterizer now cleans up partial files on failure using idiomatic `Using`

---

## [0.3.0] - 2025-12-13

### Added

- **Financial Functions** (PR #77): 2 new date-aware financial functions
  - `XNPV` - Net present value with irregular dates
  - `XIRR` - Internal rate of return with irregular dates

- **Security Hardening** (PR #78): Protection against malicious files
  - ZIP bomb detection with configurable thresholds
  - Formula injection guards (escape `=`, `+`, `-`, `@` prefixes)
  - Configurable via `SecurityConfig` in read/write operations

### Fixed

- **Style components not rendering** (PR #80): When adding new styles (bold, fill color) to existing files, font/fill/border components weren't being added to `styles.xml`. New styles now correctly include their component definitions.

- **Column widths lost on save** (PR #80): Column properties set via API/CLI were overwritten by preserved XML on subsequent operations. Domain properties now take priority over preserved XML.

---

## [0.2.3] - 2025-12-07

### Changed

- **GraalVM upgraded to 25.0.1 LTS**: Native image builds now use JDK 25
  - 22% heap reduction from Compact Object Headers (JEP 519)
  - Up to 8x faster String::hashCode for constant keys
  - Better Vector API support for numeric operations

### Added

- **Formula validation**: `CellValue.formula()` smart constructor validates:
  - Non-empty expression (OOXML requirement)
  - Cached value is not a nested Formula (illegal state)
- **Render constants documented**: Magic numbers extracted to `RenderUtils`
  - `IndentPxPerLevel` (21px) - Excel indent spacing
  - `InterRunGapPx` (4px) - Rich text run spacing
- **Enhanced scaladocs**: Union return type pattern explained
  - Documents `transparent inline` type narrowing behavior
  - Examples for compile-time vs runtime validation paths

---

## [0.2.2] - 2025-12-06

### Fixed

- **WartRemover warnings eliminated**: All 82 compile-time warnings resolved
  - Source code: Refactored `isInstanceOf` to pattern matching in `PutLiteral.scala`
  - Source code: Added `@nowarn` for unreachable `PolyRef` case in `FormulaShifter.scala`
  - Test code: Added class-level `@SuppressWarnings` annotations following project policy

### Changed

- **Release workflow improved**: `/release-prep` command now creates annotated tags with CHANGELOG extraction
  - Ensures GitHub releases display proper release notes instead of commit messages
  - Tags are now annotated (`git tag -a`) with message extracted from CHANGELOG.md

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
