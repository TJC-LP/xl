# Architectural Decisions (ADRs)

## ADR-001: Pure core, effect interpreters
**Date**: 2024-11 (P0)
**Status**: ✅ Implemented

- **Decision**: All IO isolated in `xl-cats-effect`. Core & OOXML mapping are pure.
- **Rationale**: Easier testing, determinism, referential transparency
- **Consequence**: Slightly more plumbing in interpreters, but massive testing/reasoning benefits

## ADR-002: BigDecimal over Double
**Date**: 2024-11 (P1)
**Status**: ✅ Implemented

- **Decision**: Use `BigDecimal` for cell values to preserve precision; offer `Double` codecs separately.
- **Rationale**: Financial calculations require exact decimal arithmetic
- **Consequence**: Slower arithmetic; mitigated with optional `Double` fast path for performance-critical code

## ADR-003: Named tuples for ad-hoc schemas
**Date**: 2024-11 (P6 future)
**Status**: ⬜ Deferred to P6b

- **Decision**: Scala 3.7 named tuples enable type-safe, ergonomic row access.
- **Rationale**: Better DX than HList or raw tuples
- **Consequence**: Add binder macro and tests; great DX win

## ADR-004: StyleRegistry per-sheet
**Date**: 2025-01 (P4 extension)
**Status**: ✅ Implemented

- **Decision**: Each `Sheet` has optional `StyleRegistry` for tracking cell styles
- **Rationale**: Enables bidirectional style mapping without global mutable state; maintains purity
- **Consequence**: Explicit initialization required (`Sheet.withStyleRegistry`), but guarantees referential transparency and correct round-trips

## ADR-005: Optics over implicits for updates
**Date**: 2025-01 (P31)
**Status**: ✅ Implemented

- **Decision**: Provide `Lens[S, A]` and `Optional[S, A]` for functional updates instead of implicit conversions
- **Rationale**: Composable, law-governed, explicit control flow
- **Consequence**: More verbose than lens macros, but zero magic and full type safety

## ADR-006: RichText as CellValue variant
**Date**: 2025-01 (P31)
**Status**: ✅ Implemented

- **Decision**: Add `CellValue.RichText(richtext.RichText)` enum case for multi-formatted text
- **Rationale**: Common Excel feature; enables bold/italic/color within single cell
- **Consequence**: Adds complexity to SST serialization, but worth it for feature parity

## ADR-007: fs2-data-xml for streaming
**Date**: 2025-01 (P5)
**Status**: ✅ Implemented

- **Decision**: Use fs2-data-xml for event-based XML parsing/writing instead of scala-xml
- **Rationale**: Constant-memory streaming (O(1)), 4.5x faster than Apache POI
- **Consequence**: More complex implementation, but enables 1M+ row files without OOM

## ADR-008: CellCodec primitives over derivation
**Date**: 2025-01 (P6)
**Status**: ✅ Implemented (primitives), ⬜ Deferred (derivation)

- **Decision**: Ship 9 inline `given CellCodec[A]` instances for primitives; defer Magnolia/Shapeless derivation to P6b
- **Rationale**: 80% use case covered with zero dependencies; derivation is complex and can wait
- **Consequence**: Users write case class codecs manually for now, but library stays lean

## ADR-009: HTML export as core feature
**Date**: 2025-01 (P31)
**Status**: ✅ Implemented

- **Decision**: Include `sheet.toHtml(range)` in core library, not as plugin
- **Rationale**: Common reporting use case; minimal code (~100 LOC); high value
- **Consequence**: Slight scope creep, but justified by utility

## ADR-010: DateTime as Excel serial numbers
**Date**: 2025-01 (P4 extension)
**Status**: ✅ Implemented

- **Decision**: Convert `LocalDateTime` ↔ Excel serial numbers (Double) at OOXML boundary
- **Rationale**: Excel stores dates as numbers; maintain compatibility
- **Consequence**: Must handle 1900 leap year bug; implemented correctly with comprehensive tests

## ADR-011: Two I/O modes (in-memory and streaming)
**Date**: 2025-11 (P5 review)
**Status**: ✅ Implemented (in-memory + streaming read/write)

- **Decision**: Maintain both in-memory and streaming implementations instead of unified approach
- **Rationale**: No single implementation satisfies both "full features" and "constant memory" requirements. Small files need SST/styles, large files need O(1) memory.
- **Alternatives Considered**:
  - Streaming-only: Would lose SST/styles (unacceptable for <100k row use cases)
  - In-memory only: Would OOM on large files (unacceptable for ETL pipelines)
  - Two-phase streaming: Deferred to P7.5 (complex, not MVP-critical)
- **Consequences**:
  - ✅ Best-of-both-worlds (full features OR constant memory)
  - ✅ Users choose based on needs
  - ❌ Two implementations to maintain
  - ❌ Potential confusion about which to use
- **Mitigation**: Clear documentation in README and performance-guide.md; streaming read was fixed in P6.6 (fs2.io.readInputStream) and now matches streaming write on O(1) memory.

## ADR-012: Compression defaults to DEFLATED
**Date**: 2025-11 (P6.7 planned)
**Status**: ✅ Implemented

- **Decision**: Default to DEFLATED compression with prettyPrint=false for production use
- **Current State**: `WriterConfig.default` already sets `Compression.Deflated` + compact XML; `WriterConfig.debug` keeps STORED + prettyPrint for inspection
- **Rationale**:
  - STORED requires precomputing CRC/size (overhead)
  - Uncompressed files are 5-10x larger
  - Pretty-printed XML is only useful for debugging
  - Production workflows need small files and fast compression
- **Consequences**:
  - ✅ 5-10x smaller files
  - ✅ Faster overall (no CRC computation)
  - ✅ Configurable for debugging (WriterConfig)
  - ❌ Breaking change for users expecting pretty XML
- **Migration**: No migration needed for defaults; set `WriterConfig.debug` explicitly for pretty/ STORED output.

## ADR-013: Streaming reader bug acknowledged and scheduled
**Date**: 2025-11 (P6.6 planned)
**Status**: ✅ Resolved in P6.6

- **Decision**: Replace `readAllBytes()` with `fs2.io.readInputStream` for SST + worksheet entries
- **Outcome**: Streaming read now matches streaming write on constant memory; failure mode documented in archive only.

## ADR-014: TExpr GADT for typed formulas
**Date**: 2025-11-21 (WI-07)
**Status**: ✅ Implemented

- **Decision**: Use GADT (Generalized Algebraic Data Type) for formula AST with type parameter A capturing result type
- **Context**: Formula parser needs to represent Excel formula syntax as typed AST for future evaluation
- **Rationale**:
  - **Type safety**: `TExpr[BigDecimal]` vs `TExpr[Boolean]` prevents mixing incompatible operations at compile time
  - **Totality**: GADT structure enables exhaustive pattern matching, ensuring all cases handled
  - **Evaluation safety**: Type parameter flows through to evaluation result
  - **Comparison operators**: Lt, Lte, Gt, Gte, Eq, Neq integrated as first-class constructors
- **Alternatives Considered**:
  - **Untyped AST** (`TExpr` without type parameter): Rejected due to loss of type safety, would require runtime type checking
  - **Separate ASTs per type** (`NumExpr`, `BoolExpr`, `StrExpr`): Rejected due to code duplication and inability to express polymorphic operations like `If[A]`
  - **HList-based encoding**: Rejected due to complexity and poor error messages
- **Implementation Details**:
  - 17 constructors: Lit, Ref, If, Add/Sub/Mul/Div, Lt/Lte/Gt/Gte/Eq/Neq, And/Or/Not, FoldRange
  - Extension methods for ergonomic construction: `expr1 + expr2`, `expr1 < expr2`, `expr1 && expr2`
  - Smart constructors: `TExpr.sum(range)`, `TExpr.count(range)`, `TExpr.average(range)`
- **Consequences**:
  - ✅ Type-safe formula construction prevents `=SUM(TRUE, FALSE)` at compile time
  - ✅ Parser/printer with round-trip verification (51 tests passing)
  - ✅ Scientific notation support (1.5E10, 3E-5)
  - ✅ FormulaParser produces Either[ParseError, TExpr[?]] (total function)
  - ❌ More complex pattern matching (need asInstanceOf for Eq/Neq when type info lost in runtime parsing)
  - ❌ Requires decode functions for Ref/FoldRange (explicit type conversions via CellCodec)
- **Integration**: Works seamlessly with opaque ARef type (64-bit packing), CellCodec for decoding cell values
- **Testing**: 51 tests verify round-trip laws, operator precedence, error handling

## ADR-015: BigDecimal for formula numeric operations
**Date**: 2025-11-21 (WI-07)
**Status**: ✅ Implemented

- **Decision**: Use `scala.math.BigDecimal` for all numeric formula operations (not Double)
- **Context**: TExpr arithmetic nodes (Add, Sub, Mul, Div) and literals need numeric type
- **Rationale**:
  - **Financial precision**: Excel uses decimal semantics (not binary floating point)
  - **Exact arithmetic**: No rounding errors in calculations like `0.1 + 0.2`
  - **Consistency**: Matches ADR-002 decision (CellValue.Number uses BigDecimal)
  - **Excel parity**: BigDecimal behavior closer to Excel's decimal arithmetic than Double
- **Alternatives Considered**:
  - **Double**: Rejected due to floating-point rounding errors, incompatible with financial calculations
  - **Rational numbers**: Rejected due to complexity and performance overhead for division
  - **Decimal128**: Not available in Scala stdlib, would require external dependency
- **Implementation Details**:
  - All TExpr arithmetic constructors: `case Add(x: TExpr[BigDecimal], y: TExpr[BigDecimal]) extends TExpr[BigDecimal]`
  - Literals: `TExpr.Lit(BigDecimal(42))`
  - Scientific notation parsed to BigDecimal: `BigDecimal("1.5E10")`
- **Consequences**:
  - ✅ Exact precision matches Excel behavior (critical for financial models)
  - ✅ No floating-point rounding errors
  - ✅ Supports arbitrary precision (Excel limit: 15 significant digits)
  - ❌ Slower than Double (~2-10x for arithmetic operations)
  - ❌ Requires explicit BigDecimal construction (not implicit from literals)
- **Performance Impact**: Acceptable - formula complexity in typical workbooks is low (< 100 operations per formula). For performance-critical code, consider caching evaluation results (future: WI-09b dependency graph with memoization)
- **Testing**: Round-trip tests verify BigDecimal values preserved exactly, including scientific notation edge cases
