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
