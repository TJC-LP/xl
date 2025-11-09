
# Architectural Decisions (ADRs)

## ADR‑001: Pure core, effect interpreters
- Decision: All IO isolated in `xl-cats-effect`. Core & OOXML mapping are pure.
- Consequence: Easier testing; determinism; slightly more plumbing in interpreters.

## ADR‑002: BigDecimal over Double
- Decision: Use `BigDecimal` to preserve precision; offer `Double` codecs separately.
- Consequence: Slower arithmetic; mitigated with optional `Double` fast path.

## ADR‑003: Named tuples for ad‑hoc schemas
- Decision: Scala 3.7 named tuples enable type‑safe, ergonomic row access.
- Consequence: Add binder macro and tests; great DX win.
