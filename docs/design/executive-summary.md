
# Executive Summary — XL (Scala 3.7, com.tjclp.xl)

**Vision:** A purely functional, mathematically rigorous Excel library that manipulates OOXML directly and exposes an idiomatic Scala 3.7 API with **opaque types, enums, named tuples, inline, macros, and derivation**. Effects are **fully isolated**; the core is **100% pure**. We target **elegant syntax** and **POI‑level feature breadth** with **deterministic, law‑governed semantics**.

**Non‑negotiables**
- Purity and totality (no `null`, no partial functions, no thrown exceptions).
- Strong typing (addresses, ranges, styles, formulas, schemas).
- Deterministic printers/parsers; canonicalization for stable diffs.
- Performance: streaming, zero‑overhead newtypes, inlined codecs, macro‑fused syntax.
- Safety: zip/XXE limits, XLSM preservation (never execute macros), formula‑injection guards.

**MVP scope**
- One‑sheet read/write with SST + styles subset.
- Address/range macros; `Patch` edits; codecs, row derivation; named‑tuple binder.
- Streaming reader/writer (constant memory).

**Roadmap**: progress from core + IO to styling, drawings, charts, tables, conditional formatting, evaluator.
