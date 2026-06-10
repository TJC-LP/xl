# FAQ & Glossary

**Last Updated**: 2026-06-10

---

## Frequently Asked Questions

**Q: Why not reuse POI objects?**
A: We want immutability, type safety, and compile-time guarantees; POI's design makes that difficult.

**Q: What's the fastest way to use XL in a script?**
A: One import — `import com.tjclp.xl.scripting.{*, given}` with scala-cli (since 0.11.0). See [reference/scripting.md](reference/scripting.md).

**Q: Will charts look identical to Excel's?**
A: Charts are not implemented yet; when added they will target ChartML with deterministic output.

**Q: Do you support `.xls`?**
A: Not initially. Focus is `.xlsx`/`.xlsm`. BIFF can be a separate module later.

---

## Glossary of Terms

- **ARef** — Absolute cell reference (opaque 64-bit packing of (row, col)).
- **Canonicalization** — Transforming a structure into a stable, unique representative.
- **EMU** — English Metric Units (DrawingML length unit).
- **GADT** — Generalized Algebraic Data Type (typed formulas).
- **Named tuple** — Scala 3.7 tuples with field names.
- **Patch** — First-class edit value with a monoid, acting lawfully on a target.
- **RecalcResult** — Result of `Workbook.recalculate` (0.11.0): recalculated workbook with cached formula values, plus per-sheet values and per-cell `CellEvalError`s.
- **Scripting prelude** — `com.tjclp.xl.scripting.{*, given}` (0.11.0): one import bundling core API, DSL, compile-time literals, formula evaluation, sync `Excel`, and streaming `ExcelIO` for scripts.
- **SST** — Shared Strings Table; deduplicates strings in OOXML.
