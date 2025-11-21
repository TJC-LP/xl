# FAQ & Glossary

**Last Updated**: 2025-11-20

---

## Frequently Asked Questions

**Q: Why not reuse POI objects?**
A: We want immutability, type safety, and compile-time guarantees; POI's design makes that difficult.

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
- **SST** — Shared Strings Table; deduplicates strings in OOXML.
