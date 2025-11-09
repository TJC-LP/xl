
# XL — Pure Scala 3.7 Excel (OOXML) Library

**Package root:** `com.tjclp.xl`  
**Scala:** 3.7.x (named tuples, opaque types, enums, inline, macros, derivation)  
**Design pillars:** Mathematical purity · Determinism · Ease‑of‑use · Performance · Beautiful syntax

## Why XL?
- **Predictable & safe:** A law‑governed core with equational reasoning; effect isolation yields testability.
- **Type‑directed semantics:** Addresses, ranges, styles, formulas, rows/tables are encoded with precise types.
- **Performance by construction:** Streaming, canonicalization, constant‑time address math, and compile‑time DSLs.

## What’s in this spec
This repository of Markdown documents is the **source of truth** for implementation. It contains formal models, mapping tables to OOXML parts, macro designs, test plans, benchmarks, and security notes.

### Table of Contents
- [00-executive-summary.md](./00-executive-summary.md)
- [01-purity-charter.md](./01-purity-charter.md)
- [02-domain-model.md](./02-domain-model.md)
- [03-addressing-and-dsl.md](./03-addressing-and-dsl.md)
- [04-formula-system.md](./04-formula-system.md)
- [05-styles-and-themes.md](./05-styles-and-themes.md)
- [06-drawings.md](./06-drawings.md)
- [07-charts.md](./07-charts.md)
- [08-tables-and-pivots.md](./08-tables-and-pivots.md)
- [09-codecs-and-named-tuples.md](./09-codecs-and-named-tuples.md)
- [10-patches-and-optics.md](./10-patches-and-optics.md)
- [11-ooxml-mapping.md](./11-ooxml-mapping.md)
- [12-api-surface.md](./12-api-surface.md)
- [13-streaming-and-performance.md](./13-streaming-and-performance.md)
- [14-error-model-and-safety.md](./14-error-model-and-safety.md)
- [15-testing-and-laws.md](./15-testing-and-laws.md)
- [16-build-and-modules.md](./16-build-and-modules.md)
- [17-macros-and-syntax.md](./17-macros-and-syntax.md)
- [18-roadmap.md](./18-roadmap.md)
- [19-examples.md](./19-examples.md)
- [20-glossary.md](./20-glossary.md)
- [21-appendix-ooxml-cheatsheet.md](./21-appendix-ooxml-cheatsheet.md)
- [22-appendix-benchmarks-plan.md](./22-appendix-benchmarks-plan.md)
- [23-security.md](./23-security.md)
- [24-contributing.md](./24-contributing.md)
- [25-style-guide.md](./25-style-guide.md)
- [26-decisions-adr.md](./26-decisions-adr.md)
- [27-faq.md](./27-faq.md)

*Generated on 2025-11-09*
