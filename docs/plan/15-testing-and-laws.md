
# Testing & Laws — Property Suites, Goldens, and Visual Checks

## Property suites
- **AddressLaws:** A1 ↔ index round‑trips; range normalization idempotence.
- **PatchLaws:** monoid + action; idempotent setters.
- **StyleLaws:** canonicalization idempotence; equivalence vs printed.
- **FormulaLaws:** If‑fusion; ring/boolean properties; printer/parse subsets.

## Golden files
- Curated `.xlsx` corpus across features; read→write→read stable.
- Canonical diff on normalized XML (attributes ordered, whitespace normalized).

## Optional visual tests
- Pure Chart evaluator → SVG; compare normalized SVG DOM.
