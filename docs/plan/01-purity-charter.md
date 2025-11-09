
# Purity Charter (Expanded)

## 1) Core purity & effect isolation
- `xl-core` and `xl-ooxml` expose **pure functions only**; no side‑effects, no clocks, no randomness.
- `xl-cats-effect` contains **the only interpreters** for: ZIP I/O, file system, streams.
- All IO return types are **`F[_]`** with typeclass constraints (`Sync`, `Async`) and **no hidden global state**.

## 2) Totality & defensive programming
- Abolish `null`; use `Option`.
- No partial functions; exhaustive `match` on enums.
- All validation is **first class**: return `Either[XLError, A]` or `ValidatedNec[Validation, A]`.

## 3) Laws & reasoning
- **Lens laws** for optics, **Monoid laws** for patches & style patches, **Action law** for `applyPatch`.
- **Determinism:** same inputs → same outputs including **byte‑level equivalence** after canonicalization.
- **No surprises:** the `Workbook` model is a finite algebra with explicit defaults; nothing implicit.

## 4) Canonicalization & equality
- Canonicalize style structures, ranges, anchors, and XML printing order.
- Provide `Eq`/`Ordering` instances based on canonical forms; forbid referential variants.

## 5) Explicit non‑goals (initially)
- No legacy `.xls` BIFF; `.xlsx/.xlsm` only.
- Charts editing fidelity starts with core chart types; advanced options staged.

## Law check inventory
- Address round‑trips; range normalization.
- Patch monoid + action; style canonicalization equivalence.
- Formula ring/boolean laws; If‑fusion; printer/parse round‑trips (subset).
