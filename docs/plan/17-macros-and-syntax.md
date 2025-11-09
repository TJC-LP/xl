
# Macros & Syntax — Zero‑Overhead Ergonomics

## Implemented literals
- `cell"…"` and `range"…"` with **CT validation**, friendly diagnostics.
- `nf"…"`, `rgb"…"`, `theme"…"`: CT tokenization → ADTs.

## Batch put macro
```scala
sheet.put(
  cell"A1" -> "Revenue",
  cell"B1" -> BigDecimal(42),
  cell"C1" -> fx"SUM(A2:A10)"
)
```
**Expands to** a nested chain of `updateCell` calls without allocating intermediate collections.

## Row writers (case class & named tuples)
- Inline derivation avoids typeclass dictionary lookups on hot paths.
- Named tuple binder macro extracts label tuple at compile time and compiles header indices.

## Chart block macro
- Validates required fields inside `series { … }` and axes.
- Normalizes series ordering; folds constants.

## Path macro
- Type‑checks the path string against the ADT, then expands to nested `.copy` calls.

## Diagnostics
- Clear, actionable error messages with **precise spans**; suggestions for common mistakes.
