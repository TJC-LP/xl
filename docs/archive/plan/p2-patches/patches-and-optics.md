
# Patches & Optics — Monoids, Actions, and Lawful Updates

## Patch ADT
```scala
enum Patch derives CanEqual:
  case Put(ref: addr.ARef, value: core.CellValue)
  case SetStyle(ref: addr.ARef, style: core.style.CellStyle)
  case Merge(range: addr.CellRange)
  case Unmerge(range: addr.CellRange)
  case Batch(ps: Vector[Patch])
```
**Monoid:** `empty = Batch(Vector.empty)`, `combine = concat (with flattening)`.

**Action:** `applyPatch: (Sheet, Patch) => Either[SemanticError, Sheet]`

### Laws (proof sketches)
- **Associativity:** concatenation of immutable vectors is associative → composed application yields identical final map.
- **Identity:** empty batch applies no changes → sheet unchanged.
- **Idempotence:** for setters on same key/value, repeating yields same sheet (check equals by canonical style).

## Optics
- Provide lawful `Optional`/`Lens`; or generate **path macros** for zero cost:
```
path.update(sheet, "cells[A1].style.font.sizePt")(_.map(_ => 14.0))
```
- Macro expands to nested `.copy` preserving structural sharing.
