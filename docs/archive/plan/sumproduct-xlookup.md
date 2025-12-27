> **📁 ARCHIVED** — This plan was completed. SUMPRODUCT and XLOOKUP are implemented in xl-evaluator (81 functions total).

# WI-09h: SUMPRODUCT and XLOOKUP Functions

**Status**: ✅ Complete
**Priority**: Medium-High (array operations + modern lookup)
**Estimated Effort**: 1 day
**Last Updated**: 2025-11-27

---

## Metadata

| Field | Value |
|-------|-------|
| **Owner Module** | `xl-evaluator` |
| **Touches Files** | `TExpr.scala`, `FunctionParser.scala`, `Evaluator.scala`, `FormulaPrinter.scala`, `DependencyGraph.scala` |
| **Dependencies** | WI-09f complete (conditional functions - for CriteriaMatcher) |
| **Enables** | Array formulas, modern lookups, weighted calculations |
| **Parallelizable With** | WI-11 (Charts), WI-19 (Row/Col Props) — different modules |
| **Merge Risk** | Low (additive changes within xl-evaluator module) |

---

## Work Items

| ID | Description | Type | Files | Status | PR |
|----|-------------|------|-------|--------|----|
| `WI-09h-1` | Add SUMPRODUCT TExpr case + smart constructor | Feature | `TExpr.scala` | ⏳ Not Started | - |
| `WI-09h-2` | Add SUMPRODUCT parser | Feature | `FunctionParser.scala` | ⏳ Not Started | - |
| `WI-09h-3` | Implement SUMPRODUCT evaluator | Feature | `Evaluator.scala` | ⏳ Not Started | - |
| `WI-09h-4` | Add XLOOKUP TExpr case + smart constructor | Feature | `TExpr.scala` | ⏳ Not Started | - |
| `WI-09h-5` | Add XLOOKUP parser (3-6 args) | Feature | `FunctionParser.scala` | ⏳ Not Started | - |
| `WI-09h-6` | Implement XLOOKUP evaluator (all 4 match modes) | Feature | `Evaluator.scala` | ⏳ Not Started | - |
| `WI-09h-7` | Add FormulaPrinter cases | Feature | `FormulaPrinter.scala` | ⏳ Not Started | - |
| `WI-09h-8` | Update DependencyGraph | Feature | `DependencyGraph.scala` | ⏳ Not Started | - |
| `WI-09h-9` | Test suite (~45 tests) | Tests | `SumProductXLookupSpec.scala` (new) | ⏳ Not Started | - |

---

## Functions to Implement

| Function | Signature | Description |
|----------|-----------|-------------|
| **SUMPRODUCT** | `SUMPRODUCT(array1, [array2], ...)` | Multiply corresponding elements, sum products |
| **XLOOKUP** | `XLOOKUP(lookup, lookup_array, return_array, [if_not_found], [match_mode], [search_mode])` | Modern flexible lookup |

---

## Function 1: SUMPRODUCT

### Semantics

```
SUMPRODUCT(A1:A3, B1:B3) = A1*B1 + A2*B2 + A3*B3
```

1. Multiply corresponding elements across all arrays
2. Sum all products
3. All arrays must have same dimensions

### TExpr Case

```scala
case SumProduct(ranges: List[CellRange]) extends TExpr[BigDecimal]
```

### Key Behaviors

| Scenario | Behavior |
|----------|----------|
| Non-numeric cells | Treat as 0 |
| Boolean TRUE/FALSE | Coerce to 1/0 |
| Empty cells | Treat as 0 |
| Single array | Sum of values (1 * each) |
| Dimension mismatch | Error |

### Arity

`Arity.AtLeast(1)` — one or more arrays

---

## Function 2: XLOOKUP

### Signature

```
XLOOKUP(lookup_value, lookup_array, return_array, [if_not_found], [match_mode], [search_mode])
```

### Advantages over VLOOKUP

| Feature | VLOOKUP | XLOOKUP |
|---------|---------|---------|
| Direction | Right only | Any direction |
| Arrays | Combined table | Separate lookup/return |
| Not found | #N/A error | Custom default |
| Match modes | 2 (exact/approx) | 4 modes |

### Match Modes

| Value | Mode | Description |
|-------|------|-------------|
| 0 | Exact | Must match exactly (default) |
| -1 | Next smaller | Exact, or next smaller value |
| 1 | Next larger | Exact, or next larger value |
| 2 | Wildcard | Use `*` and `?` patterns |

### Search Modes

| Value | Mode | Description |
|-------|------|-------------|
| 1 | First-to-last | Search from top/left (default) |
| -1 | Last-to-first | Search from bottom/right |
| 2 | Binary ascending | Binary search, sorted ascending |
| -2 | Binary descending | Binary search, sorted descending |

### TExpr Case

```scala
case XLookup(
  lookupValue: TExpr[?],
  lookupArray: CellRange,
  returnArray: CellRange,
  ifNotFound: Option[TExpr[?]],
  matchMode: TExpr[Int],
  searchMode: TExpr[Int]
) extends TExpr[Any]
```

### Arity

`Arity.Range(3, 6)` — 3 required, 3 optional

**Defaults**:
- `if_not_found`: None (returns #N/A error)
- `match_mode`: 0 (exact)
- `search_mode`: 1 (first-to-last)

---

## Implementation Details

### SUMPRODUCT Evaluator

```scala
case TExpr.SumProduct(ranges) =>
  val refLists = ranges.map(_.cells.toList)
  val firstLen = refLists.headOption.map(_.length).getOrElse(0)

  if refLists.forall(_.length == firstLen) then
    val sum = (0 until firstLen).foldLeft(BigDecimal(0)) { (acc, idx) =>
      val product = ranges.foldLeft(BigDecimal(1)) { (prod, range) =>
        val ref = range.cells.toList(idx)
        val value = decodeNumericOrBool(sheet(ref)).getOrElse(BigDecimal(0))
        prod * value
      }
      acc + product
    }
    Right(sum)
  else
    Left(EvalError.EvalFailed("SUMPRODUCT: all arrays must have same dimensions", None))
```

### XLOOKUP Evaluator (Pseudocode)

```scala
case TExpr.XLookup(lookupExpr, lookupArray, returnArray, ifNotFound, matchModeExpr, searchModeExpr) =>
  for
    lookup <- eval(lookupExpr, sheet, clock)
    matchMode <- eval(matchModeExpr, sheet, clock)
    searchMode <- eval(searchModeExpr, sheet, clock)
    _ <- validateDimensions(lookupArray, returnArray)
    result <- matchMode match
      case 0 => exactMatch(lookup, lookupArray, returnArray, searchMode)
      case -1 => nextSmallerMatch(lookup, lookupArray, returnArray, searchMode)
      case 1 => nextLargerMatch(lookup, lookupArray, returnArray, searchMode)
      case 2 => wildcardMatch(lookup, lookupArray, returnArray, searchMode)  // Uses CriteriaMatcher
  yield result
```

---

## Dependencies

### Prerequisites

- ✅ WI-07: Formula Parser
- ✅ WI-08: Formula Evaluator
- ✅ WI-09e: VLOOKUP (reference implementation)
- ✅ WI-09f: CriteriaMatcher (for wildcard mode)

### Enables

- Array formula patterns
- Modern lookup replacement for VLOOKUP
- Weighted average calculations
- Matrix operations foundation

---

## Test Coverage

### SUMPRODUCT Tests (~20)

| Category | Tests |
|----------|-------|
| Basic | 2-array, 3-array, single array |
| Type coercion | Non-numeric→0, TRUE→1, FALSE→0 |
| Edge cases | Empty cells, empty range |
| Errors | Dimension mismatch |
| Round-trip | Parse → print → parse |

### XLOOKUP Tests (~25)

| Category | Tests |
|----------|-------|
| Match modes | Exact, smaller, larger, wildcard |
| Search modes | First, last, binary asc, binary desc |
| Not found | With/without default |
| Edge cases | Single cell, horizontal, vertical |
| Errors | Dimension mismatch, invalid mode |
| Round-trip | Parse → print → parse |

---

## Estimated Effort

| Phase | Effort |
|-------|--------|
| SUMPRODUCT (TExpr + Parser + Evaluator + Printer) | 2 hours |
| XLOOKUP (TExpr + Parser + Evaluator + Printer) | 3 hours |
| Tests | 2 hours |
| **Total** | **~7 hours** |

---

## Definition of Done

- [ ] SUMPRODUCT: TExpr case + smart constructor
- [ ] SUMPRODUCT: FunctionParser with AtLeast(1) arity
- [ ] SUMPRODUCT: Evaluator with dimension validation
- [ ] XLOOKUP: TExpr case + smart constructor
- [ ] XLOOKUP: FunctionParser with Range(3,6) arity + defaults
- [ ] XLOOKUP: Evaluator with all 4 match modes + 4 search modes
- [ ] FormulaPrinter: Both functions (plain + debug)
- [ ] DependencyGraph: Both functions
- [ ] Tests: ~45 tests covering all scenarios
- [ ] CLAUDE.md: Function count updated (28 → 30)
