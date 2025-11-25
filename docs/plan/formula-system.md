# Typed Formula System

**Status**: ✅ Complete (WI-07, WI-08, WI-09a/b/c/d all complete)
**Priority**: Delivered
**Actual Effort**: ~4 weeks (WI-07: 1 week, WI-08: 1.5 weeks, WI-09a/b/c/d: 1.5 weeks)
**Last Updated**: 2025-11-24

> **Archive Notice**: This plan has been fully implemented. The formula system is production-ready with 169+ tests, 24 built-in functions, dependency graph with cycle detection, and safe evaluation APIs. This document is retained for historical reference.

---

## Metadata

| Field | Value |
|-------|-------|
| **Owner Modules** | `xl-evaluator` (new module), `xl-core` (formula types) |
| **Touches Files** | New files in `xl-evaluator/`, minimal core changes |
| **Dependencies** | P0-P8 complete |
| **Enables** | Advanced formulas, cross-sheet references, dependency graphs |
| **Parallelizable With** | WI-10 (tables), WI-11 (charts), WI-15 (benchmarks) — completely independent |
| **Merge Risk** | Low (new module, isolated from existing code) |

---

## Work Items

| ID | Description | Type | Files | Status | PR |
|----|-------------|------|-------|--------|----|
| `WI-07` | Formula Parser (string → AST) | Core | `FormulaParser.scala`, `TExpr.scala`, `FormulaPrinter.scala`, `ParseError.scala` | ✅ Complete | Merged |
| `WI-08` | Formula Evaluator (AST → value) | Core | `Evaluator.scala`, `EvalError.scala`, `SheetEvaluator.scala` | ✅ Complete | Merged |
| `WI-09a` | Aggregate Functions (SUM, COUNT, etc.) | Core | `FunctionParser.scala` | ✅ Complete | Merged |
| `WI-09b` | Logical Functions (IF, AND, OR, NOT) | Core | `FunctionParser.scala` | ✅ Complete | Merged |
| `WI-09c` | Text & Date Functions | Core | `FunctionParser.scala` | ✅ Complete | Merged |
| `WI-09d` | Dependency Graph & Cycle Detection | Core | `DependencyGraph.scala` | ✅ Complete | Merged |
| `WI-09e` | Financial Functions (NPV, IRR, VLOOKUP) | Core | `FunctionParser.scala` | ✅ Complete | Merged |

---

## Dependencies

### Prerequisites (Complete)
- ✅ P0-P8: Foundation (Cell, Sheet, Workbook, CellCodec operational)

### Enables
- Advanced features: Cross-sheet formulas, dynamic calculations
- User-defined functions (extensibility)

### File Conflicts
- None — new xl-evaluator module, isolated from existing code

### Safe Parallelization
- ✅ WI-10 (Table Support) — different module (xl-ooxml)
- ✅ WI-11 (Chart Model) — different module (xl-ooxml)
- ✅ WI-15 (Benchmark Suite) — different module (xl-testkit)

---

## Worktree Strategy

**Branch naming**: `formula-system` or `WI-07-formula-parser`

**Merge order**:
1. WI-07 (Parser) first — foundation for evaluator
2. WI-08 (Evaluator) second — depends on AST
3. WI-09 (Functions) + WI-09b (Graph) can parallel — both use evaluator

**Conflict resolution**: None expected — new module

---

## Execution Algorithm

### Phase 1: Formula Parser (WI-07)
```
1. Create worktree: `gtr create WI-07-formula-parser`
2. Create xl-evaluator module:
   - Update build.mill with new module
   - Create xl-evaluator/src/com/tjclp/xl/formula/
3. Define TExpr GADT (see Design section below)
4. Implement parser: String → Either[ParseError, TExpr[?]]
5. Add property tests for round-trip (print ∘ parse = id)
6. Run tests: `./mill xl-evaluator.test`
7. Create PR: "feat(evaluator): add formula parser and GADT"
8. Merge to main
9. Update roadmap: WI-07 → ✅ Complete, WI-08 → 🔵 Available
```

### Phase 2: Formula Evaluator (WI-08)
```
1. Create worktree: `gtr create WI-08-evaluator` OR continue in WI-07 branch
2. Fix compile issues in Evaluator.scala:
   a. Add import: `import com.tjclp.xl.syntax.*` for Sheet.get extension
   b. Fix variable shadowing in evalFoldRange: rename `cell` parameter to `currentCell`
   c. Verify Sheet.get returns Option[Cell] (handle None case)
   d. Ensure all TExpr cases (17 total) have eval implementations
3. Implement evaluation for remaining cases:
   - All arithmetic: Add, Sub, Mul, Div (div-by-zero detection)
   - All comparison: Lt, Lte, Gt, Gte, Eq, Neq
   - All logical: And (short-circuit), Or (short-circuit), Not
   - Conditionals: If (evaluate condition, branch based on result)
   - Ranges: FoldRange (iterate cells, aggregate with step function)
4. Add comprehensive evaluator tests (~50 tests):
   a. Property tests for laws:
      - Literal identity: eval(Lit(x), _) == Right(x)
      - Arithmetic: eval(Add(Lit(a), Lit(b)), _) == Right(a + b)
      - Short-circuit: And(Lit(false), _) doesn't eval second arg
   b. Unit tests for edge cases:
      - Division by zero → EvalError.DivByZero
      - Missing cell reference → EvalError.RefError
      - Codec failure (text where number expected) → EvalError.CodecFailed
      - Empty range → aggregation returns zero value
   c. Integration tests:
      - Complex nested expressions: IF(AND(A1>0, B1<100), SUM(C1:C10), 0)
      - Real Excel formulas from test fixtures
5. Add Sheet extension methods (optional, for ergonomics):
   - sheet.evalFormula(formula: String): Either[XLError, CellValue]
   - sheet.evalCell(ref: ARef): Either[XLError, CellValue]
6. Run tests: `./mill xl-evaluator.test` (target: 101+ tests total = 51 parser + 50 evaluator)
7. Update documentation:
   - STATUS.md: Formula evaluation complete
   - formula-system.md: WI-08 → ✅ Complete
   - examples.md: Add evaluator examples
8. Create PR: "feat(evaluator): implement pure formula evaluator"
9. Merge to main
10. Update roadmap: WI-08 → ✅ Complete, WI-09/WI-09b → 🔵 Available
```

### Phase 3: Function Library (WI-09)
```
1. Create worktree: `gtr create WI-09-functions`
2. Implement built-in functions:
   - Arithmetic: SUM, AVERAGE, MIN, MAX, COUNT
   - Logic: IF, AND, OR, NOT
   - Text: CONCATENATE, LEFT, RIGHT, LEN
   - Date: TODAY, NOW, DATE, YEAR, MONTH, DAY
3. Add function registry (extensibility)
4. Add tests for each function (400+ tests target)
5. Run tests: `./mill xl-evaluator.test`
6. Create PR: "feat(evaluator): add built-in function library"
```

### Phase 4: Dependency Graph (WI-09b)
```
1. Create worktree: `gtr create WI-09b-dependency-graph`
2. Implement dependency graph builder
3. Add cycle detection (DirectedGraph with SCC algorithm)
4. Add topological sort for evaluation order
5. Add tests for circular reference detection
6. Run tests: `./mill xl-evaluator.test`
7. Create PR: "feat(evaluator): add dependency graph and cycle detection"
```

---

## Lessons Learned from WI-07 (Parser Implementation)

**Completion Date**: 2025-11-21
**Effort**: ~1 week (5-7 days actual)
**Deliverables**: TExpr GADT (16 constructors), FormulaParser (620 LOC), FormulaPrinter (250 LOC), ParseError ADT, 51 passing tests

### Key Implementation Patterns

#### 1. Opaque Type Handling

**Challenge**: ARef is opaque type wrapping Long (64-bit packing: `(row << 32) | col`)

**Solution Patterns**:
```scala
// Extract components from ARef
val arefLong: Long = aref.asInstanceOf[Long]  // Safe: ARef is opaque type = Long
val colIndex = (arefLong & 0xffffffffL).toInt
val rowIndex = (arefLong >> 32).toInt

// Format to A1 notation
val col = Column.from0(colIndex)
val colLetter = Column.toLetter(col)
val rowNum = rowIndex + 1
s"$colLetter$rowNum"
```

**Lesson**: Extension methods like `.toA1` use inline expansion with variable name `ref`. When using pattern matching or parameters, variable shadowing causes compile errors. **Recommendation**: Add `ARef.toA1Helper(ref: ARef): String` utility in xl-core to avoid code duplication (EvalError, FormulaPrinter, Evaluator all need this).

#### 2. Extension Methods & Imports

**Challenge**: Sheet.get, CellRange.cells are extension methods requiring specific imports.

**Solution**:
```scala
// Always import at top of formula files:
import com.tjclp.xl.*            // Gets core types + unified exports
import com.tjclp.xl.syntax.*     // Gets Sheet.get, other extensions
```

**Key APIs**:
- `Sheet.get(ref: ARef): Option[Cell]` — Returns None if cell doesn't exist
- `CellRange.cells: Iterator[ARef]` — Iterates all refs in range (row-major order)
- `ARef.toA1: String` — Formats as A1 notation (extension method, inlined)

#### 3. Testing Strategy (51 Tests Achieved)

**Breakdown**:
- 7 property-based tests (round-trip laws using ScalaCheck)
- 26 parser unit tests (literals, operators, functions, parentheses, whitespace)
- 10 scientific notation tests (1E10, 3E-5, edge cases)
- 5 error handling tests (empty formula, unbalanced parens, unknown functions)
- 3 integration tests (complex expressions, nested IFs, operator precedence)

**Pattern**: Generators at top of test file for reusability:
```scala
val genNumericLit: Gen[TExpr[BigDecimal]] = Gen.choose(-1000.0, 1000.0).map(TExpr.Lit.apply)
val genARef: Gen[ARef] = for
  col <- Gen.choose(0, 100)
  row <- Gen.choose(0, 100)
yield ARef.from0(col, row)
```

**Lesson**: Separate property tests (laws) from unit tests (specific behaviors) from error tests for clarity.

#### 4. Scientific Notation Support

**Challenge**: BigDecimal.toString() emits scientific notation for very small/large numbers (|x| < 1E-7 or |x| > 1E7).

**Solution**: Parser must handle E notation:
```scala
// Added to parseNumberLiteral:
case Some('E' | 'e') if !hasExponent =>
  val s2 = s.advance()
  s2.currentChar match
    case Some('+' | '-') => // Optional sign
      loop(s2.advance(), hasDecimal, hasExponent = true)
    case Some(c) if c.isDigit =>
      loop(s2, hasDecimal, hasExponent = true)
```

**Edge cases tested**: `1E10`, `1.5E-5`, `3.14e2`, `-5.2E-8`, invalid: `1E` (no digits), `1E2E3` (multiple E)

#### 5. Round-Trip Property Test

**Challenge**: Negative numbers print with leading `-` (e.g., `=-5`), which parser interprets as unary minus: `Sub(Lit(0), Lit(5))`, not `Lit(-5)`.

**Solution**: Property test must handle both:
```scala
parsed.exists {
  case TExpr.Lit(value: BigDecimal) =>
    (value - original).abs < BigDecimal("1E-15")
  case TExpr.Sub(TExpr.Lit(zero), TExpr.Lit(value)) if zero == BigDecimal(0) =>
    ((-value) - original).abs < BigDecimal("1E-15")  // Unary minus case
  case _ => false
}
```

**Lesson**: Parser design decision (unary minus as `Sub(0, x)`) affects test design. Document in ADR.

#### 6. WartRemover Compliance

**Warnings encountered** (37 total, all acceptable):
- `asInstanceOf` (Tier 2): Required for runtime parsing where GADT type info is lost
- `return` (Tier 2): Used in early returns for readability (could refactor to nested match)
- Fixed `head` (Tier 1): Changed `args.head` → `case head :: _` pattern

**Lesson**: Runtime parsing loses compile-time type info, requiring `asInstanceOf` for Eq/Neq/If branches. This is acceptable per WartRemover policy for parsers.

### Expected WI-08 Challenges

#### 1. Cell Reference Resolution

**Pattern** (from Evaluator.scala skeleton):
```scala
case TExpr.Ref(at, decode) =>
  sheet.get(at) match
    case Some(cell) =>
      decode(cell).left.map(codecErr => EvalError.CodecFailed(at, codecErr))
    case None =>
      Left(EvalError.RefError(at, "cell not found"))
```

**Key points**:
- Sheet.get returns `Option[Cell]`, not `Cell`
- Handle None → RefError
- Codec errors → CodecFailed (wrap, don't throw)

#### 2. Short-Circuit Evaluation

**Must implement correctly**:
```scala
case TExpr.And(x, y) =>
  eval(x, sheet) match
    case Right(false) => Right(false)  // Don't evaluate y
    case Right(true) => eval(y, sheet)
    case Left(err) => Left(err)
```

**Test verification**:
```scala
test("And(Lit(false), Ref(missing)) doesn't error") {
  // Even though missing cell would error, And short-circuits
  val expr = TExpr.And(TExpr.Lit(false), TExpr.Ref(ref"Z999", decode))
  evaluator.eval(expr, sheet) == Right(false)  // ✓ No error
}
```

#### 3. Range Aggregation

**Current issue**: Variable shadowing in fold:
```scala
// BROKEN:
cells.foldLeft[Either[EvalError, B]](Right(zero)) { case (accEither, cell) =>
  decode(cell) match  // 'cell' from fold parameter
    case Left(codecErr) =>
      Left(EvalError.CodecFailed(cell.ref, codecErr))  // Shadows outer 'cell'
}

// FIXED:
cells.foldLeft[Either[EvalError, B]](Right(zero)) { case (accEither, currentCell) =>
  decode(currentCell) match
    case Left(codecErr) =>
      Left(EvalError.CodecFailed(currentCell.ref, codecErr))
}
```

#### 4. Division by Zero

**Pattern**:
```scala
case TExpr.Div(x, y) =>
  for
    xv <- eval(x, sheet)
    yv <- eval(y, sheet)
    result <-
      if yv == BigDecimal(0) then
        Left(EvalError.DivByZero(
          FormulaPrinter.print(x, includeEquals = false),
          FormulaPrinter.print(y, includeEquals = false)
        ))
      else Right(xv / yv)
  yield result
```

**Test cases**:
- Direct: `Div(Lit(10), Lit(0))`
- Nested: `Div(Lit(10), Sub(Lit(5), Lit(5)))`  — Zero from evaluation, not literal
- Range: `Div(Lit(10), count(empty_range))` — Zero from COUNT

### Design Decisions Made in WI-07

#### Decision 1: BigDecimal for All Numeric Operations

**Rationale**: Financial precision, matches Excel decimal semantics

**Impact on WI-08**: Evaluator must use BigDecimal arithmetic (not Double). Comparison operators work naturally.

#### Decision 2: Comparison Operators in GADT

**Decision**: Added Lt, Lte, Gt, Gte, Eq, Neq as first-class TExpr constructors (not functions)

**Alternative considered**: `Compare(op: CompareOp, x, y)` with enum `CompareOp { Lt, Lte, Gt, Gte, Eq, Neq }`

**Rationale**: Direct constructors provide better type safety and clearer AST representation

**Impact on WI-08**: Each comparison operator needs separate eval case (6 cases). Acceptable trade-off.

#### Decision 3: FoldRange as Universal Aggregation

**Decision**: Single `FoldRange[A, B]` constructor for all aggregations (SUM, COUNT, AVERAGE, MIN, MAX)

**Rationale**:
- Functional abstraction (DRY)
- Extensible (any aggregation via step function)
- Type-safe (decode function ensures type A)

**Impact on WI-08**:
- Single eval case handles all aggregations ✅
- Cell iteration with decode function
- Empty ranges → returns zero value (not error)

---

## Design

### Two-Level Formula System

We provide two levels:
1. **`TExpr[A]`**: a typed, total expression GADT for programmatic formulas with pure evaluation
2. **`formula.Expr[V]`**: an interop AST for Excel string formulas (`V` captures validation state)

### GADT (selected constructors)

```scala
package com.tjclp.xl.formula

import com.tjclp.xl.core.*, com.tjclp.xl.core.addr.*

enum TExpr[A]:
  case Lit[A](value: A)
  case Ref[A](at: ARef, decode: Cell => Either[codec.ReadError, A])
  case If[A](cond: TExpr[Boolean], ifTrue: TExpr[A], ifFalse: TExpr[A])
  case Add(x: TExpr[BigDecimal], y: TExpr[BigDecimal]) extends TExpr[BigDecimal]
  case Sub(x: TExpr[BigDecimal], y: TExpr[BigDecimal]) extends TExpr[BigDecimal]
  case Mul(x: TExpr[BigDecimal], y: TExpr[BigDecimal]) extends TExpr[BigDecimal]
  case Div(x: TExpr[BigDecimal], y: TExpr[BigDecimal]) extends TExpr[BigDecimal]
  case And(x: TExpr[Boolean], y: TExpr[Boolean])        extends TExpr[Boolean]
  case Or (x: TExpr[Boolean], y: TExpr[Boolean])        extends TExpr[Boolean]
  case Not(x: TExpr[Boolean])                           extends TExpr[Boolean]
  case FoldRange[A,B](range: addr.CellRange, z: B, step: (B, A) => B, decode: Cell => Either[codec.ReadError, A]) extends TExpr[B]
```

### Laws (sketch)

- **If-fusion:** `If(c, Lit(x), Lit(y)) ≡ Lit(if ⟦c⟧ then x else y)`
- **Ring laws:** `Add/Mul` form a commutative semiring over `BigDecimal` nodes modulo printer parentheses
- **Short-circuit:** `And/Or` evaluator respects left-to-right semantics

### Printer (Excel string)

- We provide a **total printer** from `TExpr[A]` to Excel text with correct precedence & parentheses
- Where Excel lacks a direct function, we use idioms (e.g., `SUM` + `COUNT` for average)

### Evaluator

```scala
trait Eval:
  def eval[A](e: TExpr[A], s: Sheet): Either[EvalError, A]
```

- Pure; cycle detection happens on the dependency graph before calling `eval`
- Deterministic; no external state

---

## Definition of Done

### Functional
- [ ] Parser converts Excel formula strings to TExpr AST
- [ ] Evaluator computes values from TExpr
- [ ] 50+ built-in functions (SUM, IF, VLOOKUP, etc.)
- [ ] Dependency graph detects circular references
- [ ] Integration with Sheet (formula cells auto-evaluate)

### Code Quality
- [ ] Zero WartRemover errors
- [ ] Property tests for all laws (If-fusion, ring laws, etc.)
- [ ] 400+ tests (parser + evaluator + functions)
- [ ] Scalafmt applied

### Documentation
- [ ] CLAUDE.md updated with formula patterns
- [ ] STATUS.md updated with evaluator capability
- [ ] Examples in docs/reference/examples.md

### Integration
- [ ] Roadmap.md: WI-07, WI-08, WI-09, WI-09b marked Complete
- [ ] No conflicts with parallel work

---

## Module Ownership

**Primary**: `xl-evaluator` (new module — ~3000 LOC)

**Secondary**: `xl-core` (minimal — TExpr types only)

**Test Files**: `xl-evaluator/test` (~2000 LOC tests)

---

## Merge Risk Assessment

**Risk Level**: Low

**Rationale**:
- New module (xl-evaluator), isolated from existing code
- Core changes minimal (just type definitions)
- No changes to xl-ooxml or xl-cats-effect
- Integration is additive (opt-in evaluation)

---

## Related Documentation

- **Roadmap**: `docs/plan/roadmap.md` (WI-07 through WI-09b)
- **Design**: `docs/design/domain-model.md` (Cell, CellValue.Formula)
- **Implementation**: `docs/reference/implementation-scaffolds.md` (formula examples)
- **Strategic**: `docs/plan/strategic-implementation-plan.md` (Phase 4: Formula/Graph)

---

## Notes

- **Complexity**: High (400+ Excel functions to implement)
- **Timeline**: 6-8 weeks for full implementation
- **Alternative**: Let Excel recalculate formulas (current approach — adequate for most use cases)
- **User value**: High for computational spreadsheets, low for data export/import use cases
