# Typed Formula System

**Status**: ⬜ Not Started
**Priority**: Medium
**Estimated Effort**: 6-8 weeks
**Last Updated**: 2025-11-20

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
| `WI-07` | Formula Parser (string → AST) | Core | `FormulaParser.scala`, `TExpr.scala` | ⏳ Not Started | - |
| `WI-08` | Formula Evaluator (AST → value) | Core | `Evaluator.scala`, `EvalError.scala` | ⏳ Not Started | - |
| `WI-09` | Function Library (SUM, IF, etc.) | Core | `Functions.scala` | ⏳ Not Started | - |
| `WI-09b` | Dependency Graph & Cycle Detection | Core | `DependencyGraph.scala` | ⏳ Not Started | - |

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
1. Create worktree: `gtr create WI-08-evaluator`
2. Implement Eval trait (see Design section below)
3. Add evaluation for all TExpr cases
4. Add EvalError ADT (DivByZero, TypeError, RefError, etc.)
5. Add property tests for evaluation laws
6. Run tests: `./mill xl-evaluator.test`
7. Create PR: "feat(evaluator): add pure formula evaluator"
8. Merge to main
9. Update roadmap: WI-08 → ✅ Complete, WI-09/WI-09b → 🔵 Available
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
