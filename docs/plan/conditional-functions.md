# WI-09f: Conditional Aggregation Functions (SUMIF/COUNTIF)

**Status**: 🔵 Ready to Start
**Priority**: High (80% of reporting use cases)
**Estimated Effort**: 3-4 days
**Last Updated**: 2025-11-25

---

## Metadata

| Field | Value |
|-------|-------|
| **Owner Module** | `xl-evaluator` |
| **Touches Files** | `TExpr.scala`, `FunctionParser.scala`, `Evaluator.scala`, `FormulaPrinter.scala`, `DependencyGraph.scala` |
| **Dependencies** | WI-09a/b/c/d complete (formula system foundation) |
| **Enables** | Financial reporting, data analysis, conditional summaries |
| **Parallelizable With** | WI-11 (Charts), WI-19 (Row/Col Props), WI-18 (Merged Cells) — different modules |
| **Merge Risk** | Low (additive changes within xl-evaluator module) |

---

## Work Items

| ID | Description | Type | Files | Status | PR |
|----|-------------|------|-------|--------|----|
| `WI-09f-1` | Create CriteriaMatcher helper | Feature | `CriteriaMatcher.scala` (new) | ⏳ Not Started | - |
| `WI-09f-2` | Add TExpr cases + smart constructors | Feature | `TExpr.scala` | ⏳ Not Started | - |
| `WI-09f-3` | Add function parsers (variable arity) | Feature | `FunctionParser.scala` | ⏳ Not Started | - |
| `WI-09f-4` | Implement evaluator logic | Feature | `Evaluator.scala` | ⏳ Not Started | - |
| `WI-09f-5` | Add FormulaPrinter cases | Feature | `FormulaPrinter.scala` | ⏳ Not Started | - |
| `WI-09f-6` | Update DependencyGraph | Feature | `DependencyGraph.scala` | ⏳ Not Started | - |
| `WI-09f-7` | Comprehensive test suite | Tests | `ConditionalFunctionsSpec.scala` (new) | ⏳ Not Started | - |

---

## Functions to Implement

| Function | Signature | Description |
|----------|-----------|-------------|
| **SUMIF** | `SUMIF(range, criteria, [sum_range])` | Sum cells matching criteria |
| **COUNTIF** | `COUNTIF(range, criteria)` | Count cells matching criteria |
| **SUMIFS** | `SUMIFS(sum_range, criteria_range1, criteria1, ...)` | Sum with multiple criteria (AND logic) |
| **COUNTIFS** | `COUNTIFS(criteria_range1, criteria1, ...)` | Count with multiple criteria (AND logic) |

---

## Criteria Matching (Excel-Compatible)

Excel criteria support 5 patterns:

| Pattern | Example | Matches |
|---------|---------|---------|
| **Exact** | `"Apple"` | Cells equal to "Apple" (case-insensitive) |
| **Wildcard** | `"A*"`, `"*pple"`, `"A?ple"` | Pattern matching (* = any chars, ? = single char) |
| **Greater** | `">100"` | Numbers > 100 |
| **Less** | `"<=50"` | Numbers <= 50 |
| **Not Equal** | `"<>0"` | Numbers != 0 |

**Escape sequences**: `~*` for literal asterisk, `~?` for literal question mark.

---

## Dependencies

### Prerequisites (Complete)
- ✅ WI-07: Formula Parser (FormulaParser, FunctionParser pattern)
- ✅ WI-08: Formula Evaluator (Evaluator, TExpr GADT)
- ✅ WI-09a/b/c/d: Core functions (SUM, COUNT, AVERAGE, etc.)
- ✅ WI-09e: Financial functions (demonstrates range iteration pattern)

### Enables
- Complex financial reporting (budget vs actual)
- Data analysis with filters
- Conditional KPI calculations
- Foundation for AVERAGEIF, MAXIF, MINIF (future)

### File Conflicts
- **Low risk**: `TExpr.scala` (additive — new enum cases)
- **Low risk**: `FunctionParser.scala` (additive — new given instances)
- **Low risk**: `Evaluator.scala` (additive — new match cases)
- **None**: `CriteriaMatcher.scala` is new file

### Safe Parallelization
- ✅ WI-11 (Charts) — different module
- ✅ WI-19 (Row/Col Props) — different module
- ✅ WI-18 (Merged Cells) — different module
- ✅ WI-30 (Security) — different concern

---

## Worktree Strategy

**Branch naming**: `feat/conditional-functions` or `WI-09f-sumif-countif`

**Merge order**:
1. WI-09f-1 (CriteriaMatcher) — foundation for all functions
2. WI-09f-2 (TExpr cases) — depends on CriteriaMatcher design
3. WI-09f-3 (FunctionParser) — needs TExpr cases
4. WI-09f-4 (Evaluator) — needs TExpr + CriteriaMatcher
5. WI-09f-5 (FormulaPrinter) — needs TExpr cases
6. WI-09f-6 (DependencyGraph) — needs TExpr cases
7. WI-09f-7 (Tests) — continuous throughout

**Conflict resolution**: Minimal — mostly new code and additive changes

---

## Execution Algorithm

### Phase 1: CriteriaMatcher (WI-09f-1) — 3-4 hours

```
1. Create worktree: `gtr create feat/conditional-functions`
2. Create xl-evaluator/src/com/tjclp/xl/formula/CriteriaMatcher.scala:
   - Define Criterion sealed trait (Exact, Compare, Wildcard)
   - Define CompareOp enum (Gt, Gte, Lt, Lte, Neq)
   - Implement parse(value: Any): Criterion
   - Implement matches(cellValue: CellValue, criterion: Criterion): Boolean
3. Handle all 5 criteria patterns:
   - Exact: string/number/boolean equality (case-insensitive for strings)
   - Compare: parse ">100", ">=50", "<10", "<=5", "<>0"
   - Wildcard: convert * and ? to regex, handle ~* and ~? escapes
4. Handle type coercion:
   - Text "42" matches Number 42
   - Boolean TRUE matches text "TRUE"
   - Empty cells don't match (except exact match on "")
5. Create xl-evaluator/test/src/.../CriteriaMatcherSpec.scala:
   - Test all 5 pattern types
   - Test type coercion scenarios
   - Test wildcard escapes (~* and ~?)
   - Test case insensitivity
6. Run tests: `./mill xl-evaluator.test.testOnly *.CriteriaMatcherSpec`
7. Commit: "feat(evaluator): add CriteriaMatcher for Excel criteria parsing"
```

### Phase 2: TExpr Cases (WI-09f-2) — 2-3 hours

```
1. Edit xl-evaluator/src/com/tjclp/xl/formula/TExpr.scala:
2. Add 4 new enum cases after existing range functions:

   case SumIf(
     range: CellRange,
     criteria: TExpr[?],           // Evaluated at runtime to get criterion
     sumRange: Option[CellRange]   // If None, sum the range itself
   ) extends TExpr[BigDecimal]

   case CountIf(
     range: CellRange,
     criteria: TExpr[?]
   ) extends TExpr[BigDecimal]    // Returns count as BigDecimal (Excel convention)

   case SumIfs(
     sumRange: CellRange,
     conditions: List[(CellRange, TExpr[?])]  // (criteria_range, criteria) pairs
   ) extends TExpr[BigDecimal]

   case CountIfs(
     conditions: List[(CellRange, TExpr[?])]
   ) extends TExpr[BigDecimal]

3. Add smart constructors in TExpr object:
   def sumIf(range: CellRange, criteria: TExpr[?], sumRange: Option[CellRange] = None)
   def countIf(range: CellRange, criteria: TExpr[?])
   def sumIfs(sumRange: CellRange, conditions: List[(CellRange, TExpr[?])])
   def countIfs(conditions: List[(CellRange, TExpr[?])])

4. Run tests: `./mill xl-evaluator.test`
5. Commit: "feat(evaluator): add TExpr cases for SUMIF/COUNTIF/SUMIFS/COUNTIFS"
```

### Phase 3: FunctionParser (WI-09f-3) — 3-4 hours

```
1. Edit xl-evaluator/src/com/tjclp/xl/formula/FunctionParser.scala:
2. Add parser for SUMIF (2-3 arguments):
   - Arg 1: range (CellRange)
   - Arg 2: criteria (any TExpr)
   - Arg 3: optional sum_range (CellRange)

   given sumIfParser: FunctionParser with
     def name = "SUMIF"
     def arity = ArityRange(2, 3)
     def parse(args: List[TExpr[?]]): Either[ParseError, TExpr[?]] = args match
       case List(rangeExpr, criteriaExpr) =>
         extractRange(rangeExpr).map(r => TExpr.SumIf(r, criteriaExpr, None))
       case List(rangeExpr, criteriaExpr, sumRangeExpr) =>
         for
           range <- extractRange(rangeExpr)
           sumRange <- extractRange(sumRangeExpr)
         yield TExpr.SumIf(range, criteriaExpr, Some(sumRange))
       case _ => Left(ParseError.InvalidArity("SUMIF", 2, 3, args.length))

3. Add parser for COUNTIF (exactly 2 arguments):
   given countIfParser: FunctionParser with
     def name = "COUNTIF"
     def arity = ArityRange(2, 2)
     def parse(args: List[TExpr[?]]): Either[ParseError, TExpr[?]] = ...

4. Add parser for SUMIFS (3+ arguments, odd count):
   - Arg 1: sum_range
   - Args 2,3: criteria_range1, criteria1
   - Args 4,5: criteria_range2, criteria2
   - ...etc

   given sumIfsParser: FunctionParser with
     def name = "SUMIFS"
     def arity = ArityRange(3, Int.MaxValue)
     def parse(args: List[TExpr[?]]): Either[ParseError, TExpr[?]] =
       if args.length < 3 || args.length % 2 == 0 then
         Left(ParseError.InvalidArity("SUMIFS", "3, 5, 7, ...", args.length))
       else
         val sumRangeExpr = args.head
         val pairs = args.tail.grouped(2).toList
         // Build conditions list from pairs
         ...

5. Add parser for COUNTIFS (2+ arguments, even count):
   given countIfsParser: FunctionParser with
     def name = "COUNTIFS"
     def arity = ArityRange(2, Int.MaxValue)
     def parse(args: List[TExpr[?]]): Either[ParseError, TExpr[?]] =
       if args.length < 2 || args.length % 2 != 0 then
         Left(ParseError.InvalidArity("COUNTIFS", "2, 4, 6, ...", args.length))
       else ...

6. Add helper for range extraction:
   private def extractRange(expr: TExpr[?]): Either[ParseError, CellRange] = ...

7. Run tests: `./mill xl-evaluator.test`
8. Commit: "feat(evaluator): add FunctionParser for SUMIF/COUNTIF/SUMIFS/COUNTIFS"
```

### Phase 4: Evaluator (WI-09f-4) — 3-4 hours

```
1. Edit xl-evaluator/src/com/tjclp/xl/formula/Evaluator.scala:
2. Add case for SumIf:

   case TExpr.SumIf(range, criteriaExpr, sumRangeOpt) =>
     for
       criteriaValue <- eval(criteriaExpr, sheet, clock)
       criterion = CriteriaMatcher.parse(criteriaValue)
       effectiveRange = sumRangeOpt.getOrElse(range)
       _ <- validateSameDimensions(range, effectiveRange, "SUMIF")
       pairs = range.cells.zip(effectiveRange.cells)
       sum = pairs.foldLeft(BigDecimal(0)) { case (acc, (testRef, sumRef)) =>
         val testCell = sheet(testRef)
         if CriteriaMatcher.matches(testCell.value, criterion) then
           sheet(sumRef).value match
             case CellValue.Number(n) => acc + n
             case _ => acc  // Skip non-numeric (Excel behavior)
         else acc
       }
     yield sum

3. Add case for CountIf:

   case TExpr.CountIf(range, criteriaExpr) =>
     for
       criteriaValue <- eval(criteriaExpr, sheet, clock)
       criterion = CriteriaMatcher.parse(criteriaValue)
       count = range.cells.count { ref =>
         CriteriaMatcher.matches(sheet(ref).value, criterion)
       }
     yield BigDecimal(count)

4. Add case for SumIfs (multiple criteria AND logic):

   case TExpr.SumIfs(sumRange, conditions) =>
     // Evaluate all criteria first
     val criteriaResults = conditions.map { case (range, criteriaExpr) =>
       eval(criteriaExpr, sheet, clock).map(v => (range, CriteriaMatcher.parse(v)))
     }
     // ... validate dimensions, compute sum where ALL criteria match

5. Add case for CountIfs (multiple criteria AND logic):
   Similar to SumIfs but counting instead of summing

6. Add helper for dimension validation:
   private def validateSameDimensions(
     range1: CellRange,
     range2: CellRange,
     funcName: String
   ): Either[EvalError, Unit] = ...

7. Run tests: `./mill xl-evaluator.test`
8. Commit: "feat(evaluator): implement SUMIF/COUNTIF/SUMIFS/COUNTIFS evaluation"
```

### Phase 5: FormulaPrinter + DependencyGraph (WI-09f-5, WI-09f-6) — 2 hours

```
1. Edit xl-evaluator/src/com/tjclp/xl/formula/FormulaPrinter.scala:
2. Add print cases for all 4 functions:

   case TExpr.SumIf(range, criteria, sumRangeOpt) =>
     sumRangeOpt match
       case Some(sumRange) => s"SUMIF(${range.toA1}, ${print(criteria)}, ${sumRange.toA1})"
       case None => s"SUMIF(${range.toA1}, ${print(criteria)})"

   case TExpr.CountIf(range, criteria) =>
     s"COUNTIF(${range.toA1}, ${print(criteria)})"

   case TExpr.SumIfs(sumRange, conditions) =>
     val condStrs = conditions.map { case (r, c) => s"${r.toA1}, ${print(c)}" }.mkString(", ")
     s"SUMIFS(${sumRange.toA1}, $condStrs)"

   case TExpr.CountIfs(conditions) =>
     val condStrs = conditions.map { case (r, c) => s"${r.toA1}, ${print(c)}" }.mkString(", ")
     s"COUNTIFS($condStrs)"

3. Edit xl-evaluator/src/com/tjclp/xl/formula/DependencyGraph.scala:
4. Update extractDependencies for new cases:

   case TExpr.SumIf(range, _, sumRangeOpt) =>
     range.cells.toSet ++ sumRangeOpt.map(_.cells.toSet).getOrElse(Set.empty)
   case TExpr.CountIf(range, _) =>
     range.cells.toSet
   case TExpr.SumIfs(sumRange, conditions) =>
     sumRange.cells.toSet ++ conditions.flatMap(_._1.cells).toSet
   case TExpr.CountIfs(conditions) =>
     conditions.flatMap(_._1.cells).toSet

5. Run tests: `./mill xl-evaluator.test`
6. Commit: "feat(evaluator): add FormulaPrinter and DependencyGraph for conditional functions"
```

### Phase 6: Comprehensive Tests (WI-09f-7) — 4-5 hours

```
1. Create xl-evaluator/test/src/com/tjclp/xl/formula/ConditionalFunctionsSpec.scala:
2. Write 40+ tests covering:

   // === SUMIF Tests ===
   test("SUMIF: sum values where criteria matches") { ... }
   test("SUMIF: sum with separate sum_range") { ... }
   test("SUMIF: sum range same as criteria range (2-arg form)") { ... }
   test("SUMIF: no matches returns 0") { ... }
   test("SUMIF: case-insensitive text matching") { ... }
   test("SUMIF: wildcard * matches any characters") { ... }
   test("SUMIF: wildcard ? matches single character") { ... }
   test("SUMIF: escaped wildcards ~* and ~?") { ... }
   test("SUMIF: numeric comparison >") { ... }
   test("SUMIF: numeric comparison >=") { ... }
   test("SUMIF: numeric comparison <") { ... }
   test("SUMIF: numeric comparison <=") { ... }
   test("SUMIF: numeric comparison <>") { ... }
   test("SUMIF: skip non-numeric in sum_range") { ... }

   // === COUNTIF Tests ===
   test("COUNTIF: count matching text") { ... }
   test("COUNTIF: count with wildcard") { ... }
   test("COUNTIF: count with numeric comparison") { ... }
   test("COUNTIF: count numbers matching exact value") { ... }
   test("COUNTIF: no matches returns 0") { ... }

   // === SUMIFS Tests ===
   test("SUMIFS: multiple criteria AND logic") { ... }
   test("SUMIFS: three criteria") { ... }
   test("SUMIFS: mixed text and numeric criteria") { ... }
   test("SUMIFS: all criteria must match for row to count") { ... }

   // === COUNTIFS Tests ===
   test("COUNTIFS: two criteria") { ... }
   test("COUNTIFS: three criteria") { ... }
   test("COUNTIFS: wildcards in multiple criteria") { ... }

   // === Round-trip Tests ===
   test("SUMIF: parse -> print -> parse roundtrip") { ... }
   test("COUNTIF: parse -> print -> parse roundtrip") { ... }
   test("SUMIFS: parse -> print -> parse roundtrip") { ... }
   test("COUNTIFS: parse -> print -> parse roundtrip") { ... }

   // === Error Cases ===
   test("SUMIF: mismatched range dimensions") { ... }
   test("SUMIFS: mismatched range dimensions") { ... }
   test("COUNTIF: invalid arity") { ... }
   test("SUMIFS: even number of args (invalid)") { ... }

   // === Edge Cases ===
   test("SUMIF: empty range") { ... }
   test("SUMIF: all cells empty") { ... }
   test("SUMIF: formula cells with cached values") { ... }
   test("COUNTIF: boolean criteria") { ... }
   test("SUMIF: date comparison") { ... }

3. Run tests: `./mill xl-evaluator.test.testOnly *.ConditionalFunctionsSpec`
4. Run full suite: `./mill __.test`
5. Commit: "test(evaluator): comprehensive tests for SUMIF/COUNTIF/SUMIFS/COUNTIFS"
```

### Phase 7: Documentation Update — 30 minutes

```
1. Update CLAUDE.md:
   - Change function count: "24 functions" → "28 functions"
   - Add SUMIF, COUNTIF, SUMIFS, COUNTIFS to function list

2. Update docs/STATUS.md:
   - Add conditional functions to feature list

3. Update docs/plan/roadmap.md:
   - Mark WI-09f as complete
   - Update "Next Available Work" section

4. Run format check: `./mill __.checkFormat`
5. Commit: "docs: update function count and roadmap for WI-09f"
6. Create PR: "feat(evaluator): SUMIF/COUNTIF/SUMIFS/COUNTIFS conditional aggregation"
```

---

## Design

### CriteriaMatcher Architecture

```scala
// xl-evaluator/src/com/tjclp/xl/formula/CriteriaMatcher.scala
package com.tjclp.xl.formula

import com.tjclp.xl.cells.CellValue
import scala.util.matching.Regex

object CriteriaMatcher:

  /** Parsed criterion for cell matching */
  sealed trait Criterion derives CanEqual
  case class Exact(value: Any) extends Criterion
  case class Compare(op: CompareOp, value: BigDecimal) extends Criterion
  case class Wildcard(pattern: String) extends Criterion

  /** Comparison operators for numeric criteria */
  enum CompareOp derives CanEqual:
    case Gt   // >
    case Gte  // >=
    case Lt   // <
    case Lte  // <=
    case Neq  // <>

  /**
   * Parse a criterion from an evaluated expression result.
   *
   * Handles:
   *   - Strings with operators: ">100", ">=50", "<10", "<=5", "<>0"
   *   - Strings with wildcards: "A*", "*pple", "A?ple"
   *   - Plain strings: exact match
   *   - Numbers/Booleans: exact match
   */
  def parse(value: Any): Criterion = value match
    case s: String => parseString(s)
    case n: BigDecimal => Exact(n)
    case b: Boolean => Exact(b)
    case _ => Exact(value)

  private def parseString(s: String): Criterion =
    s match
      case _ if s.startsWith("<>") => parseCompare(s.drop(2), CompareOp.Neq)
      case _ if s.startsWith(">=") => parseCompare(s.drop(2), CompareOp.Gte)
      case _ if s.startsWith("<=") => parseCompare(s.drop(2), CompareOp.Lte)
      case _ if s.startsWith(">")  => parseCompare(s.drop(1), CompareOp.Gt)
      case _ if s.startsWith("<")  => parseCompare(s.drop(1), CompareOp.Lt)
      case _ if s.contains("*") || s.contains("?") => Wildcard(s)
      case _ => Exact(s)

  private def parseCompare(numStr: String, op: CompareOp): Criterion =
    scala.util.Try(BigDecimal(numStr.trim)).toOption match
      case Some(n) => Compare(op, n)
      case None => Exact(s"${opToString(op)}$numStr")  // Couldn't parse, treat as literal

  /**
   * Test if a cell value matches a criterion.
   */
  def matches(cellValue: CellValue, criterion: Criterion): Boolean =
    criterion match
      case Exact(expected) => matchesExact(cellValue, expected)
      case Compare(op, threshold) => matchesCompare(cellValue, op, threshold)
      case Wildcard(pattern) => matchesWildcard(cellValue, pattern)

  private def matchesExact(cellValue: CellValue, expected: Any): Boolean = ...
  private def matchesCompare(cellValue: CellValue, op: CompareOp, threshold: BigDecimal): Boolean = ...
  private def matchesWildcard(cellValue: CellValue, pattern: String): Boolean = ...
```

### TExpr Cases

```scala
// In TExpr.scala enum

/** Sum cells where criteria matches */
case SumIf(
  range: CellRange,
  criteria: TExpr[?],           // Evaluated at runtime
  sumRange: Option[CellRange]   // If None, sum the range itself
) extends TExpr[BigDecimal]

/** Count cells where criteria matches */
case CountIf(
  range: CellRange,
  criteria: TExpr[?]
) extends TExpr[BigDecimal]     // Returns count as BigDecimal (Excel convention)

/** Sum with multiple criteria (AND logic) */
case SumIfs(
  sumRange: CellRange,
  conditions: List[(CellRange, TExpr[?])]  // (criteria_range, criteria) pairs
) extends TExpr[BigDecimal]

/** Count with multiple criteria (AND logic) */
case CountIfs(
  conditions: List[(CellRange, TExpr[?])]
) extends TExpr[BigDecimal]
```

### Evaluator Logic

```scala
// In Evaluator.scala

case TExpr.SumIf(range, criteriaExpr, sumRangeOpt) =>
  for
    criteriaValue <- eval(criteriaExpr, sheet, clock)
    criterion = CriteriaMatcher.parse(criteriaValue)
    effectiveRange = sumRangeOpt.getOrElse(range)
    pairs = range.cells.zip(effectiveRange.cells)
    sum = pairs.foldLeft(BigDecimal(0)) { case (acc, (testRef, sumRef)) =>
      val testCell = sheet(testRef)
      if CriteriaMatcher.matches(testCell.value, criterion) then
        sheet(sumRef).value match
          case CellValue.Number(n) => acc + n
          case _ => acc  // Skip non-numeric (Excel behavior)
      else acc
    }
  yield sum

case TExpr.SumIfs(sumRange, conditions) =>
  // Evaluate all criteria expressions first
  conditions.traverse { case (range, criteriaExpr) =>
    eval(criteriaExpr, sheet, clock).map(v => (range, CriteriaMatcher.parse(v)))
  }.map { parsedConditions =>
    // For each cell index, check if ALL criteria match
    sumRange.cells.zipWithIndex.foldLeft(BigDecimal(0)) { case (acc, (sumRef, idx)) =>
      val allMatch = parsedConditions.forall { case (criteriaRange, criterion) =>
        val testRef = criteriaRange.cells.toIndexedSeq(idx)
        CriteriaMatcher.matches(sheet(testRef).value, criterion)
      }
      if allMatch then
        sheet(sumRef).value match
          case CellValue.Number(n) => acc + n
          case _ => acc
      else acc
    }
  }
```

---

## Test Cases

```scala
// Basic SUMIF
test("SUMIF: sum values where criteria matches") {
  val sheet = sheetWith(
    ref"A1" -> "Apple", ref"B1" -> 10,
    ref"A2" -> "Banana", ref"B2" -> 20,
    ref"A3" -> "Apple", ref"B3" -> 30
  )
  assertEquals(eval("=SUMIF(A1:A3, \"Apple\", B1:B3)"), Right(40))
}

// Wildcard
test("SUMIF: wildcard matching with *") {
  val sheet = sheetWith(
    ref"A1" -> "Apple", ref"B1" -> 10,
    ref"A2" -> "Apricot", ref"B2" -> 20
  )
  assertEquals(eval("=SUMIF(A1:A2, \"A*\", B1:B2)"), Right(30))
}

// Numeric comparison
test("COUNTIF: count values > 100") {
  val sheet = sheetWith(
    ref"A1" -> 50, ref"A2" -> 150, ref"A3" -> 200
  )
  assertEquals(eval("=COUNTIF(A1:A3, \">100\")"), Right(2))
}

// Multiple criteria
test("SUMIFS: multiple criteria AND logic") {
  val sheet = sheetWith(
    ref"A1" -> "Apple", ref"B1" -> "Red", ref"C1" -> 10,
    ref"A2" -> "Apple", ref"B2" -> "Green", ref"C2" -> 20,
    ref"A3" -> "Banana", ref"B3" -> "Yellow", ref"C3" -> 30
  )
  assertEquals(eval("=SUMIFS(C1:C3, A1:A3, \"Apple\", B1:B3, \"Red\")"), Right(10))
}
```

---

## Definition of Done

### Phase 1: CriteriaMatcher
- [ ] `CriteriaMatcher.scala` created in xl-evaluator
- [ ] `Criterion` sealed trait with Exact, Compare, Wildcard
- [ ] `CompareOp` enum with Gt, Gte, Lt, Lte, Neq
- [ ] `parse()` handles all 5 criteria patterns
- [ ] `matches()` handles all CellValue types
- [ ] Case-insensitive text matching
- [ ] Wildcard escape sequences (~* and ~?)
- [ ] `CriteriaMatcherSpec` with 15+ tests

### Phase 2: TExpr
- [ ] `SumIf` case added to TExpr enum
- [ ] `CountIf` case added to TExpr enum
- [ ] `SumIfs` case added to TExpr enum
- [ ] `CountIfs` case added to TExpr enum
- [ ] Smart constructors in TExpr object
- [ ] Existing tests still pass

### Phase 3: FunctionParser
- [ ] `sumIfParser` given instance (2-3 args)
- [ ] `countIfParser` given instance (2 args)
- [ ] `sumIfsParser` given instance (3+ odd args)
- [ ] `countIfsParser` given instance (2+ even args)
- [ ] Arity validation for all functions
- [ ] Range extraction helper

### Phase 4: Evaluator
- [ ] `SumIf` evaluation implemented
- [ ] `CountIf` evaluation implemented
- [ ] `SumIfs` evaluation with AND logic
- [ ] `CountIfs` evaluation with AND logic
- [ ] Dimension validation for mismatched ranges
- [ ] Non-numeric cells skipped (Excel behavior)

### Phase 5: FormulaPrinter + DependencyGraph
- [ ] Print cases for all 4 functions
- [ ] Round-trip: parse(print(expr)) == expr
- [ ] DependencyGraph extracts all range references
- [ ] No regression in existing tests

### Phase 6: Tests
- [ ] 40+ tests in `ConditionalFunctionsSpec`
- [ ] Coverage: basic, wildcards, comparisons, multiple criteria
- [ ] Coverage: round-trip, error cases, edge cases
- [ ] All tests pass: `./mill __.test`

### Phase 7: Documentation
- [ ] CLAUDE.md function count: 24 → 28
- [ ] docs/STATUS.md updated
- [ ] docs/plan/roadmap.md WI-09f marked complete
- [ ] Format check passes: `./mill __.checkFormat`
- [ ] PR created and merged

---

## Files Summary

### New Files
| File | Purpose | LOC |
|------|---------|-----|
| `xl-evaluator/src/.../formula/CriteriaMatcher.scala` | Criteria parsing + matching | ~200 |
| `xl-evaluator/test/.../formula/CriteriaMatcherSpec.scala` | CriteriaMatcher unit tests | ~150 |
| `xl-evaluator/test/.../formula/ConditionalFunctionsSpec.scala` | Integration tests | ~400 |

### Modified Files
| File | Changes | LOC |
|------|---------|-----|
| `xl-evaluator/src/.../formula/TExpr.scala` | 4 new cases + smart constructors | ~60 |
| `xl-evaluator/src/.../formula/FunctionParser.scala` | 4 new given instances | ~80 |
| `xl-evaluator/src/.../formula/Evaluator.scala` | 4 new match cases | ~80 |
| `xl-evaluator/src/.../formula/FormulaPrinter.scala` | 4 new print cases | ~20 |
| `xl-evaluator/src/.../formula/DependencyGraph.scala` | 4 new dependency cases | ~15 |
| `CLAUDE.md` | Update function count | ~2 |

**Total**: ~1000 LOC new/modified

---

## Module Ownership

**Primary**: `xl-evaluator` (all implementation)

**Test Files**:
- `xl-evaluator/test` (CriteriaMatcherSpec, ConditionalFunctionsSpec)

**Documentation**:
- `CLAUDE.md` (function count)
- `docs/STATUS.md` (feature list)
- `docs/plan/roadmap.md` (work item status)

---

## Merge Risk Assessment

**Risk Level**: Low

**Rationale**:
- All changes within single module (xl-evaluator)
- TExpr changes are additive (new enum cases)
- FunctionParser uses given instances (no modification to existing)
- Evaluator is additive (new match cases, no changes to existing)
- CriteriaMatcher is entirely new file
- No breaking API changes

**Potential Conflicts**:
- WI-09g (INDEX/MATCH) — same files, sequence after this PR
- Any other formula work — coordinate on TExpr.scala

---

## Related Documentation

- **Roadmap**: `docs/plan/roadmap.md` (WI-09f)
- **Formula System**: `docs/design/formula-system.md` (if exists)
- **STATUS**: `docs/STATUS.md` (update after completion)
- **CLAUDE.md**: Function count update

---

## Future Extensions

After WI-09f is complete, these functions become easy to add:
- **AVERAGEIF/AVERAGEIFS**: Same pattern, divide sum by count
- **MAXIF/MAXIFS**: Same pattern, track max instead of sum
- **MINIF/MINIFS**: Same pattern, track min instead of sum

The CriteriaMatcher is reusable for all conditional functions.
