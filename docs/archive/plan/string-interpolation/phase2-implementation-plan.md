# Phase 2 Implementation Plan: Compile-Time Optimization for String Interpolation

**Status**: Ready for Implementation
**Parent**: `string-interpolation.md` (Design Specification)
**Prerequisite**: Phase 1 Complete (PR #12 merged)
**Timeline**: 10-14 developer-days (sequential) or 6-8 days (2 devs parallel)
**PRs**: 3 independently reviewable pull requests

## Executive Summary

Phase 2 adds **compile-time constant detection** to string interpolation. When all interpolated values are compile-time constants, we reconstruct the string at compile time, parse it, validate it, and emit a constant directly - achieving **zero runtime overhead**.

**Phase 1 Status** (Merged):
- All 4 macros support runtime interpolation
- Returns `Either[XLError, T]` for dynamic values
- No compile-time optimization yet

**Phase 2 Goal**:
- Detect compile-time constants using `Expr.value` API
- Emit unwrapped types (`ARef`, `Formatted`) for constants
- Preserve Phase 1 runtime path for dynamic values
- Verify zero-overhead via property tests

### Strategic Approach

**3 PRs in Sequential Order**:

1. **PR #1: Infrastructure** (1-2 days) - `MacroUtil` detection/reconstruction utilities
2. **PR #2: ref"$literal" optimization** (3-4 days) - Most complex, demonstrates pattern
3. **PR #3: Formatted + fx optimization** (2-3 days) - Apply pattern to remaining macros

### Key Architecture: Three-Branch Pattern

All macros will use this structure:

```scala
args match
  case Varargs(exprs) if exprs.isEmpty =>
    // Branch 1: No interpolation (Phase 1)
    compileTimeLiteral(sc)

  case Varargs(exprs) =>
    MacroUtil.allLiterals(args) match
      case Some(literals) =>
        // Branch 2: All compile-time constants (Phase 2 NEW)
        compileTimeOptimized(sc, literals)

      case None =>
        // Branch 3: Has runtime variables (Phase 1)
        runtimeParsing(sc, args)
```

**Decision point**: `MacroUtil.allLiterals(args)` determines optimization path.

---

## Current Phase 1 Architecture

### Existing Implementation (Post-PR #12)

**File**: `xl-core/src/com/tjclp/xl/macros/RefLiteral.scala` (lines 109-125)

```scala
private def refImplN(
  sc: Expr[StringContext],
  args: Expr[Seq[Any]]
)(using Quotes): Expr[ARef | CellRange | RefType | Either[XLError, RefType]] =
  import quotes.reflect.*

  args match
    case Varargs(exprs) if exprs.isEmpty =>
      // No interpolation - compile-time literal
      refImpl0(sc)

    case Varargs(exprs) =>
      // Has interpolation - ALWAYS runtime parsing (Phase 1)
      '{
        val str = $sc.s($args*)
        RefType.parseToXLError(str)
      }.asExprOf[Either[XLError, RefType]]
```

**Observation**: Two branches currently (no-args, with-args). Phase 2 splits the with-args branch into **two paths** (all-literals, any-runtime).

### Expr.value API (Already Used in XL)

**File**: `xl-core/src/com/tjclp/xl/style/dsl.scala` (lines 286-307)

The `.hex()` method demonstrates the pattern:

```scala
code.value match
  case Some(literal) =>
    // Compile-time constant - validate now!
    Color.fromHex(literal) match
      case Right(c) => emitConstant(c, style)
      case Left(err) => report.errorAndAbort(s"Invalid hex: $err")

  case None =>
    // Runtime variable - runtime validation
    '{ Color.fromHex($code).fold(_ => $style, c => applyColor($style, c)) }
```

**Key API**: `Expr[T].value: Option[T]`
- Returns `Some(literal)` for compile-time constants
- Returns `None` for runtime expressions (function calls, variables, etc.)

---

## PR #1: Infrastructure - Detection & Reconstruction Utilities

**Priority**: Must merge first
**Estimated Effort**: 8-12 hours (1-2 days)
**Risk Level**: VERY LOW (purely additive)

### 1. Objective

Create shared utilities for Phase 2 optimization:
- Compile-time constant detection (`MacroUtil.allLiterals`)
- String reconstruction (`MacroUtil.reconstructString`)
- Error formatting helpers
- Test utilities for verification

### 2. Files Created

**2 files created**:
- `xl-core/src/com/tjclp/xl/macros/MacroUtil.scala` (~120 lines)
- `xl-core/test/src/com/tjclp/xl/macros/MacroUtilSpec.scala` (~100 lines)

### 3. Implementation: MacroUtil.scala

**Location**: `xl-core/src/com/tjclp/xl/macros/MacroUtil.scala`

```scala
package com.tjclp.xl.macros

import scala.quoted.*

/**
 * Shared utilities for macro compile-time optimization (Phase 2).
 *
 * Provides detection and reconstruction helpers used by all macros.
 */
object MacroUtil:

  /**
   * Detect if all interpolation arguments are compile-time constants.
   *
   * Returns None if any argument is a runtime expression, or Some(literals) if all are constants.
   *
   * Uses Expr.value to check if each argument is a compile-time constant.
   *
   * Example:
   * {{{
   * val x = "A1"              // Compile-time constant
   * ref"$x"                   // allLiterals → Some(Seq("A1"))
   *
   * def getUserInput() = "A1" // Runtime expression
   * ref"${getUserInput()}"    // allLiterals → None
   * }}}
   */
  def allLiterals(args: Expr[Seq[Any]])(using Quotes): Option[Seq[Any]] =
    import quotes.reflect.*

    args match
      case Varargs(exprs) =>
        val literalValues = exprs.map(_.value)
        // Check if ALL are defined (all are compile-time constants)
        if literalValues.forall(_.isDefined) then Some(literalValues.flatten.toSeq)
        else None
      case _ => None

  /**
   * Reconstruct interpolated string from parts and literal values.
   *
   * Interleaves StringContext.parts with argument values to produce the full string.
   *
   * Algorithm: parts(0) + literals(0) + parts(1) + literals(1) + ... + parts(n)
   *
   * Invariant: parts.length == literals.length + 1 (enforced by Scala's StringContext)
   *
   * Example:
   * {{{
   * parts = Seq("", "!A1")
   * literals = Seq("Sales")
   * reconstructString(parts, literals) → "Sales!A1"
   *
   * parts = Seq("", "", "")
   * literals = Seq("B", 42)
   * reconstructString(parts, literals) → "B42"
   * }}}
   */
  @SuppressWarnings(Array("org.wartremover.warts.Var", "org.wartremover.warts.While"))
  def reconstructString(parts: Seq[String], literals: Seq[Any]): String =
    require(
      parts.length == literals.length + 1,
      s"parts.length (${parts.length}) must equal literals.length + 1 (${literals.length + 1})"
    )

    val sb = new StringBuilder
    var i = 0
    val n = literals.length

    while i < n do
      sb.append(parts(i))
      sb.append(literals(i).toString)
      i += 1

    sb.append(parts(i)) // Final part (always exists)
    sb.toString

  /**
   * Format compile error for invalid compile-time interpolation.
   *
   * Provides helpful error message with the reconstructed string and parse error.
   *
   * Example output:
   * {{{
   * Invalid ref literal in interpolation: 'INVALID!@#$'
   * Error: Invalid characters in sheet name
   * Hint: Check that all interpolated parts form a valid Excel ref
   * }}}
   */
  def formatCompileError(macroName: String, fullString: String, parseError: String): String =
    s"Invalid $macroName literal in interpolation: '$fullString'\n" +
      s"Error: $parseError\n" +
      s"Hint: Check that all interpolated parts form a valid Excel $macroName"

end MacroUtil
```

### 4. Testing: MacroUtilSpec.scala

**Location**: `xl-core/test/src/com/tjclp/xl/macros/MacroUtilSpec.scala`

```scala
package com.tjclp.xl.macros

import munit.FunSuite

class MacroUtilSpec extends FunSuite:

  // ===== String Reconstruction Tests =====

  test("reconstructString: empty literals (no interpolation)") {
    val parts = Seq("A1")
    val literals = Seq.empty
    assertEquals(MacroUtil.reconstructString(parts, literals), "A1")
  }

  test("reconstructString: single literal") {
    val parts = Seq("", "!A1")
    val literals = Seq("Sales")
    assertEquals(MacroUtil.reconstructString(parts, literals), "Sales!A1")
  }

  test("reconstructString: multiple literals") {
    val parts = Seq("", "!", "")
    val literals = Seq("Sales", "A1")
    assertEquals(MacroUtil.reconstructString(parts, literals), "Sales!A1")
  }

  test("reconstructString: literals with numbers") {
    val parts = Seq("", "", "")
    val literals = Seq("B", 42)
    assertEquals(MacroUtil.reconstructString(parts, literals), "B42")
  }

  test("reconstructString: complex interpolation") {
    val parts = Seq("", "!", "", ":")
    val literals = Seq("Sheet1", "A1", "B10")
    assertEquals(MacroUtil.reconstructString(parts, literals), "Sheet1!A1:B10")
  }

  test("reconstructString: with prefix") {
    val parts = Seq("Sales!", "")
    val literals = Seq("A1")
    assertEquals(MacroUtil.reconstructString(parts, literals), "Sales!A1")
  }

  test("reconstructString: with suffix") {
    val parts = Seq("", ":B10")
    val literals = Seq("A1")
    assertEquals(MacroUtil.reconstructString(parts, literals), "A1:B10")
  }

  test("reconstructString: all empty parts") {
    val parts = Seq("", "")
    val literals = Seq("A1")
    assertEquals(MacroUtil.reconstructString(parts, literals), "A1")
  }

  test("reconstructString: invariant violation fails") {
    val parts = Seq("A", "B")
    val literals = Seq("1", "2") // parts.length != literals.length + 1
    intercept[IllegalArgumentException] {
      MacroUtil.reconstructString(parts, literals)
    }
  }

  // ===== Error Formatting Tests =====

  test("formatCompileError: includes all components") {
    val err = MacroUtil.formatCompileError("ref", "INVALID!@#$", "Invalid characters in sheet name")
    assert(err.contains("Invalid ref literal"))
    assert(err.contains("INVALID!@#$"))
    assert(err.contains("Invalid characters"))
    assert(err.contains("Hint"))
  }

  test("formatCompileError: includes macro name") {
    val err = MacroUtil.formatCompileError("money", "$ABC", "non-numeric")
    assert(err.contains("Invalid money literal"))
    assert(err.contains("$ABC"))
  }
```

### 5. Success Criteria

- [ ] `./mill xl-core.compile` - Zero warnings
- [ ] `./mill xl-core.test.testOnly com.tjclp.xl.macros.MacroUtilSpec` - 11/11 tests pass
- [ ] `./mill __.test` - All 374 existing tests + 11 new = 385 total
- [ ] String reconstruction handles all edge cases
- [ ] No functional changes to existing macros

### 6. Rollback Plan

Simple revert (zero risk, purely additive).

---

## PR #2: Compile-Time Optimization for ref"$literal"

**Priority**: Depends on PR #1
**Estimated Effort**: 20-28 hours (3-4 days)
**Risk Level**: MEDIUM

### Implementation Summary

**Three-branch refactor of `refImplN`**:
1. No args → existing `refImpl0` (Phase 1)
2. All literals → new `compileTimeOptimizedPath` (Phase 2)
3. Any runtime → existing `runtimePath` (Phase 1)

**Key Functions Added**:
- `compileTimeOptimizedPath(sc, literals)` - Reconstruct + parse + emit constant
- `runtimePath(sc, args)` - Extract existing runtime logic

**Testing**: 30 new tests verifying optimization + identity law

**Documentation**: Update examples.md with optimization behavior

---

## PR #3: Formatted + fx Optimization

**Priority**: Depends on PR #1 and PR #2
**Estimated Effort**: 12-20 hours (2-3 days)
**Risk Level**: LOW

### Implementation Summary

Apply three-branch pattern to:
- `money`, `percent`, `date`, `accounting` (FormattedLiterals.scala)
- `fx` (CellRangeLiterals.scala)

**Pattern** (same for all):
```scala
MacroUtil.allLiterals(args) match
  case Some(literals) =>
    val fullString = MacroUtil.reconstructString(parts, literals)
    parseAndEmitConstant(fullString)
  case None =>
    emitRuntimeParsing(sc, args)
```

**Testing**: 24 new tests (6 per macro type)

---

## Phase 2 Completion Criteria

After all 3 PRs merged:

- [ ] 438 total tests (374 existing + 64 new)
- [ ] Zero warnings across all modules
- [ ] Compile-time constants emit unwrapped types
- [ ] Runtime variables still return Either
- [ ] Property tests verify identity law
- [ ] Documentation updated
- [ ] Benchmarks show zero overhead for constants

---

## Timeline

**Sequential (1 dev)**: 10-14 days (2-3 weeks)
**Parallel (2 devs)**: 6-8 days per dev (1.5-2 weeks calendar)

**Milestones**:
- M1: Infrastructure (End of Week 1, Day 2)
- M2: ref Optimization (End of Week 2, Day 2)
- M3: All Macros Optimized (End of Week 2, Day 5)
- M4: Phase 2 Complete (End of Week 3, Day 3)

---

**Document Metadata**:
- **Created**: 2025-11-16
- **Author**: Generated by Claude Code
- **Version**: 1.0
- **Status**: Ready for Implementation
