# Phase 2 Feasibility Issue: Compile-Time Constant Detection in String Interpolation

**Date**: 2025-11-16
**Status**: Research Required
**Priority**: Blocking for Phase 2

## Executive Summary

Phase 2 compile-time optimization is currently **not working** due to inability to detect compile-time constants in string interpolation arguments. The `MacroUtil.allLiterals` detection function returns `None` for all interpolations, even those with `inline val` or direct string literals, causing all interpolations to use the runtime path instead of being optimized.

**Impact**: Without this working, Phase 2 provides no value (no zero-overhead optimization).

**Next Steps**: Research alternative constant detection mechanisms in Scala 3 macros.

---

## Goal: What We're Trying to Achieve

### Desired Behavior (Phase 2)

Enable zero-overhead compile-time optimization when all interpolation arguments are compile-time constants:

```scala
// Example 1: inline val constants
inline val sheet = "Sales"
inline val cell = "A1"
ref"$sheet!$cell"  // Should be optimized at compile time
                   // Expected: RefType.QualifiedCell (unwrapped)
                   // Actual: Either[XLError, RefType] (runtime path)

// Example 2: Direct literals
ref"${"Sales"}!${"A1"}"  // Should be optimized
                          // Expected: RefType.QualifiedCell
                          // Actual: Either[XLError, RefType]

// Example 3: Non-constant (correctly uses runtime)
def getUserSheet() = "Sales"
ref"${getUserSheet()}!A1"  // Should use runtime path
                           // Expected: Either[XLError, RefType] ✓
                           // Actual: Either[XLError, RefType] ✓
```

### Expected vs Actual Results

| Case | Input | Expected Return Type | Actual Return Type | Status |
|------|-------|---------------------|-------------------|--------|
| No interpolation | `ref"A1"` | `ARef` (unwrapped) | `ARef` ✓ | ✅ Works (Phase 1) |
| Direct literals | `ref"${"A"}${1}"` | `ARef` (optimized) | `Either[XLError, RefType]` | ❌ **NOT optimized** |
| inline val | `inline val x="A"; ref"$x"` | `ARef` (optimized) | `Either[XLError, RefType]` | ❌ **NOT optimized** |
| Function call | `ref"${f()}"` | `Either[...]` (runtime) | `Either[...]` ✓ | ✅ Works (Phase 1) |

**Bottom line**: Only "no interpolation" case works. All interpolations (even with constants) use runtime path.

---

## The Problem: MacroUtil.allLiterals Not Detecting Constants

### Current Implementation

**File**: `xl-core/src/com/tjclp/xl/macros/MacroUtil.scala` (lines 29-44)

```scala
def allLiterals(args: Expr[Seq[Any]])(using Quotes): Option[Seq[Any]] =
  import quotes.reflect.*

  args match
    case Varargs(exprs) =>
      val literalValues = exprs.map { expr =>
        // Check if expression is a literal constant using reflection
        expr.asTerm match
          case Inlined(_, _, Literal(constant)) => Some(constant.value)
          case Literal(constant) => Some(constant.value)
          case _ => None
      }
      // Check if ALL are defined (all are compile-time constants)
      if literalValues.forall(_.isDefined) then Some(literalValues.flatten.toSeq)
      else None
    case _ => None
```

###  What We're Checking

Using `quotes.reflect` to check if `Expr` is a `Literal` or `Inlined(..., Literal(...))`.

**Pattern matching on**:
- `Inlined(_, _, Literal(constant))` - Inlined literal constant
- `Literal(constant)` - Direct literal constant
- `_` - Everything else (runtime expressions)

### What's Failing

Even for cases like:
```scala
inline val x = "A"
ref"$x"
```

The `expr.asTerm` is **NOT** matching `Literal(...)` or `Inlined(..., Literal(...))`.

**Hypothesis**: The macro sees a *reference* to the inline val, not the inlined literal itself. Scala 3's inlining may happen after macro expansion, or the reference isn't being beta-reduced before the macro sees it.

---

## Technical Details

### How String Interpolation Works

When you write:
```scala
val sheet = "Sales"
ref"$sheet!A1"
```

The compiler transforms it to:
```scala
StringContext("", "!A1").ref(sheet)
```

**Macro receives**:
- `sc: Expr[StringContext]` with `parts = Seq("", "!A1")`
- `args: Expr[Seq[Any]]` containing the expressions for each interpolated value

### What the Macro Sees

**For `ref"$sheet!A1"` where `val sheet = "Sales"`**:

```scala
args match
  case Varargs(exprs) =>
    // exprs(0) is an Expr[Any] representing the variable reference to `sheet`
    exprs(0).asTerm  // What is this?
```

### Debugging: What Does asTerm Actually Return?

We need to inspect what `expr.asTerm` actually is for different cases:

```scala
// Case 1: Direct string literal
ref"${"Sales"}!A1"
// expr.asTerm = ???

// Case 2: inline val
inline val x = "Sales"
ref"$x!A1"
// expr.asTerm = ???

// Case 3: Regular val
val y = "Sales"
ref"$y!A1"
// expr.asTerm = ???

// Case 4: Function call
ref"${getSheet()}!A1"
// expr.asTerm = Apply(...) presumably
```

**We need to add debug logging to see what these actually are!**

---

## Research Directions

### 1. Use Expr.valueOrAbort / Expr.value (Type-Safe Version)

The working `.hex()` method uses:
```scala
code.value match  // code: Expr[String]
  case Some(literal) => ...
  case None => ...
```

**Key difference**: `.value` is called on `Expr[String]`, not `Expr[Any]`.

**Question**: Can we somehow cast/convert `Expr[Any]` to a typed `Expr[String]` or `Expr[Int]` and then use `.value`?

**Possible approach**:
```scala
exprs.map { expr =>
  // Try to extract as various types
  expr.asExprOf[String].value
    .orElse(expr.asExprOf[Int].value.map(_.toString))
    .orElse(expr.asExprOf[Long].value.map(_.toString))
    // ... etc
}
```

### 2. Check for Ident Referring to Inline Val

Maybe we need to check if the `Term` is an `Ident` (identifier reference) and then look up its definition to see if it's an inline val?

**Possible approach**:
```scala
expr.asTerm match
  case Ident(name) =>
    // Look up the symbol
    val symbol = expr.asTerm.symbol
    // Check if it's an inline val
    if symbol.flags.is(Flags.Inline) then
      // Get the RHS of the inline val?
      ???
```

### 3. Use Staging / Explicit Inlining

Scala 3 has staging capabilities (`scala.quoted.staging`). Maybe we need to explicitly request inlining?

**Research**: Can we force inline expansion before macro sees the arguments?

### 4. Check Scala 3 Documentation

**Resources to consult**:
- [Scala 3 Macros Documentation](https://docs.scala-lang.org/scala3/guides/macros/)
- [TASTy Reflect API](https://docs.scala-lang.org/scala3/reference/metaprogramming/tasty-inspect.html)
- Scala 3 Macro Examples (esp. string interpolation macros)
- [Contextual library source](https://github.com/propensive/contextual) - They solve this!

### 5. Consult Existing Libraries

**Contextual library** by propensive successfully does compile-time validation of string interpolations. How do they do it?

**Other libraries** that might have solved this:
- sourcecode library (compile-time value extraction)
- iron (refined types with compile-time validation)

---

## What We've Tried

### Attempt 1: asTerm Pattern Matching

```scala
expr.asTerm match
  case Inlined(_, _, Literal(constant)) => Some(constant.value)
  case Literal(constant) => Some(constant.value)
  case _ => None
```

**Result**: Always matches `case _` (not a literal)

### Attempt 2: inline val in Tests

```scala
inline val sheet = "Sales"
ref"$sheet!A1"
```

**Result**: Still uses runtime path, returns `Either`

### Attempt 3: Direct Literals

```scala
ref"${"Sales"}!${"A1"}"
```

**Result**: (Not yet tested, but likely same issue)

---

## Diagnostic Approach

### Add Debug Logging to MacroUtil

To understand what the macro actually sees, add debug output:

```scala
def allLiterals(args: Expr[Seq[Any]])(using Quotes): Option[Seq[Any]] =
  import quotes.reflect.*

  args match
    case Varargs(exprs) =>
      val literalValues = exprs.zipWithIndex.map { case (expr, idx) =>
        val term = expr.asTerm

        // DEBUG: Print what we see
        println(s"[DEBUG] Arg $idx:")
        println(s"  expr.show: ${expr.show}")
        println(s"  term: $term")
        println(s"  term.getClass: ${term.getClass}")
        println(s"  term.tpe: ${term.tpe}")

        term match
          case Inlined(_, _, Literal(constant)) =>
            println(s"  ✓ Matched Inlined Literal: ${constant.value}")
            Some(constant.value)
          case Literal(constant) =>
            println(s"  ✓ Matched Direct Literal: ${constant.value}")
            Some(constant.value)
          case Ident(name) =>
            println(s"  ✗ Matched Ident: $name (not a literal)")
            None
          case Apply(_, _) =>
            println(s"  ✗ Matched Apply (function call)")
            None
          case other =>
            println(s"  ✗ Matched other: ${other.getClass.getSimpleName}")
            None
      }

      if literalValues.forall(_.isDefined) then Some(literalValues.flatten.toSeq)
      else None
    case _ => None
```

**Action**: Add this debug version, run tests, examine output to see what the macro actually receives.

---

## Comparison: How .hex() Works (IT DOES WORK)

### Working Example

**File**: `xl-core/src/com/tjclp/xl/style/dsl.scala` (lines 281-307)

```scala
// In style DSL
transparent inline def hex(code: String): CellStyle = ${ validateHex('code, 'style, false) }

// Macro implementation
def validateHex(code: Expr[String], ...)(using Quotes): Expr[CellStyle] =
  code.value match  // This WORKS!
    case Some(literal) => // Compile-time constant
      Color.fromHex(literal) match
        case Right(c) => emitConstant(c, style)
        case Left(err) => report.errorAndAbort(s"Invalid hex: $err")
    case None => // Runtime variable
      '{ Color.fromHex($code).fold(...) }
```

**Why it works**:
- `code: Expr[String]` is a typed expression
- `.value: Option[String]` works because there's a `FromExpr[String]` instance
- The macro receives a typed `Expr[String]`, not `Expr[Any]`

### Why String Interpolation is Different

**String interpolation signature**:
```scala
extension (inline sc: StringContext)
  transparent inline def ref(inline args: Any*): ... = ${ refImplN('sc, 'args) }

// Macro receives:
def refImplN(sc: Expr[StringContext], args: Expr[Seq[Any]]): ...
```

**Key difference**: `args: Expr[Seq[Any]]` is:
- A sequence of `Any` (not typed)
- Each element is `Expr[Any]` (no `FromExpr[Any]` instance)
- Can't use `.value` on `Expr[Any]`

### Potential Solution: Type-Based Extraction

**Idea**: Try to cast each `Expr[Any]` to common types and extract value:

```scala
def extractValue(expr: Expr[Any])(using Quotes): Option[Any] =
  expr.asExprOf[String].value
    .orElse(expr.asExprOf[Int].value.map(_.toString))
    .orElse(expr.asExprOf[Long].value.map(_.toString))
    .orElse(expr.asExprOf[Double].value.map(_.toString))
    .orElse(expr.asExprOf[Boolean].value.map(_.toString))
    // ... try all reasonable types
```

**Status**: NOT YET TESTED - this could be the solution!

---

## Code Examples for Testing

### Test Case 1: Direct String Literal

```scala
// User code
ref"${"Sales"}!${"A1"}"

// What macro sees:
// sc.parts = Seq("", "!", "")
// args = Varargs(Seq(Expr("Sales"), Expr("A1")))

// Expected: Both exprs should match Literal("Sales") and Literal("A1")
// Actual: ???
```

### Test Case 2: inline val

```scala
// User code
inline val sheet = "Sales"
ref"$sheet!A1"

// What macro sees:
// sc.parts = Seq("", "!A1")
// args = Varargs(Seq(Expr(sheet)))

// Expected: expr should be inlined to Literal("Sales")
// Actual: expr is Ident("sheet") reference?
```

### Test Case 3: Regular val (Should NOT Optimize)

```scala
// User code
val sheet = "Sales"
ref"$sheet!A1"

// What macro sees:
// sc.parts = Seq("", "!A1")
// args = Varargs(Seq(Expr(sheet)))

// Expected: Runtime path (can't optimize)
// Actual: Runtime path ✓ (correct)
```

---

## Diagnostic Test Suite

### Minimal Reproduction

Create a minimal test to diagnose the issue:

**File**: `xl-core/test/src/com/tjclp/xl/macros/DiagnosticSpec.scala`

```scala
package com.tjclp.xl.macros

import munit.FunSuite
import com.tjclp.xl.*

class DiagnosticSpec extends FunSuite:

  test("Diagnostic 1: Direct string literal in interpolation") {
    val result = ref"${"A"}${1}"
    println(s"Direct literal result type: ${result.getClass}")

    result match
      case aref: ARef =>
        println("✓ SUCCESS: Optimized to ARef")
      case Right(RefType.Cell(_)) =>
        println("✗ FAILURE: Not optimized, using runtime path")
      case other =>
        println(s"✗ UNEXPECTED: $other")
  }

  test("Diagnostic 2: inline val in function scope") {
    inline val col = "A"
    inline val row = 1
    val result = ref"$col$row"
    println(s"inline val result type: ${result.getClass}")

    result match
      case aref: ARef =>
        println("✓ SUCCESS: Optimized to ARef")
      case Right(RefType.Cell(_)) =>
        println("✗ FAILURE: Not optimized, using runtime path")
      case other =>
        println(s"✗ UNEXPECTED: $other")
  }

  test("Diagnostic 3: Regular val (should NOT optimize)") {
    val col = "A"
    val row = 1
    val result = ref"$col$row"
    println(s"Regular val result type: ${result.getClass}")

    result match
      case Right(RefType.Cell(_)) =>
        println("✓ CORRECT: Using runtime path as expected")
      case aref: ARef =>
        println("✗ UNEXPECTED: Should use runtime path, not optimize")
      case other =>
        println(s"✗ UNEXPECTED: $other")
  }
```

**Run with**: `./mill xl-core.test.testOnly com.tjclp.xl.macros.DiagnosticSpec`

**Expected output**: Will show exactly what type we're getting for each case.

---

## Alternative Approaches to Research

### Approach 1: Type-Safe Expr.value with asExprOf

Instead of pattern matching on `asTerm`, try using typed `.value`:

```scala
def allLiterals(args: Expr[Seq[Any]])(using Quotes): Option[Seq[Any]] =
  args match
    case Varargs(exprs) =>
      val values = exprs.map { expr =>
        // Try common types
        tryExtractValue[String](expr)
          .orElse(tryExtractValue[Int](expr))
          .orElse(tryExtractValue[Long](expr))
          .orElse(tryExtractValue[Double](expr))
          .orElse(tryExtractValue[Boolean](expr))
      }
      if values.forall(_.isDefined) then Some(values.flatten.toSeq)
      else None

def tryExtractValue[T: Type: FromExpr](expr: Expr[Any])(using Quotes): Option[Any] =
  try
    expr.asExprOf[T].value
  catch
    case _: Exception => None
```

**Status**: Needs testing, might work!

### Approach 2: Symbol Introspection for inline vals

Check if the expression is an Ident referring to an inline val, then extract its RHS:

```scala
expr.asTerm match
  case Ident(name) =>
    val symbol = expr.asTerm.symbol
    if symbol.flags.is(Flags.Inline) then
      // Get the inline val's definition
      val tree = symbol.tree
      tree match
        case ValDef(_, _, Some(rhs)) =>
          rhs match
            case Literal(constant) => Some(constant.value)
            case _ => None
        case _ => None
    else None
```

**Status**: Speculative, needs research on Symbol API

### Approach 3: Consult Contextual Library Implementation

**Contextual** (propensive/contextual) successfully does compile-time string interpolation validation.

**Research**:
1. Clone contextual repo
2. Find their constant detection mechanism
3. Adapt it to XL's use case

**Likely location**: `src/core/contextual.scala` in the Contextual library

### Approach 4: Use Macros Differently

Maybe we need to change the macro signature itself:

**Current**:
```scala
transparent inline def ref(inline args: Any*): ... = ${ refImplN('sc, 'args) }
```

**Alternative**: Use individual typed parameters?
```scala
transparent inline def ref2[T1: Type](inline arg1: T1): ... = ${ ref2Impl[T1]('arg1) }
```

**Problem**: This doesn't scale (can't have variable number of typed parameters)

### Approach 5: Require inline at Call Site

Document that users must use direct literals, not variables:

```scala
// ✓ Works (direct literals)
ref"${"Sales"}!${"A1"}"

// ✗ Doesn't optimize (even with inline val)
inline val x = "Sales"
ref"$x!A1"  // Still runtime path
```

**Status**: Defeats the purpose of Phase 2, but documents reality

---

## Questions for Research

### Q1: Does asExprOf[T].value Work for Any?

**Test**:
```scala
val expr: Expr[Any] = '{  "Sales" }
expr.asExprOf[String].value  // Does this work?
```

**Expected**: Might throw or return None if types don't match
**If works**: Could iterate through common types

### Q2: Can We Inspect Inline Val Definitions?

**Test**:
```scala
inline val x = "Sales"

// In macro:
expr.asTerm match
  case Ident(name) =>
    expr.asTerm.symbol.flags.is(Flags.Inline)  // true?
    expr.asTerm.symbol.tree  // ValDef with RHS?
```

**Research**: Symbol API documentation, inspect tree structure

### Q3: How Does Contextual Do It?

**Action**: Study Contextual library source code
- Find constant detection mechanism
- Understand their approach
- Adapt to XL

**Link**: https://github.com/propensive/contextual

### Q4: Is There a Beta-Reduction API?

**Question**: Can we ask the compiler to beta-reduce inline defs before macro expansion?

**Research**: Check if there's a `quotes.reflect` API to force inlining/reduction

---

## Workarounds if Phase 2 Proves Infeasible

### Workaround 1: Phase 1 is Actually Fine

**Reality check**: Phase 1 runtime interpolation is:
- Pure and total (returns Either)
- Fast enough for user-facing operations (parsing user input)
- Already better than most libraries

**Conclusion**: If Phase 2 doesn't work, Phase 1 is still a huge win. Don't let perfect be enemy of good.

### Workaround 2: Provide Separate Optimized Macro

If we can't auto-detect, provide explicit optimized versions:

```scala
// Runtime (Phase 1)
ref"$sheet!$cell"  // Always runtime, returns Either

// Compile-time (explicit)
refConst("Sales", "A1")  // Compile-time only, returns ARef
// Or: refC"Sales", "A1"  // Different macro name
```

**Pro**: Clear separation, users know what they get
**Con**: Two APIs to learn, verbose

### Workaround 3: Document Limitation

If optimization only works for direct literals `ref"${"A"}${1}"` (useless), document this clearly:

> **Phase 2 Limitation**: Compile-time optimization only applies to direct string/number literals in interpolation positions. Variables (even `inline val`) always use runtime path. For zero-overhead, use non-interpolated literals: `ref"Sales!A1"` instead of `ref"$sheet!$cell"`.

---

## Current Code Status

### What's Implemented (Potentially Not Working)

- [x] `MacroUtil.allLiterals` - Compiles but may always return None
- [x] `MacroUtil.reconstructString` - Works (tested, 11 tests pass)
- [x] Three-branch pattern in `refImplN` - Compiles but optimization branch may never execute
- [x] `compileTimeOptimizedPath` - Compiles but may never be called

### What's NOT Working

- [ ] Compile-time constant detection (allLiterals always returns None)
- [ ] Tests expect optimization but get runtime path
- [ ] Identity law tests fail (type mismatch)

### What's Definitely Working

- [x] Phase 1 runtime path (unchanged, all tests pass)
- [x] MacroUtil utility functions (reconstruction, error formatting)
- [x] No regressions to existing functionality

---

## Recommended Next Steps

### Step 1: Add Debug Logging

Add debug prints to `MacroUtil.allLiterals` to see what the macro actually receives.

### Step 2: Run Diagnostic Tests

Create and run minimal diagnostic tests to understand behavior.

### Step 3: Research Approach 1 (asExprOf[T].value)

Test if we can use typed `.value` extraction:
```scala
expr.asExprOf[String].value.orElse(expr.asExprOf[Int].value)
```

### Step 4: Study Contextual Library

Examine how propensive/contextual solves this problem.

### Step 5: Consult Scala 3 Community

If still stuck:
- Post on Scala Users forum
- Ask in Scala Discord #macros channel
- Consult Scala 3 macro documentation maintainers

---

## Success Criteria for Phase 2

Phase 2 is only viable if we can:

1. ✅ Detect when **some** interpolations are compile-time constants
2. ✅ Extract their values reliably
3. ✅ Emit optimized code (unwrapped types)
4. ✅ Provide meaningful value over Phase 1

**Current Status**: ❌ Step 1 is failing (detection doesn't work)

**Decision Point**: If research doesn't find a solution, **abandon Phase 2** and mark Phase 1 as complete, or **redefine Phase 2 scope** to something achievable.

---

## File Locations for Further Investigation

**Current implementation**:
- `xl-core/src/com/tjclp/xl/macros/MacroUtil.scala` - Literal detection (broken)
- `xl-core/src/com/tjclp/xl/macros/RefLiteral.scala` - Three-branch pattern (implemented)

**Working reference**:
- `xl-core/src/com/tjclp/xl/style/dsl.scala:286` - `.hex()` method (works!)

**Test failures**:
- `xl-core/test/src/com/tjclp/xl/macros/RefCompileTimeOptimizationSpec.scala` - All expect optimization, all fail

**Resources**:
- [Scala 3 Macros Guide](https://docs.scala-lang.org/scala3/guides/macros/macros.html)
- [TASTy Reflect API](https://docs.scala-lang.org/scala3/reference/metaprogramming/tasty-inspect.html)
- [Contextual Library](https://github.com/propensive/contextual)

---

## Document Metadata

**Created**: 2025-11-16
**Author**: Generated by Claude Code (Anthropic)
**Status**: Research Required
**Blocking**: Phase 2 Implementation
**Severity**: High (Phase 2 may not be feasible)

**Next Action**: Research alternative constant detection mechanisms, particularly Approach 1 (asExprOf[T].value) and studying Contextual library.
