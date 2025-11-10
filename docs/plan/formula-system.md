
# Typed Formula System — GADT, Laws, Printer & Evaluator

We provide two levels:
1. **`TExpr[A]`**: a typed, total expression GADT for programmatic formulas with pure evaluation.
2. **`formula.Expr[V]`**: an interop AST for Excel string formulas (`V` captures validation state).

## GADT (selected constructors)
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

## Laws (sketch)
- **If‑fusion:** `If(c, Lit(x), Lit(y)) ≡ Lit(if ⟦c⟧ then x else y)`.
- **Ring laws:** `Add/Mul` form a commutative semiring over `BigDecimal` nodes modulo printer parentheses.
- **Short‑circuit:** `And/Or` evaluator respects left‑to‑right semantics.

## Printer (Excel string)
- We provide a **total printer** from `TExpr[A]` to Excel text with correct precedence & parentheses.
- Where Excel lacks a direct function, we use idioms (e.g., `SUM` + `COUNT` for average).

## Evaluator
```scala
trait Eval:
  def eval[A](e: TExpr[A], s: Sheet): Either[EvalError, A]
```
- Pure; cycle detection happens on the dependency graph before calling `eval`.
- Deterministic; no external state.
