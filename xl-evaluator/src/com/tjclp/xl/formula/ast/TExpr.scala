package com.tjclp.xl.formula.ast

import com.tjclp.xl.formula.functions.{FunctionSpec, FunctionSpecs}
import com.tjclp.xl.formula.eval.EvalError
import com.tjclp.xl.formula.functions.EvalContext

import com.tjclp.xl.{ARef, Anchor, CellRange, SheetName}
import com.tjclp.xl.cells.Cell
import com.tjclp.xl.codec.CodecError

import scala.math.BigDecimal

/**
 * Typed expression GADT for Excel formulas.
 *
 * TExpr[A] represents a formula expression that evaluates to type A. This is a purely functional,
 * total representation - all operations return Either for error handling, no exceptions thrown.
 *
 * Type parameter A captures the result type:
 *   - TExpr[BigDecimal] for numeric expressions
 *   - TExpr[Boolean] for boolean/logical expressions
 *   - TExpr[String] for text expressions
 *
 * Laws satisfied:
 *   1. If-fusion: If(c, Lit(x), Lit(y)) ≡ Lit(if ⟦c⟧ then x else y)
 *   2. Ring laws: Add/Mul form commutative semiring over BigDecimal
 *   3. Short-circuit: And/Or respect left-to-right semantics
 *   4. Round-trip: parse(print(expr)) == Right(expr)
 */
enum TExpr[A] derives CanEqual:
  /**
   * Literal value - compile-time constant.
   *
   * Example: Lit(42) represents the number 42
   */
  case Lit[A](value: A) extends TExpr[A]

  /**
   * Cell reference - reads value from a cell at runtime.
   *
   * @param at
   *   The cell address to read from
   * @param anchor
   *   Anchoring mode for formula dragging (default: Relative)
   * @param decode
   *   Function to decode the cell's value to type A
   *
   * Example: Ref(ARef("A1"), Anchor.Relative, decodeNumber) reads numeric value from A1
   */
  case Ref[A](at: ARef, anchor: Anchor, decode: Cell => Either[CodecError, A]) extends TExpr[A]

  /**
   * Polymorphic cell reference - defers type commitment until function context.
   *
   * This is used during parsing when the expected type of a cell reference is not yet known.
   * Function parsers convert PolyRef to Ref with appropriate type-coercing decoder based on
   * context.
   *
   * Matches Excel's type coercion semantics: A1=42 in =LEFT(A1,1) converts 42→"4".
   *
   * @param at
   *   The cell address to read from
   * @param anchor
   *   Anchoring mode for formula dragging (default: Relative)
   *
   * Example: PolyRef(ARef("A1"), Anchor.Relative) - type determined by enclosing function
   */
  case PolyRef(at: ARef, anchor: Anchor = Anchor.Relative) extends TExpr[Nothing]

  /**
   * Sheet-qualified cell reference - reads value from a cell in another sheet.
   *
   * @param sheet
   *   The target sheet name
   * @param at
   *   The cell address within the target sheet
   * @param anchor
   *   Anchoring mode for formula dragging
   * @param decode
   *   Function to decode the cell's value to type A
   *
   * Example: SheetRef("Sales", ARef("F4"), Anchor.Relative, decodeNumber) reads from Sales!F4
   */
  case SheetRef[A](
    sheet: SheetName,
    at: ARef,
    anchor: Anchor,
    decode: Cell => Either[CodecError, A]
  ) extends TExpr[A]

  /**
   * Polymorphic sheet-qualified cell reference - defers type commitment until function context.
   *
   * @param sheet
   *   The target sheet name
   * @param at
   *   The cell address within the target sheet
   * @param anchor
   *   Anchoring mode for formula dragging
   *
   * Example: SheetPolyRef("Sales", ARef("F4"), Anchor.Relative) - type determined by context
   */
  case SheetPolyRef(sheet: SheetName, at: ARef, anchor: Anchor = Anchor.Relative)
      extends TExpr[Nothing]

  /**
   * Local range reference.
   *
   * Represents a range like A1:B10 before being consumed by a function.
   */
  case RangeRef(range: CellRange) extends TExpr[Nothing]

  /**
   * Sheet-qualified range reference - references a range in another sheet.
   *
   * @param sheet
   *   The target sheet name
   * @param range
   *   The cell range within the target sheet
   *
   * Example: SheetRange("Sales", CellRange(A1, A10)) represents Sales!A1:A10
   */
  case SheetRange(sheet: SheetName, range: CellRange) extends TExpr[Nothing]

  // Arithmetic operators (form commutative semiring over BigDecimal)

  /**
   * Addition: x + y
   *
   * Laws:
   *   - Commutative: Add(x, y) = Add(y, x)
   *   - Associative: Add(Add(x, y), z) = Add(x, Add(y, z))
   *   - Identity: Add(x, Lit(0)) = x
   */
  case Add(x: TExpr[BigDecimal], y: TExpr[BigDecimal]) extends TExpr[BigDecimal]

  /**
   * Subtraction: x - y
   */
  case Sub(x: TExpr[BigDecimal], y: TExpr[BigDecimal]) extends TExpr[BigDecimal]

  /**
   * Multiplication: x * y
   *
   * Laws:
   *   - Commutative: Mul(x, y) = Mul(y, x)
   *   - Associative: Mul(Mul(x, y), z) = Mul(x, Mul(y, z))
   *   - Identity: Mul(x, Lit(1)) = x
   *   - Distributive: Mul(x, Add(y, z)) = Add(Mul(x, y), Mul(x, z))
   */
  case Mul(x: TExpr[BigDecimal], y: TExpr[BigDecimal]) extends TExpr[BigDecimal]

  /**
   * Division: x / y
   *
   * Note: Division by zero returns evaluation error (not parse error)
   */
  case Div(x: TExpr[BigDecimal], y: TExpr[BigDecimal]) extends TExpr[BigDecimal]

  /**
   * Exponentiation: x ^ y
   *
   * Note: Right-associative (2^3^2 = 2^(3^2) = 512) Excel convention: 0^0 = 1
   */
  case Pow(x: TExpr[BigDecimal], y: TExpr[BigDecimal]) extends TExpr[BigDecimal]

  // String operators

  /**
   * String concatenation: x & y
   *
   * Excel's & operator joins two strings together.
   *
   * Examples:
   *   - "Hello" & "World" → "HelloWorld"
   *   - A1 & B1 where A1="Hello", B1="World" → "HelloWorld"
   *   - A1 & " - " & B1 → "Hello - World"
   */
  case Concat(x: TExpr[String], y: TExpr[String]) extends TExpr[String]

  // Range aggregation

  // Comparison operators (return Boolean)

  /**
   * Equality: x == y
   */
  case Eq[A](x: TExpr[A], y: TExpr[A]) extends TExpr[Boolean]

  /**
   * Inequality: x != y
   */
  case Neq[A](x: TExpr[A], y: TExpr[A]) extends TExpr[Boolean]

  /**
   * Less than: x < y (numeric comparison)
   */
  case Lt(x: TExpr[BigDecimal], y: TExpr[BigDecimal]) extends TExpr[Boolean]

  /**
   * Less than or equal: x <= y
   */
  case Lte(x: TExpr[BigDecimal], y: TExpr[BigDecimal]) extends TExpr[Boolean]

  /**
   * Greater than: x > y
   */
  case Gt(x: TExpr[BigDecimal], y: TExpr[BigDecimal]) extends TExpr[Boolean]

  /**
   * Greater than or equal: x >= y
   */
  case Gte(x: TExpr[BigDecimal], y: TExpr[BigDecimal]) extends TExpr[Boolean]

  // Type conversions

  /**
   * Convert BigDecimal to Int.
   *
   * Required for DATE(YEAR(A1), MONTH(A1), DAY(A1)) pattern where date extraction functions return
   * BigDecimal but DATE expects Int arguments.
   *
   * Validates at evaluation time that value is a valid integer.
   *
   * Example: ToInt(YEAR(A1)) where YEAR returns BigDecimal(2025) → Int(2025)
   */
  case ToInt(expr: TExpr[BigDecimal]) extends TExpr[Int]

  /**
   * Convert LocalDate to Excel serial number.
   *
   * Excel represents dates as the number of days since December 30, 1899. This enables date
   * arithmetic like TODAY()+30.
   *
   * Example: DateToSerial(TODAY()) where TODAY() = 2025-12-20 → BigDecimal(46011)
   */
  case DateToSerial(expr: TExpr[java.time.LocalDate]) extends TExpr[BigDecimal]

  /**
   * Convert LocalDateTime to Excel serial number.
   *
   * Excel represents dates/times as days since epoch + fractional day for time. This enables
   * datetime arithmetic like NOW()+0.5.
   *
   * Example: DateTimeToSerial(NOW()) where NOW() = 2025-12-20 12:00 → BigDecimal(46011.5)
   */
  case DateTimeToSerial(expr: TExpr[java.time.LocalDateTime]) extends TExpr[BigDecimal]

  /**
   * Generic aggregate function over a range using the Aggregator typeclass.
   *
   * This case unifies all simple aggregate functions (SUM, COUNT, MIN, MAX, AVERAGE) into a single
   * representation. The aggregatorId determines which Aggregator instance to use for evaluation.
   *
   * Benefits:
   *   - Single pattern match case in Evaluator, FormulaPrinter, FormulaShifter, DependencyGraph
   *   - Adding new aggregates only requires a new Aggregator given instance
   *   - Automatically supports full row/column references via RangeLocation
   *
   * @param aggregatorId
   *   The name of the aggregator (SUM, COUNT, MIN, MAX, AVERAGE)
   * @param location
   *   The range to aggregate (local or cross-sheet)
   *
   * Example: Aggregate("SUM", RangeLocation.Local(A1:A10))
   */
  case Aggregate(aggregatorId: String, location: TExpr.RangeLocation) extends TExpr[BigDecimal]

  /**
   * Generic function call backed by FunctionSpec.
   *
   * Centralizes function definitions so parser, evaluator, printer, shifter, and dependency
   * analysis can share a single spec.
   */
  case Call[A](spec: FunctionSpec[A], args: spec.Args) extends TExpr[A]

  /**
   * GH-193: LET(name1, value1, [name2, value2, ...], calculation) — Excel 365 lexical bindings.
   *
   * Each binding's value expression is evaluated against the environment of all PRIOR bindings
   * (let* semantics); the body sees every binding. The Let expression's type is the body's type.
   *
   * Range-shaped binding values (RangeRef/SheetRange) are recorded here for canonical printing, but
   * the parser substitutes them directly into the body so range-typed argument positions (e.g.
   * SUMIF's criteria range) keep working — the evaluator never materializes them.
   *
   * Example: Let(List(("x", Lit(1))), Add(BindingRef("x"), Lit(1))) ≡ LET(x, 1, x+1)
   */
  case Let[A](bindings: List[(String, TExpr[?])], body: TExpr[A]) extends TExpr[A]

  /**
   * GH-193: Reference to an in-scope LET binding by its declared name.
   *
   * Existential/Any-typed like PolyRef: the runtime value comes from the evaluation environment,
   * and surrounding context coerces it. The parser only emits BindingRef for names it resolved
   * lexically, so an unknown name at evaluation time is a programming error.
   */
  case BindingRef(name: String) extends TExpr[Nothing]

object TExpr
    extends TExprRangeLocation
    with TExprConstructors
    with TExprAggregateOps
    with TExprFinancialOps
    with TExprLookupOps
    with TExprTypeCheckOps
    with TExprMathOps
    with TExprReferenceOps
    with TExprTextOps
    with TExprDateOps
    with TExprDecoders
    with TExprCoercions
    with TExprAnalysis
    with TExprExtensions
