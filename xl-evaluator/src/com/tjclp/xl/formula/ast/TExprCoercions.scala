package com.tjclp.xl.formula.ast

import com.tjclp.xl.formula.functions.FunctionSpecs
import com.tjclp.xl.formula.eval.EvalError
import com.tjclp.xl.formula.functions.EvalContext

import com.tjclp.xl.cells.CellValue
import TExpr.*

trait TExprCoercions:
  // ===== PolyRef Conversion Helpers =====

  /**
   * GH-302/GH-306: shapes whose RUNTIME value is not pinned by their static type — function calls
   * (Any/CellValue/ArrayResult-returning), aggregates, LET bodies, operators of a different static
   * type, and already-coerced nodes. In a typed argument position these must coerce at evaluation
   * time (TExpr.Coerced) rather than be erased-cast — the cast defers a ClassCastException to the
   * consuming function. Per-target helpers below subtract the shapes that are statically SAFE for
   * that target (e.g. arithmetic in a numeric position) to preserve existing tree shapes.
   */
  private def isRuntimePolymorphic(expr: TExpr[?]): Boolean = expr match
    case _: TExpr.Call[?] | _: TExpr.Let[?] | _: TExpr.Aggregate | _: TExpr.Coerced[?] => true
    case _: TExpr.Add | _: TExpr.Sub | _: TExpr.Mul | _: TExpr.Div | _: TExpr.Pow => true
    case _: TExpr.Concat => true
    case _: TExpr.Eq[?] | _: TExpr.Neq[?] | _: TExpr.Lt | _: TExpr.Lte | _: TExpr.Gt |
        _: TExpr.Gte =>
      true
    case _: TExpr.ToInt | _: TExpr.DateToSerial | _: TExpr.DateTimeToSerial => true
    case _ => false

  @SuppressWarnings(Array("org.wartremover.warts.AsInstanceOf"))
  private def coerced[A](expr: TExpr[?], target: BindingCoercion): TExpr[A] =
    Coerced[A](expr.asInstanceOf[TExpr[Any]], target)

  /**
   * Convert any TExpr to String type with coercion.
   *
   * Used by text functions (LEFT, RIGHT, UPPER, etc.) to handle PolyRef arguments.
   */
  @SuppressWarnings(Array("org.wartremover.warts.AsInstanceOf"))
  def asStringExpr(expr: TExpr[?]): TExpr[String] = expr match
    case PolyRef(at, anchor) => Ref(at, anchor, decodeAsString)
    case SheetPolyRef(sheet, at, anchor) => SheetRef(sheet, at, anchor, decodeAsString)
    // GH-193: LET bindings are Any-typed — coerce totally at evaluation time
    case BindingRef(name) => CoercedBindingRef[String](name, BindingCoercion.Text)
    case TExpr.Lit(value: String) => TExpr.Lit(value)
    case TExpr.Lit(value: BigDecimal) => TExpr.Lit(value.toString)
    case TExpr.Lit(value: Boolean) => TExpr.Lit(if value then "TRUE" else "FALSE")
    case TExpr.Lit(value: java.time.LocalDate) => TExpr.Lit(value.toString)
    case TExpr.Lit(value: java.time.LocalDateTime) => TExpr.Lit(value.toString)
    // Concat is String by construction — the only statically text-typed operator
    case c: TExpr.Concat => c
    // GH-302/GH-306: numeric/boolean/array call results render as text at evaluation time
    // (=UPPER(SUM(A1:A2)) → "30", =LEFT(INDIRECT("B1"), 2) collapses then renders)
    case other if isRuntimePolymorphic(other) => coerced[String](other, BindingCoercion.Text)
    case other => other.asInstanceOf[TExpr[String]] // Safe: non-PolyRef already has correct type

  /**
   * Convert any TExpr to LocalDate type with coercion.
   *
   * Used by date functions (YEAR, MONTH, DAY) to handle PolyRef arguments.
   */
  @SuppressWarnings(Array("org.wartremover.warts.AsInstanceOf"))
  def asDateExpr(expr: TExpr[?]): TExpr[java.time.LocalDate] = expr match
    case PolyRef(at, anchor) => Ref(at, anchor, decodeAsDate)
    case SheetPolyRef(sheet, at, anchor) => SheetRef(sheet, at, anchor, decodeAsDate)
    // GH-193: LET bindings are Any-typed — coerce totally at evaluation time (bound dates from
    // cells are stored as Excel serial numbers, which the Date target converts back)
    case BindingRef(name) => CoercedBindingRef[java.time.LocalDate](name, BindingCoercion.Date)
    // Date-returning calls (TODAY, DATE, EDATE, ...) already produce LocalDate
    case call: TExpr.Call[?] if call.spec.flags.returnsDate =>
      call.asInstanceOf[TExpr[java.time.LocalDate]]
    // GH-306: time-returning calls (NOW), serial arithmetic (TODAY()+1), numeric calls — coerce
    // at evaluation time (LocalDateTime → date, Excel serial → date)
    case other if isRuntimePolymorphic(other) =>
      coerced[java.time.LocalDate](other, BindingCoercion.Date)
    case other =>
      other.asInstanceOf[TExpr[java.time.LocalDate]] // Safe: non-PolyRef already has correct type

  /**
   * Convert any TExpr to Int type with coercion.
   *
   * Used by functions requiring integer arguments (LEFT, RIGHT, DATE). Automatically converts
   * BigDecimal expressions (like YEAR/MONTH/DAY) to Int.
   */
  @SuppressWarnings(Array("org.wartremover.warts.AsInstanceOf"))
  def asIntExpr(expr: TExpr[?]): TExpr[Int] = expr match
    case PolyRef(at, anchor) => Ref(at, anchor, decodeAsInt)
    case SheetPolyRef(sheet, at, anchor) => SheetRef(sheet, at, anchor, decodeAsInt)
    // GH-193: LET bindings are Any-typed — coerce totally at evaluation time
    case BindingRef(name) => CoercedBindingRef[Int](name, BindingCoercion.Integer)
    case TExpr.Lit(bd: BigDecimal) if bd.isValidInt => TExpr.Lit(bd.toInt)
    // Any function call returning BigDecimal (flagged via returnsNumeric) — wrap in ToInt.
    // Covers SUM, COUNT, AVERAGE, ROUND, ABS, MOD, ROW, COLUMN, MATCH, PMT, FIND, LEN,
    // YEAR/MONTH/DAY, etc. — every numeric-returning function in the registry.
    case call: TExpr.Call[?] if call.spec.flags.returnsNumeric =>
      ToInt(call.asInstanceOf[TExpr[BigDecimal]])
    // Arithmetic expressions return BigDecimal — wrap in ToInt to avoid
    // a runtime ClassCastException when used in Int-arg positions
    // (e.g. =MID(A1, FIND("@", A1) + 1, 100)).
    case _: TExpr.Add | _: TExpr.Sub | _: TExpr.Mul | _: TExpr.Div | _: TExpr.Pow =>
      ToInt(expr.asInstanceOf[TExpr[BigDecimal]])
    case agg: TExpr.Aggregate => ToInt(agg)
    // GH-302/GH-306: non-numeric call results (text, boolean, arrays, IF branches) coerce at
    // evaluation time — fractionals truncate, numeric text parses, "abc" is a clean error
    case other if isRuntimePolymorphic(other) => coerced[Int](other, BindingCoercion.Integer)
    case other => other.asInstanceOf[TExpr[Int]] // Safe: non-PolyRef already has correct type

  /**
   * Convert any TExpr to BigDecimal type (numeric).
   *
   * Used by arithmetic functions to handle PolyRef arguments.
   */
  @SuppressWarnings(Array("org.wartremover.warts.AsInstanceOf"))
  def asNumericExpr(expr: TExpr[?]): TExpr[BigDecimal] = expr match
    case PolyRef(at, anchor) => Ref(at, anchor, decodeNumeric)
    case SheetPolyRef(sheet, at, anchor) => SheetRef(sheet, at, anchor, decodeNumeric)
    // GH-193: LET bindings are Any-typed — coerce totally at evaluation time
    case BindingRef(name) => CoercedBindingRef[BigDecimal](name, BindingCoercion.Numeric)
    // Date functions return LocalDate/LocalDateTime - convert to Excel serial number
    case call: TExpr.Call[?] if call.spec.flags.returnsDate =>
      DateToSerial(call.asInstanceOf[TExpr[java.time.LocalDate]])
    case call: TExpr.Call[?] if call.spec.flags.returnsTime =>
      DateTimeToSerial(call.asInstanceOf[TExpr[java.time.LocalDateTime]])
    // Numeric-returning calls, arithmetic, aggregates and serial conversions are already
    // BigDecimal — keep their shape
    case call: TExpr.Call[?] if call.spec.flags.returnsNumeric =>
      call.asInstanceOf[TExpr[BigDecimal]]
    case _: TExpr.Add | _: TExpr.Sub | _: TExpr.Mul | _: TExpr.Div | _: TExpr.Pow |
        _: TExpr.Aggregate | _: TExpr.DateToSerial | _: TExpr.DateTimeToSerial =>
      expr.asInstanceOf[TExpr[BigDecimal]]
    // GH-302/GH-306: text/boolean/array call results coerce at evaluation time (numeric text
    // parses, booleans are 1/0, 1×1 arrays collapse; arrays still broadcast in operand positions)
    case other if isRuntimePolymorphic(other) =>
      coerced[BigDecimal](other, BindingCoercion.Numeric)
    case other =>
      other.asInstanceOf[TExpr[BigDecimal]] // Safe: non-PolyRef already has correct type

  /**
   * Convert any TExpr to BigDecimal type, preserving RangeRef for array arithmetic.
   *
   * Unlike asNumericExpr which converts all expressions, this preserves RangeRef and SheetRange so
   * the evaluator can convert them to ArrayResult for array arithmetic with broadcasting.
   *
   * Used by binary arithmetic operators (+, -, *, /) to support array formulas.
   */
  @SuppressWarnings(Array("org.wartremover.warts.AsInstanceOf"))
  def asNumericOrRangeExpr(expr: TExpr[?]): TExpr[BigDecimal] = expr match
    case r: TExpr.RangeRef => r.asInstanceOf[TExpr[BigDecimal]] // Preserve for array eval
    case sr: TExpr.SheetRange => sr.asInstanceOf[TExpr[BigDecimal]] // Preserve for array eval
    case other => asNumericExpr(other)

  /**
   * Convert any TExpr to Boolean type.
   *
   * Used by logical functions (AND, OR, NOT, IF) to handle PolyRef arguments.
   */
  @SuppressWarnings(Array("org.wartremover.warts.AsInstanceOf"))
  def asBooleanExpr(expr: TExpr[?]): TExpr[Boolean] = expr match
    case PolyRef(at, anchor) => Ref(at, anchor, decodeBool)
    case SheetPolyRef(sheet, at, anchor) => SheetRef(sheet, at, anchor, decodeBool)
    // GH-193: LET bindings are Any-typed — coerce totally at evaluation time
    case BindingRef(name) => CoercedBindingRef[Boolean](name, BindingCoercion.Bool)
    // Comparisons are the statically boolean-typed operators — keep their shape
    case _: TExpr.Eq[?] | _: TExpr.Neq[?] | _: TExpr.Lt | _: TExpr.Lte | _: TExpr.Gt |
        _: TExpr.Gte =>
      expr.asInstanceOf[TExpr[Boolean]]
    // GH-306: numeric call results use Excel truthiness (=IF(SUM(A1:A2), 1, 2), 0 = FALSE);
    // uncoercible results (text) are a clean error instead of a ClassCastException
    case other if isRuntimePolymorphic(other) => coerced[Boolean](other, BindingCoercion.Bool)
    case other => other.asInstanceOf[TExpr[Boolean]] // Safe: non-PolyRef already has correct type

  /**
   * Convert any TExpr to CellValue type.
   *
   * Used by error handling functions (IFERROR, ISERROR) to preserve raw cell values. No Coerced
   * wrapping: CellValue positions are Any-tolerant (consumers go through ExprValue.from), and the
   * scalar argument boundary collapses ArrayResults to their top-left value (GH-302).
   */
  @SuppressWarnings(Array("org.wartremover.warts.AsInstanceOf"))
  def asCellValueExpr(expr: TExpr[?]): TExpr[CellValue] = expr match
    case PolyRef(at, anchor) => Ref(at, anchor, decodeCellValue)
    case SheetPolyRef(sheet, at, anchor) => SheetRef(sheet, at, anchor, decodeCellValue)
    case other =>
      other.asInstanceOf[TExpr[CellValue]] // Safe: non-PolyRef already has correct type

  /**
   * Convert any TExpr to resolved CellValue type.
   *
   * Used for standalone cell references (e.g., =A1, =Sheet1!B2) where we need the cell's
   * "effective" value with cached formula results extracted and empty cells converted to 0.
   */
  @SuppressWarnings(Array("org.wartremover.warts.AsInstanceOf"))
  def asResolvedValueExpr(expr: TExpr[?]): TExpr[CellValue] = expr match
    case PolyRef(at, anchor) => Ref(at, anchor, decodeResolvedValue)
    case SheetPolyRef(sheet, at, anchor) => SheetRef(sheet, at, anchor, decodeResolvedValue)
    case other =>
      other.asInstanceOf[TExpr[CellValue]] // Safe: non-PolyRef already has correct type
