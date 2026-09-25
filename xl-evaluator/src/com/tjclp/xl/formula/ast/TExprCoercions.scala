package com.tjclp.xl.formula.ast

import com.tjclp.xl.formula.functions.FunctionSpecs
import com.tjclp.xl.formula.eval.{EvalError, ScalarCoercion}
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
    // GH-384: a defined name's runtime value comes from the workbook's name table — coerce it
    // at evaluation time like a call result (=EOMONTH(named_date, 0), =IF(case=2, ...))
    case _: TExpr.NameRef => true
    // GH-394: sheet-qualified names likewise (=EOMONTH(Model!named_date, 0))
    case _: TExpr.SheetNameRef => true
    // a range in a scalar slot is a value position: a plain cell reads the implicitly intersected
    // cell, which must coerce to the slot's type (=CHOOSE(D1:D10, ...), =OFFSET(A1, B1:B5, 0))
    case _: TExpr.RangeRef | _: TExpr.SheetRange => true
    case _: TExpr.Add | _: TExpr.Sub | _: TExpr.Mul | _: TExpr.Div | _: TExpr.Pow |
        _: TExpr.Percent =>
      true
    case _: TExpr.Concat => true
    case _: TExpr.Eq[?] | _: TExpr.Neq[?] | _: TExpr.Lt[?] | _: TExpr.Lte[?] | _: TExpr.Gt[?] |
        _: TExpr.Gte[?] =>
      true
    case _: TExpr.ToInt | _: TExpr.DateToSerial | _: TExpr.DateTimeToSerial => true
    // GH-603: an omitted argument evaluates to Empty and coerces per target like a blank cell
    // (0 / "" / FALSE / the blank date) — `LEFT("abc",)` is "", `RATE(10,,-100,150)` reads pmt 0
    case TExpr.Missing => true
    case _ => false

  /**
   * The static scalar kind of a node's value — what an array it yields in array mode must collapse
   * through at a scalar position (ScalarCoercion.collapseTo). The array-mode typing invariant: any
   * `TExpr[A]` may evaluate to an ArrayResult in array mode, and every consumer either handles the
   * array or turns it into a value of its position's type or a Left, never a mistyped value.
   *
   * Complete by construction of the as*Expr tables below: a statically typed node survives BARE in
   * a typed slot only as arithmetic, Aggregate, date-serial nodes and returnsNumeric calls in
   * numeric slots, returnsDate calls in date slots, ToInt in integer slots and Concat in text
   * slots; everything else is wrapped in Coerced (isRuntimePolymorphic). None marks nodes whose
   * value is not pinned by a scalar kind — they reach only Any/CellValue slots, tolerant by design.
   * No wildcard arm: a new TExpr case must decide its kind here.
   */
  private[formula] def scalarKind(expr: TExpr[?]): Option[BindingCoercion] = expr match
    case TExpr.Add(_, _) | TExpr.Sub(_, _) | TExpr.Mul(_, _) | TExpr.Div(_, _) | TExpr.Pow(_, _) |
        TExpr.Percent(_) =>
      Some(BindingCoercion.Numeric)
    case TExpr.Aggregate(_, _) | TExpr.DateToSerial(_) | TExpr.DateTimeToSerial(_) =>
      Some(BindingCoercion.Numeric)
    case TExpr.ToInt(_) => Some(BindingCoercion.Integer)
    case TExpr.Concat(_, _) => Some(BindingCoercion.Text)
    case TExpr.Eq(_, _) | TExpr.Neq(_, _) | TExpr.Lt(_, _) | TExpr.Lte(_, _) | TExpr.Gt(_, _) |
        TExpr.Gte(_, _) =>
      Some(BindingCoercion.Bool)
    case TExpr.Coerced(_, target) => Some(target)
    case TExpr.CoercedBindingRef(_, target) => Some(target)
    case TExpr.UnaryPlus(inner) => scalarKind(inner)
    case TExpr.Call(spec, _) =>
      if spec.flags.returnsNumeric then Some(BindingCoercion.Numeric)
      else if spec.flags.returnsDate then Some(BindingCoercion.Date)
      else None
    case TExpr.Lit(_) | TExpr.Ref(_, _, _) | TExpr.PolyRef(_, _) | TExpr.SheetRef(_, _, _, _) |
        TExpr.SheetPolyRef(_, _, _) | TExpr.ExternalRef(_, _, _, _) |
        TExpr.ExternalRange(_, _, _, _) | TExpr.RangeRef(_, _) | TExpr.SheetRange(_, _, _) |
        TExpr.ErrorLit(_) | TExpr.Missing | TExpr.Let(_, _) | TExpr.BindingRef(_) |
        TExpr.NameRef(_) | TExpr.SheetNameRef(_, _) =>
      None

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
    // GH-374: unary plus is a type-preserving transparent wrapper — push the coercion through
    // so wrapped PolyRefs still resolve (and the node survives for byte-faithful printing)
    case UnaryPlus(inner) => UnaryPlus(asStringExpr(inner))
    case SheetPolyRef(sheet, at, anchor) => SheetRef(sheet, at, anchor, decodeAsString)
    // GH-193: LET bindings are Any-typed — coerce totally at evaluation time
    case BindingRef(name) => CoercedBindingRef[String](name, BindingCoercion.Text)
    case TExpr.Lit(value: String) => TExpr.Lit(value)
    // GH-665: a numeric literal in a text position renders as Excel's General text at evaluation
    // time (=2.50&"" is "2.5", =LEN(2.50) is 3) and SURVIVES in the AST, so the printer emits
    // `=2.50&""` back rather than the folded `="2.50"&""`
    case TExpr.Lit(_: BigDecimal) => coerced[String](expr, BindingCoercion.Text)
    case TExpr.Lit(value: Boolean) => TExpr.Lit(if value then "TRUE" else "FALSE")
    // GH-561: a date in a text position is its Excel serial, not ISO text
    case TExpr.Lit(value: java.time.LocalDate) => TExpr.Lit(ScalarCoercion.dateSerialText(value))
    case TExpr.Lit(value: java.time.LocalDateTime) =>
      TExpr.Lit(ScalarCoercion.dateSerialText(value))
    // Any other literal (a programmatic AST: a CellValue, an Int, an array) coerces at evaluation
    // time rather than being cast to a String it is not
    case lit: TExpr.Lit[?] => coerced[String](lit, BindingCoercion.Text)
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
    // GH-374: push through the transparent unary-plus wrapper (see asStringExpr)
    case UnaryPlus(inner) => UnaryPlus(asDateExpr(inner))
    case SheetPolyRef(sheet, at, anchor) => SheetRef(sheet, at, anchor, decodeAsDate)
    // GH-193: LET bindings are Any-typed — coerce totally at evaluation time (bound dates from
    // cells are stored as Excel serial numbers, which the Date target converts back)
    case BindingRef(name) => CoercedBindingRef[java.time.LocalDate](name, BindingCoercion.Date)
    // Date-returning calls (TODAY, DATE, EDATE, ...) already produce LocalDate
    case call: TExpr.Call[?] if call.spec.flags.returnsDate =>
      call.asInstanceOf[TExpr[java.time.LocalDate]]
    // GH-307: cross-typed literals coerce at evaluation time (Excel serial → date; booleans
    // are their serial like everywhere else in the table, so =YEAR(TRUE) is 1900 like Excel —
    // previously Lit(Boolean) fell through to the erased cast and threw ClassCastException)
    case lit: TExpr.Lit[?] => coerced[java.time.LocalDate](lit, BindingCoercion.Date)
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
    // GH-374: push through the transparent unary-plus wrapper (see asStringExpr)
    case UnaryPlus(inner) => UnaryPlus(asIntExpr(inner))
    case SheetPolyRef(sheet, at, anchor) => SheetRef(sheet, at, anchor, decodeAsInt)
    // GH-193: LET bindings are Any-typed — coerce totally at evaluation time
    case BindingRef(name) => CoercedBindingRef[Int](name, BindingCoercion.Integer)
    case TExpr.Lit(bd: BigDecimal) if bd.isValidInt => TExpr.Lit(bd.toInt)
    // GH-307: cross-typed literals coerce at evaluation time — fractionals truncate like
    // Excel (LEFT("hello", 2.7) → "he"), numeric text parses, booleans are 1/0, anything
    // else is a clean error (previously an erased cast deferring a ClassCastException)
    case lit: TExpr.Lit[?] => coerced[Int](lit, BindingCoercion.Integer)
    // Any function call returning BigDecimal (flagged via returnsNumeric) — wrap in ToInt.
    // Covers SUM, COUNT, AVERAGE, ROUND, ABS, MOD, ROW, COLUMN, MATCH, PMT, FIND, LEN,
    // YEAR/MONTH/DAY, etc. — every numeric-returning function in the registry.
    case call: TExpr.Call[?] if call.spec.flags.returnsNumeric =>
      ToInt(call.asInstanceOf[TExpr[BigDecimal]])
    // Arithmetic expressions return BigDecimal — wrap in ToInt to avoid
    // a runtime ClassCastException when used in Int-arg positions
    // (e.g. =MID(A1, FIND("@", A1) + 1, 100)).
    case _: TExpr.Add | _: TExpr.Sub | _: TExpr.Mul | _: TExpr.Div | _: TExpr.Pow |
        _: TExpr.Percent =>
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
    // GH-385: scalar numeric positions treat a blank cell as 0 (=A1+1 with blank A1 is 1, like
    // Excel); range FOLDS keep strict decodeNumeric (TExpr.Aggregate, cashflow collection) so
    // MIN/MEDIAN skip blanks and NPV periods don't shift
    case PolyRef(at, anchor) => Ref(at, anchor, decodeNumericScalar)
    // GH-374: push through the transparent unary-plus wrapper (see asStringExpr)
    case UnaryPlus(inner) => UnaryPlus(asNumericExpr(inner))
    case SheetPolyRef(sheet, at, anchor) => SheetRef(sheet, at, anchor, decodeNumericScalar)
    // GH-193: LET bindings are Any-typed — coerce totally at evaluation time
    case BindingRef(name) => CoercedBindingRef[BigDecimal](name, BindingCoercion.Numeric)
    // GH-307: cross-typed literals coerce at evaluation time (=SQRT("16") → 4, TRUE → 1, and a
    // programmatic date, CellValue or array literal likewise); numeric literals keep their shape
    case TExpr.Lit(_: BigDecimal) => expr.asInstanceOf[TExpr[BigDecimal]]
    case lit: TExpr.Lit[?] => coerced[BigDecimal](lit, BindingCoercion.Numeric)
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
        _: TExpr.Percent | _: TExpr.Aggregate | _: TExpr.DateToSerial | _: TExpr.DateTimeToSerial =>
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
    // GH-374: push through the transparent unary-plus wrapper so a wrapped range stays
    // range-shaped for array arithmetic (=+A1:A3*10 broadcasts)
    case UnaryPlus(inner) => UnaryPlus(asNumericOrRangeExpr(inner))
    case other => asNumericExpr(other)

  /**
   * Convert any TExpr to Boolean type.
   *
   * Used by logical functions (AND, OR, NOT, IF) to handle PolyRef arguments.
   */
  @SuppressWarnings(Array("org.wartremover.warts.AsInstanceOf"))
  def asBooleanExpr(expr: TExpr[?]): TExpr[Boolean] = expr match
    case PolyRef(at, anchor) => Ref(at, anchor, decodeBool)
    // GH-374: push through the transparent unary-plus wrapper (see asStringExpr)
    case UnaryPlus(inner) => UnaryPlus(asBooleanExpr(inner))
    case SheetPolyRef(sheet, at, anchor) => SheetRef(sheet, at, anchor, decodeBool)
    // GH-193: LET bindings are Any-typed — coerce totally at evaluation time
    case BindingRef(name) => CoercedBindingRef[Boolean](name, BindingCoercion.Bool)
    // GH-307: cross-typed literals coerce at evaluation time (=IF(1, ...) uses Excel
    // truthiness; text is a clean error; a programmatic CellValue or array literal likewise);
    // boolean literals keep their shape
    case TExpr.Lit(_: Boolean) => expr.asInstanceOf[TExpr[Boolean]]
    case lit: TExpr.Lit[?] => coerced[Boolean](lit, BindingCoercion.Bool)
    // GH-333: comparisons are statically Boolean but runtime-polymorphic — over a range operand
    // in array mode they yield an ArrayResult of booleans. Coerce at evaluation time so scalar
    // boolean positions (IFS/AND/OR/NOT conditions) collapse-and-coerce totally instead of
    // deferring a ClassCastException; array/operand positions still pass the array through.
    // (No dedicated case: isRuntimePolymorphic already includes the comparison operators.)
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
    // GH-374: push through the transparent unary-plus wrapper (see asStringExpr)
    case UnaryPlus(inner) => UnaryPlus(asCellValueExpr(inner))
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
    // GH-374: push through the transparent unary-plus wrapper (see asStringExpr)
    case UnaryPlus(inner) => UnaryPlus(asResolvedValueExpr(inner))
    case SheetPolyRef(sheet, at, anchor) => SheetRef(sheet, at, anchor, decodeResolvedValue)
    case other =>
      other.asInstanceOf[TExpr[CellValue]] // Safe: non-PolyRef already has correct type

  /**
   * Convert any TExpr to a comparison operand (GH-335).
   *
   * Like asResolvedValueExpr but cell references PRESERVE emptiness (decodeComparableValue) so
   * ArrayArithmetic.compareCellValues can coerce an empty cell relative to the other operand (0 vs
   * numbers, "" vs text, FALSE vs booleans) — the Excel comparison semantics for = <> < <= > >=.
   */
  @SuppressWarnings(Array("org.wartremover.warts.AsInstanceOf"))
  def asComparableValueExpr(expr: TExpr[?]): TExpr[CellValue] = expr match
    case PolyRef(at, anchor) => Ref(at, anchor, decodeComparableValue)
    // GH-374: push through the transparent unary-plus wrapper (see asStringExpr)
    case UnaryPlus(inner) => UnaryPlus(asComparableValueExpr(inner))
    case SheetPolyRef(sheet, at, anchor) => SheetRef(sheet, at, anchor, decodeComparableValue)
    case other =>
      other.asInstanceOf[TExpr[CellValue]] // Safe: non-PolyRef already has correct type
