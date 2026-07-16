package com.tjclp.xl.formula.eval

import com.tjclp.xl.addressing.{ARef, CellRange}
import com.tjclp.xl.cells.{CellError, CellValue}
import com.tjclp.xl.sheets.Sheet

/**
 * Array arithmetic operations with NumPy-style broadcasting.
 *
 * Broadcasting rules:
 *   - scalar * array -> element-wise
 *   - 1xN * MxN -> broadcast row across M rows
 *   - Mx1 * MxN -> broadcast column across N columns
 *   - MxN * MxN -> element-wise (dimensions must match)
 *
 * GH-337: errors propagate ELEMENTWISE through the broadcasts (compare, arithmetic, IF), matching
 * Excel: a carried `CellValue.Error` input element (left operand's error wins) or an element-local
 * failure (non-numeric text, division by zero) becomes an error OUTPUT element via
 * [[EvalError.toCellError]] instead of failing the whole formula. The broadcasts return Left only
 * for broadcast-dimension mismatch; aggregators decide whether consumed error elements propagate.
 */
object ArrayArithmetic:

  /** GH-196: Convert boolean to numeric (TRUE→1, FALSE→0). */
  def boolToNumeric(b: Boolean): BigDecimal = if b then BigDecimal(1) else BigDecimal(0)

  /** Sealed ADT for operand types in array arithmetic */
  sealed trait ArrayOperand
  object ArrayOperand:
    case class Scalar(value: BigDecimal) extends ArrayOperand
    case class Array(value: ArrayResult) extends ArrayOperand

  /** Binary operation type */
  type BinaryOp = (BigDecimal, BigDecimal) => Either[EvalError, BigDecimal]

  /** Safe division with zero check */
  def safeDivide(x: BigDecimal, y: BigDecimal): Either[EvalError, BigDecimal] =
    if y == BigDecimal(0) then Left(EvalError.DivByZero(x.toString, y.toString))
    else Right(x / y)

  /** Standard binary operations */
  val add: BinaryOp = (x, y) => Right(x + y)
  val sub: BinaryOp = (x, y) => Right(x - y)
  val mul: BinaryOp = (x, y) => Right(x * y)
  val div: BinaryOp = safeDivide

  /** Exponentiation with Excel conventions (0^0 = 1) */
  val pow: BinaryOp = (x, y) =>
    try
      if y.isValidInt && y >= 0 then
        // Exact precision for non-negative integer exponents
        Right(x.pow(y.toInt))
      else
        // Fall back to Double for fractional/negative exponents
        Right(BigDecimal(scala.math.pow(x.toDouble, y.toDouble)))
    catch
      case _: ArithmeticException => Left(EvalError.EvalFailed("Power overflow", None))
      case e: Exception => Left(EvalError.EvalFailed(s"Power failed: ${e.getMessage}", None))

  /**
   * Convert a CellValue to BigDecimal for arithmetic.
   *
   *   - Number -> value
   *   - Empty -> 0
   *   - Bool -> 1/0
   *   - Text containing number -> parsed value
   *   - Other -> error
   */
  def cellValueToNumeric(cv: CellValue): Either[EvalError, BigDecimal] = cv match
    case CellValue.Number(n) => Right(n)
    case CellValue.Empty => Right(BigDecimal(0))
    case CellValue.Bool(b) => Right(boolToNumeric(b))
    case CellValue.Formula(_, Some(cached)) => cellValueToNumeric(cached)
    case CellValue.Text(s) =>
      scala.util.Try(BigDecimal(s.trim)).toOption match
        case Some(n) => Right(n)
        case None => Left(EvalError.TypeMismatch("arithmetic", "number", s"text: $s"))
    case other =>
      Left(EvalError.TypeMismatch("arithmetic", "number", other.toString))

  /**
   * Convert a range to ArrayResult.
   */
  def rangeToArray(range: CellRange, sheet: Sheet): ArrayResult =
    val values = (range.rowStart.index0 to range.rowEnd.index0).map { rowIdx =>
      (range.colStart.index0 to range.colEnd.index0).map { colIdx =>
        val ref = ARef.from0(colIdx, rowIdx)
        val cell = sheet(ref)
        cell.value match
          case CellValue.Formula(_, Some(cached)) => cached
          case other => other
      }.toVector
    }.toVector
    ArrayResult(values)

  /**
   * Convert ArrayResult to numeric matrix.
   */
  def arrayToNumeric(arr: ArrayResult): Either[EvalError, Vector[Vector[BigDecimal]]] =
    traverseVV(arr.values)(cellValueToNumeric)

  /**
   * Perform broadcasting binary operation.
   *
   * GH-337: the array arms operate per CellValue element, carrying errors elementwise (see
   * [[applyOpCarrying]]) — only the scalar*scalar fast path keeps the whole-result Left for element
   * failures (scalar `=10/0` stays a formula error, matching Excel's scalar semantics).
   *
   * @param left
   *   Left operand (scalar or array)
   * @param right
   *   Right operand (scalar or array)
   * @param op
   *   Binary operation to apply element-wise
   * @return
   *   Either error or result (BigDecimal for scalar*scalar, ArrayResult otherwise)
   */
  def broadcast(
    left: ArrayOperand,
    right: ArrayOperand,
    op: BinaryOp
  ): Either[EvalError, ArrayOperand.Scalar | ArrayOperand.Array] =
    (left, right) match
      // scalar * scalar -> scalar (fast path)
      case (ArrayOperand.Scalar(l), ArrayOperand.Scalar(r)) =>
        op(l, r).map(ArrayOperand.Scalar(_))

      // scalar * array -> the scalar broadcasts as a 1x1 array (keeps operand order for op)
      case (ArrayOperand.Scalar(s), ArrayOperand.Array(arr)) =>
        broadcastArrays(ArrayResult.single(CellValue.Number(s)), arr, op).map(ArrayOperand.Array(_))

      // array * scalar -> the scalar broadcasts as a 1x1 array (keeps operand order for op)
      case (ArrayOperand.Array(arr), ArrayOperand.Scalar(s)) =>
        broadcastArrays(arr, ArrayResult.single(CellValue.Number(s)), op).map(ArrayOperand.Array(_))

      // array * array -> broadcast with dimension matching
      case (ArrayOperand.Array(l), ArrayOperand.Array(r)) =>
        broadcastArrays(l, r, op).map(ArrayOperand.Array(_))

  /**
   * GH-335: Excel's total order for the comparison operators, shared by the scalar and
   * array/broadcast paths so `A1<B1` and `(A1:A10<B1)` agree elementwise.
   *
   * Within a type: numbers compare numerically (dates ARE numbers via their Excel serial), text
   * case-insensitively and lexicographically, booleans FALSE < TRUE. Across types Excel ranks
   * number < text < logical — text never parses as a number under comparison (unlike arithmetic).
   * Empty cells coerce to the other operand's zero value (0 against numbers, "" against text, FALSE
   * against booleans; two empties are equal). GH-344: error values propagate with
   * Left(ErrorValue(code)), left operand first — on the SCALAR path this surfaces as the error
   * VALUE at the result boundary (`=1<#REF!` is #REF!, Excel-exact); the array broadcasts pre-check
   * both operands with [[carriedError]] and carry error elements through instead (GH-337), so this
   * arm fires solely for scalar comparisons like `=A1<B1` on an error cell.
   *
   * @return
   *   the comparison sign (negative, zero, positive), or Left for incomparable values
   */
  def compareCellValues(a: CellValue, b: CellValue): Either[EvalError, Int] =
    (normalizeForCompare(a), normalizeForCompare(b)) match
      // GH-344: an error OPERAND propagates as that error VALUE, left operand first (matching
      // equalityElement). The whole-formula refusal became Excel's absorption: `=1<#REF!` is
      // #REF! at the boundary. Array paths still pre-check carriedError and never reach here.
      case (CellValue.Error(err), _) =>
        Left(EvalError.ErrorValue(err, Some("comparison")))
      case (_, CellValue.Error(err)) =>
        Left(EvalError.ErrorValue(err, Some("comparison")))
      case (CellValue.Number(x), CellValue.Number(y)) => Right(x.compare(y))
      case (CellValue.Text(x), CellValue.Text(y)) => Right(x.compareToIgnoreCase(y))
      case (CellValue.Bool(x), CellValue.Bool(y)) => Right(java.lang.Boolean.compare(x, y))
      case (CellValue.Empty, CellValue.Empty) => Right(0)
      // Empty coerces to the other operand's zero value
      case (CellValue.Empty, CellValue.Number(y)) => Right(BigDecimal(0).compare(y))
      case (CellValue.Number(x), CellValue.Empty) => Right(x.compare(BigDecimal(0)))
      case (CellValue.Empty, CellValue.Text(y)) => Right("".compareToIgnoreCase(y))
      case (CellValue.Text(x), CellValue.Empty) => Right(x.compareToIgnoreCase(""))
      case (CellValue.Empty, CellValue.Bool(y)) => Right(java.lang.Boolean.compare(false, y))
      case (CellValue.Bool(x), CellValue.Empty) => Right(java.lang.Boolean.compare(x, false))
      // Cross-type: number < text < logical
      case (x, y) =>
        for
          rx <- typeRank(x)
          ry <- typeRank(y)
        yield Integer.compare(rx, ry)

  /** Excel's cross-type comparison rank: number < text < logical. */
  private def typeRank(cv: CellValue): Either[EvalError, Int] = cv match
    case CellValue.Number(_) => Right(0)
    case CellValue.Text(_) => Right(1)
    case CellValue.Bool(_) => Right(2)
    case other => Left(EvalError.TypeMismatch("comparison", "comparable value", other.toString))

  /**
   * Normalize a CellValue for comparison: cached formula values extract, dates become their Excel
   * serial number, rich text flattens to its plain text.
   */
  private def normalizeForCompare(cv: CellValue): CellValue = cv match
    case CellValue.Formula(_, Some(cached)) => normalizeForCompare(cached)
    case CellValue.DateTime(dt) =>
      CellValue.Number(BigDecimal(CellValue.dateTimeToExcelSerial(dt)))
    case CellValue.RichText(rt) => CellValue.Text(rt.toPlainText)
    case other => other

  /**
   * GH-335: broadcast an ordered comparison over two arrays elementwise (scalar operands wrap as
   * 1x1 arrays and broadcast like any other dimension-1 axis).
   *
   * GH-337: error elements carry through elementwise — a carried error on either operand (left
   * wins) becomes the OUTPUT element, and a residual per-element comparison refusal (e.g. an
   * uncomparable value) demotes to #VALUE!. Returns Left ONLY for broadcast-dimension mismatch.
   *
   * @param op
   *   interprets the comparison sign from [[compareCellValues]] (e.g. `_ < 0` for Lt)
   * @return
   *   ArrayResult of CellValue.Bool values with carried CellValue.Error elements
   */
  def broadcastOrderedCompare(
    left: ArrayResult,
    right: ArrayResult,
    op: Int => Boolean
  ): Either[EvalError, ArrayResult] =
    for
      outRows <- broadcastDim(left.rows, right.rows, "rows")
      outCols <- broadcastDim(left.cols, right.cols, "columns")
      result <- traverseVV(
        (0 until outRows).toVector.map { row =>
          (0 until outCols).toVector.map { col =>
            val lVal = getWithBroadcastCV(left.values, row, col, left.rows, left.cols)
            val rVal = getWithBroadcastCV(right.values, row, col, right.rows, right.cols)
            (lVal, rVal)
          }
        }
      ) { case (l, r) =>
        carriedError(l).orElse(carriedError(r)) match
          case Some(err) => Right(CellValue.Error(err))
          case None =>
            compareCellValues(l, r) match
              case Right(sign) => Right(CellValue.Bool(op(sign)))
              case Left(e) => toErrorElement(e)
      }
    yield ArrayResult(result)

  /**
   * GH-197: Element-wise equality/inequality comparison with broadcasting.
   *
   * Works on CellValue directly for polymorphic equality (strings, numbers, booleans). Used by
   * Eq/Neq operators.
   *
   * GH-337: a carried error on either operand (left wins) becomes the OUTPUT element — errors are
   * never "equal" or "unequal", they propagate (negate does not flip them). Returns Left ONLY for
   * broadcast-dimension mismatch.
   */
  def broadcastEqualityCompare(
    left: ArrayResult,
    right: Either[ArrayResult, Any],
    negate: Boolean
  ): Either[EvalError, ArrayResult] =
    right match
      case Right(scalar) =>
        // Array vs scalar - compare each element to the scalar
        val scalarCV = anyToCellValue(scalar)
        Right(ArrayResult(left.values.map(_.map(cv => equalityElement(cv, scalarCV, negate)))))
      case Left(rightArr) =>
        // Array vs array - broadcast
        for
          outRows <- broadcastDim(left.rows, rightArr.rows, "rows")
          outCols <- broadcastDim(left.cols, rightArr.cols, "columns")
        yield
          val result = (0 until outRows).toVector.map { row =>
            (0 until outCols).toVector.map { col =>
              val lVal = getWithBroadcastCV(left.values, row, col, left.rows, left.cols)
              val rVal = getWithBroadcastCV(rightArr.values, row, col, rightArr.rows, rightArr.cols)
              equalityElement(lVal, rVal, negate)
            }
          }
          ArrayResult(result)

  /** GH-337: one equality output element — carried errors win (left first), then Bool equality. */
  private def equalityElement(l: CellValue, r: CellValue, negate: Boolean): CellValue =
    carriedError(l).orElse(carriedError(r)) match
      case Some(err) => CellValue.Error(err)
      case None =>
        val eq = cellValueEquals(l, r)
        CellValue.Bool(if negate then !eq else eq)

  /** Get CellValue with broadcasting. */
  private def getWithBroadcastCV(
    arr: Vector[Vector[CellValue]],
    row: Int,
    col: Int,
    arrRows: Int,
    arrCols: Int
  ): CellValue =
    val r = if arrRows == 1 then 0 else row
    val c = if arrCols == 1 then 0 else col
    arr(r)(c)

  /**
   * GH-337: the Excel error an element carries, if any — a `CellValue.Error` directly or cached
   * inside a formula value. Package-visible so the aggregate error guards (FunctionSpecsAggregate)
   * share the exact matching semantics of the broadcasts.
   */
  private[formula] def carriedError(cv: CellValue): Option[CellError] = cv match
    case CellValue.Error(err) => Some(err)
    case CellValue.Formula(_, Some(cached)) => carriedError(cached)
    case _ => None

  /**
   * GH-337: demote an element-local failure to the error ELEMENT it carries, via the
   * [[EvalError.toCellError]] table. Failures with no element form (CircularRef — structural, never
   * element-local) stay a fatal Left.
   */
  private def toErrorElement(e: EvalError): Either[EvalError, CellValue] =
    EvalError.toCellError(e) match
      case Some(err) => Right(CellValue.Error(err))
      case None => Left(e)

  /**
   * GH-337: apply a numeric binary op to two elements, carrying errors elementwise: a carried error
   * on either operand wins (left first), and element-local failures (non-numeric text → #VALUE!,
   * division by zero → #DIV/0!) become error elements rather than failing the whole broadcast.
   * Accepted micro-divergence: pow overflow arrives as a generic EvalFailed and surfaces as #VALUE!
   * (Excel: #NUM!).
   */
  private def applyOpCarrying(
    l: CellValue,
    r: CellValue,
    op: BinaryOp
  ): Either[EvalError, CellValue] =
    carriedError(l).orElse(carriedError(r)) match
      case Some(err) => Right(CellValue.Error(err))
      case None =>
        val computed =
          for
            ln <- cellValueToNumeric(l)
            rn <- cellValueToNumeric(r)
            v <- op(ln, rn)
          yield CellValue.Number(v)
        computed match
          case Right(cv) => Right(cv)
          case Left(e) => toErrorElement(e)

  /**
   * GH-333: elementwise IF over an array condition (Excel CSE semantics).
   *
   * The condition broadcasts against both branch arrays (scalar branches arrive as 1x1 arrays and
   * repeat like any dimension-1 axis): out(i)(j) = if cond(i)(j) then ifTrue(i)(j) else
   * ifFalse(i)(j). This is the classic MIN(IF(...))/MAX(IF(...)) shape — the aggregate then folds
   * the resulting array.
   *
   * Condition elements follow Excel truthiness (booleans as-is, numbers zero/non-zero, empty is
   * FALSE). GH-337: a condition element carrying an error emits THAT error at its output position
   * (masking whatever the branches hold there); a text condition element demotes to a #VALUE!
   * element. Returns Left ONLY for broadcast-dimension mismatch.
   */
  def broadcastIf(
    cond: ArrayResult,
    ifTrue: ArrayResult,
    ifFalse: ArrayResult
  ): Either[EvalError, ArrayResult] =
    for
      rowsCT <- broadcastDim(cond.rows, ifTrue.rows, "rows")
      outRows <- broadcastDim(rowsCT, ifFalse.rows, "rows")
      colsCT <- broadcastDim(cond.cols, ifTrue.cols, "columns")
      outCols <- broadcastDim(colsCT, ifFalse.cols, "columns")
      result <- traverseVV(
        (0 until outRows).toVector.map { row =>
          (0 until outCols).toVector.map(col => (row, col))
        }
      ) { case (row, col) =>
        val condCV = getWithBroadcastCV(cond.values, row, col, cond.rows, cond.cols)
        carriedError(condCV) match
          case Some(err) => Right(CellValue.Error(err))
          case None =>
            conditionTruthy("IF condition", condCV) match
              case Right(truthy) =>
                Right(
                  if truthy then
                    getWithBroadcastCV(ifTrue.values, row, col, ifTrue.rows, ifTrue.cols)
                  else getWithBroadcastCV(ifFalse.values, row, col, ifFalse.rows, ifFalse.cols)
                )
              case Left(e) => toErrorElement(e)
      }
    yield ArrayResult(result)

  /**
   * GH-333/GH-338: Excel truthiness for a condition element, shared by broadcastIf (IF over an
   * array condition) and the logical functions' array paths (AND/OR aggregation, NOT broadcast):
   * booleans as-is, numbers zero/non-zero, empty is FALSE (the ScalarCoercion Bool conventions);
   * text refuses cleanly with a Left naming the position via `label`.
   *
   * GH-344: an error element propagates as its Excel error VALUE (Left(ErrorValue) — the AND/OR
   * folds short-circuit with it and the boundary promotes it to the error cell, Excel-exact).
   * broadcastIf and broadcastNot pre-check errors with [[carriedError]] and demote their residual
   * text refusals to #VALUE! elements, so their condition failures stay positional; the AND/OR
   * folds keep the text refusal loud (full #VALUE! parity arrives with the TypeMismatch boundary
   * demotion follow-up).
   */
  def conditionTruthy(label: String, cv: CellValue): Either[EvalError, Boolean] = cv match
    case CellValue.Bool(b) => Right(b)
    case CellValue.Number(n) => Right(n.signum != 0)
    case CellValue.Empty => Right(false)
    case CellValue.Formula(_, Some(cached)) => conditionTruthy(label, cached)
    case CellValue.Error(err) =>
      Left(EvalError.ErrorValue(err, Some(label)))
    case other => Left(EvalError.TypeMismatch(label, "boolean", other.toString))

  /**
   * GH-338: truthiness of every element (row-major), for the AND/OR aggregation folds (AND =
   * forall, OR = exists). The first refusing element short-circuits with its Left — GH-344: an
   * error element's Left carries its Excel error VALUE (the fold's result at the boundary), a text
   * element's stays a loud TypeMismatch.
   */
  def truthyElements(label: String, arr: ArrayResult): Either[EvalError, Vector[Boolean]] =
    traverseV(arr.values.flatten)(cv => conditionTruthy(label, cv))

  /**
   * GH-338: elementwise NOT over an array condition, preserving shape (Excel broadcasts NOT rather
   * than aggregating). GH-344: mirrors broadcastIf's positional carriage — an element carrying an
   * error emits THAT error at its output position, and a residual refusal (text) demotes to a
   * #VALUE! element. Total in Right.
   */
  def broadcastNot(label: String, arr: ArrayResult): Either[EvalError, ArrayResult] =
    traverseVV(arr.values) { cv =>
      carriedError(cv) match
        case Some(err) => Right(CellValue.Error(err))
        case None =>
          conditionTruthy(label, cv) match
            case Right(b) => Right(CellValue.Bool(!b))
            case Left(e) => toErrorElement(e)
    }.map(ArrayResult(_))

  /** Convert any value to CellValue for comparison. */
  def anyToCellValue(v: Any): CellValue = v match
    case cv: CellValue => cv
    case s: String => CellValue.Text(s)
    case n: BigDecimal => CellValue.Number(n)
    case n: Int => CellValue.Number(BigDecimal(n))
    case n: Long => CellValue.Number(BigDecimal(n))
    case n: Double => CellValue.Number(BigDecimal(n))
    case b: Boolean => CellValue.Bool(b)
    // GH-335: date-returning calls (TODAY, DATE, ...) compare as dates, not as their toString
    case ld: java.time.LocalDate => CellValue.DateTime(ld.atStartOfDay())
    case ldt: java.time.LocalDateTime => CellValue.DateTime(ldt)
    case _ => CellValue.Text(v.toString)

  /**
   * Compare CellValues for equality (case-insensitive for text), matching Excel.
   *
   * Shared with the scalar equality fast path in Evaluator (GH-234) so scalar `=A1=B1` and array
   * `=A1:A3=B1:B3` equality use identical semantics. GH-335: defined via [[compareCellValues]] so
   * = and < agree — empty cells equal 0 / "" / FALSE, dates equal their serial number, and
   * cross-type values (which never coerce under comparison) are simply unequal. Error values
   * equal nothing on this scalar path; the equality broadcast pre-checks them with
   * [[carriedError]] and carries error elements instead (GH-337).
   */
  def cellValueEquals(a: CellValue, b: CellValue): Boolean =
    compareCellValues(a, b).fold(_ => false, _ == 0)

  /**
   * Broadcast two arrays together, operating per CellValue element (GH-337: errors carry
   * elementwise via [[applyOpCarrying]]; Left ONLY for broadcast-dimension mismatch).
   */
  private def broadcastArrays(
    left: ArrayResult,
    right: ArrayResult,
    op: BinaryOp
  ): Either[EvalError, ArrayResult] =
    // Determine output dimensions
    for
      outRows <- broadcastDim(left.rows, right.rows, "rows")
      outCols <- broadcastDim(left.cols, right.cols, "columns")
      result <- traverseVV(
        (0 until outRows).toVector.map { row =>
          (0 until outCols).toVector.map { col =>
            val lVal = getWithBroadcastCV(left.values, row, col, left.rows, left.cols)
            val rVal = getWithBroadcastCV(right.values, row, col, right.rows, right.cols)
            (lVal, rVal)
          }
        }
      ) { case (l, r) => applyOpCarrying(l, r, op) }
    yield ArrayResult(result)

  /**
   * Compute broadcast output dimension for a single axis.
   */
  private def broadcastDim(l: Int, r: Int, dimName: String): Either[EvalError, Int] =
    if l == r then Right(l)
    else if l == 1 then Right(r)
    else if r == 1 then Right(l)
    else
      Left(
        EvalError.EvalFailed(
          s"Cannot broadcast arrays: $dimName mismatch ($l vs $r). Broadcasting requires dimensions to match or be 1.",
          None
        )
      )

  // ===== Helper: traverse for Vector[Vector[A]] =====
  // We don't have Cats, so implement manually

  private def traverseV[A, E, B](vec: Vector[A])(f: A => Either[E, B]): Either[E, Vector[B]] =
    vec.foldLeft[Either[E, Vector[B]]](Right(Vector.empty)) { (acc, a) =>
      acc.flatMap(bs => f(a).map(b => bs :+ b))
    }

  private def traverseVV[A, E, B](
    vv: Vector[Vector[A]]
  )(f: A => Either[E, B]): Either[E, Vector[Vector[B]]] =
    traverseV(vv)(row => traverseV(row)(f))
