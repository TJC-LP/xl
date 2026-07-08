package com.tjclp.xl.formula.eval

import com.tjclp.xl.addressing.{ARef, CellRange}
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.sheets.Sheet

/**
 * Array arithmetic operations with NumPy-style broadcasting.
 *
 * Broadcasting rules:
 *   - scalar * array -> element-wise
 *   - 1xN * MxN -> broadcast row across M rows
 *   - Mx1 * MxN -> broadcast column across N columns
 *   - MxN * MxN -> element-wise (dimensions must match)
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

      // scalar * array -> broadcast scalar to all elements
      case (ArrayOperand.Scalar(s), ArrayOperand.Array(arr)) =>
        arrayToNumeric(arr).flatMap { nums =>
          traverseVV(nums)(v => op(s, v)).map(toArrayResult).map(ArrayOperand.Array(_))
        }

      // array * scalar -> broadcast scalar to all elements
      case (ArrayOperand.Array(arr), ArrayOperand.Scalar(s)) =>
        arrayToNumeric(arr).flatMap { nums =>
          traverseVV(nums)(v => op(v, s)).map(toArrayResult).map(ArrayOperand.Array(_))
        }

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
   * against booleans; two empties are equal). Error values refuse to compare with a clean Left
   * naming the Excel error code.
   *
   * @return
   *   the comparison sign (negative, zero, positive), or Left for incomparable values
   */
  def compareCellValues(a: CellValue, b: CellValue): Either[EvalError, Int] =
    (normalizeForCompare(a), normalizeForCompare(b)) match
      case (CellValue.Error(err), _) =>
        Left(EvalError.EvalFailed(s"comparison: cannot compare ${err.toExcel} error value", None))
      case (_, CellValue.Error(err)) =>
        Left(EvalError.EvalFailed(s"comparison: cannot compare ${err.toExcel} error value", None))
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
   * @param op
   *   interprets the comparison sign from [[compareCellValues]] (e.g. `_ < 0` for Lt)
   * @return
   *   ArrayResult of CellValue.Bool values
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
      ) { case (l, r) => compareCellValues(l, r).map(c => CellValue.Bool(op(c))) }
    yield ArrayResult(result)

  /**
   * GH-197: Element-wise equality/inequality comparison with broadcasting.
   *
   * Works on CellValue directly for polymorphic equality (strings, numbers, booleans). Used by
   * Eq/Neq operators.
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
        Right(ArrayResult(left.values.map(_.map { cv =>
          val eq = cellValueEquals(cv, scalarCV)
          CellValue.Bool(if negate then !eq else eq)
        })))
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
              val eq = cellValueEquals(lVal, rVal)
              CellValue.Bool(if negate then !eq else eq)
            }
          }
          ArrayResult(result)

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
   * GH-333: elementwise IF over an array condition (Excel CSE semantics).
   *
   * The condition broadcasts against both branch arrays (scalar branches arrive as 1x1 arrays and
   * repeat like any dimension-1 axis): out(i)(j) = if cond(i)(j) then ifTrue(i)(j) else
   * ifFalse(i)(j). This is the classic MIN(IF(...))/MAX(IF(...)) shape — the aggregate then folds
   * the resulting array.
   *
   * Condition elements follow Excel truthiness (booleans as-is, numbers zero/non-zero, empty is
   * FALSE); text or error elements refuse with a clean Left.
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
        cellValueTruthy(condCV).map { truthy =>
          if truthy then getWithBroadcastCV(ifTrue.values, row, col, ifTrue.rows, ifTrue.cols)
          else getWithBroadcastCV(ifFalse.values, row, col, ifFalse.rows, ifFalse.cols)
        }
      }
    yield ArrayResult(result)

  /**
   * Excel truthiness for an IF-condition element: booleans as-is, numbers zero/non-zero, empty is
   * FALSE (the ScalarCoercion Bool conventions); text and errors refuse cleanly.
   */
  private def cellValueTruthy(cv: CellValue): Either[EvalError, Boolean] = cv match
    case CellValue.Bool(b) => Right(b)
    case CellValue.Number(n) => Right(n.signum != 0)
    case CellValue.Empty => Right(false)
    case CellValue.Formula(_, Some(cached)) => cellValueTruthy(cached)
    case CellValue.Error(err) =>
      Left(EvalError.EvalFailed(s"IF condition: cannot coerce ${err.toExcel} error value", None))
    case other => Left(EvalError.TypeMismatch("IF condition", "boolean", other.toString))

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
   * equal nothing.
   */
  def cellValueEquals(a: CellValue, b: CellValue): Boolean =
    compareCellValues(a, b).fold(_ => false, _ == 0)

  /**
   * Broadcast two arrays together.
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
      leftNums <- arrayToNumeric(left)
      rightNums <- arrayToNumeric(right)
      result <- traverseVV(
        (0 until outRows).toVector.map { row =>
          (0 until outCols).toVector.map { col =>
            val lVal = getWithBroadcast(leftNums, row, col, left.rows, left.cols)
            val rVal = getWithBroadcast(rightNums, row, col, right.rows, right.cols)
            (lVal, rVal)
          }
        }
      ) { case (l, r) => op(l, r) }
    yield toArrayResult(result)

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

  /**
   * Get value with broadcasting (dimensions of 1 repeat).
   */
  private def getWithBroadcast(
    arr: Vector[Vector[BigDecimal]],
    row: Int,
    col: Int,
    arrRows: Int,
    arrCols: Int
  ): BigDecimal =
    val r = if arrRows == 1 then 0 else row
    val c = if arrCols == 1 then 0 else col
    arr(r)(c)

  /**
   * Convert numeric matrix to ArrayResult.
   */
  private def toArrayResult(nums: Vector[Vector[BigDecimal]]): ArrayResult =
    ArrayResult(nums.map(_.map(n => CellValue.Number(n))))

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
