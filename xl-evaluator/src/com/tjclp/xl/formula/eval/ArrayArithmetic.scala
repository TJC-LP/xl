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
    case CellValue.Bool(b) => Right(if b then BigDecimal(1) else BigDecimal(0))
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
