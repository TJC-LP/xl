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
    try Right(BigDecimal(scala.math.pow(x.toDouble, y.toDouble)))
    catch case e: Exception => Left(EvalError.EvalFailed(s"Power failed: ${e.getMessage}", None))

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

  /** Comparison operation type (returns Boolean, not Either) */
  type CompareOp = (BigDecimal, BigDecimal) => Boolean

  /**
   * GH-197: Perform broadcasting comparison operation.
   *
   * Like `broadcast`, but for comparison operators (>, <, =, etc.). Returns ArrayResult of
   * booleans.
   *
   * @param left
   *   Left operand (scalar or array)
   * @param right
   *   Right operand (scalar or array)
   * @param op
   *   Comparison operation to apply element-wise
   * @return
   *   ArrayResult of CellValue.Bool values
   */
  def broadcastCompare(
    left: ArrayOperand,
    right: ArrayOperand,
    op: CompareOp
  ): Either[EvalError, ArrayResult] =
    (left, right) match
      // scalar vs scalar -> 1x1 boolean array
      case (ArrayOperand.Scalar(l), ArrayOperand.Scalar(r)) =>
        Right(ArrayResult.single(CellValue.Bool(op(l, r))))

      // scalar vs array -> broadcast scalar to all elements
      case (ArrayOperand.Scalar(s), ArrayOperand.Array(arr)) =>
        arrayToNumeric(arr).map { nums =>
          ArrayResult(nums.map(_.map(v => CellValue.Bool(op(s, v)))))
        }

      // array vs scalar -> broadcast scalar to all elements
      case (ArrayOperand.Array(arr), ArrayOperand.Scalar(s)) =>
        arrayToNumeric(arr).map { nums =>
          ArrayResult(nums.map(_.map(v => CellValue.Bool(op(v, s)))))
        }

      // array vs array -> broadcast with dimension matching
      case (ArrayOperand.Array(l), ArrayOperand.Array(r)) =>
        broadcastCompareArrays(l, r, op)

  /**
   * GH-197: Broadcast compare two arrays together.
   */
  private def broadcastCompareArrays(
    left: ArrayResult,
    right: ArrayResult,
    op: CompareOp
  ): Either[EvalError, ArrayResult] =
    for
      outRows <- broadcastDim(left.rows, right.rows, "rows")
      outCols <- broadcastDim(left.cols, right.cols, "columns")
      leftNums <- arrayToNumeric(left)
      rightNums <- arrayToNumeric(right)
    yield
      val result = (0 until outRows).toVector.map { row =>
        (0 until outCols).toVector.map { col =>
          val lVal = getWithBroadcast(leftNums, row, col, left.rows, left.cols)
          val rVal = getWithBroadcast(rightNums, row, col, right.rows, right.cols)
          CellValue.Bool(op(lVal, rVal))
        }
      }
      ArrayResult(result)

  /**
   * GH-197: Element-wise equality/inequality comparison with broadcasting.
   *
   * Unlike `broadcastCompare`, this works on CellValue directly for polymorphic equality (strings,
   * numbers, booleans). Used by Eq/Neq operators.
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

  /** Convert any value to CellValue for comparison. */
  def anyToCellValue(v: Any): CellValue = v match
    case cv: CellValue => cv
    case s: String => CellValue.Text(s)
    case n: BigDecimal => CellValue.Number(n)
    case n: Int => CellValue.Number(BigDecimal(n))
    case n: Long => CellValue.Number(BigDecimal(n))
    case n: Double => CellValue.Number(BigDecimal(n))
    case b: Boolean => CellValue.Bool(b)
    case _ => CellValue.Text(v.toString)

  /** Compare CellValues for equality (case-insensitive for text). */
  private def cellValueEquals(a: CellValue, b: CellValue): Boolean = (a, b) match
    case (CellValue.Text(t1), CellValue.Text(t2)) =>
      t1.equalsIgnoreCase(t2)
    case (CellValue.Number(n1), CellValue.Number(n2)) =>
      n1 == n2
    case (CellValue.Bool(b1), CellValue.Bool(b2)) =>
      b1 == b2
    case (CellValue.Empty, CellValue.Empty) =>
      true
    case (CellValue.Formula(_, Some(c1)), other) =>
      cellValueEquals(c1, other)
    case (other, CellValue.Formula(_, Some(c2))) =>
      cellValueEquals(other, c2)
    case _ =>
      false

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
