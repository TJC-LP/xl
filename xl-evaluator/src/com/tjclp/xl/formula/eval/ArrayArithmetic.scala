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
 * [[EvalError.toCellError]] instead of failing the whole formula. GH-344 4b: the broadcasts are
 * TOTAL in Right — mismatched dimensions pad with #N/A elements ([[getWithBroadcastCV]]);
 * aggregators decide whether consumed error elements propagate.
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
      else if x == 0 && y < 0 then
        // Excel: 0 raised to a negative power is #DIV/0!
        Left(EvalError.ErrorValue(CellError.Div0, Some(s"$x^$y")))
      else
        // Fall back to Double for fractional/negative exponents; a NaN (a negative base under a
        // fractional exponent) or an overflow is Excel's #NUM!
        val d = scala.math.pow(x.toDouble, y.toDouble)
        if d.isNaN || d.isInfinite then
          Left(EvalError.ErrorValue(CellError.Num, Some(s"$x^$y is not a finite number")))
        else Right(BigDecimal(d))
    catch
      // GH-344: a genuine magnitude/scale overflow is Excel's #NUM! (op-level classification);
      // any other failure stays a generic loud EvalFailed
      case _: ArithmeticException =>
        Left(EvalError.ErrorValue(CellError.Num, Some("power overflow")))
      case e: Exception => Left(EvalError.EvalFailed(s"Power failed: ${e.getMessage}", None))

  /**
   * Convert a CellValue to BigDecimal for arithmetic.
   *
   *   - Number -> value
   *   - Empty -> 0
   *   - Bool -> 1/0
   *   - DateTime -> Excel serial number (GH-449: dates are numbers in arithmetic, the
   *     [[normalizeForCompare]] convention extended to the broadcast operand path so `=A1:A3-B1`
   *     over date cells agrees with the scalar `=A1-B1`)
   *   - Text containing number -> parsed value
   *   - Other -> error
   */
  def cellValueToNumeric(cv: CellValue): Either[EvalError, BigDecimal] = cv match
    case CellValue.Number(n) => Right(n)
    case CellValue.Empty => Right(BigDecimal(0))
    case CellValue.Bool(b) => Right(boolToNumeric(b))
    case CellValue.DateTime(dt) => Right(BigDecimal(CellValue.dateTimeToExcelSerial(dt)))
    case CellValue.Formula(_, Some(cached), _) => cellValueToNumeric(cached)
    case CellValue.Text(s) =>
      ScalarCoercion.parseNumericText(s) match
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
          case CellValue.Formula(_, Some(cached), _) => cached
          case other => other
      }.toVector
    }.toVector
    ArrayResult(values)

  /** GH-499: retain every position while resolving formulas, and keep host failures explicit. */
  private[formula] def rangeToArrayEval(
    range: CellRange,
    readCell: ARef => Either[EvalError, CellValue]
  ): Either[EvalError, ArrayResult] =
    rangeToMatrixEval(range, readCell).map(ArrayResult(_))

  /** [[rangeToArrayEval]]'s rows, for the consumers that work on the whole grid. */
  private[formula] def rangeToMatrixEval(
    range: CellRange,
    readCell: ARef => Either[EvalError, CellValue]
  ): Either[EvalError, Vector[Vector[CellValue]]] =
    (range.rowStart.index0 to range.rowEnd.index0)
      .foldLeft[Either[EvalError, Vector[Vector[CellValue]]]](Right(Vector.empty)) { (rows, row) =>
        rows.flatMap { accumulated =>
          (range.colStart.index0 to range.colEnd.index0)
            .foldLeft[Either[EvalError, Vector[CellValue]]](Right(Vector.empty)) { (values, col) =>
              values.flatMap(cells => readCell(ARef.from0(col, row)).map(cells :+ _))
            }
            .map(accumulated :+ _)
        }
      }

  /**
   * How [[readRange]] reads a range: by its size (`Auto`, every production evaluation), or always
   * one way. `Sparse` (every range through its populated box) and `Dense` (every range position by
   * position) exist for the sparse-equals-dense laws, which set them on the pass memo
   * ([[Evaluator.EvalMemo.rangeRead]]) through [[Evaluator.arrayInstanceReading]].
   */
  private[formula] enum RangeRead derives CanEqual:
    case Auto, Sparse, Dense

  /**
   * #750: a range's values in array mode, at the range's full shape, read in time and memory
   * proportional to the smaller of the range and the sheet's populated cells.
   *
   * A range with no more positions than the sheet has cells is read position by position, as always
   * (a bounded `$A$2:$A$50` never scans the sheet). A larger one (`R:KHI`, `A:XFD`, the generous
   * `INDEX($A$1:$A$100000,0)`) is read from the other side: one pass over the sheet's cell keys
   * finds the box of populated cells inside it, only that box is read, and every other position is
   * the array's Empty fill, which reads exactly as an absent cell does. Every formula cell the
   * range holds is populated, so it lies in the box and is evaluated as before. The box found for a
   * (sheet, range) is kept in the pass memo when there is one.
   *
   * Only cells that hold a value widen the box: a style-only record (kept in `Sheet.cells` as
   * Empty, the rule `QualifiedGraph` and the streaming reader use) reads as Empty, exactly as the
   * fill does, so a formatted blank in XFD1 does not stretch `A:XFD`'s box across the sheet.
   *
   * The box is always inside the range, so reading it is never more work than the positional read
   * of the whole range that 0.23.1 did. Only a box past [[ArrayResult.MaxBoxCells]] (and past the
   * sheet's own cell count, whichever is larger) is refused, as a loud `EvalFailed`: two cells at
   * opposite corners of `A:XFD` make a 17-billion-cell box that no read could build.
   */
  private[formula] def readRange(
    range: CellRange,
    sheet: Sheet,
    readCell: ARef => Either[EvalError, CellValue],
    memo: Option[Evaluator.EvalMemo]
  ): Either[EvalError, ArrayResult] =
    val height = range.height
    val width = range.width
    if height <= 0 || width <= 0 then Right(ArrayResult.empty)
    else
      val area = height.toLong * width.toLong
      val positional = memo.fold(RangeRead.Auto)(_.rangeRead) match
        case RangeRead.Dense => true
        case RangeRead.Sparse => false
        case RangeRead.Auto => area <= sheet.cells.size.toLong
      if positional then rangeToMatrixEval(range, readCell).map(ArrayResult(_))
      else
        val found =
          memo.fold(populatedBox(range, sheet))(
            _.populatedBox(sheet, range)(populatedBox(range, sheet))
          )
        found match
          case None => Right(ArrayResult.constant(height, width, CellValue.Empty))
          case Some(box) =>
            val boxArea = box.height.toLong * box.width.toLong
            val bound = math.max(ArrayResult.MaxBoxCells, sheet.cells.size.toLong)
            if boxArea > bound then
              Left(
                EvalError.EvalFailed(
                  s"${range.toA1}: its populated cells span ${box.toA1} " +
                    s"(${box.height}×${box.width} cells), more than the " +
                    s"${ArrayResult.count(bound)}-cell bound on a range's populated box",
                  None
                )
              )
            else
              rangeToMatrixEval(box, readCell).map { values =>
                ArrayResult.sparse(
                  height,
                  width,
                  box.rowStart.index0 - range.rowStart.index0,
                  box.colStart.index0 - range.colStart.index0,
                  values,
                  ArrayResult.Fill.Const(CellValue.Empty)
                )
              }

  /**
   * The smallest range holding every cell of `sheet` inside `range` that holds a value (a
   * style-only record, stored as Empty, reads as the Empty fill does): one scan of the cells.
   */
  private def populatedBox(range: CellRange, sheet: Sheet): Option[CellRange] =
    val r0 = range.rowStart.index0
    val r1 = range.rowEnd.index0
    val c0 = range.colStart.index0
    val c1 = range.colEnd.index0
    sheet.cells.iterator
      .foldLeft(Option.empty[(Int, Int, Int, Int)]) { case (acc, (at, cell)) =>
        val r = at.row.index0
        val c = at.col.index0
        if r < r0 || r > r1 || c < c0 || c > c1 then acc
        else if !holdsValue(cell.value) then acc
        else
          Some(acc.fold((r, r, c, c)) { case (top, bottom, left, right) =>
            (math.min(top, r), math.max(bottom, r), math.min(left, c), math.max(right, c))
          })
      }
      .map((top, bottom, left, right) =>
        CellRange(ARef.from0(left, top), ARef.from0(right, bottom))
      )

  private def holdsValue(value: CellValue): Boolean = value match
    case CellValue.Empty => false
    case _ => true

  /**
   * Convert ArrayResult to numeric matrix (every cell: refused past [[ArrayResult.MaxCells]]).
   */
  def arrayToNumeric(arr: ArrayResult): Either[EvalError, Vector[Vector[BigDecimal]]] =
    arr.toDense("arrayToNumeric").flatMap(traverseVV(_)(cellValueToNumeric))

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
    case CellValue.Formula(_, Some(cached), _) => normalizeForCompare(cached)
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
   * uncomparable value) demotes to #VALUE!. GH-344 4b: total in Right — mismatched dimensions pad
   * with #N/A elements.
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
    broadcastN(Vector(left, right)) { elements =>
      val l = elements.headOption.getOrElse(CellValue.Empty)
      val r = elements.lift(1).getOrElse(CellValue.Empty)
      carriedError(l).orElse(carriedError(r)) match
        case Some(err) => Right(CellValue.Error(err))
        case None =>
          compareCellValues(l, r) match
            case Right(sign) => Right(CellValue.Bool(op(sign)))
            case Left(e) => toErrorElement(e)
    }

  /**
   * GH-197: Element-wise equality/inequality comparison with broadcasting.
   *
   * Works on CellValue directly for polymorphic equality (strings, numbers, booleans). Used by
   * Eq/Neq operators.
   *
   * GH-337: a carried error on either operand (left wins) becomes the OUTPUT element — errors are
   * never "equal" or "unequal", they propagate (negate does not flip them). GH-344 4b: total in
   * Right — mismatched dimensions pad with #N/A elements.
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
        Right(left.mapPure(cv => equalityElement(cv, scalarCV, negate)))
      case Left(rightArr) =>
        // Array vs array - broadcast
        broadcastN(Vector(left, rightArr)) { elements =>
          Right(
            equalityElement(
              elements.headOption.getOrElse(CellValue.Empty),
              elements.lift(1).getOrElse(CellValue.Empty),
              negate
            )
          )
        }

  /** GH-337: one equality output element — carried errors win (left first), then Bool equality. */
  private def equalityElement(l: CellValue, r: CellValue, negate: Boolean): CellValue =
    carriedError(l).orElse(carriedError(r)) match
      case Some(err) => CellValue.Error(err)
      case None =>
        val eq = cellValueEquals(l, r)
        CellValue.Bool(if negate then !eq else eq)

  /**
   * Get CellValue with broadcasting. GH-344 4b: a position beyond an operand's extent (extent ≠ 1,
   * so not a broadcast axis) reads as a #N/A element — Excel pads mismatched array dimensions with
   * #N/A rather than failing the operation.
   */
  private def getWithBroadcastCV(arr: ArrayResult, row: Int, col: Int): CellValue =
    val arrRows = arr.rows
    val arrCols = arr.cols
    val r = if arrRows == 1 then 0 else row
    val c = if arrCols == 1 then 0 else col
    if r >= arrRows || c >= arrCols then CellValue.Error(CellError.NA)
    else arr(r, c)

  /**
   * GH-337: the Excel error an element carries, if any — a `CellValue.Error` directly or cached
   * inside a formula value. Package-visible so the aggregate error guards (FunctionSpecsAggregate)
   * share the exact matching semantics of the broadcasts.
   */
  private[formula] def carriedError(cv: CellValue): Option[CellError] = cv match
    case CellValue.Error(err) => Some(err)
    case CellValue.Formula(_, Some(cached), _) => carriedError(cached)
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
   * division by zero → #DIV/0!, pow overflow → #NUM! via its GH-344 op-level classification) become
   * error elements rather than failing the whole broadcast.
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
   * element. GH-344 4b: total in Right — mismatched dimensions pad with #N/A elements.
   */
  def broadcastIf(
    cond: ArrayResult,
    ifTrue: ArrayResult,
    ifFalse: ArrayResult
  ): Either[EvalError, ArrayResult] =
    broadcastN(Vector(cond, ifTrue, ifFalse)) { elements =>
      val condCV = elements.headOption.getOrElse(CellValue.Empty)
      carriedError(condCV) match
        case Some(err) => Right(CellValue.Error(err))
        case None =>
          conditionTruthy("IF condition", condCV) match
            case Right(truthy) =>
              Right(elements.lift(if truthy then 1 else 2).getOrElse(CellValue.Empty))
            case Left(e) => toErrorElement(e)
    }

  /**
   * GH-333/GH-338: Excel truthiness for a condition element, shared by broadcastIf (IF over an
   * array condition) and the logical functions' array paths (AND/OR aggregation, NOT broadcast):
   * booleans as-is, numbers zero/non-zero, empty is FALSE (the ScalarCoercion Bool conventions);
   * text other than TRUE/FALSE is `#VALUE!` (#709), naming the position via `label`.
   *
   * GH-344: an error element propagates as its Excel error VALUE (Left(ErrorValue) — the AND/OR
   * folds short-circuit with it and the boundary promotes it to the error cell, Excel-exact).
   * broadcastIf and broadcastNot pre-check errors with [[carriedError]] and demote their residual
   * refusals to #VALUE! elements, so their condition failures stay positional.
   */
  def conditionTruthy(label: String, cv: CellValue): Either[EvalError, Boolean] = cv match
    case CellValue.Bool(b) => Right(b)
    case CellValue.Number(n) => Right(n.signum != 0)
    // GH-564: dates are numbers in Excel's value model — their serial is zero/non-zero
    case CellValue.DateTime(dt) => Right(CellValue.dateTimeToExcelSerial(dt) != 0.0)
    case CellValue.Empty => Right(false)
    case CellValue.Formula(_, Some(cached), _) => conditionTruthy(label, cached)
    case CellValue.Error(err) =>
      Left(EvalError.ErrorValue(err, Some(label)))
    // GH-344 item 5: exactly "TRUE"/"FALSE" (case-insensitive, no trim) coerce via the shared
    // ScalarCoercion.boolTextValue table — aligned with coerceBool and decodeBool (L6 parity)
    case CellValue.Text(s) =>
      ScalarCoercion.boolTextValue(s) match
        case Some(b) => Right(b)
        case None =>
          Left(
            EvalError.ErrorValue(
              CellError.Value,
              Some(s"$label: text '${ScalarCoercion.quotedText(s)}' is not TRUE or FALSE")
            )
          )
    // rich text is text: read through its plain text, like the aggregates' directValue
    case CellValue.RichText(rt) => conditionTruthy(label, CellValue.Text(rt.toPlainText))
    case other => Left(EvalError.TypeMismatch(label, "boolean", other.toString))

  /**
   * GH-338: truthiness of every element (row-major), for the AND/OR aggregation folds (AND =
   * forall, OR = exists). The first refusing element short-circuits with its Left — GH-344: an
   * error element's Left carries its Excel error VALUE (the fold's result at the boundary), a text
   * element's is `#VALUE!` (#709).
   */
  def truthyElements(label: String, arr: ArrayResult): Either[EvalError, Vector[Boolean]] =
    arr.toDense(label).flatMap(rows => traverseV(rows.flatten)(cv => conditionTruthy(label, cv)))

  /**
   * #750: [[truthyElements]] folded with `combine` from `seed` (AND = `&&`, OR = `||`), walking the
   * array's [[ArrayResult.tally]]: an idempotent `combine` needs each group's value once, so a
   * whole-axis array folds in the time of its populated cells. The first refusing element, in
   * row-major order, short-circuits with its Left.
   */
  def truthyFold(
    label: String,
    arr: ArrayResult,
    seed: Boolean,
    combine: (Boolean, Boolean) => Boolean
  ): Either[EvalError, Boolean] =
    arr.tally.foldLeft[Either[EvalError, Boolean]](Right(seed)) {
      case (Left(err), _) => Left(err)
      case (Right(acc), (cv, _)) => conditionTruthy(label, cv).map(combine(acc, _))
    }

  /**
   * GH-338: elementwise NOT over an array condition, preserving shape (Excel broadcasts NOT rather
   * than aggregating). GH-344: mirrors broadcastIf's positional carriage — an element carrying an
   * error emits THAT error at its output position, and a residual refusal (text) demotes to a
   * #VALUE! element. Total in Right.
   */
  def broadcastNot(label: String, arr: ArrayResult): Either[EvalError, ArrayResult] =
    arr.map { cv =>
      carriedError(cv) match
        case Some(err) => Right(CellValue.Error(err))
        case None =>
          conditionTruthy(label, cv) match
            case Right(b) => Right(CellValue.Bool(!b))
            case Left(e) => toErrorElement(e)
    }

  /** Convert any value to CellValue for comparison. */
  def anyToCellValue(v: Any): CellValue = v match
    case cv: CellValue => cv
    case s: String => CellValue.Text(s)
    case n: BigDecimal => CellValue.Number(n)
    case n: Int => CellValue.Number(BigDecimal(n))
    case n: Long => CellValue.Number(BigDecimal(n))
    case n: Double if n.isFinite => CellValue.Number(BigDecimal(n))
    // #681: NaN and ±Infinity have no Excel value (BigDecimal(Double) throws on them) — the
    // #NUM! element, which `&` and every carried-error check then propagate
    case _: Double => CellValue.Error(CellError.Num)
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
   * elementwise via [[applyOpCarrying]]; GH-344 4b: total in Right — mismatched dimensions pad with
   * #N/A elements).
   */
  private def broadcastArrays(
    left: ArrayResult,
    right: ArrayResult,
    op: BinaryOp
  ): Either[EvalError, ArrayResult] =
    broadcastN(Vector(left, right)) { elements =>
      applyOpCarrying(
        elements.headOption.getOrElse(CellValue.Empty),
        elements.lift(1).getOrElse(CellValue.Empty),
        op
      )
    }

  /**
   * #750: what one operand of a broadcast contributes outside the part of the output computed cell
   * by cell. Its value varies over the output rows `[top, bottom)` and columns `[left, right)` (an
   * empty rectangle when it varies nowhere), and is `outside` everywhere else: one value, or one
   * value per output row or per output column.
   */
  private final case class Footprint(
    top: Int,
    bottom: Int,
    left: Int,
    right: Int,
    outside: Outside
  ):
    def isEmpty: Boolean = top >= bottom || left >= right
    def area: Long = if isEmpty then 0L else (bottom - top).toLong * (right - left).toLong

  /**
   * An operand's value outside its [[Footprint]]: one value, or a function of the row or column.
   */
  private enum Outside derives CanEqual:
    case Const(value: CellValue)
    case ByRow(at: Int => CellValue)
    case ByCol(at: Int => CellValue)

    def entry(i: Int): CellValue = this match
      case Const(v) => v
      case ByRow(at) => at(i)
      case ByCol(at) => at(i)

  /** The fill a broadcast's result takes: one value, one per row, or one per column. */
  private enum FillAxis derives CanEqual:
    case Const, ByRow, ByCol

    def admits(outside: Outside): Boolean = (this, outside) match
      case (_, Outside.Const(_)) => true
      case (ByRow, Outside.ByRow(_)) => true
      case (ByCol, Outside.ByCol(_)) => true
      case _ => false

  private def uniform(value: CellValue): Footprint = Footprint(0, 0, 0, 0, Outside.Const(value))

  /**
   * The ways an operand can be described on a `outRows` x `outCols` broadcast, each exact:
   *   - a vector broadcast along the other axis (`{1,2,3}` down a million rows) varies only by
   *     column (or by row): no footprint at all, one value per output column (or row);
   *   - its populated box, stretched along a broadcast axis, and its own fill outside it (a fill
   *     that also takes the #N/A of the positions past its extent, when it is shorter than the
   *     output along one axis and that fill does not vary along the other);
   *   - when it is shorter than the output, its whole extent, with the #N/A pad outside it.
   */
  private def footprints(op: ArrayResult, outRows: Int, outCols: Int): Vector[Footprint] =
    val na = CellValue.Error(CellError.NA)
    if op.isEmpty then Vector(uniform(na))
    else if op.rows == 1 && op.cols == 1 then Vector(uniform(op(0, 0)))
    else
      val padRows = op.rows != 1 && op.rows < outRows
      val padCols = op.cols != 1 && op.cols < outCols
      val vector =
        if op.rows == 1 then
          Vector(Footprint(0, 0, 0, 0, Outside.ByCol(c => getWithBroadcastCV(op, 0, c))))
        else if op.cols == 1 then
          Vector(Footprint(0, 0, 0, 0, Outside.ByRow(r => getWithBroadcastCV(op, r, 0))))
        else Vector.empty
      // the output span an axis varies over: the whole axis when it broadcasts, else the box's
      def span(extent: Int, out: Int, boxStart: Int, boxSize: Int): (Int, Int) =
        if extent == 1 then (0, out) else (boxStart, boxStart + boxSize)
      val hasBox = op.boxRows > 0 && op.boxCols > 0
      val (top, bottom) = if hasBox then span(op.rows, outRows, op.boxTop, op.boxRows) else (0, 0)
      val (left, right) = if hasBox then span(op.cols, outCols, op.boxLeft, op.boxCols) else (0, 0)
      val ownFill: Option[Outside] = op.fill match
        case ArrayResult.Fill.Const(v) =>
          if padRows && padCols then None
          else if padRows then Some(Outside.ByRow(r => if r < op.rows then v else na))
          else if padCols then Some(Outside.ByCol(c => if c < op.cols then v else na))
          else Some(Outside.Const(v))
        case f @ ArrayResult.Fill.ByRow(_) =>
          if padCols then None
          else Some(Outside.ByRow(r => if r < op.rows then f.at(r, 0) else na))
        case f @ ArrayResult.Fill.ByCol(_) =>
          if padRows then None
          else Some(Outside.ByCol(c => if c < op.cols then f.at(0, c) else na))
      val boxed = ownFill.map(Footprint(top, bottom, left, right, _))
      val extent =
        if padRows || padCols then
          val height = if op.rows == 1 then outRows else op.rows
          val width = if op.cols == 1 then outCols else op.cols
          Some(Footprint(0, height, 0, width, Outside.Const(na)))
        else None
      vector ++ boxed ++ extent

  /** The output rectangle a broadcast computes cell by cell, and the fill outside it. */
  private final case class BroadcastPlan(
    top: Int,
    bottom: Int,
    left: Int,
    right: Int,
    axis: FillAxis,
    outsides: Vector[Outside]
  )

  /**
   * N-ary element-wise broadcast under the same shape rules as the binary broadcasts: an axis of
   * extent 1 repeats, unequal extents take the larger and pad the missing positions with #N/A
   * (GH-344 4b). `elem` receives one element per operand, in operand order; a Left from it fails
   * the whole broadcast (element-local failures should come back as error ELEMENTS instead).
   *
   * #750: over sparse operands only the union of the operands' footprints is computed cell by cell;
   * every other cell is the element of the operands' outside values, computed once per fill entry.
   * The result's fill is one value when every operand's outside is, and one value per column (or
   * per row) when an operand's varies along that axis: `A:A*{1,2,3}` computes the populated rows of
   * A and three fill entries. An outside that would vary along both axes (a whole column times a
   * whole row, both padded) is computed cell by cell when the result is at most
   * [[ArrayResult.MaxCells]] cells, and otherwise over the smallest rectangle that leaves a fill by
   * one axis outside it, refused when that rectangle is past [[ArrayResult.MaxBoxCells]] (and past
   * the operands' boxes together).
   */
  private[formula] def broadcastN(operands: Vector[ArrayResult], perCell: Boolean = false)(
    elem: Vector[CellValue] => Either[EvalError, CellValue]
  ): Either[EvalError, ArrayResult] =
    val outRows = operands.foldLeft(1)((acc, op) => broadcastExtent(acc, op.rows))
    val outCols = operands.foldLeft(1)((acc, op) => broadcastExtent(acc, op.cols))
    val outCells = outRows.toLong * outCols.toLong
    if outRows <= 0 || outCols <= 0 then Right(ArrayResult.empty)
    else
      val whole = BroadcastPlan(0, outRows, 0, outCols, FillAxis.Const, Vector.empty)
      def cellByCell: Either[EvalError, BroadcastPlan] =
        if outCells > ArrayResult.MaxCells then
          Left(ArrayResult.tooLarge("array operation", outRows.toLong, outCols.toLong))
        else Right(whole)
      val plan: Either[EvalError, BroadcastPlan] =
        if operands.forall(_.isDense) then Right(whole)
        else if perCell then cellByCell
        else
          val options = operands.map(footprints(_, outRows, outCols))
          // the computed rectangle is built cell by cell, so it takes the bound a range's box does
          // (its operands' boxes, already built, count too)
          val budget = math.max(ArrayResult.MaxBoxCells, operands.map(_.boxArea).sum)
          // for each fill axis, every operand's smallest footprint whose outside it admits
          val candidates = FillAxis.values.toVector.flatMap { axis =>
            val picks = options.map(
              _.filter(f => axis.admits(f.outside))
                .sortBy(f =>
                  (
                    f.area,
                    f.outside match
                      case Outside.Const(_) => 0
                      case _ => 1
                  )
                )
                .headOption
            )
            if picks.exists(_.isEmpty) then None
            else
              val chosen = picks.flatten
              val rect = boundingRect(
                chosen.filterNot(_.isEmpty).map(f => (f.top, f.bottom, f.left, f.right))
              )
                .getOrElse((0, 0, 0, 0))
              val area = (rect._2 - rect._1).toLong * (rect._4 - rect._3).toLong
              val varies = chosen.exists(f =>
                !(f.outside match
                  case Outside.Const(_) => true
                  case _ => false
                )
              )
              val entries = (axis, varies) match
                case (FillAxis.ByRow, true) => outRows.toLong
                case (FillAxis.ByCol, true) => outCols.toLong
                case _ => 0L
              val effective = if varies then axis else FillAxis.Const
              Some(
                (
                  area,
                  area + entries,
                  BroadcastPlan(
                    rect._1,
                    rect._2,
                    rect._3,
                    rect._4,
                    effective,
                    chosen.map(_.outside)
                  )
                )
              )
          }
          candidates.filter(_._1 <= budget).sortBy(_._2).headOption match
            case Some((_, _, plan)) => Right(plan)
            case None if outCells <= ArrayResult.MaxCells => Right(whole)
            case None =>
              candidates.sortBy(_._1).headOption match
                case Some((_, _, p)) =>
                  Left(
                    EvalError.EvalFailed(
                      s"array operation: the populated cells of its ${outRows}×$outCols result " +
                        s"span ${p.bottom - p.top}×${p.right - p.left} cells, more than the " +
                        s"${ArrayResult.count(budget)}-cell bound on its computed box",
                      None
                    )
                  )
                case None => cellByCell
      plan.flatMap { p =>
        val height = math.max(p.bottom - p.top, 0)
        val width = math.max(p.right - p.left, 0)
        val g = ArrayResult.Geometry(outRows, outCols, p.top, p.left, height, width)
        // every cell outside the computed box: the element of the operands' outside values, once
        // per fill entry
        val fillResult = ArrayResult.computeFill(
          g,
          p.axis match
            case FillAxis.Const => None
            case FillAxis.ByRow => Some(true)
            case FillAxis.ByCol => Some(false)
        )(i => elem(p.outsides.map(_.entry(i))))
        computeBox(p.top, height, p.left, width, outCols)(r =>
          c => elem(operands.map(op => getWithBroadcastCV(op, r, c)))
        ) match
          case Left((err, at)) =>
            fillResult match
              case Left((fillErr, fillAt)) if fillAt < at => Left(fillErr)
              case _ => Left(err)
          case Right(box) =>
            fillResult match
              case Left((fillErr, _)) => Left(fillErr)
              case Right(fill) =>
                Right(ArrayResult.sparse(outRows, outCols, p.top, p.left, box, fill))
      }

  /**
   * The smallest `(top, bottom, left, right)` rectangle (bottom and right exclusive) holding every
   * one of `rects`; None for none.
   */
  private[formula] def boundingRect(
    rects: Iterable[(Int, Int, Int, Int)]
  ): Option[(Int, Int, Int, Int)] =
    rects.foldLeft(Option.empty[(Int, Int, Int, Int)]) { (acc, rect) =>
      Some(acc.fold(rect) { case (t, b, l, r) =>
        (math.min(t, rect._1), math.max(b, rect._2), math.min(l, rect._3), math.max(r, rect._4))
      })
    }

  /**
   * The `height` x `width` cells at (`top`, `left`) of an output `outCols` wide, row-major,
   * stopping at the first Left, which comes back with its row-major position in the output.
   */
  private def computeBox(top: Int, height: Int, left: Int, width: Int, outCols: Int)(
    cell: Int => Int => Either[EvalError, CellValue]
  ): Either[(EvalError, Long), Vector[Vector[CellValue]]] =
    val rowsOut = new scala.collection.immutable.VectorBuilder[Vector[CellValue]]
    @scala.annotation.tailrec
    def rowLoop(i: Int): Either[(EvalError, Long), Vector[Vector[CellValue]]] =
      if i >= height then Right(rowsOut.result())
      else
        val r = top + i
        val at = cell(r)
        val cellsOut = new scala.collection.immutable.VectorBuilder[CellValue]
        @scala.annotation.tailrec
        def colLoop(j: Int): Option[(EvalError, Long)] =
          if j >= width then None
          else
            at(left + j) match
              case Right(v) =>
                cellsOut += v
                colLoop(j + 1)
              case Left(e) => Some((e, r.toLong * outCols.toLong + (left + j).toLong))
        colLoop(0) match
          case Some(failure) => Left(failure)
          case None =>
            rowsOut += cellsOut.result()
            rowLoop(i + 1)
    rowLoop(0)

  /**
   * One `&` output element: a carried error on either side wins (left first) — errors propagate,
   * never stringify — otherwise the two texts join under [[ScalarCoercion.concatText]].
   */
  private[formula] def concatElement(l: CellValue, r: CellValue): CellValue =
    carriedError(l).orElse(carriedError(r)) match
      case Some(err) => CellValue.Error(err)
      case None => CellValue.Text(ScalarCoercion.concatText(l) + ScalarCoercion.concatText(r))

  /** The total form of [[broadcastDim]] for the N-ary fold. */
  private def broadcastExtent(l: Int, r: Int): Int =
    if l == r then l else if l == 1 then r else if r == 1 then l else math.max(l, r)

  /**
   * Compute broadcast output dimension for a single axis. GH-344 4b: unequal non-1 dimensions no
   * longer fail — the output extends to the larger extent and [[getWithBroadcastCV]] pads the
   * beyond-extent positions with #N/A (Excel's array-mismatch semantics). The Either shape is kept
   * for the callers' comprehensions; this function is total in Right. SUMPRODUCT deliberately does
   * NOT ride this law — it enforces exact dimensions as #VALUE! (Excel).
   */
  private def broadcastDim(l: Int, r: Int, dimName: String): Either[EvalError, Int] =
    Right(broadcastExtent(l, r))

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
