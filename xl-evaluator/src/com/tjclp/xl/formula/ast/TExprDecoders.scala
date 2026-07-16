package com.tjclp.xl.formula.ast

import com.tjclp.xl.formula.functions.FunctionSpecs
import com.tjclp.xl.formula.eval.EvalError
import com.tjclp.xl.formula.functions.EvalContext

import com.tjclp.xl.cells.{Cell, CellValue}
import com.tjclp.xl.codec.CodecError
import com.tjclp.xl.formula.eval.{ArrayArithmetic, ScalarCoercion}

trait TExprDecoders:
  // Decoder functions for cell coercion

  /**
   * Decode cell as numeric value (Double or BigDecimal).
   *
   * Handles Formula cells by extracting the cached numeric value when available. This enables
   * nested formula evaluation where a cell reference points to another formula cell with a cached
   * result.
   */
  def decodeNumeric(cell: Cell): Either[CodecError, BigDecimal] =
    cell.value match
      case CellValue.Number(value) => scala.util.Right(value)
      // GH-196: Coerce booleans to numeric (TRUE→1, FALSE→0)
      case CellValue.Bool(true) => scala.util.Right(BigDecimal(1))
      case CellValue.Bool(false) => scala.util.Right(BigDecimal(0))
      case CellValue.Formula(_, Some(CellValue.Number(cached))) =>
        // Extract cached numeric value from formula cell
        scala.util.Right(cached)
      // GH-196: Handle cached boolean values in formulas
      case CellValue.Formula(_, Some(CellValue.Bool(true))) =>
        scala.util.Right(BigDecimal(1))
      case CellValue.Formula(_, Some(CellValue.Bool(false))) =>
        scala.util.Right(BigDecimal(0))
      case other =>
        scala.util.Left(
          CodecError.TypeMismatch(
            expected = "Numeric",
            actual = other
          )
        )

  /**
   * GH-385: numeric decoding for SCALAR reference positions — Excel treats a direct reference to a
   * blank cell as 0 in numeric contexts (=A1+1 with blank A1 is 1, =ABS(A1) is 0), matching
   * ScalarCoercion.coerceNumeric's Empty -> 0 and ArrayArithmetic's Empty -> 0 broadcast rule.
   *
   * A separate decoder rather than an Empty case inside [[decodeNumeric]]: decodeNumeric is also
   * the range-FOLD decoder for the TExpr.Aggregate path (Evaluator) and the XNPV/XIRR/NPV/IRR
   * cashflow collector (FunctionSpecsFinancialCashflow), where Empty must stay a skip — MIN/MEDIAN
   * over a sparse range ignore blanks, and a blank cashflow must not become a period-shifting 0.
   * Error values still refuse via the delegate's catch-all (Empty and Error are distinct
   * constructors, so blank-as-zero can never mask a carried error).
   */
  def decodeNumericScalar(cell: Cell): Either[CodecError, BigDecimal] =
    cell.value match
      case CellValue.Empty => scala.util.Right(BigDecimal(0))
      case _ => decodeNumeric(cell)

  /**
   * Decode cell as LocalDate value (extracts date from DateTime).
   *
   * Handles Formula cells by extracting the cached DateTime value when available.
   */
  def decodeDate(cell: Cell): Either[CodecError, java.time.LocalDate] =
    cell.value match
      case CellValue.DateTime(value) => scala.util.Right(value.toLocalDate)
      case CellValue.Formula(_, Some(CellValue.DateTime(cached))) =>
        scala.util.Right(cached.toLocalDate)
      case other =>
        scala.util.Left(
          CodecError.TypeMismatch(
            expected = "Date",
            actual = other
          )
        )

  /**
   * Decode cell as Boolean value.
   *
   * Handles Formula cells by extracting the cached Boolean value when available.
   *
   * GH-306: numbers coerce with Excel truthiness (0 = FALSE, non-zero = TRUE) and blanks are FALSE,
   * so `=IF(A1, ...)` over a numeric or empty cell matches Excel — and matches the
   * BindingCoercion.Bool conventions for LET bindings (the LetFunctionSpec parity pin requires the
   * direct and bound forms to agree). Text still refuses (clean per-cell error).
   */
  def decodeBool(cell: Cell): Either[CodecError, Boolean] =
    cell.value match
      case CellValue.Bool(value) => scala.util.Right(value)
      case CellValue.Number(n) => scala.util.Right(n.signum != 0)
      case CellValue.Empty => scala.util.Right(false)
      case CellValue.Formula(_, Some(CellValue.Bool(cached))) =>
        scala.util.Right(cached)
      case CellValue.Formula(_, Some(CellValue.Number(cached))) =>
        scala.util.Right(cached.signum != 0)
      case other =>
        scala.util.Left(
          CodecError.TypeMismatch(
            expected = "Boolean",
            actual = other
          )
        )

  /**
   * Decode cell as CellValue (always succeeds).
   *
   * Used for IFERROR/ISERROR which need to preserve the raw cell value.
   */
  def decodeCellValue(cell: Cell): Either[CodecError, CellValue] =
    scala.util.Right(cell.value)

  /**
   * Decode cell for comparison operands (GH-335): extracts cached formula values but PRESERVES
   * emptiness, unlike decodeResolvedValue's Empty -> 0.
   *
   * Excel coerces an empty cell relative to the OTHER comparison operand (0 against numbers, ""
   * against text, FALSE against booleans) — that coercion lives in
   * ArrayArithmetic.compareCellValues and needs to see the Empty. An uncached formula value keeps
   * the decodeResolvedValue zero convention.
   */
  def decodeComparableValue(cell: Cell): Either[CodecError, CellValue] =
    val resolved = cell.value match
      case CellValue.Formula(_, Some(cached)) => cached
      case CellValue.Formula(_, None) => CellValue.Number(BigDecimal(0))
      case other => other
    scala.util.Right(resolved)

  /**
   * Decode cell as resolved CellValue (extracts cached values, converts empty to 0).
   *
   * Used for standalone cell references (e.g., =A1, =Sheet1!B2) where the formula returns the
   * cell's "effective" value:
   *   - Number, Text, Bool, DateTime, RichText -> returned as-is
   *   - Formula -> returns cached value if present, or Number(0) if no cache
   *   - Empty -> returns Number(0) (Excel treats empty as 0 in numeric contexts)
   *   - Error -> returns the error
   *
   * This matches Excel semantics for standalone cell references.
   */
  def decodeResolvedValue(cell: Cell): Either[CodecError, CellValue] =
    val resolved = cell.value match
      case CellValue.Number(n) => CellValue.Number(n)
      case CellValue.Text(s) => CellValue.Text(s)
      case CellValue.Bool(b) => CellValue.Bool(b)
      case CellValue.DateTime(dt) => CellValue.DateTime(dt)
      case CellValue.RichText(rt) => CellValue.Text(rt.toPlainText)
      case CellValue.Formula(_, cached) =>
        cached match
          case Some(CellValue.Number(n)) => CellValue.Number(n)
          case Some(CellValue.Text(s)) => CellValue.Text(s)
          case Some(CellValue.Bool(b)) => CellValue.Bool(b)
          case Some(CellValue.DateTime(dt)) => CellValue.DateTime(dt)
          case Some(CellValue.RichText(rt)) => CellValue.Text(rt.toPlainText)
          case _ => CellValue.Number(BigDecimal(0))
      case CellValue.Error(err) => CellValue.Error(err)
      case CellValue.Empty => CellValue.Number(BigDecimal(0))
    scala.util.Right(resolved)

  // ===== Type-Coercing Decoders (Excel-compatible automatic conversion) =====

  /**
   * Decode cell as String with automatic type coercion.
   *
   * Matches Excel semantics:
   *   - Text -> as-is
   *   - Number -> toString (42 -> "42")
   *   - Boolean -> toString (true -> "TRUE", false -> "FALSE")
   *   - DateTime -> ISO format
   *   - Formula -> text representation
   *   - Empty -> empty string
   */
  def decodeAsString(cell: Cell): Either[CodecError, String] =
    cell.value match
      case CellValue.Empty => scala.util.Right("")
      case CellValue.Text(s) => scala.util.Right(s)
      case CellValue.Number(n) => scala.util.Right(n.toString)
      case CellValue.Bool(b) => scala.util.Right(if b then "TRUE" else "FALSE")
      case CellValue.DateTime(dt) => scala.util.Right(dt.toString)
      case CellValue.Formula(text, _) => scala.util.Right(text)
      case CellValue.RichText(rt) => scala.util.Right(rt.toPlainText)
      case other => scala.util.Left(CodecError.TypeMismatch("String", other))

  /**
   * Decode cell as LocalDate with automatic type coercion.
   *
   * Matches Excel semantics:
   *   - DateTime -> extract date component
   *   - Number -> Excel serial number -> date (GH-385: professional workbooks store dates
   *     exclusively as serials with date numFmts; guarded to 0..MaxExcelDateSerial like
   *     ScalarCoercion.coerceDate, 1900 date system — the evaluator-wide assumption, see
   *     docs/LIMITATIONS.md)
   *   - Formula -> cached DateTime or cached serial Number, same conversions
   *   - Text -> error (Excel does not coerce arbitrary text in date positions)
   */
  def decodeAsDate(cell: Cell): Either[CodecError, java.time.LocalDate] =
    def serialToDate(serial: BigDecimal): Option[java.time.LocalDate] =
      if serial >= 0 && serial <= ScalarCoercion.MaxExcelDateSerial then
        Some(CellValue.excelSerialToDateTime(serial.toDouble).toLocalDate)
      else None

    cell.value match
      case CellValue.DateTime(dt) => scala.util.Right(dt.toLocalDate)
      // Out-of-range serials (negative or beyond 9999-12-31) fold to a clean error, mirroring
      // ScalarCoercion.coerceDate's guard
      case CellValue.Number(serial) =>
        serialToDate(serial).toRight(CodecError.TypeMismatch("Date", cell.value))
      case CellValue.Formula(_, Some(CellValue.DateTime(cached))) =>
        scala.util.Right(cached.toLocalDate)
      case CellValue.Formula(_, Some(CellValue.Number(cached))) =>
        serialToDate(cached).toRight(CodecError.TypeMismatch("Date", cell.value))
      case other =>
        scala.util.Left(
          CodecError.TypeMismatch(
            expected = "Date",
            actual = other
          )
        )

  /**
   * Decode cell as Int with automatic type coercion.
   *
   * Matches Excel semantics:
   *   - Number -> toInt if valid
   *   - Boolean -> 1 for TRUE, 0 for FALSE
   *   - Text -> parse as number (not yet implemented)
   *   - Other -> error
   */
  def decodeAsInt(cell: Cell): Either[CodecError, Int] =
    cell.value match
      case CellValue.Number(n) if n.isValidInt => scala.util.Right(n.toInt)
      case CellValue.Bool(b) => scala.util.Right(ArrayArithmetic.boolToNumeric(b).toInt)
      case CellValue.Number(n) =>
        scala.util.Left(
          CodecError.TypeMismatch(
            expected = "Int",
            actual = CellValue.Number(n)
          )
        )
      case other =>
        scala.util.Left(
          CodecError.TypeMismatch(
            expected = "Int",
            actual = other
          )
        )
