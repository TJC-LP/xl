package com.tjclp.xl.formula.ast

import com.tjclp.xl.formula.functions.FunctionSpecs
import com.tjclp.xl.formula.eval.EvalError
import com.tjclp.xl.formula.functions.EvalContext

import com.tjclp.xl.cells.{Cell, CellValue}
import com.tjclp.xl.codec.CodecError

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
   */
  def decodeBool(cell: Cell): Either[CodecError, Boolean] =
    cell.value match
      case CellValue.Bool(value) => scala.util.Right(value)
      case CellValue.Formula(_, Some(CellValue.Bool(cached))) =>
        scala.util.Right(cached)
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
   *   - Number -> interpret as Excel serial number (not yet implemented)
   *   - Text -> parse as ISO date (not yet implemented)
   *   - Other -> error
   */
  def decodeAsDate(cell: Cell): Either[CodecError, java.time.LocalDate] =
    cell.value match
      case CellValue.DateTime(dt) => scala.util.Right(dt.toLocalDate)
      case other =>
        // Future: could add Excel serial number -> date conversion for Number
        // Future: could add ISO date string parsing for Text
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
      case CellValue.Bool(b) => scala.util.Right(if b then 1 else 0)
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
