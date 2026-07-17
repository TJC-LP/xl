package com.tjclp.xl.formula.eval

import com.tjclp.xl.formula.ast.BindingCoercion

import com.tjclp.xl.cells.CellValue
import scala.math.BigDecimal

/**
 * GH-302/GH-306/GH-307: the single total scalar-coercion table for typed argument positions.
 *
 * Shared by [[TExpr.CoercedBindingRef]] (LET bindings, the GH-193 precedent this generalizes),
 * [[TExpr.Coerced]] (runtime-polymorphic expressions: call results, aggregates, cross-typed
 * operators) and `FunctionSpecsBase.toIntArg` (EDATE/EOMONTH/WORKDAY-family integer arguments).
 *
 * Each target mirrors the conventions of the corresponding cell decoder (decodeAsString,
 * decodeAsInt, decodeBool, decodeNumeric, decodeAsDate) plus Excel's value coercions where the
 * decoders are stricter than Excel:
 *   - number → text renders via toString (the decodeAsString/concatText convention)
 *   - number → boolean is zero/non-zero (Excel: 0 = FALSE, anything else = TRUE)
 *   - numeric text → number/integer parses ("3" coerces, "abc" is a clean error — Excel #VALUE!)
 *   - fractional → integer TRUNCATES toward zero (Excel truncates months/days/num_chars)
 *   - boolean → number is TRUE=1/FALSE=0; dates ARE numbers (Excel serial)
 *   - Empty → ""/0/FALSE per target (the decodeResolvedValue zero convention)
 *
 * CellValue inputs unwrap first (cached formula values extracted), so call results that surface raw
 * cell values (IF branches, lookup results) coerce identically to primitives. Error VALUES (#REF!
 * and friends) refuse to coerce with a clean per-cell error naming the Excel error code.
 * ArrayResult is deliberately NOT handled here — the collapse-vs-broadcast policy belongs to the
 * evaluation positions (scalar argument positions collapse to top-left, operand positions pass
 * arrays through to the broadcasting machinery).
 *
 * Total: every input yields Right(coerced) or Left(TypeMismatch/EvalFailed); never a thrown
 * exception, never a ClassCastException deferred to the consuming function.
 */
private[formula] object ScalarCoercion:

  /**
   * Largest Excel date serial (9999-12-31); guards excelSerialToDateTime against overflow. Shared
   * with TExprDecoders.decodeAsDate (GH-385) so the direct-cell and Coerced date boundaries accept
   * the same serial domain.
   */
  val MaxExcelDateSerial: BigDecimal = BigDecimal(2958465)

  /**
   * GH-396: the date a BLANK yields in a date position. Excel renders a blank date argument as
   * serial 0 — the phantom "January 0, 1900" (=YEAR(blank) is 1900, =MONTH(blank) is 1, =DAY(blank)
   * is 0). LocalDate cannot represent a day-0 date and the CellValue serial mapping normalizes
   * serial 0 to 1899-12-31 (the pre-leap-bug shift), so neither reproduces Excel's observable
   * year/month; the closest representable date is 1900-01-01 — YEAR and MONTH are Excel-exact, DAY
   * reads 1 instead of 0 (documented divergence). Shared with TExprDecoders.decodeAsDate (the
   * tables are mirrors). NOTE: deliberately NOT the serial-0 mapping — Bool FALSE keeps the GH-307
   * serial convention (1899-12-31) while blank matches Excel's observed blank-argument outputs.
   */
  val BlankDate: java.time.LocalDate = java.time.LocalDate.of(1900, 1, 1)

  /** Collapse an ArrayResult to its scalar value: top-left, Empty when empty (GH-302). */
  def collapseArray(ar: ArrayResult): CellValue =
    if ar.isEmpty then CellValue.Empty else ar(0, 0)

  /**
   * GH-344 item 5: Excel coerces exactly the text literals "TRUE"/"FALSE" (case-insensitive, NO
   * trim — `" TRUE"` refuses; strict is loosening-safe) in condition positions. The single
   * recognition table shared by [[coerceBool]], `TExprDecoders.decodeBool` and
   * `ArrayArithmetic.conditionTruthy` — the three condition tables must stay aligned (see the L6
   * parity law). Documented accepted micro-divergence: cell-sourced boolean text coerces too (Excel
   * coerces only literals; provenance is invisible at the decode layer — the direct/bound parity
   * law wins).
   */
  private[formula] def boolTextValue(s: String): Option[Boolean] =
    if s.equalsIgnoreCase("TRUE") then Some(true)
    else if s.equalsIgnoreCase("FALSE") then Some(false)
    else None

  /**
   * Totally coerce a runtime value into a typed argument position.
   *
   * @param label
   *   Position description for error messages (e.g. "LET binding 'x'", "text argument")
   */
  def coerce(label: String, value: Any, target: BindingCoercion): Either[EvalError, Any] =
    unwrapCellValue(value) match
      // GH-344: an Excel error VALUE entering a typed position propagates AS that error (the
      // strict-position absorption rule) — the single highest-leverage arm: every typed argument
      // position, IF/IFS scalar conditions, toIntArg, LET/Coerced positions, and the scalar-entry
      // top-left collapse all funnel through here.
      case CellValue.Error(err) =>
        Left(EvalError.ErrorValue(err, Some(label)))
      case unwrapped =>
        target match
          case BindingCoercion.Text => coerceText(label, unwrapped)
          case BindingCoercion.Integer => coerceInteger(label, unwrapped)
          case BindingCoercion.Bool => coerceBool(label, unwrapped)
          case BindingCoercion.Numeric => coerceNumeric(label, unwrapped)
          case BindingCoercion.Date => coerceDate(label, unwrapped)

  /**
   * Unwrap CellValue to its primitive (cached formula extracted, RichText flattened) so cell-shaped
   * call results coerce like primitives. Empty stays CellValue.Empty (per-target conventions); an
   * uncached formula value follows the decodeResolvedValue zero convention.
   */
  private def unwrapCellValue(value: Any): Any = value match
    case CellValue.Number(n) => n
    case CellValue.Text(s) => s
    case CellValue.Bool(b) => b
    case CellValue.DateTime(dt) => dt
    case CellValue.RichText(rt) => rt.toPlainText
    case CellValue.Formula(_, Some(cached)) => unwrapCellValue(cached)
    case CellValue.Formula(_, None) => BigDecimal(0)
    case other => other

  private def mismatch(label: String, expected: String, value: Any): Either[EvalError, Any] =
    Left(EvalError.TypeMismatch(label, expected, s"$value"))

  private def coerceText(label: String, value: Any): Either[EvalError, Any] = value match
    case s: String => Right(s)
    case bd: BigDecimal => Right(bd.toString)
    case i: Int => Right(i.toString)
    case b: Boolean => Right(if b then "TRUE" else "FALSE")
    case ld: java.time.LocalDate => Right(ld.toString)
    case ldt: java.time.LocalDateTime => Right(ldt.toString)
    case CellValue.Empty => Right("")
    case other => mismatch(label, "text", other)

  private def coerceInteger(label: String, value: Any): Either[EvalError, Any] = value match
    case bd: BigDecimal => truncateToInt(label, bd)
    case i: Int => Right(i)
    case b: Boolean => Right(if b then 1 else 0)
    case s: String =>
      parseNumericText(s) match
        case Some(bd) => truncateToInt(label, bd)
        case None => mismatch(label, "integer", s)
    case CellValue.Empty => Right(0)
    case other => mismatch(label, "integer", other)

  private def coerceBool(label: String, value: Any): Either[EvalError, Any] = value match
    case b: Boolean => Right(b)
    // Excel truthiness: 0 = FALSE, any other number = TRUE (GH-306)
    case bd: BigDecimal => Right(bd.signum != 0)
    case i: Int => Right(i != 0)
    // GH-344 item 5: exactly "TRUE"/"FALSE" (case-insensitive, no trim) coerce; other text
    // refuses — the [[boolTextValue]] table, aligned with decodeBool and conditionTruthy
    case s: String =>
      boolTextValue(s) match
        case Some(b) => Right(b)
        case None => mismatch(label, "boolean", s)
    case CellValue.Empty => Right(false)
    case other => mismatch(label, "boolean", other)

  private def coerceNumeric(label: String, value: Any): Either[EvalError, Any] = value match
    case bd: BigDecimal => Right(bd)
    case i: Int => Right(BigDecimal(i))
    case b: Boolean => Right(if b then BigDecimal(1) else BigDecimal(0))
    case ld: java.time.LocalDate =>
      Right(BigDecimal(CellValue.dateTimeToExcelSerial(ld.atStartOfDay())))
    case ldt: java.time.LocalDateTime =>
      Right(BigDecimal(CellValue.dateTimeToExcelSerial(ldt)))
    case s: String =>
      parseNumericText(s) match
        case Some(bd) => Right(bd)
        case None => mismatch(label, "number", s)
    case CellValue.Empty => Right(BigDecimal(0))
    case other => mismatch(label, "number", other)

  private def coerceDate(label: String, value: Any): Either[EvalError, Any] = value match
    case ld: java.time.LocalDate => Right(ld)
    case ldt: java.time.LocalDateTime => Right(ldt.toLocalDate)
    case bd: BigDecimal if bd >= 0 && bd <= MaxExcelDateSerial =>
      Right(CellValue.excelSerialToDateTime(bd.toDouble).toLocalDate)
    case i: Int if i >= 0 && BigDecimal(i) <= MaxExcelDateSerial =>
      Right(CellValue.excelSerialToDateTime(i.toDouble).toLocalDate)
    // GH-307: booleans are their Excel serial in date positions too (TRUE=1 → 1900-01-01,
    // so =YEAR(TRUE) is 1900 like Excel) — delegate so they match the numeric branch exactly
    case b: Boolean => coerceDate(label, if b then BigDecimal(1) else BigDecimal(0))
    // GH-396: a blank in a date position renders like Excel's serial-0 year/month (see
    // [[BlankDate]]) — previously a clean error while decodeAsDate/other targets accepted Empty
    case CellValue.Empty => Right(BlankDate)
    case other => mismatch(label, "date", other)

  /** Excel truncates fractional values toward zero in integer positions; guard the Int range. */
  private def truncateToInt(label: String, bd: BigDecimal): Either[EvalError, Any] =
    val truncated = bd.setScale(0, scala.math.BigDecimal.RoundingMode.DOWN)
    if truncated.isValidInt then Right(truncated.toInt)
    else mismatch(label, "valid integer", bd)

  /**
   * Numeric-text parse (the extractNumericForMatch convention): trimmed, total. Shared with
   * TExprDecoders.decodeAsInt (GH-396) so the direct-cell and Coerced integer boundaries accept the
   * same text domain.
   */
  private[formula] def parseNumericText(s: String): Option[BigDecimal] =
    scala.util.Try(BigDecimal(s.trim)).toOption
