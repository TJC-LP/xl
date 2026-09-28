package com.tjclp.xl.formula.eval

import com.tjclp.xl.formula.ast.BindingCoercion

import com.tjclp.xl.cells.{CellError, CellValue}
import com.tjclp.xl.display.NumFmtFormatter
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
 *   - number → text renders via numberText (Excel's General text conversion, GH-665 — the
 *     decodeAsString/concatText convention)
 *   - number → boolean is zero/non-zero (Excel: 0 = FALSE, anything else = TRUE)
 *   - numeric text → number/integer parses ("3" coerces, "abc" is a clean error — Excel #VALUE!)
 *   - fractional → integer TRUNCATES toward zero (Excel truncates months/days/num_chars)
 *   - boolean → number is TRUE=1/FALSE=0; dates ARE numbers (Excel serial)
 *   - Empty → ""/0/FALSE per target (the decodeResolvedValue zero convention)
 *
 * CellValue inputs unwrap first (cached formula values extracted), so call results that surface raw
 * cell values (IF branches, lookup results) coerce identically to primitives. Error VALUES (#REF!
 * and friends) refuse to coerce with a clean per-cell error naming the Excel error code.
 * ArrayResult is deliberately NOT handled by [[coerce]] — the collapse-vs-broadcast policy belongs
 * to the evaluation positions (scalar argument positions collapse to top-left through
 * [[collapseTo]], operand positions pass arrays through to the broadcasting machinery).
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

  /**
   * GH-561: the text form of a date in a TEXT position (`&`, CONCATENATE, LEN, ...) is its Excel
   * serial number, never an ISO rendering — dates ARE numbers in Excel's value model and only
   * TEXT() formats them, so `">="&DATE(2026,1,1)` is `">=46023"` and matches numeric date cells in
   * COUNTIFS/SUMIFS criteria. Whole days print as integers ("46023"); times keep their fraction
   * ("46023.5"), rounded to Excel's 15 significant digits.
   *
   * 1900 date system, like every other serial conversion in the evaluator (docs/LIMITATIONS.md): a
   * `date1904` workbook renders the 1900-system serial here until the evaluator carries the
   * workbook's date system.
   */
  def dateSerialText(dt: java.time.LocalDateTime): String =
    numberText(BigDecimal(CellValue.dateTimeToExcelSerial(dt)))

  /**
   * GH-665: the evaluator's single number → text conversion — Excel's width-independent General
   * text (15 significant digits, trailing zeros stripped, plain up to 20 characters then E
   * notation), never `BigDecimal.toString`, which carries the stored scale (`<v>2.0</v>` would read
   * "2.0", `SUM` of scaled cells "1070.0", `1/3` 34 digits). Every implicit coercion (`&`,
   * CONCATENATE, text-typed arguments, numeric literals in text positions) routes through here; the
   * rule itself lives in xl-core so TEXT(x,"General") shares it.
   */
  def numberText(n: BigDecimal): String = NumFmtFormatter.generalText(n)

  /** GH-561: see the LocalDateTime overload — a date-only value is its whole-day serial. */
  def dateSerialText(ld: java.time.LocalDate): String = dateSerialText(ld.atStartOfDay())

  /** Collapse an ArrayResult to its scalar value: top-left, Empty when empty (GH-302). */
  def collapseArray(ar: ArrayResult): CellValue =
    if ar.isEmpty then CellValue.Empty else ar(0, 0)

  /**
   * The typed scalar collapse: an array reaching a scalar position yields its top-left value,
   * coerced to `kind` — the static scalar kind of the node that produced it (TExpr.scalarKind) — so
   * an arithmetic node's collapsed element reaches its consumer as a number or a Left, never as the
   * raw CellValue a typed function body would cast. With no kind (Any and CellValue positions,
   * tolerant by design) the raw element passes through.
   */
  def collapseTo(
    label: String,
    ar: ArrayResult,
    kind: Option[BindingCoercion]
  ): Either[EvalError, Any] =
    val top = collapseArray(ar)
    kind.fold[Either[EvalError, Any]](Right(top))(coerce(label, top, _))

  /**
   * Total text coercion for '&' operands and elements, mirroring the decodeAsString conventions:
   * Number → General text via [[numberText]] (GH-665: 2.0 → "2", never the stored scale), Bool →
   * TRUE/FALSE, DateTime → Excel serial (GH-561), Empty → "", rich text → its plain text, a cached
   * formula → its cached value. Callers check carried error values first (they propagate, never
   * stringify); arrays broadcast element-wise before reaching here.
   */
  def concatText(value: Any): String = unwrapCellValue(value) match
    case s: String => s
    case b: Boolean => if b then "TRUE" else "FALSE"
    case bd: BigDecimal => numberText(bd)
    case i: Int => i.toString
    // anyToCellValue admits Long/Double runtime values into Any positions — render them as numbers
    case l: Long => numberText(BigDecimal(l))
    // #681: a non-finite Double never gets here — both callers first read the operand through
    // ArrayArithmetic.anyToCellValue, which makes it #NUM!, and propagate that error
    case d: Double if d.isFinite => numberText(BigDecimal(d))
    // GH-561: `&` on a date yields its Excel serial ("46023"), never ISO text — dates are
    // numbers; only TEXT() formats them (the `">="&DATE(y,m,d)` criteria idiom depends on it)
    case ld: java.time.LocalDate => dateSerialText(ld)
    case ldt: java.time.LocalDateTime => dateSerialText(ldt)
    case CellValue.Empty => ""
    case other => other.toString

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
    case CellValue.Formula(_, Some(cached), _) => unwrapCellValue(cached)
    case CellValue.Formula(_, None, _) => BigDecimal(0)
    case other => other

  private def mismatch(label: String, expected: String, value: Any): Either[EvalError, Any] =
    Left(EvalError.TypeMismatch(label, expected, s"$value"))

  private def coerceText(label: String, value: Any): Either[EvalError, Any] = value match
    case s: String => Right(s)
    case bd: BigDecimal => Right(numberText(bd))
    case i: Int => Right(i.toString)
    // the same table concatText keeps: Long/Double runtime values are numbers under the one rule
    case l: Long => Right(numberText(BigDecimal(l)))
    case d: Double if d.isFinite => Right(numberText(BigDecimal(d)))
    // #681: NaN / ±Infinity have no Excel value — #NUM!, the answer `&` gives too
    case _: Double => Left(EvalError.ErrorValue(CellError.Num, Some(label)))
    case b: Boolean => Right(if b then "TRUE" else "FALSE")
    // GH-561: dates render as their Excel serial in text positions
    case ld: java.time.LocalDate => Right(dateSerialText(ld))
    case ldt: java.time.LocalDateTime => Right(dateSerialText(ldt))
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
    // GH-564: dates are numbers in Excel's value model — their serial is zero/non-zero, so
    // =IF(A1,...) and =OR(A1) on a date cell are TRUE, as in Excel
    case dt: java.time.LocalDateTime => Right(CellValue.dateTimeToExcelSerial(dt) != 0.0)
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
