package com.tjclp.xl.formula

import com.tjclp.xl.{ARef, Anchor, CellRange, SheetName}
import com.tjclp.xl.addressing.{Column, Row}

/**
 * Printer for TExpr AST to Excel formula strings.
 *
 * Produces canonical, deterministic output for round-trip verification. Pure functional - no
 * mutation, no side effects.
 *
 * Ensures proper operator precedence with minimal parentheses.
 *
 * Round-trip law: parse(print(expr)) == Right(expr)
 *
 * @note
 *   Suppression rationale:
 *   - AsInstanceOf: ARef is opaque type over Long. Cast is safe for printing.
 */
@SuppressWarnings(Array("org.wartremover.warts.AsInstanceOf"))
object FormulaPrinter:
  /**
   * Convert TExpr to Excel formula string.
   *
   * @param expr
   *   The expression to print
   * @param includeEquals
   *   Whether to include leading '=' (default: true)
   * @return
   *   Formula string in Excel syntax
   *
   * Example:
   * {{{
   * print(TExpr.Add(TExpr.Lit(1), TExpr.Lit(2))) // "=1+2"
   * print(TExpr.sum(CellRange("A1:B10")))        // "=SUM(A1:B10)"
   * }}}
   */
  def print(expr: TExpr[?], includeEquals: Boolean = true): String =
    val formula = printExpr(expr, precedence = 0)
    if includeEquals then s"=$formula" else formula

  /**
   * Expression precedence levels (higher = tighter binding).
   *
   * Used to determine when parentheses are needed.
   */
  private object Precedence:
    val Or = 1
    val And = 2
    val Comparison = 3
    val Concat = 4
    val AddSub = 5
    val MulDiv = 6
    val Unary = 7
    val Primary = 8

  /**
   * Print expression with appropriate parentheses based on precedence.
   *
   * @param expr
   *   The expression to print
   * @param precedence
   *   The precedence level of the enclosing context
   * @return
   *   Formula string (without leading '=')
   */
  private def printExpr(expr: TExpr[?], precedence: Int): String =
    expr match
      // Literals
      case TExpr.Lit(value: BigDecimal) => value.toString
      case TExpr.Lit(value: Boolean) => if value then "TRUE" else "FALSE"
      case TExpr.Lit(value: String) => s""""${escapeString(value)}""""
      case TExpr.Lit(value: Int) => value.toString
      case TExpr.Lit(value) => value.toString

      // Cell reference
      case TExpr.Ref(at, anchor, _) => formatARef(at, anchor)
      case TExpr.PolyRef(at, anchor) => formatARef(at, anchor) // PolyRef prints same as Ref

      // Sheet-qualified references
      case TExpr.SheetRef(sheet, at, anchor, _) =>
        s"${formatSheetName(sheet)}!${formatARef(at, anchor)}"
      case TExpr.SheetPolyRef(sheet, at, anchor) =>
        s"${formatSheetName(sheet)}!${formatARef(at, anchor)}"
      case TExpr.SheetRange(sheet, range) =>
        s"${formatSheetName(sheet)}!${range.toA1}"

      // Conditional
      case TExpr.If(cond, ifTrue, ifFalse) =>
        s"IF(${printExpr(cond, 0)}, ${printExpr(ifTrue, 0)}, ${printExpr(ifFalse, 0)})"

      // Arithmetic operators
      case TExpr.Add(x, y) =>
        val result = s"${printExpr(x, Precedence.AddSub)}+${printExpr(y, Precedence.AddSub)}"
        parenthesizeIf(result, precedence > Precedence.AddSub)

      case TExpr.Sub(TExpr.Lit(n: BigDecimal), y) if n == BigDecimal(0) =>
        // Unary minus: 0-x prints as -x
        val result = s"-${printExpr(y, Precedence.Unary)}"
        parenthesizeIf(result, precedence > Precedence.Unary)

      case TExpr.Sub(x, y) =>
        val result = s"${printExpr(x, Precedence.AddSub)}-${printExpr(y, Precedence.AddSub)}"
        parenthesizeIf(result, precedence > Precedence.AddSub)

      case TExpr.Mul(x, y) =>
        val result = s"${printExpr(x, Precedence.MulDiv)}*${printExpr(y, Precedence.MulDiv)}"
        parenthesizeIf(result, precedence > Precedence.MulDiv)

      case TExpr.Div(x, y) =>
        val result = s"${printExpr(x, Precedence.MulDiv)}/${printExpr(y, Precedence.MulDiv)}"
        parenthesizeIf(result, precedence > Precedence.MulDiv)

      // Boolean operators
      case TExpr.And(x, y) =>
        val result = s"AND(${printExpr(x, 0)}, ${printExpr(y, 0)})"
        result // Functions don't need precedence-based parens

      case TExpr.Or(x, y) =>
        val result = s"OR(${printExpr(x, 0)}, ${printExpr(y, 0)})"
        result

      case TExpr.Not(x) =>
        s"NOT(${printExpr(x, 0)})"

      // Comparison operators
      case TExpr.Eq(x, y) =>
        val result =
          s"${printExpr(x, Precedence.Comparison)}=${printExpr(y, Precedence.Comparison)}"
        parenthesizeIf(result, precedence > Precedence.Comparison)

      case TExpr.Neq(x, y) =>
        val result =
          s"${printExpr(x, Precedence.Comparison)}<>${printExpr(y, Precedence.Comparison)}"
        parenthesizeIf(result, precedence > Precedence.Comparison)

      case TExpr.Lt(x, y) =>
        val result =
          s"${printExpr(x, Precedence.Comparison)}<${printExpr(y, Precedence.Comparison)}"
        parenthesizeIf(result, precedence > Precedence.Comparison)

      case TExpr.Lte(x, y) =>
        val result =
          s"${printExpr(x, Precedence.Comparison)}<=${printExpr(y, Precedence.Comparison)}"
        parenthesizeIf(result, precedence > Precedence.Comparison)

      case TExpr.Gt(x, y) =>
        val result =
          s"${printExpr(x, Precedence.Comparison)}>${printExpr(y, Precedence.Comparison)}"
        parenthesizeIf(result, precedence > Precedence.Comparison)

      case TExpr.Gte(x, y) =>
        val result =
          s"${printExpr(x, Precedence.Comparison)}>=${printExpr(y, Precedence.Comparison)}"
        parenthesizeIf(result, precedence > Precedence.Comparison)

      // Type conversions (print transparently - ToInt is internal)
      case TExpr.ToInt(expr) =>
        printExpr(expr, precedence) // Print wrapped expression without conversion syntax

      // Text functions
      case TExpr.Concatenate(xs) =>
        val args = xs.map(x => printExpr(x, 0)).mkString(", ")
        s"CONCATENATE($args)"

      case TExpr.Left(text, n) =>
        s"LEFT(${printExpr(text, 0)}, ${printExpr(n, 0)})"

      case TExpr.Right(text, n) =>
        s"RIGHT(${printExpr(text, 0)}, ${printExpr(n, 0)})"

      case TExpr.Len(text) =>
        s"LEN(${printExpr(text, 0)})"

      case TExpr.Upper(text) =>
        s"UPPER(${printExpr(text, 0)})"

      case TExpr.Lower(text) =>
        s"LOWER(${printExpr(text, 0)})"

      // Date/Time functions
      case TExpr.Today() =>
        "TODAY()"

      case TExpr.Now() =>
        "NOW()"

      // Date-to-serial converters (print transparently - internal conversion)
      case TExpr.DateToSerial(expr) =>
        printExpr(expr, precedence)

      case TExpr.DateTimeToSerial(expr) =>
        printExpr(expr, precedence)

      // Math constants
      case TExpr.Pi() =>
        "PI()"

      case TExpr.Date(year, month, day) =>
        s"DATE(${printExpr(year, 0)}, ${printExpr(month, 0)}, ${printExpr(day, 0)})"

      case TExpr.Year(date) =>
        s"YEAR(${printExpr(date, 0)})"

      case TExpr.Month(date) =>
        s"MONTH(${printExpr(date, 0)})"

      case TExpr.Day(date) =>
        s"DAY(${printExpr(date, 0)})"

      // Date calculation functions
      case TExpr.Eomonth(startDate, months) =>
        s"EOMONTH(${printExpr(startDate, 0)}, ${printExpr(months, 0)})"

      case TExpr.Edate(startDate, months) =>
        s"EDATE(${printExpr(startDate, 0)}, ${printExpr(months, 0)})"

      case TExpr.Datedif(startDate, endDate, unit) =>
        s"DATEDIF(${printExpr(startDate, 0)}, ${printExpr(endDate, 0)}, ${printExpr(unit, 0)})"

      case TExpr.Networkdays(startDate, endDate, holidaysOpt) =>
        holidaysOpt match
          case Some(holidays) =>
            s"NETWORKDAYS(${printExpr(startDate, 0)}, ${printExpr(endDate, 0)}, ${formatRange(holidays)})"
          case None =>
            s"NETWORKDAYS(${printExpr(startDate, 0)}, ${printExpr(endDate, 0)})"

      case TExpr.Workday(startDate, days, holidaysOpt) =>
        holidaysOpt match
          case Some(holidays) =>
            s"WORKDAY(${printExpr(startDate, 0)}, ${printExpr(days, 0)}, ${formatRange(holidays)})"
          case None =>
            s"WORKDAY(${printExpr(startDate, 0)}, ${printExpr(days, 0)})"

      case TExpr.Yearfrac(startDate, endDate, basis) =>
        // Only print basis if non-default (0)
        basis match
          case TExpr.Lit(0) =>
            s"YEARFRAC(${printExpr(startDate, 0)}, ${printExpr(endDate, 0)})"
          case _ =>
            s"YEARFRAC(${printExpr(startDate, 0)}, ${printExpr(endDate, 0)}, ${printExpr(basis, 0)})"

      // Arithmetic range functions (now support cross-sheet via RangeLocation)
      case TExpr.Sum(range) =>
        s"SUM(${formatLocation(range)})"

      case TExpr.Count(range) =>
        s"COUNT(${formatLocation(range)})"

      case TExpr.Min(range) =>
        s"MIN(${formatLocation(range)})"

      case TExpr.Max(range) =>
        s"MAX(${formatLocation(range)})"

      case TExpr.Average(range) =>
        s"AVERAGE(${formatLocation(range)})"

      // Unified aggregate function (typeclass-based)
      case TExpr.Aggregate(aggregatorId, location) =>
        s"$aggregatorId(${formatLocation(location)})"

      case call: TExpr.Call[?] =>
        val printer = ArgPrinter(
          expr = expr => printExpr(expr, precedence = 0),
          location = formatLocation,
          cellRange = formatRange
        )
        call.spec.render(call.args, printer)

      // Cross-sheet aggregate functions (legacy, kept for backward compatibility)
      case TExpr.SheetSum(sheet, range) =>
        s"SUM(${formatSheetName(sheet)}!${formatRange(range)})"

      case TExpr.SheetMin(sheet, range) =>
        s"MIN(${formatSheetName(sheet)}!${formatRange(range)})"

      case TExpr.SheetMax(sheet, range) =>
        s"MAX(${formatSheetName(sheet)}!${formatRange(range)})"

      case TExpr.SheetAverage(sheet, range) =>
        s"AVERAGE(${formatSheetName(sheet)}!${formatRange(range)})"

      case TExpr.SheetCount(sheet, range) =>
        s"COUNT(${formatSheetName(sheet)}!${formatRange(range)})"

      // Financial functions (now support cross-sheet via RangeLocation)
      case TExpr.Npv(rate, values) =>
        s"NPV(${printExpr(rate, 0)}, ${formatLocation(values)})"

      case TExpr.Irr(values, guessOpt) =>
        val rangeText = formatLocation(values)
        guessOpt match
          case Some(guess) => s"IRR($rangeText, ${printExpr(guess, 0)})"
          case None => s"IRR($rangeText)"

      case TExpr.Xnpv(rate, values, dates) =>
        s"XNPV(${printExpr(rate, 0)}, ${formatLocation(values)}, ${formatLocation(dates)})"

      case TExpr.Xirr(values, dates, guessOpt) =>
        val valuesText = formatLocation(values)
        val datesText = formatLocation(dates)
        guessOpt match
          case Some(guess) => s"XIRR($valuesText, $datesText, ${printExpr(guess, 0)})"
          case None => s"XIRR($valuesText, $datesText)"

      // TVM Functions
      case TExpr.Pmt(rate, nper, pv, fv, pmtType) =>
        val args = List(printExpr(rate, 0), printExpr(nper, 0), printExpr(pv, 0)) ++
          fv.map(printExpr(_, 0)).toList ++
          pmtType.map(printExpr(_, 0)).toList
        s"PMT(${args.mkString(", ")})"

      case TExpr.Fv(rate, nper, pmt, pv, pmtType) =>
        val args = List(printExpr(rate, 0), printExpr(nper, 0), printExpr(pmt, 0)) ++
          pv.map(printExpr(_, 0)).toList ++
          pmtType.map(printExpr(_, 0)).toList
        s"FV(${args.mkString(", ")})"

      case TExpr.Pv(rate, nper, pmt, fv, pmtType) =>
        val args = List(printExpr(rate, 0), printExpr(nper, 0), printExpr(pmt, 0)) ++
          fv.map(printExpr(_, 0)).toList ++
          pmtType.map(printExpr(_, 0)).toList
        s"PV(${args.mkString(", ")})"

      case TExpr.Nper(rate, pmt, pv, fv, pmtType) =>
        val args = List(printExpr(rate, 0), printExpr(pmt, 0), printExpr(pv, 0)) ++
          fv.map(printExpr(_, 0)).toList ++
          pmtType.map(printExpr(_, 0)).toList
        s"NPER(${args.mkString(", ")})"

      case TExpr.Rate(nper, pmt, pv, fv, pmtType, guess) =>
        val args = List(printExpr(nper, 0), printExpr(pmt, 0), printExpr(pv, 0)) ++
          fv.map(printExpr(_, 0)).toList ++
          pmtType.map(printExpr(_, 0)).toList ++
          guess.map(printExpr(_, 0)).toList
        s"RATE(${args.mkString(", ")})"

      case TExpr.VLookup(lookup, table, colIndex, rangeLookup) =>
        s"VLOOKUP(${printExpr(lookup, 0)}, " +
          s"${formatLocation(table)}, " +
          s"${printExpr(colIndex, 0)}, ${printExpr(rangeLookup, 0)})"

      // Error handling functions
      case TExpr.Iferror(value, valueIfError) =>
        s"IFERROR(${printExpr(value, 0)}, ${printExpr(valueIfError, 0)})"

      case TExpr.Iserror(value) =>
        s"ISERROR(${printExpr(value, 0)})"

      case TExpr.Iserr(value) =>
        s"ISERR(${printExpr(value, 0)})"

      // Type-check functions
      case TExpr.Isnumber(value) =>
        s"ISNUMBER(${printExpr(value, 0)})"

      case TExpr.Istext(value) =>
        s"ISTEXT(${printExpr(value, 0)})"

      case TExpr.Isblank(value) =>
        s"ISBLANK(${printExpr(value, 0)})"

      // Rounding and math functions
      case TExpr.Round(value, numDigits) =>
        s"ROUND(${printExpr(value, 0)}, ${printExpr(numDigits, 0)})"

      case TExpr.RoundUp(value, numDigits) =>
        s"ROUNDUP(${printExpr(value, 0)}, ${printExpr(numDigits, 0)})"

      case TExpr.RoundDown(value, numDigits) =>
        s"ROUNDDOWN(${printExpr(value, 0)}, ${printExpr(numDigits, 0)})"

      case TExpr.Abs(value) =>
        s"ABS(${printExpr(value, 0)})"

      case TExpr.Sqrt(value) =>
        s"SQRT(${printExpr(value, 0)})"

      case TExpr.Mod(number, divisor) =>
        s"MOD(${printExpr(number, 0)}, ${printExpr(divisor, 0)})"

      case TExpr.Power(number, power) =>
        s"POWER(${printExpr(number, 0)}, ${printExpr(power, 0)})"

      case TExpr.Log(number, base) =>
        s"LOG(${printExpr(number, 0)}, ${printExpr(base, 0)})"

      case TExpr.Ln(value) =>
        s"LN(${printExpr(value, 0)})"

      case TExpr.Exp(value) =>
        s"EXP(${printExpr(value, 0)})"

      case TExpr.Floor(number, significance) =>
        s"FLOOR(${printExpr(number, 0)}, ${printExpr(significance, 0)})"

      case TExpr.Ceiling(number, significance) =>
        s"CEILING(${printExpr(number, 0)}, ${printExpr(significance, 0)})"

      case TExpr.Trunc(number, numDigits) =>
        s"TRUNC(${printExpr(number, 0)}, ${printExpr(numDigits, 0)})"

      case TExpr.Sign(value) =>
        s"SIGN(${printExpr(value, 0)})"

      case TExpr.Int_(value) =>
        s"INT(${printExpr(value, 0)})"

      // Reference functions
      case TExpr.Row_(ref) =>
        s"ROW(${printExpr(ref, 0)})"

      case TExpr.Column_(ref) =>
        s"COLUMN(${printExpr(ref, 0)})"

      case TExpr.Rows(rangeExpr) =>
        // ROWS takes a range argument, extract range from FoldRange if needed
        rangeExpr match
          case TExpr.FoldRange(range, _, _, _) => s"ROWS(${formatRange(range)})"
          case TExpr.SheetFoldRange(sheet, range, _, _, _) =>
            s"ROWS(${formatSheetName(sheet)}!${formatRange(range)})"
          case other => s"ROWS(${printExpr(other, 0)})"

      case TExpr.Columns(rangeExpr) =>
        // COLUMNS takes a range argument, extract range from FoldRange if needed
        rangeExpr match
          case TExpr.FoldRange(range, _, _, _) => s"COLUMNS(${formatRange(range)})"
          case TExpr.SheetFoldRange(sheet, range, _, _, _) =>
            s"COLUMNS(${formatSheetName(sheet)}!${formatRange(range)})"
          case other => s"COLUMNS(${printExpr(other, 0)})"

      case TExpr.Address(row, col, absNum, a1Style, sheetName) =>
        val args = sheetName match
          case Some(sheet) =>
            s"${printExpr(row, 0)}, ${printExpr(col, 0)}, ${printExpr(absNum, 0)}, ${printExpr(a1Style, 0)}, ${printExpr(sheet, 0)}"
          case None =>
            s"${printExpr(row, 0)}, ${printExpr(col, 0)}, ${printExpr(absNum, 0)}, ${printExpr(a1Style, 0)}"
        s"ADDRESS($args)"

      // Lookup functions (now support cross-sheet via RangeLocation)
      case TExpr.Index(array, rowNum, colNum) =>
        val arrayStr = formatLocation(array)
        colNum match
          case Some(col) => s"INDEX($arrayStr, ${printExpr(rowNum, 0)}, ${printExpr(col, 0)})"
          case None => s"INDEX($arrayStr, ${printExpr(rowNum, 0)})"

      case TExpr.Match(lookupValue, lookupArray, matchType) =>
        s"MATCH(${printExpr(lookupValue, 0)}, ${formatLocation(lookupArray)}, ${printExpr(matchType, 0)})"

      // Conditional aggregation functions (now support cross-sheet via RangeLocation)
      case TExpr.SumIf(range, criteria, sumRangeOpt) =>
        val rangeStr = formatLocation(range)
        val criteriaStr = printExpr(criteria, 0)
        sumRangeOpt match
          case Some(sumRange) =>
            s"SUMIF($rangeStr, $criteriaStr, ${formatLocation(sumRange)})"
          case None =>
            s"SUMIF($rangeStr, $criteriaStr)"

      case TExpr.CountIf(range, criteria) =>
        s"COUNTIF(${formatLocation(range)}, ${printExpr(criteria, 0)})"

      case TExpr.SumIfs(sumRange, conditions) =>
        val condStrs = conditions
          .map { case (r, criteria) =>
            s"${formatLocation(r)}, ${printExpr(criteria, 0)}"
          }
          .mkString(", ")
        s"SUMIFS(${formatLocation(sumRange)}, $condStrs)"

      case TExpr.CountIfs(conditions) =>
        val condStrs = conditions
          .map { case (r, criteria) =>
            s"${formatLocation(r)}, ${printExpr(criteria, 0)}"
          }
          .mkString(", ")
        s"COUNTIFS($condStrs)"

      case TExpr.AverageIf(range, criteria, avgRangeOpt) =>
        val rangeStr = formatLocation(range)
        val criteriaStr = printExpr(criteria, 0)
        avgRangeOpt match
          case Some(avgRange) =>
            s"AVERAGEIF($rangeStr, $criteriaStr, ${formatLocation(avgRange)})"
          case None =>
            s"AVERAGEIF($rangeStr, $criteriaStr)"

      case TExpr.AverageIfs(avgRange, conditions) =>
        val condStrs = conditions
          .map { case (r, criteria) =>
            s"${formatLocation(r)}, ${printExpr(criteria, 0)}"
          }
          .mkString(", ")
        s"AVERAGEIFS(${formatLocation(avgRange)}, $condStrs)"

      // Array and advanced lookup functions (now support cross-sheet via RangeLocation)
      case TExpr.SumProduct(arrays) =>
        s"SUMPRODUCT(${arrays.map(formatLocation).mkString(", ")})"

      case TExpr.XLookup(
            lookupValue,
            lookupArray,
            returnArray,
            ifNotFound,
            matchMode,
            searchMode
          ) =>
        val lookupStr = printExpr(lookupValue, 0)
        val lookupArrayStr = formatLocation(lookupArray)
        val returnArrayStr = formatLocation(returnArray)

        // Determine which optional args to include based on non-default values
        val hasNonDefaultMatchMode = matchMode match
          case TExpr.Lit(0) => false
          case _ => true
        val hasNonDefaultSearchMode = searchMode match
          case TExpr.Lit(1) => false
          case _ => true

        (ifNotFound, hasNonDefaultMatchMode || hasNonDefaultSearchMode) match
          case (None, false) =>
            // 3-arg form: just lookup, lookupArray, returnArray
            s"XLOOKUP($lookupStr, $lookupArrayStr, $returnArrayStr)"
          case (Some(notFound), false) =>
            // 4-arg form: with if_not_found only
            s"XLOOKUP($lookupStr, $lookupArrayStr, $returnArrayStr, ${printExpr(notFound, 0)})"
          case (ifNotFoundOpt, _) =>
            // Full form: include all args
            val notFoundStr = ifNotFoundOpt.map(nf => printExpr(nf, 0)).getOrElse("\"\"")
            s"XLOOKUP($lookupStr, $lookupArrayStr, $returnArrayStr, $notFoundStr, ${printExpr(matchMode, 0)}, ${printExpr(searchMode, 0)})"

      // Range aggregation
      case TExpr.FoldRange(range, z, step, decode) =>
        // Detect common aggregation patterns
        detectAggregation(TExpr.FoldRange(range, z, step, decode))

      // Cross-sheet range aggregation
      case TExpr.SheetFoldRange(sheet, range, z, step, decode) =>
        // Detect common aggregation patterns for cross-sheet ranges
        s"SUM(${formatSheetName(sheet)}!${formatRange(range)})"

  /**
   * Detect and print common aggregation patterns (SUM, COUNT, AVERAGE).
   */
  private def detectAggregation(fold: TExpr.FoldRange[?, ?]): String =
    // For now, assume SUM (most common)
    // Future: analyze step function to detect COUNT, AVERAGE, etc.
    s"SUM(${formatRange(fold.range)})"

  /**
   * Format ARef to A1 notation with default Relative anchor.
   *
   * Used for standalone cell refs where anchor info isn't available.
   */
  private def formatARef(aref: ARef): String = formatARef(aref, Anchor.Relative)

  /**
   * Format CellRange to A1:B2 notation with per-endpoint anchor support.
   *
   * Uses the anchor info stored in the CellRange (e.g., $A$1:B10).
   */
  private def formatRange(range: CellRange): String =
    s"${formatARef(range.start, range.startAnchor)}:${formatARef(range.end, range.endAnchor)}"

  /**
   * Format RangeLocation (local or cross-sheet) to A1 notation.
   */
  private def formatLocation(location: TExpr.RangeLocation): String =
    location match
      case TExpr.RangeLocation.Local(range) => formatRange(range)
      case TExpr.RangeLocation.CrossSheet(sheet, range) =>
        s"${sheet.value}!${formatRange(range)}"

  /**
   * Format ARef to A1 notation with anchor support.
   *
   * Helper function to avoid opaque type extension method issues. Manually extracts col/row from
   * packed Long representation and adds $ prefixes based on anchor mode.
   */
  @SuppressWarnings(Array("org.wartremover.warts.AsInstanceOf"))
  private def formatARef(aref: ARef, anchor: Anchor): String =
    // ARef is opaque type = Long with (row << 32) | col packing
    // Extract col (low 32 bits) and row (high 32 bits)
    val arefLong: Long = aref.asInstanceOf[Long] // Safe: ARef is opaque type = Long
    val colIndex = (arefLong & 0xffffffffL).toInt
    val rowIndex = (arefLong >> 32).toInt

    // Convert to A1 notation
    val col = Column.from0(colIndex)

    // Format column letter with optional $ prefix
    val colLetter = Column.toLetter(col)
    val rowNum = rowIndex + 1

    // Add $ prefixes based on anchor mode
    val colStr = if anchor.isColAbsolute then s"$$$colLetter" else colLetter
    val rowStr = if anchor.isRowAbsolute then s"$$$rowNum" else rowNum.toString

    s"$colStr$rowStr"

  /**
   * Add parentheses if condition is true.
   */
  private def parenthesizeIf(s: String, condition: Boolean): String =
    if condition then s"($s)" else s

  /**
   * Escape string literal for Excel (double quotes).
   */
  private def escapeString(s: String): String =
    s.replace("\"", "\"\"")

  /**
   * Format sheet name for Excel formula.
   *
   * Sheet names with spaces or special characters need to be quoted with single quotes. Any single
   * quotes within the name are doubled (Excel escape convention).
   */
  private def formatSheetName(sheet: SheetName): String =
    val name = sheet.value
    // Sheet names need quoting if they contain spaces, special chars, or look like cell refs
    val needsQuoting = name.contains(' ') ||
      name.contains('\'') ||
      name.contains('-') ||
      name.exists(c => !c.isLetterOrDigit && c != '_') ||
      name.headOption.exists(_.isDigit) // Names starting with digit need quoting

    if needsQuoting then
      // Escape single quotes by doubling them
      val escaped = name.replace("'", "''")
      s"'$escaped'"
    else name

  /**
   * Print with minimal whitespace (compact format).
   */
  def printCompact(expr: TExpr[?]): String =
    print(expr, includeEquals = false)

  /**
   * Print with whitespace for readability (pretty format).
   *
   * Example: "= SUM( A1:B10 )" instead of "=SUM(A1:B10)"
   */
  def printPretty(expr: TExpr[?]): String =
    val compact = print(expr, includeEquals = false)
    // Add spaces around operators and after commas
    val withSpaces = compact
      .replace("+", " + ")
      .replace("-", " - ")
      .replace("*", " * ")
      .replace("/", " / ")
      .replace("=", " = ")
      .replace("<>", " <> ")
      .replace("<=", " <= ")
      .replace(">=", " >= ")
      .replace("<", " < ")
      .replace(">", " > ")
      .replace(",", ", ")
      .replace("(", "( ")
      .replace(")", " )")

    s"=$withSpaces"

  /**
   * Print multiple expressions as a list (for debugging).
   */
  def printList(exprs: List[TExpr[?]]): String =
    exprs.map(e => print(e, includeEquals = false)).mkString(", ")

  /**
   * Print with type information (for debugging).
   *
   * Example: "Add[BigDecimal](Lit(1), Lit(2))"
   */
  def printWithTypes(expr: TExpr[?]): String =
    expr match
      case TExpr.Lit(value) =>
        s"Lit($value: ${value.getClass.getSimpleName})"
      case TExpr.Ref(at, anchor, _) =>
        s"Ref($at, $anchor)"
      case TExpr.PolyRef(at, anchor) =>
        s"PolyRef($at, $anchor)"
      case TExpr.SheetRef(sheet, at, anchor, _) =>
        s"SheetRef(${sheet.value}, $at, $anchor)"
      case TExpr.SheetPolyRef(sheet, at, anchor) =>
        s"SheetPolyRef(${sheet.value}, $at, $anchor)"
      case TExpr.SheetRange(sheet, range) =>
        s"SheetRange(${sheet.value}, $range)"
      case TExpr.If(cond, ifTrue, ifFalse) =>
        s"If(${printWithTypes(cond)}, ${printWithTypes(ifTrue)}, ${printWithTypes(ifFalse)})"
      case TExpr.Add(x, y) =>
        s"Add(${printWithTypes(x)}, ${printWithTypes(y)})"
      case TExpr.Sub(x, y) =>
        s"Sub(${printWithTypes(x)}, ${printWithTypes(y)})"
      case TExpr.Mul(x, y) =>
        s"Mul(${printWithTypes(x)}, ${printWithTypes(y)})"
      case TExpr.Div(x, y) =>
        s"Div(${printWithTypes(x)}, ${printWithTypes(y)})"
      case TExpr.And(x, y) =>
        s"And(${printWithTypes(x)}, ${printWithTypes(y)})"
      case TExpr.Or(x, y) =>
        s"Or(${printWithTypes(x)}, ${printWithTypes(y)})"
      case TExpr.Not(x) =>
        s"Not(${printWithTypes(x)})"
      case TExpr.Eq(x, y) =>
        s"Eq(${printWithTypes(x)}, ${printWithTypes(y)})"
      case TExpr.Neq(x, y) =>
        s"Neq(${printWithTypes(x)}, ${printWithTypes(y)})"
      case TExpr.Lt(x, y) =>
        s"Lt(${printWithTypes(x)}, ${printWithTypes(y)})"
      case TExpr.Lte(x, y) =>
        s"Lte(${printWithTypes(x)}, ${printWithTypes(y)})"
      case TExpr.Gt(x, y) =>
        s"Gt(${printWithTypes(x)}, ${printWithTypes(y)})"
      case TExpr.Gte(x, y) =>
        s"Gte(${printWithTypes(x)}, ${printWithTypes(y)})"
      case TExpr.ToInt(expr) =>
        s"ToInt(${printWithTypes(expr)})"
      case TExpr.Concatenate(xs) =>
        s"Concatenate(${xs.map(printWithTypes).mkString(", ")})"
      case TExpr.Left(text, n) =>
        s"Left(${printWithTypes(text)}, ${printWithTypes(n)})"
      case TExpr.Right(text, n) =>
        s"Right(${printWithTypes(text)}, ${printWithTypes(n)})"
      case TExpr.Len(text) =>
        s"Len(${printWithTypes(text)})"
      case TExpr.Upper(text) =>
        s"Upper(${printWithTypes(text)})"
      case TExpr.Lower(text) =>
        s"Lower(${printWithTypes(text)})"
      case TExpr.Today() =>
        "Today()"
      case TExpr.Now() =>
        "Now()"
      case TExpr.DateToSerial(expr) =>
        s"DateToSerial(${printWithTypes(expr)})"
      case TExpr.DateTimeToSerial(expr) =>
        s"DateTimeToSerial(${printWithTypes(expr)})"
      case TExpr.Pi() =>
        "Pi()"
      case TExpr.Date(year, month, day) =>
        s"Date(${printWithTypes(year)}, ${printWithTypes(month)}, ${printWithTypes(day)})"
      case TExpr.Year(date) =>
        s"Year(${printWithTypes(date)})"
      case TExpr.Month(date) =>
        s"Month(${printWithTypes(date)})"
      case TExpr.Day(date) =>
        s"Day(${printWithTypes(date)})"
      case TExpr.Eomonth(startDate, months) =>
        s"Eomonth(${printWithTypes(startDate)}, ${printWithTypes(months)})"
      case TExpr.Edate(startDate, months) =>
        s"Edate(${printWithTypes(startDate)}, ${printWithTypes(months)})"
      case TExpr.Datedif(startDate, endDate, unit) =>
        s"Datedif(${printWithTypes(startDate)}, ${printWithTypes(endDate)}, ${printWithTypes(unit)})"
      case TExpr.Networkdays(startDate, endDate, holidaysOpt) =>
        holidaysOpt match
          case Some(holidays) =>
            s"Networkdays(${printWithTypes(startDate)}, ${printWithTypes(endDate)}, ${formatRange(holidays)})"
          case None =>
            s"Networkdays(${printWithTypes(startDate)}, ${printWithTypes(endDate)})"
      case TExpr.Workday(startDate, days, holidaysOpt) =>
        holidaysOpt match
          case Some(holidays) =>
            s"Workday(${printWithTypes(startDate)}, ${printWithTypes(days)}, ${formatRange(holidays)})"
          case None =>
            s"Workday(${printWithTypes(startDate)}, ${printWithTypes(days)})"
      case TExpr.Yearfrac(startDate, endDate, basis) =>
        s"Yearfrac(${printWithTypes(startDate)}, ${printWithTypes(endDate)}, ${printWithTypes(basis)})"
      case TExpr.Sum(range) =>
        s"Sum(${formatLocation(range)})"
      case TExpr.Count(range) =>
        s"Count(${formatLocation(range)})"
      case TExpr.Min(range) =>
        s"Min(${formatLocation(range)})"
      case TExpr.Max(range) =>
        s"Max(${formatLocation(range)})"
      case TExpr.Average(range) =>
        s"Average(${formatLocation(range)})"
      case TExpr.Aggregate(aggregatorId, location) =>
        s"Aggregate($aggregatorId, ${formatLocation(location)})"
      case call: TExpr.Call[?] =>
        val printer = ArgPrinter(
          expr = expr => printWithTypes(expr),
          location = formatLocation,
          cellRange = formatRange
        )
        s"Call(${call.spec.name}, ${call.spec.argSpec.render(call.args, printer).mkString(", ")})"
      case TExpr.SheetSum(sheet, range) =>
        s"SheetSum(${formatSheetName(sheet)}!${formatRange(range)})"
      case TExpr.SheetMin(sheet, range) =>
        s"SheetMin(${formatSheetName(sheet)}!${formatRange(range)})"
      case TExpr.SheetMax(sheet, range) =>
        s"SheetMax(${formatSheetName(sheet)}!${formatRange(range)})"
      case TExpr.SheetAverage(sheet, range) =>
        s"SheetAverage(${formatSheetName(sheet)}!${formatRange(range)})"
      case TExpr.SheetCount(sheet, range) =>
        s"SheetCount(${formatSheetName(sheet)}!${formatRange(range)})"
      case TExpr.Npv(rate, values) =>
        s"Npv(${printWithTypes(rate)}, ${formatLocation(values)})"
      case TExpr.Irr(values, guessOpt) =>
        guessOpt match
          case Some(guess) => s"Irr(${formatLocation(values)}, ${printWithTypes(guess)})"
          case None => s"Irr(${formatLocation(values)})"
      case TExpr.Xnpv(rate, values, dates) =>
        s"Xnpv(${printWithTypes(rate)}, ${formatLocation(values)}, ${formatLocation(dates)})"
      case TExpr.Xirr(values, dates, guessOpt) =>
        guessOpt match
          case Some(guess) =>
            s"Xirr(${formatLocation(values)}, ${formatLocation(dates)}, ${printWithTypes(guess)})"
          case None => s"Xirr(${formatLocation(values)}, ${formatLocation(dates)})"
      case TExpr.Pmt(rate, nper, pv, fv, pmtType) =>
        val args = List(printWithTypes(rate), printWithTypes(nper), printWithTypes(pv)) ++
          fv.map(printWithTypes).toList ++ pmtType.map(printWithTypes).toList
        s"Pmt(${args.mkString(", ")})"
      case TExpr.Fv(rate, nper, pmt, pv, pmtType) =>
        val args = List(printWithTypes(rate), printWithTypes(nper), printWithTypes(pmt)) ++
          pv.map(printWithTypes).toList ++ pmtType.map(printWithTypes).toList
        s"Fv(${args.mkString(", ")})"
      case TExpr.Pv(rate, nper, pmt, fv, pmtType) =>
        val args = List(printWithTypes(rate), printWithTypes(nper), printWithTypes(pmt)) ++
          fv.map(printWithTypes).toList ++ pmtType.map(printWithTypes).toList
        s"Pv(${args.mkString(", ")})"
      case TExpr.Nper(rate, pmt, pv, fv, pmtType) =>
        val args = List(printWithTypes(rate), printWithTypes(pmt), printWithTypes(pv)) ++
          fv.map(printWithTypes).toList ++ pmtType.map(printWithTypes).toList
        s"Nper(${args.mkString(", ")})"
      case TExpr.Rate(nper, pmt, pv, fv, pmtType, guess) =>
        val args = List(printWithTypes(nper), printWithTypes(pmt), printWithTypes(pv)) ++
          fv.map(printWithTypes).toList ++ pmtType.map(printWithTypes).toList ++
          guess.map(printWithTypes).toList
        s"Rate(${args.mkString(", ")})"
      case TExpr.VLookup(lookup, table, colIndex, rangeLookup) =>
        s"VLookup(${printWithTypes(lookup)}, ${formatLocation(table)}, " +
          s"${printWithTypes(colIndex)}, ${printWithTypes(rangeLookup)})"
      case TExpr.Iferror(value, valueIfError) =>
        s"Iferror(${printWithTypes(value)}, ${printWithTypes(valueIfError)})"
      case TExpr.Iserror(value) =>
        s"Iserror(${printWithTypes(value)})"
      case TExpr.Iserr(value) =>
        s"Iserr(${printWithTypes(value)})"
      case TExpr.Isnumber(value) =>
        s"Isnumber(${printWithTypes(value)})"
      case TExpr.Istext(value) =>
        s"Istext(${printWithTypes(value)})"
      case TExpr.Isblank(value) =>
        s"Isblank(${printWithTypes(value)})"
      case TExpr.Round(value, numDigits) =>
        s"Round(${printWithTypes(value)}, ${printWithTypes(numDigits)})"
      case TExpr.RoundUp(value, numDigits) =>
        s"RoundUp(${printWithTypes(value)}, ${printWithTypes(numDigits)})"
      case TExpr.RoundDown(value, numDigits) =>
        s"RoundDown(${printWithTypes(value)}, ${printWithTypes(numDigits)})"
      case TExpr.Abs(value) =>
        s"Abs(${printWithTypes(value)})"
      case TExpr.Sqrt(value) =>
        s"Sqrt(${printWithTypes(value)})"
      case TExpr.Mod(number, divisor) =>
        s"Mod(${printWithTypes(number)}, ${printWithTypes(divisor)})"
      case TExpr.Power(number, power) =>
        s"Power(${printWithTypes(number)}, ${printWithTypes(power)})"
      case TExpr.Log(number, base) =>
        s"Log(${printWithTypes(number)}, ${printWithTypes(base)})"
      case TExpr.Ln(value) =>
        s"Ln(${printWithTypes(value)})"
      case TExpr.Exp(value) =>
        s"Exp(${printWithTypes(value)})"
      case TExpr.Floor(number, significance) =>
        s"Floor(${printWithTypes(number)}, ${printWithTypes(significance)})"
      case TExpr.Ceiling(number, significance) =>
        s"Ceiling(${printWithTypes(number)}, ${printWithTypes(significance)})"
      case TExpr.Trunc(number, numDigits) =>
        s"Trunc(${printWithTypes(number)}, ${printWithTypes(numDigits)})"
      case TExpr.Sign(value) =>
        s"Sign(${printWithTypes(value)})"
      case TExpr.Int_(value) =>
        s"Int_(${printWithTypes(value)})"
      case TExpr.Row_(ref) =>
        s"Row_(${printWithTypes(ref)})"
      case TExpr.Column_(ref) =>
        s"Column_(${printWithTypes(ref)})"
      case TExpr.Rows(range) =>
        s"Rows(${printWithTypes(range)})"
      case TExpr.Columns(range) =>
        s"Columns(${printWithTypes(range)})"
      case TExpr.Address(row, col, absNum, a1Style, sheetName) =>
        val sheetStr = sheetName.map(s => s", ${printWithTypes(s)}").getOrElse("")
        s"Address(${printWithTypes(row)}, ${printWithTypes(col)}, ${printWithTypes(absNum)}, ${printWithTypes(a1Style)}$sheetStr)"
      case TExpr.Index(array, rowNum, colNum) =>
        colNum match
          case Some(col) =>
            s"Index(${formatLocation(array)}, ${printWithTypes(rowNum)}, ${printWithTypes(col)})"
          case None =>
            s"Index(${formatLocation(array)}, ${printWithTypes(rowNum)})"
      case TExpr.Match(lookupValue, lookupArray, matchType) =>
        s"Match(${printWithTypes(lookupValue)}, ${formatLocation(lookupArray)}, ${printWithTypes(matchType)})"
      case TExpr.SumIf(range, criteria, sumRangeOpt) =>
        sumRangeOpt match
          case Some(sumRange) =>
            s"SumIf(${formatLocation(range)}, ${printWithTypes(criteria)}, ${formatLocation(sumRange)})"
          case None =>
            s"SumIf(${formatLocation(range)}, ${printWithTypes(criteria)})"
      case TExpr.CountIf(range, criteria) =>
        s"CountIf(${formatLocation(range)}, ${printWithTypes(criteria)})"
      case TExpr.SumIfs(sumRange, conditions) =>
        val condStrs = conditions
          .map { case (r, criteria) =>
            s"(${formatLocation(r)}, ${printWithTypes(criteria)})"
          }
          .mkString(", ")
        s"SumIfs(${formatLocation(sumRange)}, $condStrs)"
      case TExpr.CountIfs(conditions) =>
        val condStrs = conditions
          .map { case (r, criteria) =>
            s"(${formatLocation(r)}, ${printWithTypes(criteria)})"
          }
          .mkString(", ")
        s"CountIfs($condStrs)"
      case TExpr.AverageIf(range, criteria, avgRangeOpt) =>
        avgRangeOpt match
          case Some(avgRange) =>
            s"AverageIf(${formatLocation(range)}, ${printWithTypes(criteria)}, ${formatLocation(avgRange)})"
          case None =>
            s"AverageIf(${formatLocation(range)}, ${printWithTypes(criteria)})"
      case TExpr.AverageIfs(avgRange, conditions) =>
        val condStrs = conditions
          .map { case (r, criteria) =>
            s"(${formatLocation(r)}, ${printWithTypes(criteria)})"
          }
          .mkString(", ")
        s"AverageIfs(${formatLocation(avgRange)}, $condStrs)"
      case TExpr.SumProduct(arrays) =>
        s"SumProduct(${arrays.map(formatLocation).mkString(", ")})"
      case TExpr.XLookup(
            lookupValue,
            lookupArray,
            returnArray,
            ifNotFound,
            matchMode,
            searchMode
          ) =>
        val ifNotFoundStr = ifNotFound.map(nf => s", ${printWithTypes(nf)}").getOrElse("")
        s"XLookup(${printWithTypes(lookupValue)}, ${formatLocation(lookupArray)}, ${formatLocation(returnArray)}$ifNotFoundStr, ${printWithTypes(matchMode)}, ${printWithTypes(searchMode)})"
      case fold @ TExpr.FoldRange(range, _, _, _) =>
        s"FoldRange(${formatRange(range)})"
      case TExpr.SheetFoldRange(sheet, range, _, _, _) =>
        s"SheetFoldRange(${formatSheetName(sheet)}!${formatRange(range)})"
