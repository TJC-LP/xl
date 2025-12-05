package com.tjclp.xl.formula

import com.tjclp.xl.{ARef, Anchor, CellRange}
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

      case TExpr.Date(year, month, day) =>
        s"DATE(${printExpr(year, 0)}, ${printExpr(month, 0)}, ${printExpr(day, 0)})"

      case TExpr.Year(date) =>
        s"YEAR(${printExpr(date, 0)})"

      case TExpr.Month(date) =>
        s"MONTH(${printExpr(date, 0)})"

      case TExpr.Day(date) =>
        s"DAY(${printExpr(date, 0)})"

      // Arithmetic range functions
      case TExpr.Min(range) =>
        s"MIN(${formatRange(range)})"

      case TExpr.Max(range) =>
        s"MAX(${formatRange(range)})"

      case TExpr.Average(range) =>
        s"AVERAGE(${formatRange(range)})"

      // Financial functions
      case TExpr.Npv(rate, values) =>
        s"NPV(${printExpr(rate, 0)}, ${formatRange(values)})"

      case TExpr.Irr(values, guessOpt) =>
        val rangeText = formatRange(values)
        guessOpt match
          case Some(guess) => s"IRR($rangeText, ${printExpr(guess, 0)})"
          case None => s"IRR($rangeText)"

      case TExpr.VLookup(lookup, table, colIndex, rangeLookup) =>
        s"VLOOKUP(${printExpr(lookup, 0)}, " +
          s"${formatRange(table)}, " +
          s"${printExpr(colIndex, 0)}, ${printExpr(rangeLookup, 0)})"

      // Conditional aggregation functions
      case TExpr.SumIf(range, criteria, sumRangeOpt) =>
        val rangeStr = formatRange(range)
        val criteriaStr = printExpr(criteria, 0)
        sumRangeOpt match
          case Some(sumRange) =>
            s"SUMIF($rangeStr, $criteriaStr, ${formatRange(sumRange)})"
          case None =>
            s"SUMIF($rangeStr, $criteriaStr)"

      case TExpr.CountIf(range, criteria) =>
        s"COUNTIF(${formatRange(range)}, ${printExpr(criteria, 0)})"

      case TExpr.SumIfs(sumRange, conditions) =>
        val condStrs = conditions
          .map { case (r, criteria) =>
            s"${formatRange(r)}, ${printExpr(criteria, 0)}"
          }
          .mkString(", ")
        s"SUMIFS(${formatRange(sumRange)}, $condStrs)"

      case TExpr.CountIfs(conditions) =>
        val condStrs = conditions
          .map { case (r, criteria) =>
            s"${formatRange(r)}, ${printExpr(criteria, 0)}"
          }
          .mkString(", ")
        s"COUNTIFS($condStrs)"

      // Array and advanced lookup functions
      case TExpr.SumProduct(arrays) =>
        s"SUMPRODUCT(${arrays.map(formatRange).mkString(", ")})"

      case TExpr.XLookup(
            lookupValue,
            lookupArray,
            returnArray,
            ifNotFound,
            matchMode,
            searchMode
          ) =>
        val lookupStr = printExpr(lookupValue, 0)
        val lookupArrayStr = formatRange(lookupArray)
        val returnArrayStr = formatRange(returnArray)

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
      case TExpr.Date(year, month, day) =>
        s"Date(${printWithTypes(year)}, ${printWithTypes(month)}, ${printWithTypes(day)})"
      case TExpr.Year(date) =>
        s"Year(${printWithTypes(date)})"
      case TExpr.Month(date) =>
        s"Month(${printWithTypes(date)})"
      case TExpr.Day(date) =>
        s"Day(${printWithTypes(date)})"
      case TExpr.Min(range) =>
        s"Min(${formatRange(range)})"
      case TExpr.Max(range) =>
        s"Max(${formatRange(range)})"
      case TExpr.Average(range) =>
        s"Average(${formatRange(range)})"
      case TExpr.Npv(rate, values) =>
        s"Npv(${printWithTypes(rate)}, ${formatRange(values)})"
      case TExpr.Irr(values, guessOpt) =>
        guessOpt match
          case Some(guess) => s"Irr(${formatRange(values)}, ${printWithTypes(guess)})"
          case None => s"Irr(${formatRange(values)})"
      case TExpr.VLookup(lookup, table, colIndex, rangeLookup) =>
        s"VLookup(${printWithTypes(lookup)}, ${formatRange(table)}, " +
          s"${printWithTypes(colIndex)}, ${printWithTypes(rangeLookup)})"
      case TExpr.SumIf(range, criteria, sumRangeOpt) =>
        sumRangeOpt match
          case Some(sumRange) =>
            s"SumIf(${formatRange(range)}, ${printWithTypes(criteria)}, ${formatRange(sumRange)})"
          case None =>
            s"SumIf(${formatRange(range)}, ${printWithTypes(criteria)})"
      case TExpr.CountIf(range, criteria) =>
        s"CountIf(${formatRange(range)}, ${printWithTypes(criteria)})"
      case TExpr.SumIfs(sumRange, conditions) =>
        val condStrs = conditions
          .map { case (r, criteria) =>
            s"(${formatRange(r)}, ${printWithTypes(criteria)})"
          }
          .mkString(", ")
        s"SumIfs(${formatRange(sumRange)}, $condStrs)"
      case TExpr.CountIfs(conditions) =>
        val condStrs = conditions
          .map { case (r, criteria) =>
            s"(${formatRange(r)}, ${printWithTypes(criteria)})"
          }
          .mkString(", ")
        s"CountIfs($condStrs)"
      case TExpr.SumProduct(arrays) =>
        s"SumProduct(${arrays.map(formatRange).mkString(", ")})"
      case TExpr.XLookup(
            lookupValue,
            lookupArray,
            returnArray,
            ifNotFound,
            matchMode,
            searchMode
          ) =>
        val ifNotFoundStr = ifNotFound.map(nf => s", ${printWithTypes(nf)}").getOrElse("")
        s"XLookup(${printWithTypes(lookupValue)}, ${formatRange(lookupArray)}, ${formatRange(returnArray)}$ifNotFoundStr, ${printWithTypes(matchMode)}, ${printWithTypes(searchMode)})"
      case fold @ TExpr.FoldRange(range, _, _, _) =>
        s"FoldRange(${formatRange(range)})"
