package com.tjclp.xl.formula.printer

import com.tjclp.xl.formula.ast.TExpr
import com.tjclp.xl.formula.functions.{FunctionSpec, FunctionSpecs, ArgPrinter}

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
      case TExpr.RangeRef(range) =>
        formatRange(range)
      case TExpr.SheetRange(sheet, range) =>
        s"${formatSheetName(sheet)}!${formatRange(range)}"

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

      // String operators
      case TExpr.Concat(x, y) =>
        val result = s"${printExpr(x, Precedence.Concat)}&${printExpr(y, Precedence.Concat)}"
        parenthesizeIf(result, precedence > Precedence.Concat)

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

      // Date/Time conversions
      case TExpr.DateToSerial(expr) =>
        printExpr(expr, precedence)

      case TExpr.DateTimeToSerial(expr) =>
        printExpr(expr, precedence)

      // Aggregate functions
      case TExpr.Aggregate(aggregatorId, location) =>
        s"$aggregatorId(${formatLocation(location)})"

      case call: TExpr.Call[?] =>
        val printer = ArgPrinter(
          expr = expr => printExpr(expr, precedence = 0),
          location = formatLocation,
          cellRange = formatRange
        )
        call.spec.render(call.args, printer)

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
        s"${formatSheetName(sheet)}!${formatRange(range)}"

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
      case TExpr.RangeRef(range) =>
        s"RangeRef(${formatRange(range)})"
      case TExpr.SheetRange(sheet, range) =>
        s"SheetRange(${sheet.value}, ${formatRange(range)})"
      case TExpr.Add(x, y) =>
        s"Add(${printWithTypes(x)}, ${printWithTypes(y)})"
      case TExpr.Sub(x, y) =>
        s"Sub(${printWithTypes(x)}, ${printWithTypes(y)})"
      case TExpr.Mul(x, y) =>
        s"Mul(${printWithTypes(x)}, ${printWithTypes(y)})"
      case TExpr.Div(x, y) =>
        s"Div(${printWithTypes(x)}, ${printWithTypes(y)})"
      case TExpr.Concat(x, y) =>
        s"Concat(${printWithTypes(x)}, ${printWithTypes(y)})"
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
      case TExpr.DateToSerial(expr) =>
        s"DateToSerial(${printWithTypes(expr)})"
      case TExpr.DateTimeToSerial(expr) =>
        s"DateTimeToSerial(${printWithTypes(expr)})"
      case TExpr.Aggregate(aggregatorId, location) =>
        s"Aggregate($aggregatorId, ${formatLocation(location)})"
      case call: TExpr.Call[?] =>
        val printer = ArgPrinter(
          expr = expr => printWithTypes(expr),
          location = formatLocation,
          cellRange = formatRange
        )
        s"Call(${call.spec.name}, ${call.spec.argSpec.render(call.args, printer).mkString(", ")})"
