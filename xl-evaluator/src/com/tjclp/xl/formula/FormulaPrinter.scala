package com.tjclp.xl.formula

import com.tjclp.xl.{ARef, CellRange}
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
      case TExpr.Ref(at, _) => formatARef(at)

      // Conditional
      case TExpr.If(cond, ifTrue, ifFalse) =>
        s"IF(${printExpr(cond, 0)}, ${printExpr(ifTrue, 0)}, ${printExpr(ifFalse, 0)})"

      // Arithmetic operators
      case TExpr.Add(x, y) =>
        val result = s"${printExpr(x, Precedence.AddSub)}+${printExpr(y, Precedence.AddSub)}"
        parenthesizeIf(result, precedence > Precedence.AddSub)

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
        s"MIN(${formatARef(range.start)}:${formatARef(range.end)})"

      case TExpr.Max(range) =>
        s"MAX(${formatARef(range.start)}:${formatARef(range.end)})"

      // Range aggregation
      case TExpr.FoldRange(range, z, step, decode) =>
        // Detect common aggregation patterns
        detectAggregation(TExpr.FoldRange(range, z, step, decode))

  /**
   * Detect and print common aggregation patterns (SUM, COUNT, AVERAGE).
   */
  private def detectAggregation(fold: TExpr.FoldRange[?, ?]): String =
    val range = fold.range
    val rangeStr = s"${formatARef(range.start)}:${formatARef(range.end)}"

    // For now, assume SUM (most common)
    // Future: analyze step function to detect COUNT, AVERAGE, etc.
    s"SUM($rangeStr)"

  /**
   * Format ARef to A1 notation.
   *
   * Helper function to avoid opaque type extension method issues. Manually extracts col/row from
   * packed Long representation.
   */
  @SuppressWarnings(Array("org.wartremover.warts.AsInstanceOf"))
  private def formatARef(aref: ARef): String =
    // ARef is opaque type = Long with (row << 32) | col packing
    // Extract col (low 32 bits) and row (high 32 bits)
    val arefLong: Long = aref.asInstanceOf[Long] // Safe: ARef is opaque type = Long
    val colIndex = (arefLong & 0xffffffffL).toInt
    val rowIndex = (arefLong >> 32).toInt

    // Convert to A1 notation
    val col = Column.from0(colIndex)
    val row = Row.from1(rowIndex + 1)

    // Format column letter
    val colLetter = Column.toLetter(col)
    val rowNum = rowIndex + 1

    s"$colLetter$rowNum"

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
      case TExpr.Ref(at, _) =>
        s"Ref($at)"
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
        s"Min(${formatARef(range.start)}:${formatARef(range.end)})"
      case TExpr.Max(range) =>
        s"Max(${formatARef(range.start)}:${formatARef(range.end)})"
      case fold @ TExpr.FoldRange(range, _, _, _) =>
        s"FoldRange(${formatARef(range.start)}:${formatARef(range.end)})"
