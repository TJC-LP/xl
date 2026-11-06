package com.tjclp.xl.formula.functions

import com.tjclp.xl.addressing.{ARef, CellRange}
import com.tjclp.xl.cells.CellError
import com.tjclp.xl.formula.Arity
import com.tjclp.xl.formula.ast.TExpr
import com.tjclp.xl.formula.eval.{ArrayResult, EvalError, RangeOperand}
import com.tjclp.xl.formula.parser.ParseError
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.workbooks.Workbook

/**
 * GH-669: the bounded, qualified plain-column form `Table1[Column]`.
 *
 * A private, unregistered reference call keeps table identifiers out of A1 shifting and the
 * function registry. Resolution uses OOXML displayName (the formula name), not the object-model
 * name or the Sheet.tables map key. No sheet-name lookup policy is introduced here.
 */
private[formula] object TableColumnReferences extends FunctionSpecsBase:
  final val Name = "(table-column)"

  final case class Column(table: String, column: String) derives CanEqual

  private val columnArgs: ArgSpec[Column] = new ArgSpec[Column]:
    def describeParts: List[String] = List("qualified table column")
    def parse(
      args: List[TExpr[?]],
      pos: Int,
      fnName: String
    ): Either[ParseError, (Column, List[TExpr[?]])] =
      Left(ParseError.InvalidArguments(fnName, pos, describe, "internal reference"))
    // Names are metadata identifiers, never cell expressions or scalar slots.
    def toValues(args: Column): List[ArgValue] = Nil
    def map(args: Column)(
      mapExpr: TExpr[?] => TExpr[?],
      mapRange: TExpr.RangeLocation => TExpr.RangeLocation,
      mapCells: CellRange => CellRange
    ): Column = args

  private val spec: FunctionSpec[ArrayResult] { type Args = Column } =
    FunctionSpec.referencing[ArrayResult, Column](
      Name,
      Arity.none,
      renderFn = Some((column, _) => s"${column.table}[${column.column}]")
    )((column, ctx) => resolve(column, ctx.sheet, ctx.workbook)) { (column, ctx) =>
      resolve(column, ctx.sheet, ctx.workbook).flatMap { ref =>
        referenceResult(ref.range, ref.sheet, ctx)(readRangeArray(ref.range, ref.sheet, ctx))
      }
    }(using columnArgs)

  def apply(table: String, column: String): TExpr[ArrayResult] =
    TExpr.Call(spec, Column(table, column))

  object Ref:
    def unapply(expr: TExpr[?]): Option[Column] = expr match
      case TExpr.Coerced(inner, _) => unapply(inner)
      case call: TExpr.Call[?] if call.spec eq spec =>
        // The spec identity establishes the path-dependent Args type.
        Some(arguments(call))
      case _ => None

  @SuppressWarnings(Array("org.wartremover.warts.AsInstanceOf"))
  private def arguments(call: TExpr.Call[?]): Column = call.args.asInstanceOf[Column]

  def resolve(
    column: Column,
    current: Sheet,
    workbook: Option[Workbook]
  ): Either[EvalError, RangeOperand] =
    def invalid(detail: String): Left[EvalError, Nothing] =
      Left(EvalError.ErrorValue(CellError.Ref, Some(detail)))
    val matches = workbook.fold(Vector(current))(_.sheets).flatMap { sheet =>
      sheet.tables.valuesIterator
        .filter(_.displayName.equalsIgnoreCase(column.table))
        .map(table => (sheet, table))
        .toVector
    }
    matches match
      case Vector((sheet, table)) =>
        val columns = table.columns.zipWithIndex.filter(_._1.name.equalsIgnoreCase(column.column))
        // Guard before constructing CellRange: its normalizing constructor would otherwise turn
        // a header-only table into a range that reads the header, or overflow at the sheet edge.
        val first = table.range.rowStart.index0 + (if table.showHeaderRow then 1 else 0)
        val last = table.range.rowEnd.index0 - (if table.showTotalsRow then 1 else 0)
        columns match
          case Vector((_, index)) if table.isValid && first <= last =>
            val col = table.range.colStart.index0 + index
            Right(RangeOperand(sheet, CellRange(ARef.from0(col, first), ARef.from0(col, last))))
          case _ => invalid(s"Invalid or missing table column ${column.table}[${column.column}]")
      case _ => invalid(s"Missing or ambiguous table ${column.table}")

  def isPlainColumnChar(c: Char): Boolean =
    c.isLetterOrDigit || c == '_' || c == ' ' || c == '.'

  /**
   * Cheap scheduling prefilter; literals, external refs and unsupported nested selectors skip it.
   */
  def mayContain(text: String): Boolean =
    @annotation.tailrec
    def quoted(i: Int, quote: Char): Int =
      if i >= text.length then i
      else if text.charAt(i) != quote then quoted(i + 1, quote)
      else if text.lift(i + 1).contains(quote) then quoted(i + 2, quote)
      else i + 1
    @annotation.tailrec
    def bracket(i: Int, depth: Int): Int =
      if i >= text.length || depth == 0 then i
      else
        text.charAt(i) match
          case '[' => bracket(i + 1, depth + 1)
          case ']' => bracket(i + 1, depth - 1)
          case _ => bracket(i + 1, depth)
    @annotation.tailrec
    def endWhile(i: Int, accepts: Char => Boolean): Int =
      if i < text.length && accepts(text.charAt(i)) then endWhile(i + 1, accepts) else i
    @annotation.tailrec
    def scan(i: Int): Boolean =
      if i >= text.length then false
      else
        text.charAt(i) match
          case '"' | '\'' => scan(quoted(i + 1, text.charAt(i)))
          case '[' => scan(bracket(i + 1, 1))
          case c if com.tjclp.xl.formula.parser.ReferenceChars.isToken(c) =>
            val end = endWhile(i, com.tjclp.xl.formula.parser.ReferenceChars.isToken)
            val columnEnd = endWhile(end + 1, isPlainColumnChar)
            if text.lift(end).contains('[') && text.lift(columnEnd).contains(']') &&
              text.substring(end + 1, columnEnd).trim.nonEmpty && !text
                .substring(i, end)
                .contains('$')
            then true
            else scan(end)
          case _ => scan(i + 1)
    scan(0)
