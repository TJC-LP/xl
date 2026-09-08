package com.tjclp.xl.formula.eval

import com.tjclp.xl.addressing.{Column, Row, SheetName}
import com.tjclp.xl.error.{XLError, XLResult}
import com.tjclp.xl.formula.parser.{FormulaParser, ParseError}
import com.tjclp.xl.formula.printer.FormulaOps
import com.tjclp.xl.ops.FormulaSupport
import com.tjclp.xl.workbooks.Workbook

/**
 * The evaluator-backed [[com.tjclp.xl.ops.FormulaSupport]] (ADR-017 §2.12): what the core `Edit`
 * interpreter cannot do without a parser. `validate` is `FormulaParser.parse` (the leading `=`
 * optional, as the CLI's `putf` has always accepted it), `shift` is `FormulaOps.shift`, the four
 * structural edits are `StructuralEditor.*Checked` (refuse-before-mutate: an out-of-bounds shift, a
 * torn data table or a name that cannot be rewritten leaves the workbook untouched; a sheet the
 * book lacks is `SheetNotFound` here, where the editor alone would no-op) and `renameSheet` is
 * `SheetRenamer.rename`. The scripting prelude and `import com.tjclp.xl.{*, given}` provide it as
 * the `given FormulaSupport`; the CLI passes it explicitly.
 */
object EvalFormulaSupport extends FormulaSupport:

  /** `=`-prefixed, whitespace-trimmed: the form the parser and its error positions expect. */
  private def full(formula: String): String =
    val trimmed = formula.trim
    if trimmed.startsWith("=") then trimmed else s"=$trimmed"

  def validate(formula: String): XLResult[Unit] =
    val text = full(formula)
    FormulaParser.parse(text).map(_ => ()).left.map(ParseError.toXLError(_, text))

  def shift(formula: String, colDelta: Int, rowDelta: Int): XLResult[String] =
    FormulaOps.shift(formula, colDelta, rowDelta)

  /** The editor no-ops on a sheet the book lacks; the seam refuses instead. */
  private def onSheet(wb: Workbook, sheet: SheetName)(
    edit: => XLResult[Workbook]
  ): XLResult[Workbook] =
    if wb.sheets.exists(_.name == sheet) then edit else Left(XLError.SheetNotFound(sheet.value))

  def insertRows(wb: Workbook, sheet: SheetName, at: Row, count: Int): XLResult[Workbook] =
    onSheet(wb, sheet)(StructuralEditor.insertRowsChecked(wb, sheet, at.index0, count))

  def deleteRows(wb: Workbook, sheet: SheetName, at: Row, count: Int): XLResult[Workbook] =
    onSheet(wb, sheet)(StructuralEditor.deleteRowsChecked(wb, sheet, at.index0, count))

  def insertCols(wb: Workbook, sheet: SheetName, at: Column, count: Int): XLResult[Workbook] =
    onSheet(wb, sheet)(StructuralEditor.insertColumnsChecked(wb, sheet, at.index0, count))

  def deleteCols(wb: Workbook, sheet: SheetName, at: Column, count: Int): XLResult[Workbook] =
    onSheet(wb, sheet)(StructuralEditor.deleteColumnsChecked(wb, sheet, at.index0, count))

  def renameSheet(wb: Workbook, from: SheetName, to: SheetName): XLResult[Workbook] =
    SheetRenamer.rename(wb, from, to)
