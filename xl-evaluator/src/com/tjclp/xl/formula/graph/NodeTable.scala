package com.tjclp.xl.formula.graph

import com.tjclp.xl.addressing.{ARef, Column, Row, SheetName}
import com.tjclp.xl.cells.{CellValue, FormulaKind}
import com.tjclp.xl.formula.graph.DependencyGraph.QualifiedRef
import com.tjclp.xl.workbooks.Workbook

/**
 * The formula cells of a workbook as dense `Int` ids in grid order (#767).
 *
 * Id `i` is the `i`-th formula cell under [[QualifiedRef.gridKey]]: sheet NAME, then row, then
 * column. Comparing two ids therefore compares their grid positions, which is the order
 * `QualifiedGraph.byPosition`, the SCC member listings and the sorted error vectors already use.
 *
 * The node set is the legacy graph's: every `CellValue.Formula` except `FormulaKind.DataTable`
 * records (GH-430), unparseable formulas included (they read nothing). Sheet ids index
 * [[sheetNames]], which is sorted by name.
 */
private[formula] final class NodeTable private (
  val sheetNames: Vector[SheetName],
  sheetIndex: Map[SheetName, Int],
  formulaSheet: IArray[Int],
  formulaRow: IArray[Int],
  formulaCol: IArray[Int],
  byCell: LongIntMap,
  offGrid: Map[(Int, Int, Int), Int]
):
  /** Number of formula nodes. */
  def formulaCount: Int = formulaSheet.length

  /** The id of the sheet named exactly `name`, or -1. */
  def sheetOf(name: SheetName): Int = sheetIndex.getOrElse(name, -1)

  def sheet(n: Int): Int = formulaSheet(n)
  def row(n: Int): Int = formulaRow(n)
  def col(n: Int): Int = formulaCol(n)

  /** The formula id at (sheet id, 0-based row, 0-based column), or -1 when no formula is there. */
  private[graph] def indexOf(sheet: Int, row: Int, col: Int): Int =
    val key = NodeTable.cellKey(sheet, row, col)
    if key >= 0L then byCell.indexOf(key) else offGrid.getOrElse((sheet, row, col), -1)

  /** The formula id of `q`, if `q` is a formula node. */
  def lookup(q: QualifiedRef): Option[Int] =
    val s = sheetOf(q.sheet)
    if s < 0 then None
    else
      val n = indexOf(s, q.ref.row.index0, q.ref.col.index0)
      if n < 0 then None else Some(n)

  /** The cell of formula id `n`. */
  def qualified(n: Int): QualifiedRef =
    QualifiedRef(sheetNames(formulaSheet(n)), ARef.from0(formulaCol(n), formulaRow(n)))

private[formula] object NodeTable:

  /**
   * One `Long` per cell of Excel's grid: `sheet << 34 | row << 14 | col`, unique for rows below
   * 2^20 and columns below 2^14 (Excel's limits) and ordered like the grid. -1 for a cell outside
   * the grid, which only an unchecked `ARef.from0` can make.
   */
  def cellKey(sheet: Int, row: Int, col: Int): Long =
    if sheet < 0 || row < 0 || row > Row.MaxIndex0 || col < 0 || col > Column.MaxIndex0 then -1L
    else (sheet.toLong << 34) | (row.toLong << 14) | col.toLong

  private final class Entry(val sheet: Int, val row: Int, val col: Int, val text: String)

  private val gridOrder: java.util.Comparator[Entry] = (x, y) =>
    if x.sheet != y.sheet then Integer.compare(x.sheet, y.sheet)
    else if x.row != y.row then Integer.compare(x.row, y.row)
    else Integer.compare(x.col, y.col)

  /**
   * The table of `wb`'s formula cells and each one's formula text, indexed by id. Two cells at one
   * position (only a workbook holding two sheets of one name has them) keep the later one, as the
   * legacy graph's `toMap` does.
   */
  @SuppressWarnings(Array("org.wartremover.warts.Var", "org.wartremover.warts.While"))
  def build(wb: Workbook): (NodeTable, IArray[String]) =
    val names: Vector[SheetName] = wb.sheets.map(_.name).distinct.sortBy(_.value)
    val sheetIndex: Map[SheetName, Int] = names.iterator.zipWithIndex.toMap
    val entries = Array.newBuilder[Entry]
    wb.sheets.foreach { sheet =>
      val s = sheetIndex.getOrElse(sheet.name, -1)
      sheet.cells.foreach { case (ref, cell) =>
        cell.value match
          case CellValue.Formula(_, _, _: FormulaKind.DataTable) => ()
          case CellValue.Formula(text, _, _) =>
            entries += Entry(s, ref.row.index0, ref.col.index0, text)
          case _ => ()
      }
    }
    val all = entries.result()
    // stable: equal positions keep their input order, so the later duplicate is last
    java.util.Arrays.sort(all, gridOrder)
    val kept = Array.newBuilder[Entry]
    var i = 0
    while i < all.length do
      val e = all(i)
      val last = i + 1 >= all.length || gridOrder.compare(e, all(i + 1)) != 0
      if last then kept += e
      i += 1
    val formulas = kept.result()
    val n = formulas.length
    val keys = new Array[Long](n)
    val ids = new Array[Int](n)
    var onGrid = 0
    val offGrid = Map.newBuilder[(Int, Int, Int), Int]
    i = 0
    while i < n do
      val e = formulas(i)
      val key = cellKey(e.sheet, e.row, e.col)
      if key >= 0L then
        keys(onGrid) = key
        ids(onGrid) = i
        onGrid += 1
      else offGrid += ((e.sheet, e.row, e.col) -> i)
      i += 1
    val table = new NodeTable(
      names,
      sheetIndex,
      IArray.unsafeFromArray(formulas.map(_.sheet)),
      IArray.unsafeFromArray(formulas.map(_.row)),
      IArray.unsafeFromArray(formulas.map(_.col)),
      LongIntMap.from(keys, ids, onGrid),
      offGrid.result()
    )
    (table, IArray.unsafeFromArray(formulas.map(_.text)))
