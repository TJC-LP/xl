package com.tjclp.xl.formula.graph

import com.tjclp.xl.CellRange
import com.tjclp.xl.addressing.{ARef, Column, SheetName}
import com.tjclp.xl.cells.{Cell, CellValue, FormulaKind}
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.workbooks.Workbook
import org.scalacheck.Gen

/**
 * #767: random workbooks for the compact-graph laws. Small grids (rows 1..Rows, columns A..D on
 * three sheets whose names sort differently from their workbook order) whose formulas read single
 * cells, windowed, whole-column, whole-row, wide, overlapping and running-total ranges, defined
 * names, VLOOKUP strips and their own cells; with constants, DataTable records and unparseable
 * formulas mixed in. Cycles arise from the random references.
 */
object CompactGraphGens:

  val Rows: Int = 10
  val Cols: Int = 4

  /** Workbook order differs from name order ("b" sorts after "Z"). */
  val sheetNames: Vector[String] = Vector("b", "Z", "Mid")

  private def col(c: Int): String = Column.from0(c).toLetter

  private val genSheet: Gen[String] = Gen.oneOf(sheetNames)
  private val genRow: Gen[Int] = Gen.choose(1, Rows)
  private val genCol: Gen[Int] = Gen.choose(0, Cols - 1)
  private def qualify(sheet: Option[String], ref: String): String =
    sheet.fold(ref)(s => s"'$s'!$ref")

  private val genQualifier: Gen[Option[String]] =
    Gen.frequency(3 -> Gen.const(None), 2 -> genSheet.map(Some(_)))

  private val genPoint: Gen[String] =
    for
      q <- genQualifier
      c <- genCol
      r <- genRow
    yield qualify(q, s"${col(c)}$r")

  private def genRange(row: Int, column: Int): Gen[String] = Gen.oneOf(
    // windowed and overlapping
    for
      q <- genQualifier
      c0 <- genCol
      c1 <- Gen.choose(c0, Cols - 1)
      r0 <- genRow
      r1 <- Gen.choose(r0, Rows)
    yield qualify(q, s"${col(c0)}$r0:${col(c1)}$r1"),
    // whole column, possibly the formula's own
    for
      q <- genQualifier
      c <- Gen.frequency(1 -> Gen.const(column), 2 -> genCol)
    yield qualify(q, s"${col(c)}:${col(c)}"),
    // running total ending at (or just above) the formula's row
    for
      c <- genCol
      end <- Gen.oneOf(row, math.max(1, row - 1))
    yield s"$$${col(c)}$$1:${col(c)}$end",
    // wider than CompactGraph.Wide columns, and a whole row
    Gen.oneOf(s"A1:CZ$Rows", s"$row:$row", "1:3"),
    Gen.oneOf("nmCol", "nmBox", "nmOther")
  )

  private def genFormula(row: Int, column: Int): Gen[String] =
    Gen.frequency(
      3 -> genPoint.map(p => s"$p+1"),
      2 -> Gen.zip(genPoint, genPoint).map((a, b) => s"$a*2+$b"),
      4 -> genRange(row, column).map(r => s"SUM($r)"),
      2 -> Gen.zip(genRange(row, column), genPoint).map((r, p) => s"COUNT($r)+$p"),
      1 -> Gen.const("nmCell*2"),
      1 -> genCol.map(c => s"VLOOKUP(1,A1:${col(Cols - 1)}$Rows,${c + 1},FALSE)"),
      1 -> Gen.const("SUM("),
      1 -> Gen.const("1+2")
    )

  private def genCell(sheet: String, row: Int, column: Int): Gen[Option[CellValue]] =
    Gen.frequency(
      3 -> Gen.const(None),
      3 -> Gen.choose(0, 9).map(n => Some(CellValue.Number(BigDecimal(n)))),
      5 -> genFormula(row, column).map(f => Some(CellValue.Formula(f))),
      1 -> Gen.const(
        Some(
          CellValue.Formula(
            "TABLE(,A1)",
            Some(CellValue.Number(BigDecimal(1))),
            FormulaKind.DataTable(
              CellRange(ARef.from0(column, row - 1), ARef.from0(column, row - 1)),
              dt2D = false,
              dtr = false,
              r1 = Some(ARef.from0(0, 0)),
              r2 = None
            )
          )
        )
      )
    )

  private def genSheetCells(name: String): Gen[Sheet] =
    val positions = for r <- 1 to Rows; c <- 0 until Cols yield (r, c)
    Gen
      .sequence[Vector[Option[(ARef, Cell)]], Option[(ARef, Cell)]](positions.map { (r, c) =>
        genCell(name, r, c).map(_.map { v =>
          val ref = ARef.from0(c, r - 1)
          ref -> Cell(ref, v)
        })
      })
      .map(cells => Sheet(SheetName.unsafe(name), cells = cells.flatten.toMap))

  /** A random three-sheet book with the names its formulas read. */
  val genWorkbook: Gen[Workbook] =
    for
      sheets <- Gen.sequence[Vector[Sheet], Sheet](sheetNames.map(genSheetCells))
      colSheet <- genSheet
      boxSheet <- genSheet
    yield Workbook(sheets)
      .withDefinedName("nmCol", s"'$colSheet'!$$A:$$A")
      .withDefinedName("nmBox", s"'$boxSheet'!$$B$$2:$$C$$7")
      .withDefinedName("nmOther", "'Z'!$D$1:$D$9")
      .withDefinedName("nmCell", "'Mid'!$B$3")

  /** A book of `formulas` (sheet, A1 ref, text) on the named sheets, constants elsewhere unset. */
  def book(formulas: (String, String, String)*): Workbook =
    val bySheet = formulas.groupBy(_._1)
    val names = formulas.map(_._1).distinct.toVector
    Workbook(names.map { name =>
      val cells = bySheet.getOrElse(name, Nil).map { (_, a1, text) =>
        val ref = ARef.parse(a1).getOrElse(ARef.from0(0, 0))
        ref -> Cell(ref, CellValue.Formula(text))
      }
      Sheet(SheetName.unsafe(name), cells = cells.toMap)
    })
