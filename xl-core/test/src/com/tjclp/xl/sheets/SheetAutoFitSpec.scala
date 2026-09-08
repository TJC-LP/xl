package com.tjclp.xl.sheets

import com.tjclp.xl.Generators
import com.tjclp.xl.addressing.{ARef, Column, SheetName}
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.ops.{ColSpan, Edit, FormulaSupport, Scope}
import com.tjclp.xl.patch.Patch
import com.tjclp.xl.sheets.styleSyntax.withCellStyle
import com.tjclp.xl.workbooks.Workbook
import munit.ScalaCheckSuite
import org.scalacheck.Gen
import org.scalacheck.Prop.forAll

/**
 * `SheetEdits.autoFit` groups the sheet's cells by column once and fits every column from that
 * grouping. This suite pins it to the per-column implementation it replaced — `autoFitWidth(col)`
 * scans the cells for that one column — so the single pass cannot drift by a digit: same widths,
 * same properties, same sheet.
 */
class SheetAutoFitSpec extends ScalaCheckSuite:

  // ===== GH-613: a column is fitted to what it displays, never to uncached formula text =====

  test("GH-613: an uncached formula contributes nothing; a cached one is measured by its value") {
    val long = "COUNTIF('Deal Pipeline'!$E$2:$E$1000,A2)"
    val uncachedOnly =
      Sheet(SheetName.unsafe("S")).put(ARef.from0(0, 0), CellValue.Formula(long, None))
    // nothing measurable in the column: the floor, not the 40-odd characters of formula text
    assertEquals(uncachedOnly.autoFitWidth(Column.from0(0)), 5.0)
    val cached = Sheet(SheetName.unsafe("S"))
      .put(ARef.from0(0, 0), CellValue.Formula(long, Some(CellValue.Number(BigDecimal(6816)))))
      .put(ARef.from0(0, 1), CellValue.Formula(long, None))
    val width = cached.autoFitWidth(Column.from0(0))
    assert(width < 12.0, s"a cached 6816 must size like a four-digit number, got $width")
    assertEquals(
      width,
      Sheet(SheetName.unsafe("S"))
        .put(ARef.from0(0, 0), CellValue.Number(BigDecimal(6816)))
        .autoFitWidth(Column.from0(0))
    )
  }

  override def scalaCheckTestParameters =
    super.scalaCheckTestParameters.withMinSuccessfulTests(100)

  private given FormulaSupport = FormulaSupport.textOnly

  /** The per-column fold `autoFit` used to be: one `autoFitWidth` scan per column, in order. */
  private def perColumn(sheet: Sheet, columns: Iterable[Column]): Sheet =
    columns.foldLeft(sheet) { (s, col) =>
      s.setColumnProperties(col, s.getColumnProperties(col).copy(width = Some(s.autoFitWidth(col))))
    }

  private def usedColumns(sheet: Sheet): Vector[Column] =
    sheet.usedRange.fold(Vector.empty[Column])(r =>
      (r.colStart.index0 to r.colEnd.index0).map(Column.from0).toVector
    )

  /** Sheets whose cells cluster in a few columns, so most fitted columns have several cells. */
  private val genClusteredSheet: Gen[Sheet] =
    for
      n <- Gen.choose(0, 30)
      cells <- Gen.listOfN(
        n,
        for
          col <- Gen.choose(0, 5)
          row <- Gen.choose(0, 19)
          value <- Generators.genCellValue
        yield ARef.from0(col, row) -> value
      )
      styled <- Gen.listOf(Gen.zip(Gen.choose(0, 5), Gen.choose(0, 19), Generators.genCellStyle))
    yield
      val base = cells.foldLeft(Sheet(SheetName.unsafe("Fit"))) { case (s, (ref, v)) =>
        s.put(ref, v)
      }
      styled.take(4).foldLeft(base) { case (s, (c, r, style)) =>
        val ref = ARef.from0(c, r)
        if s.contains(ref) then s.withCellStyle(ref, style) else s
      }

  private val genColumns: Gen[Vector[Column]] =
    Gen.listOf(Gen.choose(0, 7).map(Column.from0)).map(_.toVector)

  property("autoFit(columns) equals the per-column fold, width for width, in column order") {
    forAll(genClusteredSheet, genColumns) { (sheet: Sheet, columns: Vector[Column]) =>
      assertEquals(sheet.autoFit(columns), perColumn(sheet, columns))
      assertEquals(
        SheetEdits.autoFitWidths(sheet, columns),
        columns.map(c => c -> sheet.autoFitWidth(c))
      )
      true
    }
  }

  property("autoFitAll equals the per-column fold over the used columns, on any generated sheet") {
    forAll(Gen.oneOf(genClusteredSheet, Generators.genSheet)) { (sheet: Sheet) =>
      assertEquals(sheet.autoFitAll, perColumn(sheet, usedColumns(sheet)))
      true
    }
  }

  property("Edit.AutoFit applied, Edit.lower(AutoFit) through the kernel, and autoFitAll agree") {
    forAll(genClusteredSheet) { (sheet: Sheet) =>
      val edit = Edit.AutoFit(None, None)
      val applied = Edit.applyAll(Workbook(sheet), Vector(edit), Scope.of(sheet.name))
      assertEquals(applied.flatMap(_.workbook(sheet.name)), Right(sheet.autoFitAll))
      assertEquals(Edit.lower(edit, sheet).map(Patch.applyPatch(sheet, _)), Some(sheet.autoFitAll))
      true
    }
  }

  test("an empty column fits to Excel's default; repeated and unused columns are each reported") {
    val s = Sheet(SheetName.unsafe("Fit")).put(ARef.from0(1, 0), CellValue.Text("wide enough text"))
    val cols = Vector(Column.from0(0), Column.from0(1), Column.from0(1), Column.from0(9))
    val widths = SheetEdits.autoFitWidths(s, cols)
    assertEquals(widths.map(_._1), cols)
    assertEquals(widths(0)._2, SheetEdits.DefaultColumnWidth)
    assertEquals(widths(3)._2, SheetEdits.DefaultColumnWidth)
    assertEquals(widths(1), widths(2))
    assert(widths(1)._2 > SheetEdits.DefaultColumnWidth)
    assertEquals(s.autoFit(cols), perColumn(s, cols))
    // the span form the interpreter uses
    assertEquals(
      s.autoFit(ColSpan(Column.from0(0), Column.from0(2)).columns),
      perColumn(s, ColSpan(Column.from0(0), Column.from0(2)).columns)
    )
  }
