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

  // ===== GH-869: widening on entry, as Excel does for a formatted number =====

  private val currency = com.tjclp.xl.styles.CellStyle.default
    .withNumFmt(com.tjclp.xl.styles.numfmt.NumFmt.Currency)

  private def withCurrency(sheet: Sheet, ref: ARef, n: BigDecimal): Sheet =
    sheet.put(ref, CellValue.Number(n)).withCellStyle(ref, currency)

  test("GH-869: an overflowing currency entry widens its column to the autofit width") {
    val a1 = ARef.from0(0, 0)
    val s = withCurrency(Sheet(SheetName.unsafe("S")), a1, BigDecimal("1234567.89"))
    val (widened, cols) = SheetEdits.widenForEntries(s, List(a1))
    val fit = s.autoFitWidth(Column.from0(0))
    assertEquals(cols, Vector(Column.from0(0) -> fit))
    assertEquals(widened.getColumnProperties(Column.from0(0)).width, Some(fit))
  }

  test("GH-869: only the written refs count, and a column is never narrowed") {
    val a1 = ARef.from0(0, 0)
    val a2 = ARef.from0(0, 1)
    val s = withCurrency(
      withCurrency(Sheet(SheetName.unsafe("S")), a1, BigDecimal("1234567.89")),
      a2,
      BigDecimal(1)
    )
    // writing only the short A2 leaves the column as it is, though A1 overflows
    assertEquals(SheetEdits.widenForEntries(s, List(a2))._2, Vector.empty)
    // a sheet default already wider than the entry needs is not narrowed
    val wide = s.copy(defaultColumnWidth = Some(40.0))
    assertEquals(SheetEdits.widenForEntries(wide, List(a1))._2, Vector.empty)
  }

  test("GH-869: a number inside a merge is measured by the merge, not its column") {
    val a1 = ARef.from0(0, 0)
    val s = withCurrency(Sheet(SheetName.unsafe("S")), a1, BigDecimal("1234567.89"))
      .copy(mergedRanges = Set(com.tjclp.xl.addressing.CellRange(a1, ARef.from0(2, 0))))
    assertEquals(SheetEdits.widenForEntries(s, List(a1))._2, Vector.empty)
  }

  test("GH-869 review: overflow is judged in 7 px digits, whatever font measured it") {
    val unsized = 64.0 / 7.0
    // Calibri metrics: 59 px of text plus 5 px padding is exactly the unsized column's 64 px
    assert(SheetEdits.scaledFit(59, 7) <= unsized)
    assert(SheetEdits.scaledFit(60, 7) > unsized)
    // a fallback font with a 9 px digit: 8 characters of `12/31/26` measure 64 px, which is
    // 64 * 7/9 = 49.8 px in Calibri terms and fits; unscaled it would have widened
    assert(SheetEdits.scaledFit(64, 9) <= unsized)
    assert(SheetEdits.scaledFit(81, 9) > unsized)
  }
