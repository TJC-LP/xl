package com.tjclp.xl.cli

import java.nio.file.{Files, Path}

import cats.effect.IO
import munit.CatsEffectSuite

import com.tjclp.xl.{Sheet, Workbook, given}
import com.tjclp.xl.addressing.{CellRange, Column, Row, SheetName}
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.cf.{CfOperator, CfRule}
import com.tjclp.xl.cli.commands.DiffCommands
import com.tjclp.xl.io.ExcelIO
import com.tjclp.xl.macros.ref
import com.tjclp.xl.ooxml.TestFixtures
import com.tjclp.xl.sheets.{ColumnProperties, DataValidation, RowProperties}
import com.tjclp.xl.styles.{CellStyle, Dxf}
import com.tjclp.xl.styles.color.Color
import com.tjclp.xl.styles.font.Font
import com.tjclp.xl.workbooks.DefinedName

/**
 * The no-false-positive law of `diff` over the real-file corpus (openpyxl, LibreOffice, Excel and
 * derived files): a book compared with an independent read of the same bytes, and with its own
 * surgical re-write, is identical; its full regeneration (source context dropped) differs in no
 * STRUCTURAL category — the cell-level fidelity of a regenerating write is not this law's concern.
 */
class DiffFixtureRoundTripSpec extends CatsEffectSuite:

  private val excel = ExcelIO.instance[IO]

  /**
   * Fixtures whose regenerating write loses a structural field, keyed by file name with the reason
   * and its issue. Empty: every fixture regenerates its structure faithfully. An entry exempts only
   * variant (c); the byte-copy and surgical variants still run.
   */
  private val knownRegenerationGaps: Map[String, String] = Map.empty

  private def tempXlsx(label: String): IO[Path] =
    IO.blocking {
      val p = Files.createTempFile(s"xl-diff-law-$label-", ".xlsx")
      p.toFile.deleteOnExit()
      p
    }

  private def roundTrip(wb: Workbook, label: String): IO[Workbook] =
    tempXlsx(label).flatMap(p => excel.write(wb, p) *> excel.read(p))

  private def diff(a: Workbook, b: Workbook): DiffCommands.WorkbookDiff =
    DiffCommands.computeDiff(a, b, None).fold(err => fail(err), identity)

  TestFixtures.all.foreach { name =>
    test(s"diff law: $name (byte copy, surgical write, regenerating write)") {
      val label = name.stripSuffix(".xlsx")
      for
        a <- IO.blocking(TestFixtures.copyToTemp(name)).flatMap(excel.read)
        b <- IO.blocking(TestFixtures.copyToTemp(name)).flatMap(excel.read)
        surgical <- roundTrip(a, s"$label-surgical")
        regenerated <- roundTrip(a.copy(sourceContext = None), s"$label-fresh")
      yield
        val copies = diff(a, b)
        assert(copies.identical, DiffCommands.renderMarkdown(copies, name, s"$name (copy)"))
        val rewrite = diff(a, surgical)
        assert(rewrite.identical, DiffCommands.renderMarkdown(rewrite, name, s"$name (surgical)"))
        val fresh = diff(a, regenerated)
        if !knownRegenerationGaps.contains(name) then
          assertEquals(
            fresh.structureChanges,
            0,
            DiffCommands.renderMarkdown(fresh, name, s"$name (regenerated)")
          )
    }
  }

  /**
   * The corpus carries row heights, sheet defaults, conditional formats, names and sheet states but
   * no column properties, validations, freeze panes, outlines or hidden sheets: this authored book
   * carries every compared category, so each one crosses the writer and the reader.
   */
  private def structureRichBook: Workbook =
    val range = (a1: String) => CellRange.parse(a1).fold(err => fail(err.toString), identity)
    val bold = CellStyle.default.withFont(Font.default.withBold(true))
    val italic = CellStyle.default.withFont(Font.default.withItalic(true))
    val rows = Vector(
      2 -> RowProperties(height = Some(30)),
      3 -> RowProperties(hidden = true),
      4 -> RowProperties(outlineLevel = Some(1)),
      5 -> RowProperties(outlineLevel = Some(1), hidden = true),
      6 -> RowProperties(collapsed = true)
    )
    val columns = Vector(
      1 -> ColumnProperties(width = Some(20)),
      2 -> ColumnProperties(hidden = true),
      3 -> ColumnProperties(outlineLevel = Some(2))
    )
    val tail = (10 to Column.MaxIndex0).map(_ -> ColumnProperties(hidden = true))
    val withRows = rows.foldLeft(Sheet("Model").put(ref"A1", CellValue.Number(1))) {
      case (sheet, (n, props)) => sheet.setRowProperties(Row.from1(n), props)
    }
    val model = (columns ++ tail)
      .foldLeft(withRows) { case (s, (i, props)) => s.setColumnProperties(Column.from0(i), props) }
      .withRowStyle(Row.from1(8), bold)
      .withColumnStyle(Column.from0(6), italic)
      .copy(defaultRowHeight = Some(18), defaultColumnWidth = Some(12))
      .freezeAt(ref"B2")
      .conditionalFormat(
        range("A1:A9"),
        CfRule.cellIs(CfOperator.GreaterThan, "100", Dxf.fill(Color.Rgb(0xffff0000)))
      )
      .withDataValidation(range("B2:B9"), DataValidation.listOf("Low", "High"))
    val book = Workbook(Vector(model, Sheet("Inputs").put(ref"A1", "x"), Sheet("Archive")))
    val hidden = book
      .setSheetState(SheetName.unsafe("Archive"), Some("hidden"))
      .fold(err => fail(err.message), identity)
    hidden.copy(metadata =
      hidden.metadata.copy(definedNames =
        Vector(
          DefinedName("Rate", "Model!$A$1"),
          DefinedName("LocalRate", "Model!$B$1", Some(1)),
          DefinedName("Secret", "Model!$C$1", hidden = true)
        )
      )
    )

  test("diff law: an authored book carrying every structural category survives write and read") {
    val authored = structureRichBook
    for
      read <- roundTrip(authored, "authored")
      surgical <- roundTrip(read, "authored-surgical")
      regenerated <- roundTrip(read.copy(sourceContext = None), "authored-fresh")
    yield
      val written = diff(authored, read)
      assertEquals(
        written.structureChanges,
        0,
        DiffCommands.renderMarkdown(written, "authored", "written")
      )
      assert(diff(read, surgical).identical)
      val fresh = diff(read, regenerated)
      assertEquals(fresh.structureChanges, 0, DiffCommands.renderMarkdown(fresh, "read", "fresh"))
      // and the law is not vacuous: every category differs from a plain book
      val plain = Workbook(Vector(Sheet("Inputs"), Sheet("Model"), Sheet("Archive")))
      val all = diff(plain, read)
      val model = all.sheets.find(_.name == "Model").getOrElse(fail("Model not compared"))
      assert(model.rows.nonEmpty && model.columns.nonEmpty && model.properties.nonEmpty)
      assert(model.conditionalFormatsAdded.nonEmpty && model.dataValidationsAdded.nonEmpty)
      assert(all.sheetOrder.isDefined && all.namesAdded.length == 3)
      assert(all.sheets.exists(_.properties.exists(_.property == DiffCommands.Property.Visibility)))
  }
