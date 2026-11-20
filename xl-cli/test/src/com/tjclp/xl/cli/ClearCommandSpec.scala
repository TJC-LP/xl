package com.tjclp.xl.cli

import munit.FunSuite

import java.nio.file.{Files, Path}

import cats.effect.{IO, unsafe}
import com.tjclp.xl.{Workbook, Sheet}
import com.tjclp.xl.addressing.{ARef, CellRange}
import com.tjclp.xl.cells.{CellValue, Comment}
import com.tjclp.xl.cli.commands.CellCommands
import com.tjclp.xl.io.ExcelIO
import com.tjclp.xl.ops.ClearWhat
import com.tjclp.xl.ooxml.writer.WriterConfig
import com.tjclp.xl.styles.CellStyle
import com.tjclp.xl.styles.fill.Fill
import com.tjclp.xl.styles.color.Color
import com.tjclp.xl.styles.border.{Border, BorderStyle}
import com.tjclp.xl.styles.font.Font
import com.tjclp.xl.styles.numfmt.NumFmt
import com.tjclp.xl.sheets.styleSyntax.*
import com.tjclp.xl.cli.helpers.BatchParser
import com.tjclp.xl.cli.helpers.BatchParser.BatchOp

/**
 * Integration tests for clear command functionality.
 */
@SuppressWarnings(
  Array("org.wartremover.warts.OptionPartial", "org.wartremover.warts.IterableOps")
)
class ClearCommandSpec extends FunSuite:

  given unsafe.IORuntime = unsafe.IORuntime.global

  val outputPath: Path = Files.createTempFile("test-clear", ".xlsx")
  val config: WriterConfig = WriterConfig.default

  override def afterEach(context: AfterEach): Unit =
    if Files.exists(outputPath) then Files.delete(outputPath)

  // ========== Clear Contents (default mode) ==========

  test("clear: default mode clears cell contents") {
    val sheet = Sheet("Test")
      .put(ARef.from0(0, 0), CellValue.Text("A1"))
      .put(ARef.from0(1, 0), CellValue.Text("B1"))
      .put(ARef.from0(0, 1), CellValue.Text("A2"))
    val wb = Workbook(sheet)

    val result = CellCommands
      .clear(
        wb,
        Some(sheet),
        "A1:B1",
        all = false,
        styles = false,
        comments = false,
        outputPath,
        config
      )
      .unsafeRunSync()

    assert(result.contains("Cleared contents"))

    val imported = ExcelIO.instance[IO].read(outputPath).unsafeRunSync()
    val s = imported.sheets.head
    // A1 and B1 should be cleared
    assertEquals(s.cells.get(ARef.from0(0, 0)), None)
    assertEquals(s.cells.get(ARef.from0(1, 0)), None)
    // A2 should remain
    assertEquals(s.cells.get(ARef.from0(0, 1)).map(_.value), Some(CellValue.Text("A2")))
  }

  test("clear: single cell reference works") {
    val sheet = Sheet("Test")
      .put(ARef.from0(0, 0), CellValue.Text("A1"))
      .put(ARef.from0(1, 0), CellValue.Text("B1"))
    val wb = Workbook(sheet)

    val result = CellCommands
      .clear(
        wb,
        Some(sheet),
        "A1",
        all = false,
        styles = false,
        comments = false,
        outputPath,
        config
      )
      .unsafeRunSync()

    assert(result.contains("A1:A1"))

    val imported = ExcelIO.instance[IO].read(outputPath).unsafeRunSync()
    val s = imported.sheets.head
    assertEquals(s.cells.get(ARef.from0(0, 0)), None)
    assertEquals(s.cells.get(ARef.from0(1, 0)).map(_.value), Some(CellValue.Text("B1")))
  }

  // ========== Clear Styles Only ==========

  test("clear --styles: clears styles but keeps contents and comments") {
    val style = CellStyle.default.withFill(Fill.Solid(Color.Rgb(0xffff00)))
    val ref = ARef.from0(0, 0)
    val sheet = Sheet("Test")
      .put(ref, CellValue.Text("A1"))
      .comment(ref, Comment.plainText("Note", None))
    // Apply style via styleRegistry
    val styledSheet =
      com.tjclp.xl.sheets.styleSyntax.withRangeStyle(sheet)(CellRange(ref, ref), style)
    val wb = Workbook(styledSheet)

    // Verify style is applied
    val styleId = styledSheet.cells.get(ref).flatMap(_.styleId)
    assert(styleId.isDefined, "Cell should have styleId before clearing")

    val result = CellCommands
      .clear(
        wb,
        Some(styledSheet),
        "A1",
        all = false,
        styles = true,
        comments = false,
        outputPath,
        config
      )
      .unsafeRunSync()

    assert(result.contains("Cleared styles"))

    val imported = ExcelIO.instance[IO].read(outputPath).unsafeRunSync()
    val s = imported.sheets.head
    // Content should remain
    assertEquals(s.cells.get(ref).map(_.value), Some(CellValue.Text("A1")))
    // Comment should remain
    assert(s.comments.contains(ref))
    // Style should be cleared
    assertEquals(s.cells.get(ref).flatMap(_.styleId), None)
  }

  // ========== Clear Comments Only ==========

  test("clear --comments: clears comments but keeps contents and styles") {
    val ref = ARef.from0(0, 0)
    val sheet = Sheet("Test")
      .put(ref, CellValue.Text("A1"))
      .comment(ref, Comment.plainText("Test note", Some("Author")))
    val wb = Workbook(sheet)

    // Verify comment exists
    assert(sheet.comments.contains(ref))

    val result = CellCommands
      .clear(
        wb,
        Some(sheet),
        "A1",
        all = false,
        styles = false,
        comments = true,
        outputPath,
        config
      )
      .unsafeRunSync()

    assert(result.contains("Cleared comments"))

    val imported = ExcelIO.instance[IO].read(outputPath).unsafeRunSync()
    val s = imported.sheets.head
    // Content should remain
    assertEquals(s.cells.get(ref).map(_.value), Some(CellValue.Text("A1")))
    // Comment should be cleared
    assert(!s.comments.contains(ref))
  }

  test("clear --comments: clears comments from range") {
    val ref1 = ARef.from0(0, 0)
    val ref2 = ARef.from0(0, 1)
    val ref3 = ARef.from0(0, 2)
    val sheet = Sheet("Test")
      .put(ref1, CellValue.Text("A1"))
      .put(ref2, CellValue.Text("A2"))
      .put(ref3, CellValue.Text("A3"))
      .comment(ref1, Comment.plainText("Note 1", None))
      .comment(ref2, Comment.plainText("Note 2", None))
      .comment(ref3, Comment.plainText("Note 3", None))
    val wb = Workbook(sheet)

    val result = CellCommands
      .clear(
        wb,
        Some(sheet),
        "A1:A2",
        all = false,
        styles = false,
        comments = true,
        outputPath,
        config
      )
      .unsafeRunSync()

    val imported = ExcelIO.instance[IO].read(outputPath).unsafeRunSync()
    val s = imported.sheets.head
    // A1, A2 comments cleared
    assert(!s.comments.contains(ref1))
    assert(!s.comments.contains(ref2))
    // A3 comment should remain
    assert(s.comments.contains(ref3))
  }

  // ========== Clear All ==========

  test("clear --all: clears contents, styles, and comments") {
    val ref = ARef.from0(0, 0)
    val style = CellStyle.default.withFill(Fill.Solid(Color.Rgb(0xff0000)))
    val sheet = Sheet("Test")
      .put(ref, CellValue.Text("A1"))
      .comment(ref, Comment.plainText("Note", None))
    val styledSheet =
      com.tjclp.xl.sheets.styleSyntax.withRangeStyle(sheet)(CellRange(ref, ref), style)
    val wb = Workbook(styledSheet)

    val result = CellCommands
      .clear(
        wb,
        Some(styledSheet),
        "A1",
        all = true,
        styles = false,
        comments = false,
        outputPath,
        config
      )
      .unsafeRunSync()

    assert(result.contains("contents"))
    assert(result.contains("styles"))
    assert(result.contains("comments"))

    val imported = ExcelIO.instance[IO].read(outputPath).unsafeRunSync()
    val s = imported.sheets.head
    // Cell should be completely gone
    assertEquals(s.cells.get(ref), None)
    // Comment should be gone
    assert(!s.comments.contains(ref))
  }

  // ========== Formula Clearing ==========

  test("clear: default mode clears formula cells") {
    val ref = ARef.from0(0, 0)
    val sheet = Sheet("Test")
      .put(ref, CellValue.Formula("=1+1", Some(CellValue.Number(2.0))))
      .put(ARef.from0(1, 0), CellValue.Text("B1"))
    val wb = Workbook(sheet)

    // Verify formula is present
    assert(sheet.cells.get(ref).exists { cell =>
      cell.value match
        case _: CellValue.Formula => true
        case _ => false
    })

    val result = CellCommands
      .clear(
        wb,
        Some(sheet),
        "A1",
        all = false,
        styles = false,
        comments = false,
        outputPath,
        config
      )
      .unsafeRunSync()

    assert(result.contains("Cleared contents"))

    val imported = ExcelIO.instance[IO].read(outputPath).unsafeRunSync()
    val s = imported.sheets.head
    // Formula cell should be cleared
    assertEquals(s.cells.get(ref), None)
    // Other cell should remain
    assertEquals(s.cells.get(ARef.from0(1, 0)).map(_.value), Some(CellValue.Text("B1")))
  }

  // ========== Merged Region Handling ==========

  test("clear: default mode keeps a merge it touches; only the values go (GH-886)") {
    val sheet = Sheet("Test")
      .put(ARef.from0(0, 0), CellValue.Text("Merged"))
      .put(ARef.from0(2, 0), CellValue.Text("C1"))
      .merge(CellRange(ARef.from0(0, 0), ARef.from0(1, 0))) // Merge A1:B1
    val wb = Workbook(sheet)

    assertEquals(sheet.mergedRanges.size, 1)

    // Clear A1, the merge's anchor: Excel's Clear Contents leaves A1:B1 merged
    CellCommands
      .clear(
        wb,
        Some(sheet),
        "A1",
        all = false,
        styles = false,
        comments = false,
        outputPath,
        config
      )
      .unsafeRunSync()

    val imported = ExcelIO.instance[IO].read(outputPath).unsafeRunSync()
    val s = imported.sheets.head
    assertEquals(s.mergedRanges, Set(CellRange(ARef.from0(0, 0), ARef.from0(1, 0))))
    assertEquals(
      s.cells.get(ARef.from0(0, 0)).map(_.value).getOrElse(CellValue.Empty),
      CellValue.Empty
    )
    assertEquals(s.cells.get(ARef.from0(2, 0)).map(_.value), Some(CellValue.Text("C1")))
  }

  // ========== GH-886: a contents clear keeps formatting, as Excel's Clear Contents ==========

  private val numStyle = CellStyle.default
    .withNumFmt(NumFmt.Custom("#,##0_);\\(#,##0\\);\\-\\-_)"))
    .withFont(Font.default.withName("Times New Roman").withColor(Color.Rgb(0xff0000ff)))
    .withFill(Fill.Solid(Color.Rgb(0xffffff99)))
    .withBorder(Border.all(BorderStyle.Thin))
  private val titleStyle =
    CellStyle.default.withFont(Font.default.withName("Georgia").withSize(16.0).withBold(true))
  private val dataRange = CellRange.parse("B4:E5").fold(e => fail(e), identity)
  private val titleRange = CellRange.parse("A1:F1").fold(e => fail(e), identity)

  /**
   * A formatted report: a merged bold title over A1:F1, styled numbers in B4:E5, and an unstyled
   * note in A3.
   */
  private def template: Sheet =
    val withValues = dataRange.cells.zipWithIndex.foldLeft(
      Sheet("Summary")
        .put(ARef.from0(0, 0), CellValue.Text("Quarterly report"))
        .put(ARef.from0(0, 2), CellValue.Number(1))
    ) { case (s, (r, i)) => s.put(r, CellValue.Number(BigDecimal(1000 * (i + 1)))) }
    withValues
      .withRangeStyle(dataRange, numStyle)
      .withCellStyle(ARef.from0(0, 0), titleStyle)
      .merge(titleRange)

  private def assertFormattingKept(s: Sheet): Unit =
    dataRange.cells.foreach { r =>
      val cell = s.cells.get(r)
      assert(cell.isDefined, s"${r.toA1}: the styled cell record should stay")
      assertEquals(cell.map(_.value), Some(CellValue.Empty), s"${r.toA1}: value cleared")
      val style = s.getCellStyle(r)
      assertEquals(style.map(_.numFmt), Some(numStyle.numFmt), s"${r.toA1}: numFmt")
      assertEquals(style.map(_.font.name), Some("Times New Roman"), s"${r.toA1}: font")
      assertEquals(style.flatMap(_.font.color), Some(Color.Rgb(0xff0000ff)), s"${r.toA1}: colour")
      assertEquals(style.map(_.fill), Some(numStyle.fill), s"${r.toA1}: fill")
      assertEquals(style.map(_.border), Some(numStyle.border), s"${r.toA1}: border")
    }

  private def clearVerb(
    range: String,
    all: Boolean = false,
    stream: Boolean = false,
    styles: Boolean = false,
    comments: Boolean = false
  ): Sheet =
    val sheet = template
    CellCommands
      .clear(
        Workbook(sheet),
        Some(sheet),
        range,
        all = all,
        styles = styles,
        comments = comments,
        outputPath,
        config,
        stream = stream
      )
      .unsafeRunSync()
    ExcelIO.instance[IO].read(outputPath).unsafeRunSync().sheets.head

  test("GH-886 clear: a styled range keeps numFmt, font, fill and border") {
    assertFormattingKept(clearVerb("B4:E5"))
  }

  test("GH-886 clear: clearing a merged title keeps the merge and the title's font") {
    val s = clearVerb("A1")
    assertEquals(s.mergedRanges, Set(titleRange))
    assertEquals(s.cells.get(ARef.from0(0, 0)).map(_.value), Some(CellValue.Empty))
    assertEquals(s.getCellStyle(ARef.from0(0, 0)).map(_.font.name), Some("Georgia"))
  }

  test("GH-886 clear --stream: keeps formatting and the merge too") {
    assertFormattingKept(clearVerb("B4:E5", stream = true))
    assertEquals(clearVerb("A1", stream = true).mergedRanges, Set(titleRange))
  }

  test("GH-886 clear --all: still drops the cells, their formatting and the merge it touches") {
    val data = clearVerb("B4:E5", all = true)
    dataRange.cells.foreach(r => assertEquals(data.cells.get(r), None, r.toA1))
    val title = clearVerb("A1", all = true)
    assertEquals(title.mergedRanges, Set.empty[CellRange])
    assertEquals(title.cells.get(ARef.from0(0, 0)), None)
  }

  test("GH-886 batch clear: keeps formatting and merges; all=true still resets both") {
    val wb = Workbook(template)
    val contents = BatchParser
      .applyBatchOperations(
        wb,
        None,
        Vector(
          BatchOp.Clear("Summary!B4:E5", all = false, styles = false, comments = false),
          BatchOp.Clear("Summary!A1", all = false, styles = false, comments = false)
        )
      )
      .unsafeRunSync()
      .sheets
      .head
    assertFormattingKept(contents)
    assertEquals(contents.mergedRanges, Set(titleRange))

    val all = BatchParser
      .applyBatchOperations(
        wb,
        None,
        Vector(
          BatchOp.Clear("Summary!B4:E5", all = true, styles = false, comments = false),
          BatchOp.Clear("Summary!A1", all = true, styles = false, comments = false)
        )
      )
      .unsafeRunSync()
      .sheets
      .head
    dataRange.cells.foreach(r => assertEquals(all.cells.get(r), None, r.toA1))
    assertEquals(all.mergedRanges, Set.empty[CellRange])
  }

  // ========== GH-894: --styles is Excel's Clear Formats, which unmerges ==========

  test("GH-894 clear --styles: unmerges the merge it touches and keeps the value") {
    val s = clearVerb("A1", styles = true)
    assertEquals(s.mergedRanges, Set.empty[CellRange])
    assertEquals(
      s.cells.get(ARef.from0(0, 0)).map(_.value),
      Some(CellValue.Text("Quarterly report"))
    )
    assertNotEquals(s.getCellStyle(ARef.from0(0, 0)).map(_.font.name), Some("Georgia"))
  }

  test("GH-894 clear --styles: a range that only overlaps the merge unmerges it and resets A1") {
    Vector("B1", "C1:D2").foreach { r =>
      val s = clearVerb(r, styles = true)
      assertEquals(s.mergedRanges, Set.empty[CellRange], r)
      assertNotEquals(s.getCellStyle(ARef.from0(0, 0)).map(_.font.name), Some("Georgia"), r)
      assertEquals(
        s.cells.get(ARef.from0(0, 0)).map(_.value),
        Some(CellValue.Text("Quarterly report")),
        r
      )
    }
  }

  test("GH-894 clear --styles --stream: unmerges too") {
    assertEquals(clearVerb("A1", styles = true, stream = true).mergedRanges, Set.empty[CellRange])
  }

  test("GH-894 clear --comments: keeps the merge (comments are not formatting)") {
    assertEquals(clearVerb("A1", comments = true).mergedRanges, Set(titleRange))
  }

  test("GH-894 batch clear styles=true: unmerges; comments=true keeps the merge") {
    def run(styles: Boolean, comments: Boolean): Sheet =
      BatchParser
        .applyBatchOperations(
          Workbook(template),
          None,
          Vector(BatchOp.Clear("Summary!A1", all = false, styles = styles, comments = comments))
        )
        .unsafeRunSync()
        .sheets
        .head
    val styled = run(styles = true, comments = false)
    assertEquals(styled.mergedRanges, Set.empty[CellRange])
    assertEquals(
      styled.cells.get(ARef.from0(0, 0)).map(_.value),
      Some(CellValue.Text("Quarterly report"))
    )
    assertEquals(run(styles = false, comments = true).mergedRanges, Set(titleRange))
    val partial = BatchParser
      .applyBatchOperations(
        Workbook(template),
        None,
        Vector(BatchOp.Clear("Summary!C1", all = false, styles = true, comments = false))
      )
      .unsafeRunSync()
      .sheets
      .head
    assertEquals(partial.mergedRanges, Set.empty[CellRange])
    assertNotEquals(partial.getCellStyle(ARef.from0(0, 0)).map(_.font.name), Some("Georgia"))
  }

  test("GH-886 clear: a hyperlink goes with the contents; its cell keeps its formatting") {
    // Excel's Clear Contents removes the hyperlink and keeps the underline font (checked in
    // Excel for Mac 16.115); a hyperlink cell with no style is removed outright.
    val link = CellStyle.default.withFont(Font.default.withColor(Color.Rgb(0xff0563c1)))
    val b7 = ARef.from0(1, 6)
    val c7 = ARef.from0(2, 6)
    val sheet = template
      .put(b7, CellValue.Text("Example link"))
      .put(c7, CellValue.Text("Bare link"))
      .withCellStyle(b7, link)
    val linked = sheet.copy(cells =
      sheet.cells
        .updated(sheet(b7).withHyperlink("https://example.com/"))
        .updated(sheet(c7).withHyperlink("https://example.com/"))
    )
    CellCommands
      .clear(
        Workbook(linked),
        Some(linked),
        "B7:C7",
        all = false,
        styles = false,
        comments = false,
        outputPath,
        config
      )
      .unsafeRunSync()
    val s = ExcelIO.instance[IO].read(outputPath).unsafeRunSync().sheets.head
    assertEquals(s.cells.get(b7).map(_.value), Some(CellValue.Empty))
    assertEquals(s.cells.get(b7).flatMap(_.hyperlink), None, "B7: hyperlink cleared")
    assertEquals(s.getCellStyle(b7).flatMap(_.font.color), Some(Color.Rgb(0xff0563c1)))
    assertEquals(s.cells.get(c7), None, "C7: an unstyled hyperlink cell is removed")
  }

  test("GH-886 clear: an unstyled cell with no comment is still removed outright") {
    val a3 = ARef.from0(0, 2)
    assertEquals(template.cells.get(a3).map(_.value), Some(CellValue.Number(1)), "fixture: A3")
    assertEquals(template.cells.get(a3).flatMap(_.styleId), None, "fixture: A3 unstyled")
    val s = clearVerb("A3:A3")
    assertEquals(s.cells.get(a3), None)
  }

  // ========== GH-897: an unstyled cell in a styled row or column keeps the default format ==========

  private val b4 = ARef.from0(1, 3)
  private val d6 = ARef.from0(3, 5)
  private val f8 = ARef.from0(5, 7)
  private val bandStyle = CellStyle.default.withFill(Fill.Solid(Color.Rgb(0xffffff00)))

  /**
   * Column B and row 6 carry a default style (a yellow fill); B4 and D6 hold numbers with no `s` of
   * their own, so they show the default format (xf 0). F8 is outside both.
   */
  private def banded: Sheet =
    Sheet("Summary")
      .withColumnStyle(b4.col, bandStyle)
      .withRowStyle(d6.row, bandStyle)
      .put(b4, CellValue.Number(5))
      .put(d6, CellValue.Number(7))
      .put(f8, CellValue.Number(9))

  private def sheetXml: String =
    val zf = new java.util.zip.ZipFile(outputPath.toFile)
    try
      new String(
        zf.getInputStream(zf.getEntry("xl/worksheets/sheet1.xml")).readAllBytes(),
        java.nio.charset.StandardCharsets.UTF_8
      )
    finally zf.close()

  /** The cell keeps a record with no style of its own, so it stays on xf 0, not the band's. */
  private def assertPinnedToDefault(s: Sheet, xml: String, r: ARef): Unit =
    val record = s"""<c r="${r.toA1}"[^>]*>""".r.findFirstIn(xml)
    assert(record.isDefined, s"${r.toA1}: the emptied cell needs a record to stay on xf 0: $xml")
    assert(
      record.forall(rec => !rec.contains(" s=") || rec.contains(""" s="0"""")),
      s"${r.toA1}: the record must not carry the band's style: $record"
    )
    assertEquals(s.cells.get(r).map(_.value), Some(CellValue.Empty), s"${r.toA1}: value cleared")
    assertEquals(
      s.cells
        .get(r)
        .flatMap(_.styleId)
        .flatMap(s.styleRegistry.get)
        .filterNot(_ == CellStyle.default),
      None,
      s"${r.toA1}: no style of its own"
    )

  private def clearBanded(range: String, stream: Boolean = false): Sheet =
    val sheet = banded
    CellCommands
      .clear(
        Workbook(sheet),
        Some(sheet),
        range,
        all = false,
        styles = false,
        comments = false,
        outputPath,
        config,
        stream = stream
      )
      .unsafeRunSync()
    ExcelIO.instance[IO].read(outputPath).unsafeRunSync().sheets.head

  test("GH-897 clear: an unstyled cell in a styled column keeps the default format") {
    val s = clearBanded("B4")
    assertPinnedToDefault(s, sheetXml, b4)
  }

  test("GH-897 clear: an unstyled cell in a styled row keeps the default format") {
    val s = clearBanded("D6")
    assertPinnedToDefault(s, sheetXml, d6)
  }

  test("GH-897 clear --stream: an unstyled cell in a styled row or column keeps the default") {
    val s = clearBanded("B4:D6", stream = true)
    val xml = sheetXml
    assertPinnedToDefault(s, xml, b4)
    assertPinnedToDefault(s, xml, d6)
  }

  test("GH-897 batch clear: an unstyled cell in a styled column keeps the default format") {
    val s = BatchParser
      .applyBatchOperations(
        Workbook(banded),
        None,
        Vector(BatchOp.Clear("Summary!B4", all = false, styles = false, comments = false))
      )
      .unsafeRunSync()
      .sheets
      .head
    assertEquals(s.cells.get(b4).map(_.value), Some(CellValue.Empty))
    assertEquals(s.cells.get(b4).flatMap(_.styleId), None)
  }

  test("GH-897 clear: an unstyled cell outside any styled row or column is still removed") {
    val s = clearBanded("F8")
    assertEquals(s.cells.get(f8), None)
    assert(!sheetXml.contains("""<c r="F8""""), "F8: no record left")
  }

  test("GH-897 clear --all: drops the record, so the cell shows the column's style") {
    val sheet = banded
    CellCommands
      .clear(
        Workbook(sheet),
        Some(sheet),
        "B4",
        all = true,
        styles = false,
        comments = false,
        outputPath,
        config
      )
      .unsafeRunSync()
    val s = ExcelIO.instance[IO].read(outputPath).unsafeRunSync().sheets.head
    assertEquals(s.cells.get(b4), None)
  }

  // ========== GH-908: the CLI and the library share one keep-record rule ==========

  test(
    "GH-908 clear: the CLI verb and SheetEdits.clear give the same sheet over a styled row and column"
  ) {
    // B4 (styled column) and D6 (styled row) hold unstyled numbers, C5 has its own style and a
    // hyperlink, F8 sits outside both bands: the whole range takes every branch of the rule
    val c5 = ARef.from0(2, 4)
    val ownStyle = CellStyle.default.withFont(Font.default.withBold(true))
    val styled = banded.put(c5, CellValue.Text("Note")).withCellStyle(c5, ownStyle)
    val sheet =
      styled.copy(cells =
        styled.cells.updated(styled(c5).withHyperlink("https://example.com/notes"))
      )
    val range = CellRange.parse("A1:F8").toOption.get
    val cli = BatchParser
      .applyBatchOperations(
        Workbook(sheet),
        None,
        Vector(BatchOp.Clear("Summary!A1:F8", all = false, styles = false, comments = false))
      )
      .unsafeRunSync()
      .sheets
      .head
    val library = sheet.clearRange(range, ClearWhat.contents)
    assertEquals(cli.cells, library.cells)
    assertEquals(library.cells.keySet, Set(b4, d6, c5))
  }

  // ========== Combined Flags ==========

  test("clear --styles --comments: clears both but keeps contents") {
    val ref = ARef.from0(0, 0)
    val style = CellStyle.default.withFill(Fill.Solid(Color.Rgb(0x00ff00)))
    val sheet = Sheet("Test")
      .put(ref, CellValue.Text("A1"))
      .comment(ref, Comment.plainText("Note", None))
    val styledSheet =
      com.tjclp.xl.sheets.styleSyntax.withRangeStyle(sheet)(CellRange(ref, ref), style)
    val wb = Workbook(styledSheet)

    val result = CellCommands
      .clear(
        wb,
        Some(styledSheet),
        "A1",
        all = false,
        styles = true,
        comments = true,
        outputPath,
        config
      )
      .unsafeRunSync()

    assert(result.contains("styles"))
    assert(result.contains("comments"))
    assert(!result.contains("contents"))

    val imported = ExcelIO.instance[IO].read(outputPath).unsafeRunSync()
    val s = imported.sheets.head
    // Content should remain
    assertEquals(s.cells.get(ref).map(_.value), Some(CellValue.Text("A1")))
    // Style should be cleared
    assertEquals(s.cells.get(ref).flatMap(_.styleId), None)
    // Comment should be cleared
    assert(!s.comments.contains(ref))
  }
