package com.tjclp.xl.cli

import munit.FunSuite

import java.nio.file.{Files, Path}

import cats.effect.{IO, unsafe}
import com.tjclp.xl.{Workbook, Sheet}
import com.tjclp.xl.addressing.{ARef, CellRange}
import com.tjclp.xl.cells.{CellValue, Comment}
import com.tjclp.xl.cli.commands.CellCommands
import com.tjclp.xl.io.ExcelIO
import com.tjclp.xl.ooxml.writer.WriterConfig
import com.tjclp.xl.styles.CellStyle
import com.tjclp.xl.styles.fill.Fill
import com.tjclp.xl.styles.color.Color

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
      .clear(wb, Some(sheet), "A1:B1", all = false, styles = false, comments = false, outputPath, config)
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
      .clear(wb, Some(sheet), "A1", all = false, styles = false, comments = false, outputPath, config)
      .unsafeRunSync()

    assert(result.contains("A1:A1"))

    val imported = ExcelIO.instance[IO].read(outputPath).unsafeRunSync()
    val s = imported.sheets.head
    assertEquals(s.cells.get(ARef.from0(0, 0)), None)
    assertEquals(s.cells.get(ARef.from0(1, 0)).map(_.value), Some(CellValue.Text("B1")))
  }

  // ========== Clear Styles Only ==========

  test("clear --styles: clears styles but keeps contents and comments") {
    val style = CellStyle.default.withFill(Fill.Solid(Color.Rgb(0xFFFF00)))
    val ref = ARef.from0(0, 0)
    val sheet = Sheet("Test")
      .put(ref, CellValue.Text("A1"))
      .comment(ref, Comment.plainText("Note", None))
    // Apply style via styleRegistry
    val styledSheet = com.tjclp.xl.sheets.styleSyntax.withRangeStyle(sheet)(CellRange(ref, ref), style)
    val wb = Workbook(styledSheet)

    // Verify style is applied
    val styleId = styledSheet.cells.get(ref).flatMap(_.styleId)
    assert(styleId.isDefined, "Cell should have styleId before clearing")

    val result = CellCommands
      .clear(wb, Some(styledSheet), "A1", all = false, styles = true, comments = false, outputPath, config)
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
      .clear(wb, Some(sheet), "A1", all = false, styles = false, comments = true, outputPath, config)
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
      .clear(wb, Some(sheet), "A1:A2", all = false, styles = false, comments = true, outputPath, config)
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
    val style = CellStyle.default.withFill(Fill.Solid(Color.Rgb(0xFF0000)))
    val sheet = Sheet("Test")
      .put(ref, CellValue.Text("A1"))
      .comment(ref, Comment.plainText("Note", None))
    val styledSheet = com.tjclp.xl.sheets.styleSyntax.withRangeStyle(sheet)(CellRange(ref, ref), style)
    val wb = Workbook(styledSheet)

    val result = CellCommands
      .clear(wb, Some(styledSheet), "A1", all = true, styles = false, comments = false, outputPath, config)
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
        case _                    => false
    })

    val result = CellCommands
      .clear(wb, Some(sheet), "A1", all = false, styles = false, comments = false, outputPath, config)
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

  test("clear: default mode unmerges overlapping merged regions") {
    val sheet = Sheet("Test")
      .put(ARef.from0(0, 0), CellValue.Text("Merged"))
      .put(ARef.from0(2, 0), CellValue.Text("C1"))
      .merge(CellRange(ARef.from0(0, 0), ARef.from0(1, 0))) // Merge A1:B1
    val wb = Workbook(sheet)

    // Verify merge exists
    assertEquals(sheet.mergedRanges.size, 1)

    // Clear A1 which overlaps with the merge
    val result = CellCommands
      .clear(wb, Some(sheet), "A1", all = false, styles = false, comments = false, outputPath, config)
      .unsafeRunSync()

    val imported = ExcelIO.instance[IO].read(outputPath).unsafeRunSync()
    val s = imported.sheets.head
    // Merged region should be removed
    assertEquals(s.mergedRanges.size, 0)
    // Cell outside merge should remain
    assertEquals(s.cells.get(ARef.from0(2, 0)).map(_.value), Some(CellValue.Text("C1")))
  }

  // ========== Combined Flags ==========

  test("clear --styles --comments: clears both but keeps contents") {
    val ref = ARef.from0(0, 0)
    val style = CellStyle.default.withFill(Fill.Solid(Color.Rgb(0x00FF00)))
    val sheet = Sheet("Test")
      .put(ref, CellValue.Text("A1"))
      .comment(ref, Comment.plainText("Note", None))
    val styledSheet = com.tjclp.xl.sheets.styleSyntax.withRangeStyle(sheet)(CellRange(ref, ref), style)
    val wb = Workbook(styledSheet)

    val result = CellCommands
      .clear(wb, Some(styledSheet), "A1", all = false, styles = true, comments = true, outputPath, config)
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
