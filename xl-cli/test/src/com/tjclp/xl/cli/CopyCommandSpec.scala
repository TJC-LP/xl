package com.tjclp.xl.cli

import munit.FunSuite

import java.nio.file.{Files, Path}

import cats.effect.{IO, unsafe}
import com.tjclp.xl.{Sheet, Workbook}
import com.tjclp.xl.addressing.ARef
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.cli.commands.WriteCommands
import com.tjclp.xl.io.ExcelIO
import com.tjclp.xl.ooxml.writer.WriterConfig
import com.tjclp.xl.styles.CellStyle
import com.tjclp.xl.styles.dsl.{bold, italic}
import com.tjclp.xl.sheets.styleSyntax.withCellStyle

/**
 * Integration tests for the `copy` command.
 *
 * Covers the correctness issues flagged in review:
 *   - Overlapping copies within a sheet (source cells must be snapshotted)
 *   - Cross-sheet copies (qualified target must actually write to target sheet)
 *   - Values-only mode materializes formulas instead of keeping them
 *   - Styles are preserved in both modes
 *   - Formulas shift relative references by the correct delta
 *
 * Each test gets a fresh temp output file via `outputFixture` — no cross-test coupling.
 */
@SuppressWarnings(
  Array("org.wartremover.warts.OptionPartial", "org.wartremover.warts.IterableOps")
)
class CopyCommandSpec extends FunSuite:

  given unsafe.IORuntime = unsafe.IORuntime.global

  val config: WriterConfig = WriterConfig.default

  /** Per-test temp output path; MUnit handles setup/teardown around each test body. */
  val outputFixture: FunFixture[Path] = FunFixture[Path](
    setup = _ => Files.createTempFile("copy-test-", ".xlsx"),
    teardown = path => if Files.exists(path) then Files.delete(path)
  )

  private def cellValueAt(sheet: Sheet, col: Int, row: Int): Option[CellValue] =
    sheet.cells.get(ARef.from0(col, row)).map(_.value)

  // =========================================================================
  // Overlapping copies within a sheet (P1 correctness)
  // =========================================================================

  outputFixture.test("copy: overlapping A1:A3 -> A2 preserves source snapshot (1,2,3 -> 1,1,2,3)") {
    outputPath =>
      // A1=1, A2=2, A3=3. Copy A1:A3 down by 1. Result should be A1=1, A2=1, A3=2, A4=3
      // (i.e., the original sequence, not 1,1,1,1 from reading already-copied values)
      val sheet = Sheet("Test")
        .put(ARef.from0(0, 0), CellValue.Number(1))
        .put(ARef.from0(0, 1), CellValue.Number(2))
        .put(ARef.from0(0, 2), CellValue.Number(3))
      val wb = Workbook(sheet)

      WriteCommands
        .copyRange(wb, Some(sheet), "A1:A3", "A2", valuesOnly = false, outputPath, config)
        .unsafeRunSync()

      val imported = ExcelIO.instance[IO].read(outputPath).unsafeRunSync()
      val s = imported.sheets.head
      assertEquals(cellValueAt(s, 0, 0), Some(CellValue.Number(1)), "A1 unchanged")
      assertEquals(cellValueAt(s, 0, 1), Some(CellValue.Number(1)), "A2 = source A1")
      assertEquals(cellValueAt(s, 0, 2), Some(CellValue.Number(2)), "A3 = source A2")
      assertEquals(cellValueAt(s, 0, 3), Some(CellValue.Number(3)), "A4 = source A3")
  }

  outputFixture.test("copy: overlapping B1:D1 -> A1 shifts left without corruption") {
    outputPath =>
      // B1=10, C1=20, D1=30. Copy B1:D1 left by 1. Result: A1=10, B1=20, C1=30
      val sheet = Sheet("Test")
        .put(ARef.from0(1, 0), CellValue.Number(10))
        .put(ARef.from0(2, 0), CellValue.Number(20))
        .put(ARef.from0(3, 0), CellValue.Number(30))
      val wb = Workbook(sheet)

      WriteCommands
        .copyRange(wb, Some(sheet), "B1:D1", "A1", valuesOnly = false, outputPath, config)
        .unsafeRunSync()

      val imported = ExcelIO.instance[IO].read(outputPath).unsafeRunSync()
      val s = imported.sheets.head
      assertEquals(cellValueAt(s, 0, 0), Some(CellValue.Number(10)), "A1 = source B1")
      assertEquals(cellValueAt(s, 1, 0), Some(CellValue.Number(20)), "B1 = source C1")
      assertEquals(cellValueAt(s, 2, 0), Some(CellValue.Number(30)), "C1 = source D1")
  }

  // =========================================================================
  // Cross-sheet copy (qualified target — P2 bug)
  // =========================================================================

  outputFixture.test("copy: cross-sheet with qualified target writes to target sheet, not source") {
    outputPath =>
      val s1 = Sheet("Source").put(ARef.from0(0, 0), CellValue.Text("hello"))
      val s2 = Sheet("Dest")
      val wb = Workbook(Vector(s1, s2))

      WriteCommands
        .copyRange(wb, Some(s1), "A1", "Dest!B1", valuesOnly = false, outputPath, config)
        .unsafeRunSync()

      val imported = ExcelIO.instance[IO].read(outputPath).unsafeRunSync()
      val source = imported.sheets.find(_.name.value == "Source").get
      val dest = imported.sheets.find(_.name.value == "Dest").get

      // Destination should have the value at B1
      assertEquals(cellValueAt(dest, 1, 0), Some(CellValue.Text("hello")))
      // Source should be unchanged — particularly, B1 on the source sheet must be empty
      assertEquals(cellValueAt(source, 1, 0), None, "Source B1 must not be written")
      // And the original source A1 is still intact
      assertEquals(cellValueAt(source, 0, 0), Some(CellValue.Text("hello")))
  }

  outputFixture.test("copy: fully-qualified source and destination sheets") { outputPath =>
    val s1 = Sheet("Income").put(ARef.from0(0, 0), CellValue.Number(1000))
    val s2 = Sheet("Summary")
    val wb = Workbook(Vector(s1, s2))

    // Both sides qualified, no --sheet fallback provided
    WriteCommands
      .copyRange(wb, None, "Income!A1", "Summary!C3", valuesOnly = false, outputPath, config)
      .unsafeRunSync()

    val imported = ExcelIO.instance[IO].read(outputPath).unsafeRunSync()
    val summary = imported.sheets.find(_.name.value == "Summary").get
    assertEquals(cellValueAt(summary, 2, 2), Some(CellValue.Number(1000)))
  }

  // =========================================================================
  // Values-only mode (P2 bug: batch version kept formulas)
  // =========================================================================

  outputFixture.test("copy --values-only: formula with cached value materializes to cached") {
    outputPath =>
      val formulaCell = CellValue.Formula("A1+A2", Some(CellValue.Number(30)))
      val sheet = Sheet("Test")
        .put(ARef.from0(0, 0), CellValue.Number(10))
        .put(ARef.from0(0, 1), CellValue.Number(20))
        .put(ARef.from0(0, 2), formulaCell)
      val wb = Workbook(sheet)

      WriteCommands
        .copyRange(wb, Some(sheet), "A3", "B3", valuesOnly = true, outputPath, config)
        .unsafeRunSync()

      val imported = ExcelIO.instance[IO].read(outputPath).unsafeRunSync()
      val s = imported.sheets.head
      // B3 should hold 30 (the cached value), NOT "=A1+A2" (formula)
      cellValueAt(s, 1, 2) match
        case Some(CellValue.Number(n)) =>
          assertEquals(n, BigDecimal(30), "values-only should materialize formula to cached value")
        case other => fail(s"Expected Number(30), got: $other")
  }

  outputFixture.test("copy --values-only: formula with no cache evaluates against source") {
    outputPath =>
      val sheet = Sheet("Test")
        .put(ARef.from0(0, 0), CellValue.Number(5))
        .put(ARef.from0(0, 1), CellValue.Number(7))
        // Formula with no cached value
        .put(ARef.from0(0, 2), CellValue.Formula("A1+A2", None))
      val wb = Workbook(sheet)

      WriteCommands
        .copyRange(wb, Some(sheet), "A3", "B3", valuesOnly = true, outputPath, config)
        .unsafeRunSync()

      val imported = ExcelIO.instance[IO].read(outputPath).unsafeRunSync()
      val s = imported.sheets.head
      cellValueAt(s, 1, 2) match
        case Some(CellValue.Number(n)) =>
          assertEquals(n, BigDecimal(12), "Expected evaluated result 5+7=12")
        case other => fail(s"Expected Number(12), got: $other")
  }

  // =========================================================================
  // Formula shifting (non-values-only mode)
  // =========================================================================

  outputFixture.test("copy: formulas shift relative references by the delta") { outputPath =>
    val sheet = Sheet("Test")
      .put(ARef.from0(0, 0), CellValue.Number(10)) // A1
      .put(ARef.from0(1, 0), CellValue.Number(20)) // B1
      .put(ARef.from0(0, 1), CellValue.Formula("A1+B1", Some(CellValue.Number(30)))) // A2
    val wb = Workbook(sheet)

    // Copy A2 down to A5: formula should shift A1+B1 → A4+B4
    WriteCommands
      .copyRange(wb, Some(sheet), "A2", "A5", valuesOnly = false, outputPath, config)
      .unsafeRunSync()

    val imported = ExcelIO.instance[IO].read(outputPath).unsafeRunSync()
    val s = imported.sheets.head
    cellValueAt(s, 0, 4) match
      case Some(CellValue.Formula(expr, _)) =>
        // Expr should reference A4 and B4 (shifted by +3 rows)
        assert(expr.contains("A4"), s"Expected A4 in shifted formula: $expr")
        assert(expr.contains("B4"), s"Expected B4 in shifted formula: $expr")
      case other => fail(s"Expected Formula, got: $other")
  }

  outputFixture.test("copy: copied sheet-qualified lookup sees copied formula sibling caches") {
    outputPath =>
      val sheet = Sheet("Test")
        .put(ARef.from0(0, 0), CellValue.Formula("1", None)) // A1
        .put(
          ARef.from0(1, 0),
          CellValue.Formula("VLOOKUP(1,Test!A1:C1,3,FALSE)", None)
        ) // B1
        .put(ARef.from0(2, 0), CellValue.Number(42)) // C1
      val wb = Workbook(sheet)

      WriteCommands
        .copyRange(wb, Some(sheet), "A1:C1", "D1:F1", valuesOnly = false, outputPath, config)
        .unsafeRunSync()

      val imported = ExcelIO.instance[IO].read(outputPath).unsafeRunSync()
      val s = imported.sheets.head

      cellValueAt(s, 3, 0) match
        case Some(CellValue.Formula(expr, Some(CellValue.Number(n)))) =>
          assertEquals(expr, "1")
          assertEquals(n, BigDecimal(1))
        case other =>
          fail(s"Expected D1 = Formula(1, Some(Number(1))), got: $other")

      cellValueAt(s, 4, 0) match
        case Some(CellValue.Formula(expr, Some(CellValue.Number(n)))) =>
          assert(expr.contains("D1:F1"), s"Expected shifted range D1:F1, got: $expr")
          assertEquals(
            n,
            BigDecimal(42),
            "E1 should cache after seeing the copied D1 formula cache"
          )
        case other =>
          fail(s"Expected E1 = Formula(VLOOKUP(...,Test!D1:F1,...), Some(Number(42))), got: $other")
  }

  // =========================================================================
  // Style preservation (both modes)
  // =========================================================================

  outputFixture.test("copy: preserves source cell style on target") { outputPath =>
    val boldStyle = CellStyle.default.bold
    val sheet = Sheet("Test")
      .put(ARef.from0(0, 0), CellValue.Text("Bold"))
      .withCellStyle(ARef.from0(0, 0), boldStyle)
    val wb = Workbook(sheet)

    WriteCommands
      .copyRange(wb, Some(sheet), "A1", "C1", valuesOnly = false, outputPath, config)
      .unsafeRunSync()

    val imported = ExcelIO.instance[IO].read(outputPath).unsafeRunSync()
    val s = imported.sheets.head
    val c1Cell = s.cells.get(ARef.from0(2, 0))
    val c1Style = c1Cell.flatMap(_.styleId).flatMap(s.styleRegistry.get)
    assert(c1Style.exists(_.font.bold), s"Expected bold style on C1, got: $c1Style")
  }

  outputFixture.test("copy --values-only: preserves source cell style on target") { outputPath =>
    val italicStyle = CellStyle.default.italic
    val sheet = Sheet("Test")
      .put(ARef.from0(0, 0), CellValue.Text("Italic"))
      .withCellStyle(ARef.from0(0, 0), italicStyle)
    val wb = Workbook(sheet)

    WriteCommands
      .copyRange(wb, Some(sheet), "A1", "C1", valuesOnly = true, outputPath, config)
      .unsafeRunSync()

    val imported = ExcelIO.instance[IO].read(outputPath).unsafeRunSync()
    val s = imported.sheets.head
    val c1Cell = s.cells.get(ARef.from0(2, 0))
    val c1Style = c1Cell.flatMap(_.styleId).flatMap(s.styleRegistry.get)
    assert(c1Style.exists(_.font.italic), s"Expected italic style on C1, got: $c1Style")
  }

  // =========================================================================
  // Dimension handling
  // =========================================================================

  outputFixture.test("copy: single-cell target auto-expands to source dimensions") { outputPath =>
    val sheet = Sheet("Test")
      .put(ARef.from0(0, 0), CellValue.Text("a"))
      .put(ARef.from0(1, 0), CellValue.Text("b"))
      .put(ARef.from0(0, 1), CellValue.Text("c"))
      .put(ARef.from0(1, 1), CellValue.Text("d"))
    val wb = Workbook(sheet)

    // Copy 2x2 source A1:B2 to a single-cell target D1 (expands to D1:E2)
    WriteCommands
      .copyRange(wb, Some(sheet), "A1:B2", "D1", valuesOnly = false, outputPath, config)
      .unsafeRunSync()

    val imported = ExcelIO.instance[IO].read(outputPath).unsafeRunSync()
    val s = imported.sheets.head
    assertEquals(cellValueAt(s, 3, 0), Some(CellValue.Text("a")))
    assertEquals(cellValueAt(s, 4, 0), Some(CellValue.Text("b")))
    assertEquals(cellValueAt(s, 3, 1), Some(CellValue.Text("c")))
    assertEquals(cellValueAt(s, 4, 1), Some(CellValue.Text("d")))
  }

  outputFixture.test("copy: dimension mismatch errors") { outputPath =>
    val sheet = Sheet("Test").put(ARef.from0(0, 0), CellValue.Number(1))
    val wb = Workbook(sheet)

    val err = WriteCommands
      .copyRange(wb, Some(sheet), "A1:A3", "B1:B2", valuesOnly = false, outputPath, config)
      .attempt
      .unsafeRunSync()

    assert(err.isLeft, "Expected dimension mismatch error")
    val msg = err.swap.getOrElse(throw new Exception("expected left")).getMessage
    assert(msg.contains("mismatch"), s"Expected 'mismatch' in message: $msg")
  }
