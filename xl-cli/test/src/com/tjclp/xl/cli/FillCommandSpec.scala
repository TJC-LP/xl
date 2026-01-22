package com.tjclp.xl.cli

import munit.FunSuite

import java.nio.file.{Files, Path}

import cats.effect.{IO, unsafe}
import com.tjclp.xl.{Workbook, Sheet}
import com.tjclp.xl.addressing.{ARef, CellRange}
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.cli.commands.WriteCommands
import com.tjclp.xl.io.ExcelIO
import com.tjclp.xl.ooxml.writer.WriterConfig

/**
 * Integration tests for fill command functionality.
 *
 * Note: ARef.from0(col, row) - column comes first!
 */
@SuppressWarnings(
  Array("org.wartremover.warts.OptionPartial", "org.wartremover.warts.IterableOps")
)
class FillCommandSpec extends FunSuite:

  given unsafe.IORuntime = unsafe.IORuntime.global

  val outputPath: Path = Files.createTempFile("test-fill", ".xlsx")
  val config: WriterConfig = WriterConfig.default

  override def afterEach(context: AfterEach): Unit =
    if Files.exists(outputPath) then Files.delete(outputPath)

  // Helper: ARef.from0(col, row) - create ref from A1 notation indices
  // A1 = (0,0), B1 = (1,0), A2 = (0,1), B2 = (1,1)
  private def ref(col: Int, row: Int): ARef = ARef.from0(col, row)

  // ========== Fill Down - Values ==========

  test("fill down: single value copied to column") {
    // A1 = 100
    val sheet = Sheet("Test").put(ref(0, 0), CellValue.Number(100.0))
    val wb = Workbook(sheet)

    val result = WriteCommands
      .fill(wb, Some(sheet), "A1", "A1:A5", FillDirection.Down, outputPath, config)
      .unsafeRunSync()

    assert(result.contains("Filled A1:A5"))
    assert(result.contains("down"))

    val imported = ExcelIO.instance[IO].read(outputPath).unsafeRunSync()
    val s = imported.sheets.head
    // All cells A1:A5 should have value 100 (col=0, rows 0-4)
    for row <- 0 to 4 do
      assertEquals(s.cells.get(ref(0, row)).map(_.value), Some(CellValue.Number(100.0)))
  }

  test("fill down: text value fills entire column") {
    val sheet = Sheet("Test").put(ref(0, 0), CellValue.Text("Header"))
    val wb = Workbook(sheet)

    WriteCommands
      .fill(wb, Some(sheet), "A1", "A1:A3", FillDirection.Down, outputPath, config)
      .unsafeRunSync()

    val imported = ExcelIO.instance[IO].read(outputPath).unsafeRunSync()
    val s = imported.sheets.head
    assertEquals(s.cells.get(ref(0, 0)).map(_.value), Some(CellValue.Text("Header"))) // A1
    assertEquals(s.cells.get(ref(0, 1)).map(_.value), Some(CellValue.Text("Header"))) // A2
    assertEquals(s.cells.get(ref(0, 2)).map(_.value), Some(CellValue.Text("Header"))) // A3
  }

  test("fill down: multiple columns filled together") {
    // A1=A, B1=B, C1=C
    val sheet = Sheet("Test")
      .put(ref(0, 0), CellValue.Text("A")) // A1
      .put(ref(1, 0), CellValue.Text("B")) // B1
      .put(ref(2, 0), CellValue.Text("C")) // C1
    val wb = Workbook(sheet)

    WriteCommands
      .fill(wb, Some(sheet), "A1:C1", "A1:C3", FillDirection.Down, outputPath, config)
      .unsafeRunSync()

    val imported = ExcelIO.instance[IO].read(outputPath).unsafeRunSync()
    val s = imported.sheets.head
    // Row 2: A2=A, B2=B, C2=C
    assertEquals(s.cells.get(ref(0, 1)).map(_.value), Some(CellValue.Text("A"))) // A2
    assertEquals(s.cells.get(ref(1, 1)).map(_.value), Some(CellValue.Text("B"))) // B2
    assertEquals(s.cells.get(ref(2, 1)).map(_.value), Some(CellValue.Text("C"))) // C2
    // Row 3: A3=A, B3=B, C3=C
    assertEquals(s.cells.get(ref(0, 2)).map(_.value), Some(CellValue.Text("A"))) // A3
    assertEquals(s.cells.get(ref(1, 2)).map(_.value), Some(CellValue.Text("B"))) // B3
    assertEquals(s.cells.get(ref(2, 2)).map(_.value), Some(CellValue.Text("C"))) // C3
  }

  // ========== Fill Down - Formulas ==========

  test("fill down: formula shifts references") {
    // A1=10, A2=20, A3=30, B1=A1*2, fill B1 to B1:B3
    val sheet = Sheet("Test")
      .put(ref(0, 0), CellValue.Number(10.0)) // A1
      .put(ref(0, 1), CellValue.Number(20.0)) // A2
      .put(ref(0, 2), CellValue.Number(30.0)) // A3
      .put(ref(1, 0), CellValue.Formula("A1*2", Some(CellValue.Number(20.0)))) // B1
    val wb = Workbook(sheet)

    WriteCommands
      .fill(wb, Some(sheet), "B1", "B1:B3", FillDirection.Down, outputPath, config)
      .unsafeRunSync()

    val imported = ExcelIO.instance[IO].read(outputPath).unsafeRunSync()
    val s = imported.sheets.head
    // B1 should have =A1*2
    val b1 = s.cells.get(ref(1, 0)).map(_.value)
    assert(b1.exists {
      case CellValue.Formula(f, _) => f == "A1*2"
      case _                       => false
    }, s"B1 formula should be =A1*2, got: $b1")
    // B2 should have =A2*2 (shifted down by 1)
    val b2 = s.cells.get(ref(1, 1)).map(_.value)
    assert(b2.exists {
      case CellValue.Formula(f, _) => f == "A2*2"
      case _                       => false
    }, s"B2 formula should be =A2*2, got: $b2")
    // B3 should have =A3*2 (shifted down by 2)
    val b3 = s.cells.get(ref(1, 2)).map(_.value)
    assert(b3.exists {
      case CellValue.Formula(f, _) => f == "A3*2"
      case _                       => false
    }, s"B3 formula should be =A3*2, got: $b3")
  }

  test("fill down: anchored references preserved") {
    // B1=$A$1*ROW(), fill B1 to B1:B3
    val sheet = Sheet("Test")
      .put(ref(0, 0), CellValue.Number(10.0)) // A1
      .put(ref(1, 0), CellValue.Formula("$A$1*ROW()", Some(CellValue.Number(10.0)))) // B1
    val wb = Workbook(sheet)

    WriteCommands
      .fill(wb, Some(sheet), "B1", "B1:B3", FillDirection.Down, outputPath, config)
      .unsafeRunSync()

    val imported = ExcelIO.instance[IO].read(outputPath).unsafeRunSync()
    val s = imported.sheets.head
    // All should have $A$1 (anchored)
    val b2 = s.cells.get(ref(1, 1)).map(_.value)
    assert(b2.exists {
      case CellValue.Formula(f, _) => f.contains("$A$1")
      case _                       => false
    }, s"B2 formula should contain $$A$$1, got: $b2")
  }

  // ========== Fill Right - Values ==========

  test("fill right: single value copied to row") {
    // A1 = 50
    val sheet = Sheet("Test").put(ref(0, 0), CellValue.Number(50.0))
    val wb = Workbook(sheet)

    val result = WriteCommands
      .fill(wb, Some(sheet), "A1", "A1:E1", FillDirection.Right, outputPath, config)
      .unsafeRunSync()

    assert(result.contains("Filled A1:E1"))
    assert(result.contains("right"))

    val imported = ExcelIO.instance[IO].read(outputPath).unsafeRunSync()
    val s = imported.sheets.head
    // All cells A1:E1 should have value 50 (cols 0-4, row=0)
    for col <- 0 to 4 do
      assertEquals(s.cells.get(ref(col, 0)).map(_.value), Some(CellValue.Number(50.0)))
  }

  test("fill right: multiple rows filled together") {
    // A1=1, A2=2, A3=3
    val sheet = Sheet("Test")
      .put(ref(0, 0), CellValue.Number(1.0)) // A1
      .put(ref(0, 1), CellValue.Number(2.0)) // A2
      .put(ref(0, 2), CellValue.Number(3.0)) // A3
    val wb = Workbook(sheet)

    WriteCommands
      .fill(wb, Some(sheet), "A1:A3", "A1:C3", FillDirection.Right, outputPath, config)
      .unsafeRunSync()

    val imported = ExcelIO.instance[IO].read(outputPath).unsafeRunSync()
    val s = imported.sheets.head
    // Column B: B1=1, B2=2, B3=3
    assertEquals(s.cells.get(ref(1, 0)).map(_.value), Some(CellValue.Number(1.0))) // B1
    assertEquals(s.cells.get(ref(1, 1)).map(_.value), Some(CellValue.Number(2.0))) // B2
    assertEquals(s.cells.get(ref(1, 2)).map(_.value), Some(CellValue.Number(3.0))) // B3
    // Column C: C1=1, C2=2, C3=3
    assertEquals(s.cells.get(ref(2, 0)).map(_.value), Some(CellValue.Number(1.0))) // C1
    assertEquals(s.cells.get(ref(2, 1)).map(_.value), Some(CellValue.Number(2.0))) // C2
    assertEquals(s.cells.get(ref(2, 2)).map(_.value), Some(CellValue.Number(3.0))) // C3
  }

  // ========== Fill Right - Formulas ==========

  test("fill right: formula shifts column references") {
    // A1=10, B1=20, C1=30, A2=A1*2, fill A2 to A2:C2
    val sheet = Sheet("Test")
      .put(ref(0, 0), CellValue.Number(10.0)) // A1
      .put(ref(1, 0), CellValue.Number(20.0)) // B1
      .put(ref(2, 0), CellValue.Number(30.0)) // C1
      .put(ref(0, 1), CellValue.Formula("A1*2", Some(CellValue.Number(20.0)))) // A2
    val wb = Workbook(sheet)

    WriteCommands
      .fill(wb, Some(sheet), "A2", "A2:C2", FillDirection.Right, outputPath, config)
      .unsafeRunSync()

    val imported = ExcelIO.instance[IO].read(outputPath).unsafeRunSync()
    val s = imported.sheets.head
    // A2 should have =A1*2
    val a2 = s.cells.get(ref(0, 1)).map(_.value)
    assert(a2.exists {
      case CellValue.Formula(f, _) => f == "A1*2"
      case _                       => false
    }, s"A2 formula should be =A1*2, got: $a2")
    // B2 should have =B1*2 (shifted right by 1)
    val b2 = s.cells.get(ref(1, 1)).map(_.value)
    assert(b2.exists {
      case CellValue.Formula(f, _) => f == "B1*2"
      case _                       => false
    }, s"B2 formula should be =B1*2, got: $b2")
    // C2 should have =C1*2 (shifted right by 2)
    val c2 = s.cells.get(ref(2, 1)).map(_.value)
    assert(c2.exists {
      case CellValue.Formula(f, _) => f == "C1*2"
      case _                       => false
    }, s"C2 formula should be =C1*2, got: $c2")
  }

  // ========== Validation Errors ==========

  test("fill down: rejects mismatched columns") {
    val sheet = Sheet("Test").put(ref(0, 0), CellValue.Number(100.0))
    val wb = Workbook(sheet)

    val result = WriteCommands
      .fill(wb, Some(sheet), "A1", "B1:B5", FillDirection.Down, outputPath, config)
      .attempt
      .unsafeRunSync()

    assert(result.isLeft)
    assert(result.left.exists(_.getMessage.contains("matching columns")))
  }

  test("fill right: rejects mismatched rows") {
    val sheet = Sheet("Test").put(ref(0, 0), CellValue.Number(100.0))
    val wb = Workbook(sheet)

    val result = WriteCommands
      .fill(wb, Some(sheet), "A1", "A2:E2", FillDirection.Right, outputPath, config)
      .attempt
      .unsafeRunSync()

    assert(result.isLeft)
    assert(result.left.exists(_.getMessage.contains("matching rows")))
  }

  test("fill: rejects single cell target") {
    val sheet = Sheet("Test").put(ref(0, 0), CellValue.Number(100.0))
    val wb = Workbook(sheet)

    val result = WriteCommands
      .fill(wb, Some(sheet), "A1", "B1", FillDirection.Down, outputPath, config)
      .attempt
      .unsafeRunSync()

    assert(result.isLeft)
    assert(result.left.exists(_.getMessage.contains("must be a range")))
  }

  // ========== Empty Source Cell ==========

  test("fill down: empty source cell leaves target empty") {
    // A1 has value, but source is B1 (empty)
    val sheet = Sheet("Test").put(ref(0, 0), CellValue.Number(99.0)) // A1
    val wb = Workbook(sheet)

    WriteCommands
      .fill(wb, Some(sheet), "B1", "B1:B3", FillDirection.Down, outputPath, config)
      .unsafeRunSync()

    val imported = ExcelIO.instance[IO].read(outputPath).unsafeRunSync()
    val s = imported.sheets.head
    // B2 and B3 should be empty (source was empty)
    assertEquals(s.cells.get(ref(1, 1)), None) // B2
    assertEquals(s.cells.get(ref(1, 2)), None) // B3
    // A1 should still be there
    assertEquals(s.cells.get(ref(0, 0)).map(_.value), Some(CellValue.Number(99.0)))
  }

  // ========== SUM Formula with Range ==========

  test("fill down: SUM formula with range reference shifts correctly") {
    // B1=SUM(A1:A3), fill B1 to B1:B3
    val sheet = Sheet("Test")
      .put(ref(0, 0), CellValue.Number(1.0)) // A1
      .put(ref(0, 1), CellValue.Number(2.0)) // A2
      .put(ref(0, 2), CellValue.Number(3.0)) // A3
      .put(ref(1, 0), CellValue.Formula("SUM(A1:A3)", Some(CellValue.Number(6.0)))) // B1
    val wb = Workbook(sheet)

    WriteCommands
      .fill(wb, Some(sheet), "B1", "B1:B3", FillDirection.Down, outputPath, config)
      .unsafeRunSync()

    val imported = ExcelIO.instance[IO].read(outputPath).unsafeRunSync()
    val s = imported.sheets.head
    // B2 should have =SUM(A2:A4)
    val b2 = s.cells.get(ref(1, 1)).map(_.value)
    assert(b2.exists {
      case CellValue.Formula(f, _) => f == "SUM(A2:A4)"
      case _                       => false
    }, s"B2 formula should be =SUM(A2:A4), got: $b2")
    // B3 should have =SUM(A3:A5)
    val b3 = s.cells.get(ref(1, 2)).map(_.value)
    assert(b3.exists {
      case CellValue.Formula(f, _) => f == "SUM(A3:A5)"
      case _                       => false
    }, s"B3 formula should be =SUM(A3:A5), got: $b3")
  }
