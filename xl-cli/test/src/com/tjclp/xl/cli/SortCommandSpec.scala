package com.tjclp.xl.cli

import munit.FunSuite

import java.nio.file.{Files, Path}

import cats.effect.{IO, unsafe}
import com.tjclp.xl.{Workbook, Sheet}
import com.tjclp.xl.addressing.ARef
import com.tjclp.xl.cells.{Cell, CellValue}
import com.tjclp.xl.cli.commands.WriteCommands
import com.tjclp.xl.io.ExcelIO
import com.tjclp.xl.ooxml.writer.WriterConfig

/**
 * Integration tests for sort command functionality.
 *
 * Note: ARef.from0(col, row) - column comes first!
 */
@SuppressWarnings(
  Array("org.wartremover.warts.OptionPartial", "org.wartremover.warts.IterableOps")
)
class SortCommandSpec extends FunSuite:

  given unsafe.IORuntime = unsafe.IORuntime.global

  val outputPath: Path = Files.createTempFile("test-sort", ".xlsx")
  val config: WriterConfig = WriterConfig.default

  override def afterEach(context: AfterEach): Unit =
    if Files.exists(outputPath) then Files.delete(outputPath)

  // Helper: ARef.from0(col, row) - create ref from A1 notation indices
  // A1 = (0,0), B1 = (1,0), A2 = (0,1), B2 = (1,1)
  private def ref(col: Int, row: Int): ARef = ARef.from0(col, row)

  // ========== Basic Sorting ==========

  test("sort: ascending by single column (text)") {
    // A1=C, A2=A, A3=B
    val sheet = Sheet("Test")
      .put(ref(0, 0), CellValue.Text("C"))
      .put(ref(0, 1), CellValue.Text("A"))
      .put(ref(0, 2), CellValue.Text("B"))
    val wb = Workbook(sheet)

    val key = SortKey("A", SortDirection.Ascending, SortMode.Alphanumeric)
    val result = WriteCommands
      .sort(wb, Some(sheet), "A1:A3", List(key), false, outputPath, config)
      .unsafeRunSync()

    assert(result.contains("Sorted A1:A3"))
    assert(result.contains("A asc"))

    val imported = ExcelIO.instance[IO].read(outputPath).unsafeRunSync()
    val s = imported.sheets.head
    assertEquals(s.cells.get(ref(0, 0)).map(_.value), Some(CellValue.Text("A")))
    assertEquals(s.cells.get(ref(0, 1)).map(_.value), Some(CellValue.Text("B")))
    assertEquals(s.cells.get(ref(0, 2)).map(_.value), Some(CellValue.Text("C")))
  }

  test("sort: descending by single column (text)") {
    // A1=A, A2=C, A3=B
    val sheet = Sheet("Test")
      .put(ref(0, 0), CellValue.Text("A"))
      .put(ref(0, 1), CellValue.Text("C"))
      .put(ref(0, 2), CellValue.Text("B"))
    val wb = Workbook(sheet)

    val key = SortKey("A", SortDirection.Descending, SortMode.Alphanumeric)
    WriteCommands
      .sort(wb, Some(sheet), "A1:A3", List(key), false, outputPath, config)
      .unsafeRunSync()

    val imported = ExcelIO.instance[IO].read(outputPath).unsafeRunSync()
    val s = imported.sheets.head
    assertEquals(s.cells.get(ref(0, 0)).map(_.value), Some(CellValue.Text("C")))
    assertEquals(s.cells.get(ref(0, 1)).map(_.value), Some(CellValue.Text("B")))
    assertEquals(s.cells.get(ref(0, 2)).map(_.value), Some(CellValue.Text("A")))
  }

  // ========== Numeric Sorting ==========

  test("sort: numeric mode sorts numbers correctly") {
    // Without numeric mode: "10" < "2" (string comparison)
    // With numeric mode: 2 < 10
    val sheet = Sheet("Test")
      .put(ref(0, 0), CellValue.Number(10))
      .put(ref(0, 1), CellValue.Number(2))
      .put(ref(0, 2), CellValue.Number(100))
    val wb = Workbook(sheet)

    val key = SortKey("A", SortDirection.Ascending, SortMode.Numeric)
    WriteCommands
      .sort(wb, Some(sheet), "A1:A3", List(key), false, outputPath, config)
      .unsafeRunSync()

    val imported = ExcelIO.instance[IO].read(outputPath).unsafeRunSync()
    val s = imported.sheets.head
    assertEquals(s.cells.get(ref(0, 0)).map(_.value), Some(CellValue.Number(2)))
    assertEquals(s.cells.get(ref(0, 1)).map(_.value), Some(CellValue.Number(10)))
    assertEquals(s.cells.get(ref(0, 2)).map(_.value), Some(CellValue.Number(100)))
  }

  test("sort: numeric descending") {
    val sheet = Sheet("Test")
      .put(ref(0, 0), CellValue.Number(5))
      .put(ref(0, 1), CellValue.Number(15))
      .put(ref(0, 2), CellValue.Number(10))
    val wb = Workbook(sheet)

    val key = SortKey("A", SortDirection.Descending, SortMode.Numeric)
    WriteCommands
      .sort(wb, Some(sheet), "A1:A3", List(key), false, outputPath, config)
      .unsafeRunSync()

    val imported = ExcelIO.instance[IO].read(outputPath).unsafeRunSync()
    val s = imported.sheets.head
    assertEquals(s.cells.get(ref(0, 0)).map(_.value), Some(CellValue.Number(15)))
    assertEquals(s.cells.get(ref(0, 1)).map(_.value), Some(CellValue.Number(10)))
    assertEquals(s.cells.get(ref(0, 2)).map(_.value), Some(CellValue.Number(5)))
  }

  // ========== Header Row ==========

  test("sort: header row is preserved") {
    // A1=Name (header), A2=Zoe, A3=Amy
    val sheet = Sheet("Test")
      .put(ref(0, 0), CellValue.Text("Name")) // Header
      .put(ref(0, 1), CellValue.Text("Zoe"))
      .put(ref(0, 2), CellValue.Text("Amy"))
    val wb = Workbook(sheet)

    val key = SortKey("A", SortDirection.Ascending, SortMode.Alphanumeric)
    val result = WriteCommands
      .sort(wb, Some(sheet), "A1:A3", List(key), true, outputPath, config)
      .unsafeRunSync()

    assert(result.contains("header row preserved"))

    val imported = ExcelIO.instance[IO].read(outputPath).unsafeRunSync()
    val s = imported.sheets.head
    assertEquals(s.cells.get(ref(0, 0)).map(_.value), Some(CellValue.Text("Name"))) // Header unchanged
    assertEquals(s.cells.get(ref(0, 1)).map(_.value), Some(CellValue.Text("Amy")))
    assertEquals(s.cells.get(ref(0, 2)).map(_.value), Some(CellValue.Text("Zoe")))
  }

  // ========== Multi-Column Sort ==========

  test("sort: multi-column sort with tie-breaker") {
    // Sort by A, then by B
    // Row 0: A, 2
    // Row 1: B, 1
    // Row 2: A, 1
    val sheet = Sheet("Test")
      .put(ref(0, 0), CellValue.Text("A"))
      .put(ref(1, 0), CellValue.Number(2))
      .put(ref(0, 1), CellValue.Text("B"))
      .put(ref(1, 1), CellValue.Number(1))
      .put(ref(0, 2), CellValue.Text("A"))
      .put(ref(1, 2), CellValue.Number(1))
    val wb = Workbook(sheet)

    val keys = List(
      SortKey("A", SortDirection.Ascending, SortMode.Alphanumeric),
      SortKey("B", SortDirection.Ascending, SortMode.Numeric)
    )
    WriteCommands
      .sort(wb, Some(sheet), "A1:B3", keys, false, outputPath, config)
      .unsafeRunSync()

    val imported = ExcelIO.instance[IO].read(outputPath).unsafeRunSync()
    val s = imported.sheets.head
    // Row 1: A, 1 (A with smaller B)
    assertEquals(s.cells.get(ref(0, 0)).map(_.value), Some(CellValue.Text("A")))
    assertEquals(s.cells.get(ref(1, 0)).map(_.value), Some(CellValue.Number(1)))
    // Row 2: A, 2
    assertEquals(s.cells.get(ref(0, 1)).map(_.value), Some(CellValue.Text("A")))
    assertEquals(s.cells.get(ref(1, 1)).map(_.value), Some(CellValue.Number(2)))
    // Row 3: B, 1
    assertEquals(s.cells.get(ref(0, 2)).map(_.value), Some(CellValue.Text("B")))
    assertEquals(s.cells.get(ref(1, 2)).map(_.value), Some(CellValue.Number(1)))
  }

  // ========== Multiple Columns Move Together ==========

  test("sort: all columns in row move together") {
    // Row 0: 3, X, Alpha
    // Row 1: 1, Y, Beta
    // Row 2: 2, Z, Gamma
    // Sort by column A
    val sheet = Sheet("Test")
      .put(ref(0, 0), CellValue.Number(3))
      .put(ref(1, 0), CellValue.Text("X"))
      .put(ref(2, 0), CellValue.Text("Alpha"))
      .put(ref(0, 1), CellValue.Number(1))
      .put(ref(1, 1), CellValue.Text("Y"))
      .put(ref(2, 1), CellValue.Text("Beta"))
      .put(ref(0, 2), CellValue.Number(2))
      .put(ref(1, 2), CellValue.Text("Z"))
      .put(ref(2, 2), CellValue.Text("Gamma"))
    val wb = Workbook(sheet)

    val key = SortKey("A", SortDirection.Ascending, SortMode.Numeric)
    WriteCommands
      .sort(wb, Some(sheet), "A1:C3", List(key), false, outputPath, config)
      .unsafeRunSync()

    val imported = ExcelIO.instance[IO].read(outputPath).unsafeRunSync()
    val s = imported.sheets.head
    // Row 1: 1, Y, Beta
    assertEquals(s.cells.get(ref(0, 0)).map(_.value), Some(CellValue.Number(1)))
    assertEquals(s.cells.get(ref(1, 0)).map(_.value), Some(CellValue.Text("Y")))
    assertEquals(s.cells.get(ref(2, 0)).map(_.value), Some(CellValue.Text("Beta")))
    // Row 2: 2, Z, Gamma
    assertEquals(s.cells.get(ref(0, 1)).map(_.value), Some(CellValue.Number(2)))
    assertEquals(s.cells.get(ref(1, 1)).map(_.value), Some(CellValue.Text("Z")))
    assertEquals(s.cells.get(ref(2, 1)).map(_.value), Some(CellValue.Text("Gamma")))
    // Row 3: 3, X, Alpha
    assertEquals(s.cells.get(ref(0, 2)).map(_.value), Some(CellValue.Number(3)))
    assertEquals(s.cells.get(ref(1, 2)).map(_.value), Some(CellValue.Text("X")))
    assertEquals(s.cells.get(ref(2, 2)).map(_.value), Some(CellValue.Text("Alpha")))
  }

  // ========== Edge Cases ==========

  test("sort: empty cells sort last") {
    val sheet = Sheet("Test")
      .put(ref(0, 0), CellValue.Text("B"))
      // A2 is empty
      .put(ref(0, 2), CellValue.Text("A"))
    val wb = Workbook(sheet)

    val key = SortKey("A", SortDirection.Ascending, SortMode.Alphanumeric)
    WriteCommands
      .sort(wb, Some(sheet), "A1:A3", List(key), false, outputPath, config)
      .unsafeRunSync()

    val imported = ExcelIO.instance[IO].read(outputPath).unsafeRunSync()
    val s = imported.sheets.head
    assertEquals(s.cells.get(ref(0, 0)).map(_.value), Some(CellValue.Text("A")))
    assertEquals(s.cells.get(ref(0, 1)).map(_.value), Some(CellValue.Text("B")))
    assertEquals(s.cells.get(ref(0, 2)), None) // Empty at end
  }

  test("sort: case-insensitive sorting") {
    val sheet = Sheet("Test")
      .put(ref(0, 0), CellValue.Text("banana"))
      .put(ref(0, 1), CellValue.Text("Apple"))
      .put(ref(0, 2), CellValue.Text("CHERRY"))
    val wb = Workbook(sheet)

    val key = SortKey("A", SortDirection.Ascending, SortMode.Alphanumeric)
    WriteCommands
      .sort(wb, Some(sheet), "A1:A3", List(key), false, outputPath, config)
      .unsafeRunSync()

    val imported = ExcelIO.instance[IO].read(outputPath).unsafeRunSync()
    val s = imported.sheets.head
    // apple < banana < cherry (case-insensitive)
    assertEquals(s.cells.get(ref(0, 0)).map(_.value), Some(CellValue.Text("Apple")))
    assertEquals(s.cells.get(ref(0, 1)).map(_.value), Some(CellValue.Text("banana")))
    assertEquals(s.cells.get(ref(0, 2)).map(_.value), Some(CellValue.Text("CHERRY")))
  }

  test("sort: boolean values sort as 0/1") {
    val sheet = Sheet("Test")
      .put(ref(0, 0), CellValue.Bool(true)) // 1
      .put(ref(0, 1), CellValue.Bool(false)) // 0
      .put(ref(0, 2), CellValue.Bool(true)) // 1
    val wb = Workbook(sheet)

    val key = SortKey("A", SortDirection.Ascending, SortMode.Numeric)
    WriteCommands
      .sort(wb, Some(sheet), "A1:A3", List(key), false, outputPath, config)
      .unsafeRunSync()

    val imported = ExcelIO.instance[IO].read(outputPath).unsafeRunSync()
    val s = imported.sheets.head
    // false (0) < true (1) < true (1)
    assertEquals(s.cells.get(ref(0, 0)).map(_.value), Some(CellValue.Bool(false)))
    assertEquals(s.cells.get(ref(0, 1)).map(_.value), Some(CellValue.Bool(true)))
    assertEquals(s.cells.get(ref(0, 2)).map(_.value), Some(CellValue.Bool(true)))
  }

  // ========== Validation ==========

  test("sort: rejects single cell") {
    val sheet = Sheet("Test").put(ref(0, 0), CellValue.Text("A"))
    val wb = Workbook(sheet)

    val key = SortKey("A", SortDirection.Ascending, SortMode.Alphanumeric)
    val result = WriteCommands
      .sort(wb, Some(sheet), "A1", List(key), false, outputPath, config)
      .attempt
      .unsafeRunSync()

    assert(result.isLeft)
    assert(result.left.exists(_.getMessage.contains("requires a range")))
  }

  test("sort: rejects column outside range") {
    val sheet = Sheet("Test")
      .put(ref(0, 0), CellValue.Text("A"))
      .put(ref(0, 1), CellValue.Text("B"))
    val wb = Workbook(sheet)

    val key = SortKey("C", SortDirection.Ascending, SortMode.Alphanumeric) // C not in A1:A2
    val result = WriteCommands
      .sort(wb, Some(sheet), "A1:A2", List(key), false, outputPath, config)
      .attempt
      .unsafeRunSync()

    assert(result.isLeft)
    assert(result.left.exists(_.getMessage.contains("outside range")))
  }

  test("sort: rejects invalid column letter") {
    val sheet = Sheet("Test")
      .put(ref(0, 0), CellValue.Text("A"))
      .put(ref(0, 1), CellValue.Text("B"))
    val wb = Workbook(sheet)

    val key = SortKey("1", SortDirection.Ascending, SortMode.Alphanumeric) // Invalid
    val result = WriteCommands
      .sort(wb, Some(sheet), "A1:A2", List(key), false, outputPath, config)
      .attempt
      .unsafeRunSync()

    assert(result.isLeft)
    assert(result.left.exists(_.getMessage.contains("Invalid column")))
  }

  // ========== Formulas ==========

  test("sort: formulas use cached value for sorting") {
    // A1=10, A2=5, A3=15
    // B1 has formula =A1*2 with cached value 20
    // B2 has formula =A2*2 with cached value 10
    // B3 has formula =A3*2 with cached value 30
    val sheet = Sheet("Test")
      .put(ref(0, 0), CellValue.Number(10))
      .put(ref(1, 0), CellValue.Formula("A1*2", Some(CellValue.Number(20))))
      .put(ref(0, 1), CellValue.Number(5))
      .put(ref(1, 1), CellValue.Formula("A2*2", Some(CellValue.Number(10))))
      .put(ref(0, 2), CellValue.Number(15))
      .put(ref(1, 2), CellValue.Formula("A3*2", Some(CellValue.Number(30))))
    val wb = Workbook(sheet)

    val key = SortKey("B", SortDirection.Ascending, SortMode.Numeric)
    WriteCommands
      .sort(wb, Some(sheet), "A1:B3", List(key), false, outputPath, config)
      .unsafeRunSync()

    val imported = ExcelIO.instance[IO].read(outputPath).unsafeRunSync()
    val s = imported.sheets.head
    // Row 1: B cached value 10 (smallest), so A=5, B=A2*2
    assertEquals(s.cells.get(ref(0, 0)).map(_.value), Some(CellValue.Number(5)))
    // Row 2: B cached value 20, so A=10
    assertEquals(s.cells.get(ref(0, 1)).map(_.value), Some(CellValue.Number(10)))
    // Row 3: B cached value 30 (largest), so A=15
    assertEquals(s.cells.get(ref(0, 2)).map(_.value), Some(CellValue.Number(15)))
  }

  test("sort: formula references are adjusted when rows move (GH-165)") {
    // Regression test for GitHub issue #165
    // When rows move during sort, relative formula references should be adjusted
    // so formulas continue referencing cells in their own row.
    //
    // Setup:
    //   A1=Name, B1=Q1, C1=Q2, D1=Total (header)
    //   A2=Charlie, B2=100, C2=200, D2=SUM(B2:C2) -> 300
    //   A3=Bob, B3=300, C3=400, D3=SUM(B3:C3) -> 700
    //   A4=Alice, B4=500, C4=600, D4=SUM(B4:C4) -> 1100
    //
    // After sorting by A (ascending with header):
    //   Row 2: Alice, 500, 600, =SUM(B2:C2) -> should be 1100
    //   Row 3: Bob, 300, 400, =SUM(B3:C3) -> should be 700
    //   Row 4: Charlie, 100, 200, =SUM(B4:C4) -> should be 300
    val sheet = Sheet("Test")
      // Header
      .put(ref(0, 0), CellValue.Text("Name"))
      .put(ref(1, 0), CellValue.Text("Q1"))
      .put(ref(2, 0), CellValue.Text("Q2"))
      .put(ref(3, 0), CellValue.Text("Total"))
      // Charlie (row 2)
      .put(ref(0, 1), CellValue.Text("Charlie"))
      .put(ref(1, 1), CellValue.Number(100))
      .put(ref(2, 1), CellValue.Number(200))
      .put(ref(3, 1), CellValue.Formula("SUM(B2:C2)", Some(CellValue.Number(300))))
      // Bob (row 3)
      .put(ref(0, 2), CellValue.Text("Bob"))
      .put(ref(1, 2), CellValue.Number(300))
      .put(ref(2, 2), CellValue.Number(400))
      .put(ref(3, 2), CellValue.Formula("SUM(B3:C3)", Some(CellValue.Number(700))))
      // Alice (row 4)
      .put(ref(0, 3), CellValue.Text("Alice"))
      .put(ref(1, 3), CellValue.Number(500))
      .put(ref(2, 3), CellValue.Number(600))
      .put(ref(3, 3), CellValue.Formula("SUM(B4:C4)", Some(CellValue.Number(1100))))
    val wb = Workbook(sheet)

    val key = SortKey("A", SortDirection.Ascending, SortMode.Alphanumeric)
    WriteCommands
      .sort(wb, Some(sheet), "A1:D4", List(key), true, outputPath, config)
      .unsafeRunSync()

    val imported = ExcelIO.instance[IO].read(outputPath).unsafeRunSync()
    val s = imported.sheets.head

    // Header should be unchanged
    assertEquals(s.cells.get(ref(0, 0)).map(_.value), Some(CellValue.Text("Name")))

    // Row 2 should be Alice with adjusted formula
    assertEquals(s.cells.get(ref(0, 1)).map(_.value), Some(CellValue.Text("Alice")))
    assertEquals(s.cells.get(ref(1, 1)).map(_.value), Some(CellValue.Number(500)))
    assertEquals(s.cells.get(ref(2, 1)).map(_.value), Some(CellValue.Number(600)))
    // Formula should be =SUM(B2:C2) after moving from row 4 to row 2
    s.cells.get(ref(3, 1)).map(_.value) match
      case Some(CellValue.Formula(formula, _)) =>
        assertEquals(formula, "SUM(B2:C2)")
      case other =>
        fail(s"Expected Formula, got $other")

    // Row 3 should be Bob (unchanged position, unchanged formula)
    assertEquals(s.cells.get(ref(0, 2)).map(_.value), Some(CellValue.Text("Bob")))
    s.cells.get(ref(3, 2)).map(_.value) match
      case Some(CellValue.Formula(formula, _)) =>
        assertEquals(formula, "SUM(B3:C3)")
      case other =>
        fail(s"Expected Formula, got $other")

    // Row 4 should be Charlie with adjusted formula
    assertEquals(s.cells.get(ref(0, 3)).map(_.value), Some(CellValue.Text("Charlie")))
    assertEquals(s.cells.get(ref(1, 3)).map(_.value), Some(CellValue.Number(100)))
    assertEquals(s.cells.get(ref(2, 3)).map(_.value), Some(CellValue.Number(200)))
    // Formula should be =SUM(B4:C4) after moving from row 2 to row 4
    s.cells.get(ref(3, 3)).map(_.value) match
      case Some(CellValue.Formula(formula, _)) =>
        assertEquals(formula, "SUM(B4:C4)")
      case other =>
        fail(s"Expected Formula, got $other")
  }

  test("sort: absolute formula references ($) are preserved during sort") {
    // Formulas with absolute row references should NOT shift
    // Setup:
    //   A1=Rate (header)
    //   A2=10, B2=$A$1*2 (absolute ref to A1)
    //   A3=5, B3=$A$1*3 (absolute ref to A1)
    //
    // After sorting by A ascending:
    //   A2=5, B2=$A$1*3 (formula moved but $A$1 stays as $A$1)
    //   A3=10, B3=$A$1*2 (formula moved but $A$1 stays as $A$1)
    val sheet = Sheet("Test")
      .put(ref(0, 0), CellValue.Number(10))
      .put(ref(1, 0), CellValue.Formula("$A$1*2", Some(CellValue.Number(20))))
      .put(ref(0, 1), CellValue.Number(5))
      .put(ref(1, 1), CellValue.Formula("$A$1*3", Some(CellValue.Number(15))))
    val wb = Workbook(sheet)

    val key = SortKey("A", SortDirection.Ascending, SortMode.Numeric)
    WriteCommands
      .sort(wb, Some(sheet), "A1:B2", List(key), false, outputPath, config)
      .unsafeRunSync()

    val imported = ExcelIO.instance[IO].read(outputPath).unsafeRunSync()
    val s = imported.sheets.head

    // Row 1: A=5, B=$A$1*3 (moved from row 2)
    assertEquals(s.cells.get(ref(0, 0)).map(_.value), Some(CellValue.Number(5)))
    s.cells.get(ref(1, 0)).map(_.value) match
      case Some(CellValue.Formula(formula, _)) =>
        assertEquals(formula, "$A$1*3") // Absolute ref unchanged
      case other =>
        fail(s"Expected Formula, got $other")

    // Row 2: A=10, B=$A$1*2 (moved from row 1)
    assertEquals(s.cells.get(ref(0, 1)).map(_.value), Some(CellValue.Number(10)))
    s.cells.get(ref(1, 1)).map(_.value) match
      case Some(CellValue.Formula(formula, _)) =>
        assertEquals(formula, "$A$1*2") // Absolute ref unchanged
      case other =>
        fail(s"Expected Formula, got $other")
  }

  // ========== Data Preservation (Regression Tests) ==========

  test("sort: cells outside range are preserved") {
    // Column C is outside the sort range A:B
    // Row 0: 3, X, Alpha
    // Row 1: 1, Y, Beta
    // Row 2: 2, Z, Gamma
    val sheet = Sheet("Test")
      .put(ref(0, 0), CellValue.Number(3))
      .put(ref(1, 0), CellValue.Text("X"))
      .put(ref(2, 0), CellValue.Text("Alpha")) // C1 - outside sort range
      .put(ref(0, 1), CellValue.Number(1))
      .put(ref(1, 1), CellValue.Text("Y"))
      .put(ref(2, 1), CellValue.Text("Beta")) // C2 - outside sort range
      .put(ref(0, 2), CellValue.Number(2))
      .put(ref(1, 2), CellValue.Text("Z"))
      .put(ref(2, 2), CellValue.Text("Gamma")) // C3 - outside sort range
    val wb = Workbook(sheet)

    // Sort only A:B range, column C should NOT move
    val key = SortKey("A", SortDirection.Ascending, SortMode.Numeric)
    WriteCommands
      .sort(wb, Some(sheet), "A1:B3", List(key), false, outputPath, config)
      .unsafeRunSync()

    val imported = ExcelIO.instance[IO].read(outputPath).unsafeRunSync()
    val s = imported.sheets.head

    // Verify A:B columns are sorted (1, 2, 3 order)
    assertEquals(s.cells.get(ref(0, 0)).map(_.value), Some(CellValue.Number(1)))
    assertEquals(s.cells.get(ref(1, 0)).map(_.value), Some(CellValue.Text("Y")))
    assertEquals(s.cells.get(ref(0, 1)).map(_.value), Some(CellValue.Number(2)))
    assertEquals(s.cells.get(ref(1, 1)).map(_.value), Some(CellValue.Text("Z")))
    assertEquals(s.cells.get(ref(0, 2)).map(_.value), Some(CellValue.Number(3)))
    assertEquals(s.cells.get(ref(1, 2)).map(_.value), Some(CellValue.Text("X")))

    // Verify column C is UNCHANGED (did not move with A:B)
    assertEquals(s.cells.get(ref(2, 0)).map(_.value), Some(CellValue.Text("Alpha")))
    assertEquals(s.cells.get(ref(2, 1)).map(_.value), Some(CellValue.Text("Beta")))
    assertEquals(s.cells.get(ref(2, 2)).map(_.value), Some(CellValue.Text("Gamma")))
  }

  test("sort: cells above and below range are preserved") {
    // Row 0 is above range, Row 4 is below range
    val sheet = Sheet("Test")
      .put(ref(0, 0), CellValue.Text("Header")) // Above range
      .put(ref(0, 1), CellValue.Number(3))
      .put(ref(0, 2), CellValue.Number(1))
      .put(ref(0, 3), CellValue.Number(2))
      .put(ref(0, 4), CellValue.Text("Footer")) // Below range
    val wb = Workbook(sheet)

    // Sort only A2:A4
    val key = SortKey("A", SortDirection.Ascending, SortMode.Numeric)
    WriteCommands
      .sort(wb, Some(sheet), "A2:A4", List(key), false, outputPath, config)
      .unsafeRunSync()

    val imported = ExcelIO.instance[IO].read(outputPath).unsafeRunSync()
    val s = imported.sheets.head

    // Row 0 (A1) unchanged
    assertEquals(s.cells.get(ref(0, 0)).map(_.value), Some(CellValue.Text("Header")))
    // Rows 1-3 (A2:A4) sorted
    assertEquals(s.cells.get(ref(0, 1)).map(_.value), Some(CellValue.Number(1)))
    assertEquals(s.cells.get(ref(0, 2)).map(_.value), Some(CellValue.Number(2)))
    assertEquals(s.cells.get(ref(0, 3)).map(_.value), Some(CellValue.Number(3)))
    // Row 4 (A5) unchanged
    assertEquals(s.cells.get(ref(0, 4)).map(_.value), Some(CellValue.Text("Footer")))
  }
