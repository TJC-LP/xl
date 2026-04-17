package com.tjclp.xl.cli

import munit.FunSuite

import cats.effect.unsafe
import com.tjclp.xl.{Sheet, Workbook}
import com.tjclp.xl.addressing.ARef
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.cli.helpers.BatchParser
import com.tjclp.xl.sheets.FreezePane

/**
 * Batch JSON tests for `copy`, `freeze`, and `unfreeze` operations.
 *
 * Exercises the end-to-end BatchParser flow: JSON → BatchOp → applyBatchOperations → Workbook.
 * Covers the correctness issues flagged in review:
 *   - Qualified source/target sheet refs in copy
 *   - valuesOnly copy materializes formulas (doesn't just write `CellValue.Formula` again)
 *   - Overlapping copies within a sheet use snapshotted sources
 */
@SuppressWarnings(
  Array("org.wartremover.warts.OptionPartial", "org.wartremover.warts.IterableOps")
)
class BatchCopyFreezeSpec extends FunSuite:

  given unsafe.IORuntime = unsafe.IORuntime.global

  private def cellValueAt(sheet: Sheet, col: Int, row: Int): Option[CellValue] =
    sheet.cells.get(ARef.from0(col, row)).map(_.value)

  private def runBatch(
    wb: Workbook,
    defaultSheet: Option[Sheet],
    json: String
  ): Workbook =
    val parseResult = BatchParser.parseBatchJson(json) match
      case Right(r) => r
      case Left(e) => throw e
    BatchParser.applyBatchOperations(wb, defaultSheet, parseResult.ops).unsafeRunSync()

  // =========================================================================
  // batch copy — qualified source/target sheet refs
  // =========================================================================

  test("batch copy: cross-sheet with qualified refs writes to destination sheet") {
    val s1 = Sheet("Source").put(ARef.from0(0, 0), CellValue.Text("hello"))
    val s2 = Sheet("Dest")
    val wb = Workbook(Vector(s1, s2))

    val json = """[{"op":"copy","source":"Source!A1","target":"Dest!B1"}]"""
    // No default sheet provided — must use qualified refs
    val result = runBatch(wb, None, json)

    val dest = result.sheets.find(_.name.value == "Dest").get
    val source = result.sheets.find(_.name.value == "Source").get
    assertEquals(cellValueAt(dest, 1, 0), Some(CellValue.Text("hello")))
    assertEquals(cellValueAt(source, 1, 0), None, "Source sheet must not be written to")
  }

  test("batch copy: shifted formula cache uses destination-sheet context") {
    val source = Sheet("Source")
      .put(ARef.from0(0, 0), CellValue.Number(10)) // A1
      .put(ARef.from0(1, 0), CellValue.Formula("A1", Some(CellValue.Number(10)))) // B1
      .put(ARef.from0(2, 0), CellValue.Number(5)) // C1
    val dest = Sheet("Dest")
      .put(ARef.from0(2, 0), CellValue.Number(99)) // C1
    val wb = Workbook(Vector(source, dest))

    val json = """[{"op":"copy","source":"Source!B1","target":"Dest!D1"}]"""
    val result = runBatch(wb, None, json)
    val updatedDest = result.sheets.find(_.name.value == "Dest").get

    cellValueAt(updatedDest, 3, 0) match
      case Some(CellValue.Formula(expr, Some(CellValue.Number(n)))) =>
        assert(expr.contains("C1"), s"Expected shifted formula to reference C1, got: $expr")
        assertEquals(n, BigDecimal(99), "Cache must be evaluated against Dest, not Source")
      case other =>
        fail(s"Expected Formula(C1, Some(Number(99))), got: $other")
  }

  test("batch copy: unqualified refs fall back to default sheet") {
    val s1 = Sheet("Main")
      .put(ARef.from0(0, 0), CellValue.Number(10))
    val wb = Workbook(s1)

    val json = """[{"op":"copy","source":"A1","target":"B1"}]"""
    val result = runBatch(wb, Some(s1), json)

    val s = result.sheets.head
    assertEquals(cellValueAt(s, 1, 0), Some(CellValue.Number(10)))
  }

  test("batch copy: unqualified side requires default sheet (error otherwise)") {
    val s1 = Sheet("A").put(ARef.from0(0, 0), CellValue.Number(1))
    val s2 = Sheet("B")
    val wb = Workbook(Vector(s1, s2))

    val json = """[{"op":"copy","source":"A!A1","target":"B1"}]"""
    val err = BatchParser
      .applyBatchOperations(wb, None, BatchParser.parseBatchJson(json).toOption.get.ops)
      .attempt
      .unsafeRunSync()

    assert(err.isLeft)
    val msg = err.swap.getOrElse(throw new Exception("expected left")).getMessage
    assert(msg.contains("target"), s"Expected target-side error: $msg")
  }

  // =========================================================================
  // batch copy --values-only materializes formulas (P2 bug)
  // =========================================================================

  test("batch copy valuesOnly: formula with cached value becomes the cached value") {
    val formulaCell = CellValue.Formula("A1+A2", Some(CellValue.Number(42)))
    val sheet = Sheet("Test")
      .put(ARef.from0(0, 0), CellValue.Number(20))
      .put(ARef.from0(0, 1), CellValue.Number(22))
      .put(ARef.from0(0, 2), formulaCell)
    val wb = Workbook(sheet)

    val json = """[{"op":"copy","source":"A3","target":"B3","valuesOnly":true}]"""
    val result = runBatch(wb, Some(sheet), json)

    val s = result.sheets.head
    cellValueAt(s, 1, 2) match
      case Some(CellValue.Number(n)) =>
        assertEquals(n, BigDecimal(42), "batch valuesOnly must materialize, not copy formula")
      case other => fail(s"Expected Number(42), got: $other")
  }

  test("batch copy (no valuesOnly): formula shifts relative refs") {
    val sheet = Sheet("Test")
      .put(ARef.from0(0, 0), CellValue.Number(5))
      .put(ARef.from0(0, 1), CellValue.Formula("A1*2", Some(CellValue.Number(10))))
    val wb = Workbook(sheet)

    val json = """[{"op":"copy","source":"A2","target":"B5"}]"""
    val result = runBatch(wb, Some(sheet), json)

    val s = result.sheets.head
    cellValueAt(s, 1, 4) match
      case Some(CellValue.Formula(expr, _)) =>
        // Shifted from A1*2 by (+1 col, +3 rows) → B4*2
        assert(expr.contains("B4"), s"Expected B4 in shifted formula: $expr")
      case other => fail(s"Expected Formula, got: $other")
  }

  // =========================================================================
  // batch copy — overlapping (P1 correctness)
  // =========================================================================

  test("batch copy: overlapping A1:A3 -> A2 preserves source snapshot") {
    val sheet = Sheet("Test")
      .put(ARef.from0(0, 0), CellValue.Number(1))
      .put(ARef.from0(0, 1), CellValue.Number(2))
      .put(ARef.from0(0, 2), CellValue.Number(3))
    val wb = Workbook(sheet)

    val json = """[{"op":"copy","source":"A1:A3","target":"A2"}]"""
    val result = runBatch(wb, Some(sheet), json)

    val s = result.sheets.head
    assertEquals(cellValueAt(s, 0, 0), Some(CellValue.Number(1)), "A1 unchanged")
    assertEquals(cellValueAt(s, 0, 1), Some(CellValue.Number(1)), "A2 = source A1")
    assertEquals(cellValueAt(s, 0, 2), Some(CellValue.Number(2)), "A3 = source A2")
    assertEquals(cellValueAt(s, 0, 3), Some(CellValue.Number(3)), "A4 = source A3")
  }

  test("batch copy: formula cache sees copied dependencies in the final target range") {
    val sheet = Sheet("Test")
      .put(ARef.from0(0, 0), CellValue.Formula("B1", Some(CellValue.Number(2)))) // A1
      .put(ARef.from0(1, 0), CellValue.Number(2)) // B1
    val wb = Workbook(sheet)

    val json = """[{"op":"copy","source":"A1:B1","target":"B1"}]"""
    val result = runBatch(wb, Some(sheet), json)
    val s = result.sheets.head

    cellValueAt(s, 1, 0) match
      case Some(CellValue.Formula(expr, Some(CellValue.Number(n)))) =>
        assert(expr.contains("C1"), s"Expected shifted formula to reference C1, got: $expr")
        assertEquals(n, BigDecimal(2), "Cache must see copied B1 -> C1 value in final target state")
      case other =>
        fail(s"Expected Formula(C1, Some(Number(2))), got: $other")
  }

  // =========================================================================
  // batch freeze / unfreeze
  // =========================================================================

  test("batch freeze: sets At override on default sheet") {
    val sheet = Sheet("Main")
    val wb = Workbook(sheet)

    val json = """[{"op":"freeze","ref":"C3"}]"""
    val result = runBatch(wb, Some(sheet), json)

    val s = result.sheets.head
    assertEquals(s.freezePane, Some(FreezePane.At(ARef.from0(2, 2))))
  }

  test("batch unfreeze: sets Remove override on default sheet") {
    val sheet = Sheet("Main").freezeAt(ARef.from0(1, 1))
    val wb = Workbook(sheet)

    val json = """[{"op":"unfreeze"}]"""
    val result = runBatch(wb, Some(sheet), json)

    val s = result.sheets.head
    assertEquals(s.freezePane, Some(FreezePane.Remove))
  }

  test("batch freeze: errors without --sheet if no default") {
    val wb = Workbook(Sheet("A"), Sheet("B"))
    val json = """[{"op":"freeze","ref":"B2"}]"""
    val err = BatchParser
      .applyBatchOperations(wb, None, BatchParser.parseBatchJson(json).toOption.get.ops)
      .attempt
      .unsafeRunSync()

    assert(err.isLeft, "Expected error when freeze has no sheet to target")
  }

  test("batch copy: invalid ref produces user-facing error (not bypass JVM exception)") {
    val sheet = Sheet("Test")
    val wb = Workbook(sheet)

    val json = """[{"op":"copy","source":"not-a-ref","target":"A1"}]"""
    val err = BatchParser
      .applyBatchOperations(
        wb,
        Some(sheet),
        BatchParser.parseBatchJson(json).toOption.get.ops
      )
      .attempt
      .unsafeRunSync()

    assert(err.isLeft, "Expected IO error, not JVM exception bypass")
    val msg = err.swap.getOrElse(throw new Exception("expected left")).getMessage
    assert(
      msg.contains("Invalid source ref") || msg.contains("not-a-ref"),
      s"Expected source-side parse error: $msg"
    )
  }
