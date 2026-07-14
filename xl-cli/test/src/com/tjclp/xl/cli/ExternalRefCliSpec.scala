package com.tjclp.xl.cli

import java.nio.file.{Files, Path}

import cats.effect.{IO, unsafe}
import munit.FunSuite

import com.tjclp.xl.{Workbook, Sheet}
import com.tjclp.xl.addressing.ARef
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.cli.commands.{ReadCommands, WriteCommands}
import com.tjclp.xl.io.ExcelIO
import com.tjclp.xl.ooxml.writer.WriterConfig

/**
 * GH-353: external-workbook references ([2]Book1!A1) at the CLI surface.
 *
 * Previously any formula whose precedent closure touched such a cell died with `Parse error:
 * UnexpectedChar([`, and putf rejected the formula with a caret-only message. Now: eval computes
 * from the external cell's Excel-written cache when present, returns a clear "External workbook
 * reference" error when not, and putf accepts the formula (stored uncached — Excel computes it on
 * open).
 */
@SuppressWarnings(
  Array("org.wartremover.warts.OptionPartial", "org.wartremover.warts.IterableOps")
)
class ExternalRefCliSpec extends FunSuite:

  given unsafe.IORuntime = unsafe.IORuntime.global

  val outputPath: Path = Files.createTempFile("test-external", ".xlsx")
  val config: WriterConfig = WriterConfig.default

  override def afterEach(context: AfterEach): Unit =
    if Files.exists(outputPath) then Files.delete(outputPath)

  private val d1 = ARef.from0(3, 0) // D1
  private val d2 = ARef.from0(3, 1) // D2

  test("eval: formula over a cached external precedent computes from the cache (issue repro)") {
    val sheet = Sheet("Data")
      .put(d1, CellValue.Formula("=SUM([2]Book1!A1:A2)", Some(CellValue.Number(BigDecimal(42)))))
      .put(d2, CellValue.Number(BigDecimal(8)))
    val wb = Workbook(sheet)

    val result = ReadCommands.eval(wb, Some(wb.sheets.head), "=SUM(D1:D2)", Nil).unsafeRunSync()
    assert(result.contains("50"), s"Expected 50 (42 cached + 8), got: $result")
  }

  test("eval: direct external formula returns a clear error, not UnexpectedChar") {
    val wb = Workbook(Sheet("Data").put(d2, CellValue.Number(BigDecimal(8))))

    val result = ReadCommands
      .eval(wb, Some(wb.sheets.head), "=SUM([2]Book1!A1:A2)", Nil)
      .attempt
      .unsafeRunSync()

    result match
      case Left(e) =>
        assert(
          e.getMessage.contains("External workbook reference"),
          s"Expected friendly external-ref error, got: ${e.getMessage}"
        )
        assert(!e.getMessage.contains("UnexpectedChar"), e.getMessage)
      case Right(v) => fail(s"Expected error, got: $v")
  }

  test("putf: accepts an external-workbook formula (stored uncached)") {
    val wb = Workbook(Sheet("Data"))
    val result = WriteCommands
      .putFormula(wb, Some(wb.sheets.head), "B1", List("=SUM([2]Book1!A1:A2)"), outputPath, config)
      .unsafeRunSync()

    assert(result.contains("Put: B1"), s"Expected success message, got: $result")

    val imported = ExcelIO.instance[IO].read(outputPath).unsafeRunSync()
    imported.sheets.head.cells.get(ARef.from0(1, 0)).map(_.value) match
      case Some(CellValue.Formula(formula, cachedValue)) =>
        assertEquals(formula, "SUM([2]Book1!A1:A2)")
        assertEquals(cachedValue, None, "unresolvable external formula must stay uncached")
      case other => fail(s"Expected Formula, got $other")
  }

  test("putf: preserves the Excel-written cache of untouched external cells on recalc") {
    // A2 depends on A1; D1 is an external-formula cell with an Excel-written cache. Writing A1
    // triggers recalculateDependents — the external cell's cache must survive verbatim.
    val sheet = Sheet("Data")
      .put(ARef.from0(0, 0), CellValue.Number(BigDecimal(1)))
      .put(ARef.from0(0, 1), CellValue.Formula("=A1*2", None))
      .put(d1, CellValue.Formula("=SUM([2]Book1!A1:A2)", Some(CellValue.Number(BigDecimal(42)))))
    val wb = Workbook(sheet)

    WriteCommands
      .put(wb, Some(wb.sheets.head), "A1", List("5"), outputPath, config)
      .unsafeRunSync()

    val imported = ExcelIO.instance[IO].read(outputPath).unsafeRunSync()
    imported.sheets.head.cells.get(d1).map(_.value) match
      case Some(CellValue.Formula(_, cachedValue)) =>
        assertEquals(cachedValue, Some(CellValue.Number(BigDecimal(42))))
      case other => fail(s"Expected cached external Formula, got $other")
  }

  test("put: a MIXED local+external formula keeps its Excel cache when its precedent changes") {
    // GH-353 review: D1 has a real in-workbook edge (A1), so writing A1 puts D1 in the recalc
    // set — this exercises the pinnedExternalCache guard in recalculateInOrder, which without
    // the pin would hit the eval-failure arm and CLEAR the Excel-written cache (the sole source
    // of truth). The pure-external test above never enters the recalc set at all.
    val sheet = Sheet("Data")
      .put(ARef.from0(0, 0), CellValue.Number(BigDecimal(1)))
      .put(d1, CellValue.Formula("=A1+[2]Book1!B1", Some(CellValue.Number(BigDecimal(42)))))
    val wb = Workbook(sheet)

    WriteCommands
      .put(wb, Some(wb.sheets.head), "A1", List("5"), outputPath, config)
      .unsafeRunSync()

    val imported = ExcelIO.instance[IO].read(outputPath).unsafeRunSync()
    imported.sheets.head.cells.get(d1).map(_.value) match
      case Some(CellValue.Formula(_, cachedValue)) =>
        assertEquals(
          cachedValue,
          Some(CellValue.Number(BigDecimal(42))),
          "pin guard must preserve the Excel-written cache verbatim"
        )
      case other => fail(s"Expected cached external Formula, got $other")
  }

  test("putf: accepts an external SUMIF (range-typed argument slot)") {
    val wb = Workbook(Sheet("Data"))
    val result = WriteCommands
      .putFormula(
        wb,
        Some(wb.sheets.head),
        "B1",
        List("""=SUMIF([2]Book1!A1:A9, ">0")"""),
        outputPath,
        config
      )
      .unsafeRunSync()

    assert(result.contains("Put: B1"), s"Expected success message, got: $result")
  }
