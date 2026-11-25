package com.tjclp.xl.cli

import java.nio.file.{Files, Path}

import cats.effect.{IO, Ref, unsafe}
import munit.FunSuite

import com.tjclp.xl.{Workbook, Sheet}
import com.tjclp.xl.addressing.ARef
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.cli.commands.WriteCommands
import com.tjclp.xl.cli.contract.{Warning, WarningCode}
import com.tjclp.xl.io.ExcelIO
import com.tjclp.xl.ooxml.writer.WriterConfig

/**
 * #962: `recalc` keeps a formula calling an add-in function (`_xll.`) at its saved value, the SUM
 * that reads it recalculates from that value, and one `ADDIN_CACHE_KEPT` warning names the cell.
 */
@SuppressWarnings(Array("org.wartremover.warts.IterableOps"))
class AddinCacheCliSpec extends FunSuite:

  given unsafe.IORuntime = unsafe.IORuntime.global

  val outputPath: Path = Files.createTempFile("test-addin", ".xlsx")

  override def afterEach(context: AfterEach): Unit =
    if Files.exists(outputPath) then Files.delete(outputPath)

  private def num(i: Int): CellValue = CellValue.Number(BigDecimal(i))
  private val a1 = ARef.from0(0, 0)
  private val a2 = ARef.from0(0, 1)
  private val a3 = ARef.from0(0, 2)

  private def recalcCollecting(wb: Workbook, strict: Boolean = false): (String, Vector[Warning]) =
    (for
      seen <- Ref.of[IO, Vector[Warning]](Vector.empty)
      summary <- WriteCommands.recalc(
        wb,
        outputPath,
        WriterConfig.default,
        policy = WritePolicy.default.copy(strict = strict),
        warn = w => seen.update(_ :+ w)
      )
      warnings <- seen.get
    yield (summary, warnings)).unsafeRunSync()

  private def cachedAt(ref: ARef): Option[CellValue] =
    ExcelIO
      .instance[IO]
      .read(outputPath)
      .unsafeRunSync()
      .sheets
      .head
      .cells
      .get(ref)
      .map(_.value)
      .collect { case CellValue.Formula(_, cached, _) => cached }
      .flatten

  test("recalc: the SUM reads the add-in cell's saved value; one warning names the cell") {
    val wb = Workbook(
      Sheet("Data")
        .put(a1, CellValue.Formula("""=_xll.SomeAddin.FETCH("x",1)""", Some(num(40))))
        .put(a2, num(2))
        .put(a3, CellValue.Formula("=SUM(A1:A2)", Some(num(0))))
    )
    val (summary, warnings) = recalcCollecting(wb)
    assert(!summary.contains("error"), summary)
    assertEquals(cachedAt(a1), Some(num(40)))
    assertEquals(cachedAt(a3), Some(num(42)))
    assertEquals(warnings.map(_.code), Vector(WarningCode.ADDIN_CACHE_KEPT))
    val warning = warnings.head
    assert(warning.message.contains("Data!A1"), warning.message)
    assertEquals(warning.location.flatMap(_.ref), Some("A1"))
  }

  test("recalc --strict: the add-in warning is informational and never gates") {
    val wb = Workbook(
      Sheet("Data")
        .put(a1, CellValue.Formula("""=_xll.SomeAddin.FETCH("x",1)""", Some(num(40))))
        .put(a3, CellValue.Formula("=A1+2", None))
    )
    val (_, warnings) = recalcCollecting(wb, strict = true)
    assertEquals(warnings.map(_.code), Vector(WarningCode.ADDIN_CACHE_KEPT))
    assertEquals(cachedAt(a3), Some(num(42)))
  }

  test("recalc: a book with no add-in calls raises no add-in warning") {
    val wb = Workbook(Sheet("Data").put(a1, num(1)).put(a3, CellValue.Formula("=A1+2", None)))
    val (_, warnings) = recalcCollecting(wb)
    assertEquals(warnings, Vector.empty)
  }

  test("addinCacheWarning: names the first ten cells and counts the rest") {
    val refs = (0 until 12).toVector.map(i =>
      com.tjclp.xl.formula.DependencyGraph.QualifiedRef(
        com.tjclp.xl.addressing.SheetName.unsafe("S"),
        ARef.from0(0, i)
      )
    )
    val message = WriteCommands.addinCacheWarning(refs).map(_.message).getOrElse("")
    assert(message.startsWith("12 formulas calling an add-in function"), message)
    assert(message.contains("S!A10 and 2 more"), message)
    assert(!message.contains("S!A11"), message)
  }
