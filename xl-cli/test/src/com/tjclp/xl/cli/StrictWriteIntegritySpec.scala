package com.tjclp.xl.cli

import java.nio.file.{Files, Path}

import cats.effect.{IO, unsafe}
import com.tjclp.xl.{*, given}
import com.tjclp.xl.cli.commands.WriteCommands
import com.tjclp.xl.io.ExcelIO
import com.tjclp.xl.ooxml.writer.WriterConfig

import munit.FunSuite

class StrictWriteIntegritySpec extends FunSuite:
  given unsafe.IORuntime = unsafe.IORuntime.global
  private val config = WriterConfig.default
  private val strict = WritePolicy(strict = true)
  private def num(value: Int): CellValue = CellValue.Number(BigDecimal(value))
  private def plain: Workbook = Workbook(Sheet("Data").put(ref"Z1", num(1)))

  private def withOutput(test: Path => Unit): Unit =
    val path = Files.createTempFile("strict-write-integrity", ".xlsx")
    try test(path)
    finally Files.deleteIfExists(path)

  private def read(path: Path): Workbook = ExcelIO.instance[IO].read(path).unsafeRunSync()

  private def cached(wb: Workbook, sheet: String, ref: ARef): Option[CellValue] =
    wb(SheetName.unsafe(sheet)).fold(error => fail(error.message), identity)(ref).value match
      case CellValue.Formula(_, value, _) => value
      case other => fail(s"Expected formula at $sheet!${ref.toA1}, got $other")

  private def failure(action: IO[String]): String = action.attempt.unsafeRunSync() match
    case Left(error: StrictFailure) => error.summary
    case other => fail(s"Expected strict failure, got $other")

  test(
    "GH-504: strict putf rejects a newly authored self-reference even if it could be pre-cached"
  ) {
    withOutput { out =>
      val wb = plain
      val summary = failure(
        WriteCommands.putFormula(
          wb,
          Some(wb.sheets.head),
          "C1",
          List("=C1+1"),
          out,
          config,
          policy = strict
        )
      )
      assert(summary.contains("Data!C1"), summary)
      assert(summary.contains("Circular"), summary)
      assertEquals(cached(read(out), "Data", ref"C1"), None)
    }
  }

  test("GH-504: default putf reports a failed authored formula and writes it without a cache") {
    withOutput { out =>
      val wb = plain
      val summary = WriteCommands
        .putFormula(
          wb,
          Some(wb.sheets.head),
          "C1",
          List("=C1+1"),
          out,
          config
        )
        .unsafeRunSync()
      assert(summary.contains("error"), summary)
      assertEquals(cached(read(out), "Data", ref"C1"), None)
    }
  }

  // NOT over a bare range is the host-failure fixture (GH-564: AND(range) now folds like Excel)
  test("GH-504: strict putf reports a supported formula whose evaluation fails") {
    withOutput { out =>
      val wb = plain
      val summary = failure(
        WriteCommands.putFormula(
          wb,
          Some(wb.sheets.head),
          "C1",
          List("=NOT(Z1:Z1)"),
          out,
          config,
          policy = strict
        )
      )
      assert(summary.contains("Data!C1"), summary)
      assertEquals(cached(read(out), "Data", ref"C1"), None)
    }
  }

  test("GH-504: putf evaluates a batch against the fully updated workbook") {
    withOutput { out =>
      val wb = plain
      WriteCommands
        .putFormula(
          wb,
          Some(wb.sheets.head),
          "A1:A2",
          List("=A2+1", "=2"),
          out,
          config,
          policy = strict
        )
        .unsafeRunSync()
      val result = read(out)
      assertEquals(cached(result, "Data", ref"A1"), Some(num(3)))
      assertEquals(cached(result, "Data", ref"A2"), Some(num(2)))
    }
  }

  test("GH-504: strict dragging checks every authored cell") {
    withOutput { out =>
      val wb = plain
      val summary = failure(
        WriteCommands.putFormula(
          wb,
          Some(wb.sheets.head),
          "C1:C3",
          List("=C1+1"),
          out,
          config,
          policy = strict
        )
      )
      assert(summary.contains("3 errors"), summary)
      List(ref"C1", ref"C2", ref"C3").foreach(ref =>
        assertEquals(cached(read(out), "Data", ref), None)
      )
    }
  }

  test("GH-504: strict put retains failure diagnostics from affected dependents") {
    withOutput { out =>
      val sheet = Sheet("Data")
        .put(ref"A1", num(1))
        .put(ref"B1", CellValue.Formula("NOT(A1:A1)", Some(num(1))))
        .put(ref"C1", CellValue.Formula("B1*2", Some(num(2))))
      val summary = failure(
        WriteCommands.put(
          Workbook(sheet),
          Some(sheet),
          "A1",
          List("2"),
          out,
          config,
          policy = strict
        )
      )
      assert(summary.contains("Data!B1"), summary)
      assert(summary.contains("Data!C1"), summary)
      val result = read(out)
      assertEquals(cached(result, "Data", ref"B1"), None)
      assertEquals(cached(result, "Data", ref"C1"), None)
    }
  }

  private def copySource: Sheet = Sheet("Data")
    .put(ref"A1", CellValue.Formula("$B$1+1", Some(num(1))))
    .put(ref"B1", num(0))

  test("GH-504: strict fill detects a cycle created in the destination") {
    withOutput { out =>
      val sheet = copySource
      val summary = failure(
        WriteCommands.fill(
          Workbook(sheet),
          Some(sheet),
          "A1",
          "A1:B1",
          FillDirection.Right,
          out,
          config,
          policy = strict
        )
      )
      assert(summary.contains("Data!B1"), summary)
      val result = read(out)
      assertEquals(cached(result, "Data", ref"A1"), None)
      assertEquals(cached(result, "Data", ref"B1"), None)
    }
  }

  test("GH-504: strict copy detects a cycle created in the destination") {
    withOutput { out =>
      val sheet = copySource
      val summary = failure(
        WriteCommands.copyRange(
          Workbook(sheet),
          Some(sheet),
          "A1",
          "B1",
          false,
          out,
          config,
          policy = strict
        )
      )
      assert(summary.contains("Data!B1"), summary)
      assertEquals(cached(read(out), "Data", ref"B1"), None)
    }
  }

  test("GH-504: computed Excel error values do not become strict host failures") {
    withOutput { out =>
      val wb = plain
      WriteCommands
        .putFormula(
          wb,
          Some(wb.sheets.head),
          "C1",
          List("=1/0"),
          out,
          config,
          policy = strict
        )
        .unsafeRunSync()
      assertEquals(cached(read(out), "Data", ref"C1"), Some(CellValue.Error(CellError.Div0)))
    }
  }

  test("GH-504: unaffected caches survive a targeted write") {
    withOutput { out =>
      val sheet = Sheet("Data")
        .put(ref"A1", num(1))
        .put(ref"B1", CellValue.Formula("A1*2", Some(num(2))))
      val wb =
        Workbook(sheet, Sheet("Other").put(ref"A1", CellValue.Formula("TODAY()", Some(num(123)))))
      WriteCommands
        .put(wb, Some(sheet), "A1", List("2"), out, config, policy = strict)
        .unsafeRunSync()
      val result = read(out)
      assertEquals(cached(result, "Data", ref"B1"), Some(num(4)))
      assertEquals(cached(result, "Other", ref"A1"), Some(num(123)))
    }
  }

  test("GH-504: explicit no-recalc still preserves dependent caches") {
    withOutput { out =>
      val sheet = Sheet("Data")
        .put(ref"A1", num(1))
        .put(ref"B1", CellValue.Formula("A1*2", Some(num(2))))
      val summary = WriteCommands
        .put(
          Workbook(sheet),
          Some(sheet),
          "A1",
          List("10"),
          out,
          config,
          policy = WritePolicy(noRecalc = true)
        )
        .unsafeRunSync()
      assert(summary.contains("--no-recalc"), summary)
      assertEquals(cached(read(out), "Data", ref"B1"), Some(num(2)))
    }
  }

  test("GH-504: single-cell writes honor declared iterative calculation") {
    withOutput { out =>
      val sheet = Sheet("Data")
        .put(ref"A1", num(1))
        .put(ref"B1", CellValue.Formula("A1+B1*0.5", Some(num(2))))
      val wb = Workbook(sheet).withCalcPr(
        com.tjclp.xl.workbooks.CalcPr(
          iterativeCalculation = true,
          maxIterations = Some(100),
          maxChange = Some(BigDecimal("0.000001"))
        )
      )
      WriteCommands
        .put(wb, Some(sheet), "A1", List("2"), out, config, policy = strict)
        .unsafeRunSync()
      cached(read(out), "Data", ref"B1") match
        case Some(CellValue.Number(value)) =>
          assert((value - BigDecimal(4)).abs < BigDecimal("0.00001"))
        case other => fail(s"Expected the updated fixpoint, got $other")
    }
  }

  test("GH-504: strict single-cell writes report iterative non-convergence") {
    withOutput { out =>
      val sheet = Sheet("Data")
        .put(ref"A1", num(1))
        .put(ref"B1", CellValue.Formula("A1+B1*0.5", Some(num(2))))
      val wb = Workbook(sheet).withCalcPr(
        com.tjclp.xl.workbooks.CalcPr(
          iterativeCalculation = true,
          maxIterations = Some(1),
          maxChange = Some(BigDecimal("0.000001"))
        )
      )
      val summary =
        failure(WriteCommands.put(wb, Some(sheet), "A1", List("2"), out, config, policy = strict))
      assert(summary.contains("without converging"), summary)
    }
  }

  test("GH-508: strict structural writes keep diagnostics when no cache changed in the cone") {
    List("SUM(Multi)", "Missing!A1").foreach { expression =>
      withOutput { out =>
        val sheet = Sheet("Data").put(ref"A1", num(1))
        val wb = Workbook(sheet, Sheet("Other").put(ref"B2", CellValue.Formula(expression)))
          .withDefinedName("Multi", "Other!$A$1:$A$3,Other!$A$8:$A$10")
        val summary = failure(
          WriteCommands.deleteRows(
            wb,
            Some(sheet),
            20,
            1,
            out,
            config,
            policy = strict
          )
        )
        assert(summary.contains("Other!B2"), summary)
      }
    }
  }
