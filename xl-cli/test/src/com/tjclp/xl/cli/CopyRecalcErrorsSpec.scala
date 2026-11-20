package com.tjclp.xl.cli

import java.nio.file.{Files, Path}

import cats.effect.IO
import munit.CatsEffectSuite

import com.tjclp.xl.{*, given}
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.cli.contract.CliHarness
import com.tjclp.xl.io.ExcelIO
import com.tjclp.xl.macros.ref

/**
 * #891: a `copy` recalculates through the same after-edit pass as every other write, so a pasted
 * formula that fails is named in the run's recalculation summary (it used to be dropped silently by
 * the per-sheet refresh `copy` ran on its own).
 */
class CopyRecalcErrorsSpec extends CatsEffectSuite:

  private def book(): IO[Path] = IO.blocking {
    val p = Files.createTempFile("xl-891-copy-", ".xlsx")
    Files.delete(p)
    p
  }

  private val wb = Workbook(
    Sheet("S")
      .put(ref"A1", CellValue.Number(BigDecimal(1)))
      .put(ref"B1", CellValue.Formula("NOSUCHFN(A1)"))
  )

  test("batch copy: the pasted failing formula is named in the summary") {
    val ops = """[{"op":"copy","source":"B1","target":"B5"}]"""
    for
      src <- book()
      out <- book()
      _ <- ExcelIO.instance[IO].write(wb, src)
      run <- CliHarness.run(List("-f", src.toString, "-o", out.toString, "batch", "-"), ops)
    yield
      assertEquals(run.exit, 0, run.toString)
      assert(run.stdout.contains("B5"), run.stdout)
  }

  test("copy verb: the pasted failing formula is named in the summary") {
    for
      src <- book()
      out <- book()
      _ <- ExcelIO.instance[IO].write(wb, src)
      run <- CliHarness.run("-f", src.toString, "-o", out.toString, "copy", "B1", "B5")
    yield
      assertEquals(run.exit, 0, run.toString)
      assert(run.stdout.contains("B5"), run.stdout)
  }

  test("batch copy on an iterating book keeps the cycle converged (full iterative recalc)") {
    // A1 = B1/2 + C1 and B1 = A1/2 + 1 iterate; a copy into C1 feeds the cycle. A targeted pass
    // would strip the cycle's caches; the after-edit seam recalculates the whole book instead.
    val iterating = Workbook(
      Sheet("S")
        .put(ref"A1", CellValue.Formula("B1/2+C1", Some(CellValue.Number(BigDecimal(0)))))
        .put(ref"B1", CellValue.Formula("A1/2+1", Some(CellValue.Number(BigDecimal(0)))))
        .put(ref"D1", CellValue.Number(BigDecimal(3)))
    ).copy(metadata =
      Workbook().metadata.copy(calcPr =
        Some(
          com.tjclp.xl.workbooks.CalcPr(
            iterativeCalculation = true,
            maxIterations = Some(200),
            maxChange = Some(BigDecimal("0.000001"))
          )
        )
      )
    )
    val ops = """[{"op":"copy","source":"D1","target":"C1"}]"""
    for
      src <- book()
      out <- book()
      _ <- ExcelIO.instance[IO].write(iterating, src)
      run <- CliHarness.run(List("-f", src.toString, "-o", out.toString, "batch", "-"), ops)
      written <- ExcelIO.instance[IO].read(out)
    yield
      assertEquals(run.exit, 0, run.toString)
      val s = written.sheets.head
      // the fixpoint of A = B/2 + 3, B = A/2 + 1 is A = 14/3, B = 10/3
      def cached(r: com.tjclp.xl.addressing.ARef): Option[BigDecimal] =
        s.cells.get(r).map(_.value).collect {
          case CellValue.Formula(_, Some(CellValue.Number(n)), _) => n
        }
      val a = cached(ref"A1")
      val b = cached(ref"B1")
      assert(a.exists(v => (v - BigDecimal(14) / 3).abs < BigDecimal("0.0001")), s"A1 $a")
      assert(b.exists(v => (v - BigDecimal(10) / 3).abs < BigDecimal("0.0001")), s"B1 $b")
  }
