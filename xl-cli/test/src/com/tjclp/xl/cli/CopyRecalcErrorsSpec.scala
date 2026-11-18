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
