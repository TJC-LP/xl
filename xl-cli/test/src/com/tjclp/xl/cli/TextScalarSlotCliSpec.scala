package com.tjclp.xl.cli

import java.nio.file.{Files, Path}

import cats.effect.{IO, Resource}
import munit.CatsEffectSuite

import com.tjclp.xl.{*, given}
import com.tjclp.xl.cells.{CellError, CellValue}
import com.tjclp.xl.cli.contract.CliHarness
import com.tjclp.xl.io.ExcelIO
import com.tjclp.xl.macros.ref

/**
 * #709, end to end: a sheet whose formulas read a text placeholder in a typed scalar slot
 * (`=ABS(C5)` over "n/a") is ordinary to Excel — the cell shows `#VALUE!`. `recalc --strict` used
 * to report each such cell as a formula evaluation error (RECALC_ERRORS, exit 1) and leave it
 * uncached; it now caches the error value and exits clean.
 */
class TextScalarSlotCliSpec extends CatsEffectSuite:

  /** A temp file deleted however the test ends, a failed assertion or run included. */
  private def tempXlsx(prefix: String): Resource[IO, Path] =
    Resource.make(IO.blocking(Files.createTempFile(prefix, ".xlsx")))(p =>
      IO.blocking(Files.deleteIfExists(p)).void
    )

  private val source: Resource[IO, Path] =
    val sheet = Sheet("Sheet1")
      .put(ref"C5", CellValue.Text("n/a"))
      .put(ref"D5", CellValue.Formula("ABS(C5)", None))
      .put(ref"E5", CellValue.Formula("IF(C5,1,0)", None))
      .put(ref"F5", CellValue.Formula("IFERROR(ABS(C5),\"caught\")", None))
    tempXlsx("xl-709-src-").evalTap(p => ExcelIO.instance[IO].write(Workbook(sheet), p))

  private def cached(path: Path, at: ARef): IO[Option[CellValue]] =
    ExcelIO.instance[IO].read(path).map { wb =>
      wb.sheets.headOption.getOrElse(fail("no sheet"))(at).value match
        case CellValue.Formula(_, c, _) => c
        case other => fail(s"expected a formula at ${at.toA1}, got $other")
    }

  test("#709: recalc --strict over =ABS(C5) on text exits 0 and caches #VALUE!") {
    source.flatMap(src => tempXlsx("xl-709-out-").map(out => (src, out))).use { (src, out) =>
      for
        run <- CliHarness.run("-f", src.toString, "-o", out.toString, "--strict", "recalc")
        d5 <- cached(out, ref"D5")
        e5 <- cached(out, ref"E5")
        f5 <- cached(out, ref"F5")
      yield
        assertEquals(run.exit, 0, run.toString)
        assert(!run.toString.contains("RECALC_ERRORS"), run.toString)
        assert(!run.stdout.contains("STRICT FAILURE"), run.stdout)
        assertEquals(d5, Some(CellValue.Error(CellError.Value)))
        assertEquals(e5, Some(CellValue.Error(CellError.Value)))
        assertEquals(f5, Some(CellValue.Text("caught")))
    }
  }
