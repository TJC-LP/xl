package com.tjclp.xl.cli

import java.nio.file.{Files, Path}

import cats.effect.{IO, Resource}
import munit.CatsEffectSuite

import com.tjclp.xl.{*, given}
import com.tjclp.xl.cells.{CellError, CellValue, FormulaKind}
import com.tjclp.xl.cli.contract.CliHarness
import com.tjclp.xl.io.ExcelIO
import com.tjclp.xl.macros.ref

/**
 * #684, end to end: `xl recalc` and `xl put` keep the cached values of T() formulas. The source is
 * the probe from the #684 comment with the caches LibreOffice saved for it; T used to be an unknown
 * function, so both verbs wrote every T() cell with no cached value.
 */
class TFunctionRecalcCliSpec extends CatsEffectSuite:

  /** A temp file deleted however the test ends, a failed assertion or run included. */
  private def tempXlsx(prefix: String): Resource[IO, Path] =
    Resource.make(IO.blocking(Files.createTempFile(prefix, ".xlsx")))(p =>
      IO.blocking(Files.deleteIfExists(p)).void
    )

  private def text(s: String): CellValue = CellValue.Text(s)

  private def formula(expr: String, cached: CellValue): CellValue =
    CellValue.Formula(expr, Some(cached), FormulaKind.Normal())

  /** Each T() cell with the value LibreOffice caches for it. */
  private val expected: List[(ARef, String, CellValue)] = List(
    (ref"D1", "T(A1)", text("OldLabelAlpha")),
    (ref"D2", "T(B1)", text("")),
    (ref"D3", "T(E9)", text("")),
    (ref"D4", "T(BlankName)", text("")),
    (ref"D5", "T(1/0)", CellValue.Error(CellError.Div0)),
    (ref"D6", "T(\"lit\")", text("lit"))
  )

  private val source: Resource[IO, Path] =
    val sheet = expected.foldLeft(
      Sheet("Sheet1").put(ref"A1", text("OldLabelAlpha")).put(ref"B1", CellValue.Number(42))
    ) { case (s, (at, expr, cached)) => s.put(at, formula(expr, cached)) }
    val wb = Workbook(sheet).withDefinedName("BlankName", "Sheet1!$E$10")
    tempXlsx("xl-684-src-").evalTap(p => ExcelIO.instance[IO].write(wb, p))

  /** The cached value of the formula at `at`; a non-formula cell fails the IO. */
  private def cached(path: Path, at: ARef): IO[Option[CellValue]] =
    ExcelIO.instance[IO].read(path).flatMap { wb =>
      wb.sheets.headOption match
        case None => IO.raiseError(new AssertionError(s"no sheet in $path"))
        case Some(sheet) =>
          sheet(at).value match
            case CellValue.Formula(_, c, _) => IO.pure(c)
            case other =>
              IO.raiseError(new AssertionError(s"expected a formula at ${at.toA1}, got $other"))
    }

  private val files: Resource[IO, (Path, Path)] =
    source.flatMap(src => tempXlsx("xl-684-out-").map(out => (src, out)))

  test("#684: recalc --strict keeps every T() cached value") {
    files.use { (src, out) =>
      for
        run <- CliHarness.run("-f", src.toString, "-o", out.toString, "--strict", "recalc")
        _ <- assertIO(IO.pure(run.exit), 0, run.toString)
        _ <- assertIOBoolean(IO.pure(!run.stdout.contains("RECALC_ERRORS")), run.stdout)
        _ <- assertIOBoolean(IO.pure(!run.stderr.contains("RECALC_ERRORS")), run.stderr)
        _ <- expected.foldLeft(IO.unit) { case (acc, (at, _, value)) =>
          acc >> assertIO(cached(out, at), Some(value), at.toA1)
        }
      yield ()
    }
  }

  test("#684: put on a T() precedent recomputes the dependent's cached value") {
    files.use { (src, out) =>
      for
        run <- CliHarness.run(
          "-f",
          src.toString,
          "-s",
          "Sheet1",
          "-o",
          out.toString,
          "put",
          "A1",
          "NewLabelBeta"
        )
        _ <- assertIO(IO.pure(run.exit), 0, run.toString)
        _ <- assertIO(cached(out, ref"D1"), Some(text("NewLabelBeta")))
        _ <- assertIO(cached(out, ref"D6"), Some(text("lit")))
      yield ()
    }
  }
