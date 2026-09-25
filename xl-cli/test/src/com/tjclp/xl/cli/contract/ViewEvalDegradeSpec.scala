package com.tjclp.xl.cli.contract

import java.nio.file.{Files, Path}

import cats.effect.{IO, Resource}
import cats.syntax.all.*
import munit.CatsEffectSuite

import com.tjclp.xl.{*, given}
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.cli.raster.BatikRasterizer
import com.tjclp.xl.io.ExcelIO
import com.tjclp.xl.macros.ref

/**
 * `view --eval` degrades cell by cell (Weaver 9/17): a cell xl cannot evaluate no longer fails the
 * whole render — every other formula shows its live value, the failing cell and its blocked
 * dependents show exactly what the file holds, and ONE `EVAL_FAILED` warning names them. Under
 * `--strict` a failure is the `RECALC_GATE` gate for the text renders and html/svg; the raster
 * formats never gate.
 */
class ViewEvalDegradeSpec extends CatsEffectSuite:

  private def data(sheet: Sheet): Sheet =
    List[(ARef, Int)](
      ref"B2" -> 5,
      ref"C2" -> -3,
      ref"D2" -> 1,
      ref"B3" -> -2,
      ref"C3" -> 4,
      ref"D3" -> -6,
      ref"B4" -> 7,
      ref"C4" -> 1,
      ref"D4" -> 2
    ).foldLeft(sheet.put(ref"A1", "Label").put(ref"B1", "x").put(ref"C1", "y").put(ref"D1", "z")) {
      case (s, (at, n)) => s.put(at, n)
    }

  /** The Weaver repro: F2 is the formula the 0.23.1 CLI crashed on, uncached. */
  private def weaverBook(): Workbook =
    Workbook(
      Vector(
        data(Sheet("Sheet1"))
          .put(ref"F2", CellValue.Formula("SUMPRODUCT(--(B2:B4>0),ABS(C2:C4+D2:D4))", None))
          .put(ref"F3", CellValue.Formula("B2*2", None))
          .put(ref"F4", CellValue.Formula("SUM(B2:D4)", None))
      )
    )

  /**
   * F2 cannot evaluate (a missing sheet, uncached); F3 carries a stale cache the live value
   * replaces; G2 reads F2 and carries a cache that must be shown as the file's value, never
   * recomputed from anything stale.
   */
  private def degradedBook(): Workbook =
    Workbook(
      Vector(
        data(Sheet("Sheet1"))
          .put(ref"F2", CellValue.Formula("Missing!A1+1", None))
          .put(ref"F3", CellValue.Formula("B2*2", Some(CellValue.Number(BigDecimal(99)))))
          .put(ref"F4", CellValue.Formula("SUM(B2:D4)", None))
          .put(ref"G2", CellValue.Formula("F2+1", Some(CellValue.Number(BigDecimal(7)))))
      )
    )

  private val books: Map[String, () => Workbook] =
    Map("weaver.xlsx" -> (() => weaverBook()), "degraded.xlsx" -> (() => degradedBook()))

  private val fixtures = ResourceSuiteLocalFixture(
    "view-eval-degrade",
    Resource.make(
      IO.blocking(Files.createTempDirectory("xl-view-eval-")).flatTap { dir =>
        books.toList
          .traverse_((name, build) => ExcelIO.instance[IO].write(build(), dir.resolve(name)))
      }
    )(TestFixtures.delete)
  )

  override def munitFixtures = List(fixtures)

  private def file(name: String): String = fixtures().resolve(name).toString

  private def lines(text: String): List[String] = text.linesIterator.toList

  test("the Weaver repro renders instead of exiting 3 INTERNAL") {
    for
      run <- CliHarness.run("-f", file("weaver.xlsx"), "view", "A1:F4", "--eval")
      json <- CliHarness.run("-f", file("weaver.xlsx"), "--json", "view", "A1:F4", "--eval")
    yield
      assertEquals(run.exit, 0, run.stderr)
      assert(!run.stderr.contains("INTERNAL"), run.stderr)
      assert(!run.stderr.contains("EVAL_FAILED"), "nothing failed, so nothing is reported")
      val row3 = lines(run.stdout).find(_.startsWith("| 3 ")).getOrElse(fail(run.stdout))
      assert(row3.contains("| 10 "), row3)
      val row4 = lines(run.stdout).find(_.startsWith("| 4 ")).getOrElse(fail(run.stdout))
      assert(row4.contains("| 9 "), row4)
      // interim until array lifting: ABS collapses to a scalar, so SUMPRODUCT's sizes differ
      val row2 = lines(run.stdout).find(_.startsWith("| 2 ")).getOrElse(fail(run.stdout))
      assert(row2.contains("#VALUE!"), row2)
      assertEquals(json.exit, 0, json.stderr)
      val e = ujson.read(json.stdout)
      assertEquals(e("ok"), ujson.True)
      assert(
        e("warnings").arr.forall(_("code").str != "EVAL_FAILED"),
        e("warnings").toString
      )
  }

  test("a failing cell keeps the file's value, the rest is live, and ONE warning names it") {
    for run <- CliHarness.run("-f", file("degraded.xlsx"), "view", "A1:G4", "--eval")
    yield
      assertEquals(run.exit, 0, run.stderr)
      val row2 = lines(run.stdout).find(_.startsWith("| 2 ")).getOrElse(fail(run.stdout))
      assert(row2.contains("=Missing!A1+1"), s"uncached F2 renders as plain view does: $row2")
      assert(row2.contains("| 7 "), s"blocked G2 shows the file's cached value: $row2")
      val row3 = lines(run.stdout).find(_.startsWith("| 3 ")).getOrElse(fail(run.stdout))
      assert(row3.contains("| 10 "), s"F3 is live, not the stale 99: $row3")
      assert(!row3.contains("99"), row3)
      val row4 = lines(run.stdout).find(_.startsWith("| 4 ")).getOrElse(fail(run.stdout))
      assert(row4.contains("| 9 "), row4)
      val warnings = lines(run.stderr).filter(_.startsWith("Warning[EVAL_FAILED]"))
      assertEquals(warnings.size, 1, run.stderr)
      val warning = warnings.headOption.getOrElse("")
      assert(
        warning.startsWith("Warning[EVAL_FAILED]: Formula evaluation failed: Sheet1!F2: "),
        warning
      )
      assert(warning.contains("1 dependent formula not evaluated (G2)"), warning)
      assert(warning.endsWith("; those cells show the file's values"), warning)
  }

  test("--json: ok, one EVAL_FAILED warning located at the first failing cell") {
    for run <- CliHarness.run("-f", file("degraded.xlsx"), "--json", "view", "A1:G4", "--eval")
    yield
      assertEquals(run.exit, 0, run.stderr)
      val e = ujson.read(run.stdout)
      EnvelopeSchema.assertValid(e)
      assertEquals(e("ok"), ujson.True)
      val evalFailed = e("warnings").arr.filter(_("code").str == "EVAL_FAILED")
      assertEquals(evalFailed.size, 1, e("warnings").toString)
      val location = evalFailed.headOption.map(_("location")).getOrElse(fail("no location"))
      assertEquals(location("sheet"), ujson.Str("Sheet1"))
      assertEquals(location("ref"), ujson.Str("F2"))
      assertEquals(run.stderr, "")
  }

  test("--strict turns the failure into the RECALC_GATE gate: exit 1, nothing rendered") {
    for
      run <- CliHarness.run("-f", file("degraded.xlsx"), "view", "A1:G4", "--eval", "--strict")
      json <- CliHarness
        .run("-f", file("degraded.xlsx"), "--json", "view", "A1:G4", "--eval", "--strict")
      clean <- CliHarness.run("-f", file("weaver.xlsx"), "view", "A1:F4", "--eval", "--strict")
    yield
      assertEquals(run.exit, 1, run.stderr)
      assertEquals(run.stdout, "")
      assert(
        lines(run.stderr).headOption.exists(
          _.startsWith("Error: Formula evaluation failed: Sheet1!F2: ")
        ),
        run.stderr
      )
      assert(run.stderr.contains("  code: RECALC_GATE"), run.stderr)
      // nothing is rendered under the gate, so the message says what the render WOULD show
      assert(
        lines(run.stderr).headOption
          .exists(_.endsWith("; without --strict those cells show the file's values")),
        run.stderr
      )
      assert(
        run.stderr.contains(
          "  hint: drop --strict to render the other cells live and see the failure as a warning"
        ),
        run.stderr
      )
      assertEquals(json.exit, 1, json.stderr)
      val e = ujson.read(json.stdout)
      assertEquals(e("data"), ujson.Null)
      assertEquals(e("error")("code"), ujson.Str("RECALC_GATE"))
      // a window whose closure evaluates cleanly passes the gate
      assertEquals(clean.exit, 0, clean.stderr)
  }

  test("raster formats never gate: --strict png still exports and warns once, exit 0") {
    BatikRasterizer.isAvailable.flatMap { batikAvailable =>
      assume(batikAvailable, "AWT not available - skipping the raster --strict case")
      val png = file("degraded.png")
      CliHarness
        .run(
          "-f",
          file("degraded.xlsx"),
          "view",
          "A1:G4",
          "--eval",
          "--strict",
          "--format",
          "png",
          "--raster-output",
          png
        )
        .map { run =>
          assertEquals(run.exit, 0, run.stderr)
          assert(run.stdout.contains("Exported:"), run.stdout)
          assert(Files.size(Path.of(png)) > 0, "the PNG is written")
          val warnings = lines(run.stderr).filter(_.startsWith("Warning[EVAL_FAILED]"))
          assertEquals(warnings.size, 1, run.stderr)
          assert(
            warnings.forall(_.endsWith("; those cells show the file's values")),
            run.stderr
          )
          assert(!run.stderr.contains("RECALC_GATE"), run.stderr)
        }
    }
  }

  test("every --eval render shares the per-cell evaluation (csv, html, svg)") {
    for
      csv <- CliHarness
        .run("-f", file("degraded.xlsx"), "view", "A1:G4", "--eval", "--format", "csv")
      html <- CliHarness
        .run("-f", file("degraded.xlsx"), "view", "A1:G4", "--eval", "--format", "html")
      svg <- CliHarness
        .run("-f", file("degraded.xlsx"), "view", "A1:G4", "--eval", "--format", "svg")
    yield
      val csvRow3 = lines(csv.stdout).lift(2).getOrElse(fail(csv.stdout))
      assert(csvRow3.contains(",10,"), s"csv renders the live F3, not the stale 99: $csvRow3")
      List("html" -> html, "svg" -> svg).foreach { (label, run) =>
        assert(run.stdout.contains(">10<"), s"$label renders the live F3")
        assert(!run.stdout.contains(">99<"), s"$label must not show the stale F3 cache")
      }
      List("csv" -> csv, "html" -> html, "svg" -> svg).foreach { (label, run) =>
        assertEquals(run.exit, 0, s"$label: ${run.stderr}")
        assertEquals(
          lines(run.stderr).count(_.startsWith("Warning[EVAL_FAILED]")),
          1,
          s"$label: ${run.stderr}"
        )
      }
  }

  test("evala over array-valued typed arguments answers instead of exiting INTERNAL") {
    for
      abs <- CliHarness.run("-f", file("weaver.xlsx"), "evala", "=SUMPRODUCT(ABS(C2:C4+D2:D4))")
      concat <- CliHarness.run("-f", file("weaver.xlsx"), "evala", "=(C2:C4+D2:D4)&\"x\"")
    yield
      assertEquals(abs.exit, 0, abs.stderr)
      assertEquals(concat.exit, 0, concat.stderr)
      assert(concat.stdout.contains("-2x"), concat.stdout)
      assert(concat.stdout.contains("3x"), concat.stdout)
      assert(!concat.stdout.contains("ArrayResult"), concat.stdout)
  }
