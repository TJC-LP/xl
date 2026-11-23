package com.tjclp.xl.cli.contract

import java.nio.file.Files

import cats.effect.{IO, Resource}
import cats.syntax.all.*
import munit.CatsEffectSuite

import com.tjclp.xl.{*, given}
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.io.ExcelIO
import com.tjclp.xl.macros.ref
import com.tjclp.xl.workbooks.CalcPr

/**
 * #917: `view --eval` of a book whose calcPr iterates, over a cycle that holds no caches, shows
 * what an iterative recalculation computes. When that cycle does not converge the cells show the
 * last round's values; the window says so with one `NOT_CONVERGED` warning in `recalc`'s words, and
 * `--strict` gates on it as `recalc --strict` does.
 */
class ViewEvalConvergenceSpec extends CatsEffectSuite:

  private def iterating(sheet: Sheet): Workbook =
    val wb = Workbook(Vector(sheet))
    wb.copy(metadata = wb.metadata.copy(calcPr = Some(CalcPr(iterativeCalculation = true))))

  private def f(text: String): CellValue = CellValue.Formula(text, None)

  private val books: Map[String, Workbook] = Map(
    "diverging.xlsx" -> iterating(
      Sheet("Sheet1").put(ref"A1", f("B1+1")).put(ref"B1", f("A1+1")).put(ref"C1", f("2*3"))
    ),
    "stalled.xlsx" -> iterating(
      Sheet("Sheet1").put(ref"A1", f("B1+Missing!A1")).put(ref"B1", f("A1+1"))
    ),
    "converging.xlsx" -> iterating(
      Sheet("Sheet1").put(ref"A1", f("B1/2+1")).put(ref"B1", f("A1/2"))
    )
  )

  private val fixtures = ResourceSuiteLocalFixture(
    "view-eval-convergence",
    Resource.make(
      IO.blocking(Files.createTempDirectory("xl-view-conv-")).flatTap { dir =>
        books.toList.traverse_((name, wb) => ExcelIO.instance[IO].write(wb, dir.resolve(name)))
      }
    )(TestFixtures.delete)
  )

  override def munitFixtures = List(fixtures)

  private def file(name: String): String = fixtures().resolve(name).toString
  private def notConverged(stderr: String): List[String] =
    stderr.linesIterator.filter(_.startsWith("Warning[NOT_CONVERGED]")).toList

  test("an unconverged cycle renders its last-round values under ONE warning, in recalc's words") {
    val out = fixtures().resolve("diverging-recalc.xlsx").toString
    for
      run <- CliHarness.run("-f", file("diverging.xlsx"), "view", "A1:C1", "--eval")
      json <- CliHarness.run("-f", file("diverging.xlsx"), "--json", "view", "A1:C1", "--eval")
      recalc <- CliHarness.run("-f", file("diverging.xlsx"), "-o", out, "recalc")
    yield
      assertEquals(run.exit, 0, run.stderr)
      val warnings = notConverged(run.stderr)
      assertEquals(warnings.size, 1, run.stderr)
      val verdict = "exhausted 100 round(s) without converging (last values kept)"
      assert(warnings.exists(_.contains(s"Iterative calculation $verdict")), run.stderr)
      assert(
        warnings.exists(_.contains("Sheet1!A1, Sheet1!B1: exhausted 100 round(s)")),
        run.stderr
      )
      assert(recalc.stdout.contains(s"WARNING: iterative calculation $verdict"), recalc.stdout)
      assert(!run.stderr.contains("EVAL_FAILED"), run.stderr)
      val row1 = run.stdout.linesIterator.find(_.startsWith("| 1 ")).getOrElse(fail(run.stdout))
      assert(!row1.contains("=B1+1"), s"the members show values, not formulas: $row1")
      assert(row1.contains("| 6 "), row1)
      assertEquals(json.exit, 0, json.stderr)
      val e = ujson.read(json.stdout)
      val coded = e("warnings").arr.filter(_("code").str == "NOT_CONVERGED")
      assertEquals(coded.size, 1, e("warnings").toString)
      assertEquals(coded(0)("location")("ref"), ujson.Str("A1"))
  }

  test("a window that does not depend on the cycle says nothing") {
    for run <- CliHarness.run("-f", file("diverging.xlsx"), "view", "C1:C1", "--eval", "--strict")
    yield
      assertEquals(run.exit, 0, run.stderr)
      assertEquals(notConverged(run.stderr), Nil)
  }

  test("a cycle that converges says nothing") {
    for run <- CliHarness.run("-f", file("converging.xlsx"), "view", "A1:B1", "--eval", "--strict")
    yield
      assertEquals(run.exit, 0, run.stderr)
      assertEquals(notConverged(run.stderr), Nil)
  }

  test("--strict turns it into the RECALC_GATE gate: exit 1, nothing rendered") {
    val out = fixtures().resolve("diverging-strict.xlsx").toString
    for
      run <- CliHarness.run("-f", file("diverging.xlsx"), "view", "A1:C1", "--eval", "--strict")
      json <- CliHarness
        .run("-f", file("diverging.xlsx"), "--json", "view", "A1:C1", "--eval", "--strict")
      recalc <- CliHarness.run("-f", file("diverging.xlsx"), "-o", out, "recalc", "--strict")
    yield
      assertEquals(run.exit, 1, run.stderr)
      assertEquals(run.stdout, "")
      assert(run.stderr.contains("  code: RECALC_GATE"), run.stderr)
      assert(run.stderr.contains("Iterative calculation exhausted 100 round(s)"), run.stderr)
      assertEquals(json.exit, 1, json.stderr)
      assertEquals(ujson.read(json.stdout)("error")("code"), ujson.Str("RECALC_GATE"))
      assertEquals(recalc.exit, 1, "recalc --strict gates on the same book")
  }

  test("html and svg, evaluated through the conditional-format closure, warn once too") {
    for
      html <- CliHarness
        .run("-f", file("diverging.xlsx"), "view", "A1:C1", "--eval", "--format", "html")
      svg <- CliHarness
        .run("-f", file("diverging.xlsx"), "view", "A1:C1", "--eval", "--format", "svg")
    yield
      assertEquals(html.exit, 0, html.stderr)
      assertEquals(notConverged(html.stderr).size, 1, html.stderr)
      assertEquals(svg.exit, 0, svg.stderr)
      assertEquals(notConverged(svg.stderr).size, 1, svg.stderr)
  }

  test("a stalled cycle is a failure first: both warnings, and --strict names the failure") {
    for
      run <- CliHarness.run("-f", file("stalled.xlsx"), "view", "A1:B1", "--eval")
      strict <- CliHarness.run("-f", file("stalled.xlsx"), "view", "A1:B1", "--eval", "--strict")
    yield
      assertEquals(run.exit, 0, run.stderr)
      assert(run.stderr.contains("Warning[EVAL_FAILED]"), run.stderr)
      assertEquals(notConverged(run.stderr).size, 1, run.stderr)
      assert(notConverged(run.stderr).exists(_.contains("stalled after")), run.stderr)
      assertEquals(strict.exit, 1, strict.stderr)
      assert(strict.stderr.contains("Formula evaluation failed"), strict.stderr)
  }
