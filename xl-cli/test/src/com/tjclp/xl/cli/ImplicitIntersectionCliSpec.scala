package com.tjclp.xl.cli

import java.nio.file.{Files, Path}

import cats.effect.IO
import munit.CatsEffectSuite

import com.tjclp.xl.{*, given}
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.cli.contract.{CliHarness, CliRun}
import com.tjclp.xl.io.ExcelIO

/**
 * GH-714: the informational `IMPLICIT_INTERSECTION` warning — an in-memory putf or batch putf whose
 * plain-cell value (Excel's implicit intersection) differs from its array value names the cell,
 * both values and the remedy. It never gates (`--strict` included), and streaming writes and dry
 * runs stay silent.
 */
class ImplicitIntersectionCliSpec extends CatsEffectSuite:

  /** A1:A10 = 1..10 and B1:B10 = 1..10. */
  private def numbers(): IO[Path] =
    for
      p <- IO.blocking(Files.createTempFile("xl-gh714-ii-", ".xlsx"))
      sheet = (1 to 10).foldLeft(Sheet("Sheet1")) { (s, i) =>
        s.put(ARef.from0(0, i - 1), CellValue.Number(BigDecimal(i)))
          .put(ARef.from0(1, i - 1), CellValue.Number(BigDecimal(i)))
      }
      _ <- ExcelIO.instance[IO].write(Workbook(sheet), p)
    yield p

  private def outPath(): IO[Path] = IO.blocking {
    val p = Files.createTempFile("xl-gh714-ii-out-", ".xlsx")
    Files.delete(p)
    p
  }

  private def putf(at: String, formula: String, extra: String*): IO[CliRun] =
    for
      src <- numbers()
      out <- outPath()
      run <- CliHarness.run(
        (List("-f", src.toString, "-o", out.toString, "putf", at, formula, "--json") ++ extra)*
      )
    yield run

  private def warned(run: CliRun): Boolean = run.stdout.contains("IMPLICIT_INTERSECTION")

  test("SUM over an array product in row 5 warns with both values and the remedy") {
    putf("C5", "=SUM(A1:A10*B1:B10)").map { run =>
      assertEquals(run.exit, 0, run.toString)
      assert(warned(run), run.stdout)
      assert(
        run.stdout.contains(
          "Sheet1!C5: =SUM(A1:A10*B1:B10) is 25 as a plain cell (implicit intersection) but 385 " +
            "as an array formula; use SUMPRODUCT, putf --array (batch \\\"array\\\": true), or an explicit @"
        ),
        run.stdout
      )
      assert(run.stdout.contains("\"ref\": \"C5\""), run.stdout)
    }
  }

  test("a range times a scalar warns on shape, inside and outside the range's rows") {
    for
      inside <- putf("C1", "=A1:A10*2")
      outside <- putf("C20", "=A1:A10*2")
    yield
      assert(warned(inside), inside.stdout)
      assert(inside.stdout.contains("a 10x1 spill (top-left 2)"), inside.stdout)
      assert(warned(outside), outside.stdout)
      assert(outside.stdout.contains("is #VALUE! as a plain cell"), outside.stdout)
  }

  test("a spilling function warns; an array generator warns on shape") {
    for
      sorted <- putf("D1", "=SORT(A1:A3)")
      random <- putf("D1", "=SEQUENCE(3)")
    yield
      assert(warned(sorted), sorted.stdout)
      assert(warned(random), random.stdout)
  }

  test("silent: a scalar formula, SUMPRODUCT, an explicit @, --array, RAND") {
    for
      scalar <- putf("C5", "=A1*2")
      sumproduct <- putf("C5", "=SUMPRODUCT(A1:A10*B1:B10)")
      at <- putf("C5", "=@A1:A10*2")
      array <- putf("C5", "=SUM(A1:A10*B1:B10)", "--array")
      rand <- putf("C5", "=RAND()")
      whole <- putf("C5", "=SUM(A1:A10)")
    yield List(scalar, sumproduct, at, array, rand, whole).foreach { run =>
      assertEquals(run.exit, 0, run.toString)
      assert(!warned(run), run.stdout)
    }
  }

  test("--strict exits 0 with the warning; --no-recalc still warns") {
    for
      strict <- putf("C5", "=SUM(A1:A10*B1:B10)", "--strict")
      noRecalc <- putf("C5", "=SUM(A1:A10*B1:B10)", "--no-recalc")
    yield
      assertEquals(strict.exit, 0, strict.toString)
      assert(warned(strict), strict.stdout)
      assert(warned(noRecalc), noRecalc.stdout)
  }

  test("a drag lists the first ten cells, then a count") {
    putf("C1:C12", "=A$1:A$10*2").map { run =>
      assert(warned(run), run.stdout)
      assert(run.stdout.contains("12 formulas evaluate differently"), run.stdout)
      assert(run.stdout.contains("(+2 more)"), run.stdout)
      assert(run.stdout.contains("\"ref\": \"C1\""), run.stdout)
    }
  }

  private def batch(ops: String, extra: String*): IO[CliRun] =
    for
      src <- numbers()
      out <- outPath()
      run <- CliHarness.run(
        List("-f", src.toString, "-o", out.toString) ++ extra ++ List("batch", "-", "--json"),
        ops
      )
    yield run

  test("batch putf warns, located at the cell; a cell later overwritten in the batch is silent") {
    for
      run <- batch("""[{"op":"putf","ref":"C5","value":"=SUM(A1:A10*B1:B10)"}]""")
      overwritten <- batch(
        """[{"op":"putf","ref":"C5","value":"=SUM(A1:A10*B1:B10)"},{"op":"put","ref":"C5","value":1}]"""
      )
      arrayOp <- batch(
        """[{"op":"putf","ref":"C5","value":"=SUM(A1:A10*B1:B10)","array":true}]"""
      )
    yield
      assert(warned(run), run.stdout)
      assert(run.stdout.contains("\"ref\": \"C5\""), run.stdout)
      assert(!warned(overwritten), overwritten.stdout)
      assert(!warned(arrayOp), arrayOp.stdout)
  }

  test("streaming writes and dry runs stay silent") {
    val ops = """[{"op":"putf","ref":"C5","value":"=SUM(A1:A10*B1:B10)"}]"""
    for
      streamed <- putf("C5", "=SUM(A1:A10*B1:B10)", "--stream")
      streamBatch <- batch(ops, "--stream")
      dry <- CliHarness.run(List("batch", "--dry-run", "-", "--json"), ops)
    yield List(streamed, streamBatch, dry).foreach { run =>
      assertEquals(run.exit, 0, run.toString)
      assert(!warned(run), run.stdout)
    }
  }
