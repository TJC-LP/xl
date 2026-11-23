package com.tjclp.xl.cli

import java.nio.file.{Files, Path}

import cats.effect.IO
import cats.effect.unsafe.implicits.global

import com.tjclp.xl.{*, given}
import com.tjclp.xl.cli.contract.{CliHarness, CliRun}
import com.tjclp.xl.io.ExcelIO
import com.tjclp.xl.sheets.styleSyntax.withCellStyle
import com.tjclp.xl.styles.CellStyle
import com.tjclp.xl.styles.numfmt.NumFmt

import munit.FunSuite

/**
 * GH-935: `merge` (verb, `--stream` verb, batch op, streamed batch op) keeps only the upper-left
 * value, as Excel's Merge Cells does, and the formulas reading a cleared cell are recalculated: the
 * issue's `put A1 1; put B1 2; put A2 3; merge A1:B2; putf D1 =SUM(A1:B2)` gives 1, not 6.
 */
class MergeClearsHiddenCliSpec extends FunSuite:

  private def withDir[A](f: Path => A): A =
    val dir = Files.createTempDirectory("merge-clears-hidden")
    try f(dir)
    finally
      Files.list(dir).forEach(p => Files.deleteIfExists(p))
      Files.deleteIfExists(dir)

  private def num(n: Int): CellValue = CellValue.Number(BigDecimal(n))

  /** A1=1, B1=2, A2=3 and D1 = SUM(A1:B2) cached at 6. */
  private val repro: Workbook =
    Workbook(
      Sheet("Sheet1")
        .put(ARef.from1(1, 1), num(1))
        .put(ARef.from1(2, 1), num(2))
        .put(ARef.from1(1, 2), num(3))
        .put(ARef.from1(4, 1), CellValue.Formula("SUM(A1:B2)", Some(num(6))))
    )

  private def write(wb: Workbook, path: Path): Unit =
    ExcelIO.instance[IO].write(wb, path).unsafeRunSync()

  private def read(path: Path): Sheet =
    ExcelIO.instance[IO].read(path).unsafeRunSync().sheets.headOption.getOrElse(fail("no sheet"))

  private def run(args: String*): CliRun = CliHarness.run(args*).unsafeRunSync()

  private def runIn(stdin: String, args: String*): CliRun =
    CliHarness.run(args.toList, stdin).unsafeRunSync()

  private def value(s: Sheet, a1: String): CellValue =
    ARef.parse(a1).toOption.flatMap(s.cells.get).fold(CellValue.Empty)(_.value)

  private def cached(s: Sheet, a1: String): Option[CellValue] = value(s, a1) match
    case CellValue.Formula(_, cache, _) => cache
    case _ => None

  private def assertMerged(s: Sheet): Unit =
    assertEquals(value(s, "A1"), num(1))
    assertEquals(value(s, "B1"), CellValue.Empty)
    assertEquals(value(s, "A2"), CellValue.Empty)
    assertEquals(s.mergedRanges.map(_.toA1), Set("A1:B2"))

  test("GH-935: merge verb clears the hidden cells and recalculates SUM to the anchor's value") {
    withDir { dir =>
      val in = dir.resolve("in.xlsx")
      val out = dir.resolve("out.xlsx")
      write(repro, in)
      val r = run("-f", in.toString, "-o", out.toString, "merge", "A1:B2")
      assertEquals(r.exit, 0, r.stderr)
      assert(r.stdout.contains("Merged: A1:B2"), r.stdout)
      val s = read(out)
      assertMerged(s)
      assertEquals(cached(s, "D1"), Some(num(1)))
    }
  }

  test("GH-935: the issue's steps end to end: a putf after the merge sums to 1") {
    withDir { dir =>
      val f = dir.resolve("t.xlsx")
      write(Workbook(Sheet("Sheet1")), f)
      val p = f.toString
      assertEquals(run("-f", p, "-o", p, "put", "A1", "1").exit, 0)
      assertEquals(run("-f", p, "-o", p, "put", "B1", "2").exit, 0)
      assertEquals(run("-f", p, "-o", p, "put", "A2", "3").exit, 0)
      assertEquals(run("-f", p, "-o", p, "merge", "A1:B2").exit, 0)
      assertEquals(run("-f", p, "-o", p, "putf", "D1", "=SUM(A1:B2)").exit, 0)
      assertEquals(cached(read(f), "D1"), Some(num(1)))
    }
  }

  test("GH-935: --no-recalc keeps the old cache but still clears the hidden cells") {
    withDir { dir =>
      val in = dir.resolve("in.xlsx")
      val out = dir.resolve("out.xlsx")
      write(repro, in)
      val r = run("-f", in.toString, "-o", out.toString, "--no-recalc", "merge", "A1:B2")
      assertEquals(r.exit, 0, r.stderr)
      val s = read(out)
      assertMerged(s)
      assertEquals(cached(s, "D1"), Some(num(6)))
    }
  }

  test("GH-935: merge --stream (loaded in memory) clears and recalculates the same way") {
    withDir { dir =>
      val in = dir.resolve("in.xlsx")
      val out = dir.resolve("out.xlsx")
      write(repro, in)
      val r = run("-f", in.toString, "-o", out.toString, "--stream", "merge", "A1:B2")
      assertEquals(r.exit, 0, r.stderr)
      val s = read(out)
      assertMerged(s)
      assertEquals(cached(s, "D1"), Some(num(1)))
    }
  }

  test("GH-935: a hidden cell keeps its style; a clear of the anchor leaves nothing hidden") {
    withDir { dir =>
      val in = dir.resolve("in.xlsx")
      val merged = dir.resolve("merged.xlsx")
      val cleared = dir.resolve("cleared.xlsx")
      val styled = repro.update(
        SheetName.unsafe("Sheet1"),
        _.withCellStyle(ARef.from1(2, 1), CellStyle.default.withNumFmt(NumFmt.Percent))
      )
      write(styled.fold(e => fail(e.message), identity), in)
      assertEquals(run("-f", in.toString, "-o", merged.toString, "merge", "A1:B2").exit, 0)
      val s = read(merged)
      assertEquals(value(s, "B1"), CellValue.Empty)
      assert(ARef.parse("B1").toOption.flatMap(s.cells.get).exists(_.styleId.isDefined))
      assertEquals(run("-f", merged.toString, "-o", cleared.toString, "clear", "A1").exit, 0)
      val c = read(cleared)
      assert(
        CellRange
          .parse("A1:B2")
          .toOption
          .toList
          .flatMap(_.cells)
          .forall(r => c.cells.get(r).forall(_.value == CellValue.Empty))
      )
    }
  }

  test("GH-935: batch merge clears the hidden cells and recalculates their readers") {
    withDir { dir =>
      val in = dir.resolve("in.xlsx")
      val out = dir.resolve("out.xlsx")
      write(repro, in)
      val r = runIn(
        """[{"op":"merge","range":"A1:B2"}]""",
        "-f",
        in.toString,
        "-o",
        out.toString,
        "batch",
        "-"
      )
      assertEquals(r.exit, 0, r.stderr)
      val s = read(out)
      assertMerged(s)
      assertEquals(cached(s, "D1"), Some(num(1)))
    }
  }

  test("GH-935: batch ops apply in order: a put after the merge lands, one before is cleared") {
    withDir { dir =>
      val in = dir.resolve("in.xlsx")
      write(repro, in)
      val ops =
        """[{"op":"put","ref":"B2","value":5},{"op":"merge","range":"A1:B2"},{"op":"put","ref":"B1","value":7}]"""
      for streamed <- List(false, true) do
        val out = dir.resolve(s"out-$streamed.xlsx")
        val args = List("-f", in.toString, "-o", out.toString) ++
          (if streamed then List("--stream") else Nil) ++ List("batch", "-")
        val r = runIn(ops, args*)
        assertEquals(r.exit, 0, s"stream=$streamed: ${r.stderr}")
        val s = read(out)
        assertEquals(value(s, "A1"), num(1), s"stream=$streamed")
        assertEquals(value(s, "A2"), CellValue.Empty, s"stream=$streamed")
        assertEquals(value(s, "B2"), CellValue.Empty, s"stream=$streamed")
        assertEquals(value(s, "B1"), num(7), s"stream=$streamed")
    }
  }

  test("GH-935: streamed batch merge clears hidden cells (no recalculation under --stream)") {
    withDir { dir =>
      val in = dir.resolve("in.xlsx")
      val out = dir.resolve("out.xlsx")
      write(repro, in)
      val r = runIn(
        """[{"op":"merge","range":"A1:B2"}]""",
        "-f",
        in.toString,
        "-o",
        out.toString,
        "--stream",
        "batch",
        "-"
      )
      assertEquals(r.exit, 0, r.stderr)
      assertMerged(read(out))
    }
  }

  test("GH-935: reading and rewriting a file whose merge already hides values keeps them") {
    withDir { dir =>
      val in = dir.resolve("in.xlsx")
      val out = dir.resolve("out.xlsx")
      val hidden = repro.update(
        SheetName.unsafe("Sheet1"),
        s => s.copy(mergedRanges = s.mergedRanges ++ CellRange.parse("A1:B2").toOption.toList)
      )
      write(hidden.fold(e => fail(e.message), identity), in)
      assertEquals(run("-f", in.toString, "-o", out.toString, "put", "F1", "9").exit, 0)
      val s = read(out)
      assertEquals(value(s, "B1"), num(2))
      assertEquals(value(s, "A2"), num(3))
    }
  }
