package com.tjclp.xl.cli

import java.nio.file.{Files, Path}

import scala.concurrent.duration.*

import cats.effect.IO
import munit.CatsEffectSuite

import com.tjclp.xl.{*, given}
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.cli.commands.WriteCommands
import com.tjclp.xl.cli.contract.{CliHarness, CliRun, IntersectionHit}
import com.tjclp.xl.cli.helpers.BatchParser
import com.tjclp.xl.formula.eval.ImplicitIntersection
import com.tjclp.xl.io.ExcelIO
import com.tjclp.xl.macros.ref

/**
 * GH-714: the informational `IMPLICIT_INTERSECTION` warning — an in-memory putf or batch putf whose
 * plain-cell value (Excel's implicit intersection) differs from its array value names the cell,
 * both values and the remedy. It never gates (`--strict` included), and streaming writes and dry
 * runs stay silent.
 */
class ImplicitIntersectionCliSpec extends CatsEffectSuite:

  // the 5,000-row drags below take seconds each; the budget is asserted, not this timeout
  override def munitIOTimeout: Duration = 5.minutes

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

  test("a long drag checks a bounded sample and says so") {
    putf("C1:C1100", "=A$1:A$10*2").map { run =>
      assertEquals(run.exit, 0, run.toString)
      assert(run.stdout.contains("1000 formulas evaluate differently"), run.stdout)
      assert(run.stdout.contains("(+990 more)"), run.stdout)
      // the check stops one candidate past its budget: the count is a lower bound
      assert(run.stdout.contains("(first 1000 of at least 1001 cells checked)"), run.stdout)
    }
  }

  /** A1:A`rows` = 1..rows. */
  private def column(rows: Int): IO[Path] =
    for
      p <- IO.blocking(Files.createTempFile("xl-gh714-ii-col-", ".xlsx"))
      sheet = (1 to rows).foldLeft(Sheet("Sheet1")) { (s, i) =>
        s.put(ARef.from0(0, i - 1), CellValue.Number(BigDecimal(i)))
      }
      _ <- ExcelIO.instance[IO].write(Workbook(sheet), p)
    yield p

  private def timedDrag(src: Path, formula: String): IO[(CliRun, Long)] =
    for
      out <- outPath()
      start <- IO.monotonic
      run <- CliHarness.run(
        "-f",
        src.toString,
        "-o",
        out.toString,
        "putf",
        "B1:B5000",
        formula,
        "--no-recalc",
        "--json"
      )
      end <- IO.monotonic
    yield (run, (end - start).toMillis)

  /**
   * One evaluation of each running total, as `--no-recalc` caches it: the drag's own cost by design
   * (GH-714 adds none for a formula with no array source in a value position).
   */
  private def evaluationPass(src: Path): IO[Long] =
    for
      wb <- ExcelIO.instance[IO].read(src)
      sheet <- IO.fromOption(wb.sheets.headOption)(new IllegalStateException("no sheet"))
      start <- IO.monotonic
      _ <- IO.blocking((1 to 5000).foreach { i =>
        sheet.evaluateFormula(
          s"=SUM($$A$$1:A$i)",
          workbook = Some(wb),
          currentCell = Some(ARef.from0(1, i - 1))
        )
      })
      end <- IO.monotonic
    yield (end - start).toMillis

  test(
    "a 5,000-row running-total drag under --no-recalc evaluates each cell once, not three times"
  ) {
    for
      src <- column(5000)
      _ <- timedDrag(src, "=A1+1") // warm-up
      passMs <- evaluationPass(src)
      (plain, plainMs) <- timedDrag(src, "=A1+1")
      (total, totalMs) <- timedDrag(src, "=SUM($A$1:A1)")
    yield
      assertEquals(plain.exit, 0, plain.toString)
      assertEquals(total.exit, 0, total.toString)
      assert(!warned(total), total.stdout)
      // measured alone, the drag takes about 1.6 passes (the cache pass, then the write); with
      // two more evaluations per cell for the advisory it took 4.4
      val bound = plainMs + passMs * 5 / 2 + 2000
      assert(
        totalMs <= bound,
        s"running total ${totalMs}ms, bound ${bound}ms (plain ${plainMs}ms, pass ${passMs}ms)"
      )
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

  // ---------------------------------------------------------------------------------------------
  // PR #721 review follow-ups

  private val clock = com.tjclp.xl.formula.Clock.fixed(
    java.time.LocalDate.of(2026, 1, 2),
    java.time.LocalDateTime.of(2026, 1, 2, 3, 4, 5)
  )

  private def parsed(formula: String) =
    com.tjclp.xl.formula.FormulaParser.parse(formula).fold(e => fail(e.toString), identity)

  test("a drag whose source cannot diverge is not checked cell by cell; one that can is") {
    val range = CellRange.parse("C1:C3").fold(fail(_), identity)
    val S = SheetName.unsafe("S")
    assertEquals(
      WriteCommands.dragTargets(S, range, parsed("=SUM($A$1:A1)")).toVector,
      Vector.empty
    )
    assertEquals(WriteCommands.dragTargets(S, range, parsed("=A1+1")).toVector, Vector.empty)
    assertEquals(
      WriteCommands.dragTargets(S, range, parsed("=A$1:A$10*2")).toVector,
      range.cellsRowMajor.map(S -> _).toVector
    )
    val ops =
      """[{"op":"putf","ref":"C1:C3","value":"=SUM($A$1:A1)","from":"C1"},
        |{"op":"putf","ref":"D1:D3","value":"=A$1:A$10*2","from":"D1"}]""".stripMargin
    val sheet = Sheet("S")
    BatchParser.parseBatchOperations(ops).map { result =>
      val targets = BatchParser.putfTargets(Workbook(sheet), Some(sheet), result.scoped)
      // areas, never expanded: the check walks them lazily
      assertEquals(targets.map(_._2.toA1), Vector("D1:D3"))
    }
  }

  test("batch drag copies are known candidates unless a later op writes them or a ref voids") {
    val sheet = Sheet("S")
    val probe = Vector("C1", "C2", "C3", "D1", "D2", "D3", "D4", "D5")
    def known(ops: String): IO[Set[String]] =
      BatchParser.parseBatchOperations(ops).map { result =>
        val isKnown = BatchParser.putfKnownCandidates(Workbook(sheet), Some(sheet), result.scoped)
        probe.filter(r => isKnown(sheet.name -> ARef.parse(r).fold(fail(_), identity))).toSet
      }
    for
      plain <- known(
        """[{"op":"putf","ref":"C1:C3","value":"=SUM($A$1:A1)","from":"C1"},
          |{"op":"putf","ref":"D1:D3","value":"=A$1:A$10*2","from":"D1"},
          |{"op":"put","ref":"D2","value":1}]""".stripMargin
      )
      copied <- known(
        """[{"op":"putf","ref":"D1:D3","value":"=A$1:A$10*2","from":"D1"},
          |{"op":"copy","source":"Z1","target":"Z2"}]""".stripMargin
      )
      // dragged up from D5, A1:A10 shifts off the grid at D1
      voided <- known("""[{"op":"putf","ref":"D1:D5","value":"=A1:A10*2","from":"D5"}]""")
    yield
      assertEquals(plain, Set("D1", "D3"))
      assertEquals(copied, Set.empty[String])
      assertEquals(voided, Set.empty[String])
  }

  test("a large batch: known candidates in one walk, targets never expanded past the budget") {
    // 2,000 drags of 5,000 cells each (10M cells) and 2,000 later puts: the old per-drag
    // `drop(i + 1)` over expanded cells was O(ops² x cells)
    val drags = (0 until 2000).map { i =>
      val col = Column.from0(i % 50).toLetter
      val top = (i / 50) * 5000 + 1
      s"""{"op":"putf","ref":"$col$top:$col${top + 4999}","value":"=A$$1:A$$10*2","from":"$col$top"}"""
    }
    val puts = (1 to 2000).map(i => s"""{"op":"put","ref":"B$i","value":$i}""")
    val ops = (drags ++ puts).mkString("[", ",", "]")
    // only the first area holds formulas: the check must stop inside it
    val sheet = (1 to 2000).foldLeft(Sheet("S")) { (s, row) =>
      s.put(ARef.from0(0, row - 1), CellValue.Formula("A$1:A$10*2"))
    }
    val wb = Workbook(sheet)
    for
      result <- BatchParser.parseBatchOperations(ops)
      start <- IO.monotonic
      areas = BatchParser.putfTargets(wb, Some(sheet), result.scoped)
      isKnown = BatchParser.putfKnownCandidates(wb, Some(sheet), result.scoped)
      built <- IO.monotonic
      pulled = new java.util.concurrent.atomic.AtomicInteger(0)
      targets = areas.view.flatMap((n, r) =>
        r.cellsRowMajor.map { c =>
          pulled.incrementAndGet()
          n -> c
        }
      )
      report = ImplicitIntersection.check(wb, targets, clock, known = isKnown)
    yield
      assertEquals(areas.size, 2000)
      assert(isKnown(sheet.name -> ARef.from0(0, 0)))
      assert(!isKnown(sheet.name -> ARef.from0(1, 0)), "B1 is written by a later put")
      assert(isKnown(sheet.name -> ARef.from0(1, 5000)), "B5001: a later drag, no later write")
      assertEquals((report.checked, report.candidates, report.exact), (1000, 1001, false))
      // one past the budget, out of ten million target cells
      assertEquals(pulled.get(), 1001)
      assert((built - start).toSeconds < 20, s"building took ${(built - start).toMillis}ms")
  }

  test("#739: a sampled check with no hit in its sample says it sampled, not all clear") {
    // 3 candidates, budget 2: only the third diverges (the IF picks the range in row 3)
    val sheet = (1 to 10)
      .foldLeft(Sheet("Sheet1")) { (s, i) =>
        s.put(ARef.from0(0, i - 1), CellValue.Number(BigDecimal(i)))
      }
      .put(ref"C1", CellValue.Formula("IF(ROW()>2,A$1:A$10*2,1)"))
      .put(ref"C2", CellValue.Formula("IF(ROW()>2,A$1:A$10*2,1)"))
      .put(ref"C3", CellValue.Formula("IF(ROW()>2,A$1:A$10*2,1)"))
    val wb = Workbook(sheet)
    val cells = Vector(ref"C1", ref"C2", ref"C3").map(sheet.name -> _)
    val full = ImplicitIntersection.check(wb, cells, clock)
    assertEquals(full.hits.map(_.cell), Vector(ref"C3"))
    val sampled = ImplicitIntersection.check(wb, cells, clock, budget = 2)
    assertEquals((sampled.hits, sampled.checked, sampled.candidates), (Vector.empty, 2, 3))
    assert(sampled.sampled)
    // no divergence to warn about, but an informational note that the check sampled
    assertEquals(IntersectionHit.warning(sampled), None)
    val note = IntersectionHit.sampledNote(sampled).getOrElse(fail("expected a sampled note"))
    assertEquals(note.code, "INTERSECTION_SAMPLED")
    assert(note.message.contains("the first 2 of at least 3 candidate cells"), note.message)
    assert(note.message.contains("the rest were not checked"), note.message)
    // a full check, or one with a hit, adds no note
    assertEquals(IntersectionHit.sampledNote(full), None)
    assertEquals(
      IntersectionHit.sampledNote(ImplicitIntersection.check(wb, cells, clock, budget = 3)),
      None
    )
  }

  test("#739: a long drag with no divergence in its sample says it sampled") {
    // every copy is a candidate (a range in a value position) and none diverges below row 5000
    putf("C1:C1100", "=IF(ROW()>5000,A$1:A$10*2,1)").map { run =>
      assertEquals(run.exit, 0, run.toString)
      assert(!run.stdout.contains("\"IMPLICIT_INTERSECTION\""), run.stdout)
      assert(run.stdout.contains("INTERSECTION_SAMPLED"), run.stdout)
      assert(run.stdout.contains("the first 1000 of at least 1001 candidate cells"), run.stdout)
    }
  }

  test("a text value in the warning doubles its quotes, as a formula literal does") {
    val report = ImplicitIntersection.Report(
      Vector(
        ImplicitIntersection.Divergence(
          SheetName.unsafe("S"),
          ref"C1",
          "A1:A2&\"\"",
          CellValue.Text("say \"hi\""),
          CellValue.Text("x"),
          2,
          1
        )
      ),
      checked = 1,
      candidates = 1
    )
    val message = IntersectionHit.warning(report).fold(fail("no warning"))(_.message)
    assert(message.contains("is \"say \"\"hi\"\"\" as a plain cell"), message)
    assert(message.contains("(top-left \"x\")"), message)
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
