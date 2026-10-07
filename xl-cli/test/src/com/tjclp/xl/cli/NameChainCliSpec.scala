package com.tjclp.xl.cli

import java.nio.file.{Files, Path}
import java.util.concurrent.atomic.AtomicReference

import cats.effect.{IO, Resource}
import munit.CatsEffectSuite

import com.tjclp.xl.{Sheet, Workbook, given}
import com.tjclp.xl.addressing.{ARef, SheetName}
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.cli.contract.{CliHarness, CliRun, EnvelopeSchema}
import com.tjclp.xl.formula.eval.WorkbookAudit
import com.tjclp.xl.formula.eval.WorkbookEvaluator.recalculate
import com.tjclp.xl.formula.graph.QualifiedGraph
import com.tjclp.xl.io.ExcelIO
import com.tjclp.xl.workbooks.DefinedName

/**
 * #691: `xl audit`, `deps` and `recalc` on the defined-name shapes whose walks overflowed the stack
 * or ran exponentially — a 5,000-deep chain, a depth-30 diamond, and a sheet-local name defined as
 * another sheet's same-named name — through the in-process harness. Each answers in its contract;
 * before, `audit` and `recalc` died on the chain with a raw `StackOverflowError` (exit 1).
 */
class NameChainCliSpec extends CatsEffectSuite:

  private val dir = ResourceSuiteLocalFixture(
    "name-chain-books",
    Resource.make(IO.blocking(Files.createTempDirectory("gh-691")))(path =>
      IO.blocking {
        Files.list(path).forEach(Files.delete(_))
        Files.delete(path)
      }
    )
  )

  override def munitFixtures = List(dir)

  private val S = SheetName.unsafe("S")
  private val T = SheetName.unsafe("T")

  private def at(a1: String): ARef = ARef.parse(a1).fold(err => fail(err), identity)
  private def formula(text: String): CellValue = CellValue.Formula(text, None)
  private def lvl(i: Int): String = s"Lvl_$i"

  private def withNames(wb: Workbook, names: Vector[DefinedName]): Workbook =
    wb.copy(metadata = wb.metadata.copy(definedNames = names))

  /** Names `Lvl_1 … Lvl_n`, `Lvl_i` defined as `body(i)`, `Lvl_n` as `leaf`; S!B1 = 1. */
  private def named(n: Int, body: Int => String, leaf: String, cells: (String, String)*) =
    val sheet = cells.foldLeft(Sheet(S).put(at("B1"), CellValue.Number(1))) {
      case (acc, (ref, text)) => acc.put(at(ref), formula(text))
    }
    withNames(
      Workbook(sheet),
      (1 to n).toVector.map(i => DefinedName(lvl(i), if i == n then leaf else body(i)))
    )

  /**
   * Fail fast, before the harness: a walk that overflows the stack under the harness kills the
   * cats-effect runtime (a `StackOverflowError` is fatal there), so the suite would hang rather
   * than fail. The walks behind `audit`, `deps` and `recalc` run first on a 1MB thread.
   */
  private def walksFinish(book: Workbook): IO[Unit] = IO.blocking {
    val out = new AtomicReference[Option[Throwable]](Some(new IllegalStateException("not run")))
    val thread = Thread
      .ofPlatform()
      .daemon(true)
      .stackSize(1024L * 1024L)
      .start(() =>
        out.set(
          try
            WorkbookAudit.of(book)
            QualifiedGraph.of(book)
            book.recalculate()
            None
          catch case e: Throwable => Some(e)
        )
      )
    if !thread.join(java.time.Duration.ofSeconds(30)) then fail("the name walks did not finish")
    out.get().foreach(e => fail(s"a name walk threw $e"))
  }

  private def write(book: Workbook, name: String): IO[String] =
    val path: Path = dir().resolve(name)
    ExcelIO.instance[IO].write(book, path).as(path.toString)

  private def data(run: CliRun): ujson.Value =
    val parsed = ujson.read(run.stdout)
    EnvelopeSchema.assertValid(parsed)
    assertEquals(parsed("ok"), ujson.True, run.stderr)
    parsed("data")

  private def refs(nodes: ujson.Value): Vector[String] = nodes.arr.toVector.map(_("ref").str)
  private def names(values: ujson.Value): Vector[String] = values.arr.toVector.map(_.str)

  private def cached(path: String, sheet: SheetName, a1: String): IO[Option[CellValue]] =
    ExcelIO.instance[IO].read(Path.of(path)).map { wb =>
      wb.sheets
        .find(_.name == sheet)
        .flatMap(_.cells.get(at(a1)))
        .map(_.value)
        .collect { case CellValue.Formula(_, Some(value), _) => value }
    }

  test("#691: a 5,000-deep name chain — audit, deps and recalc answer, past the cap unresolvable") {
    // A1 reads 5,000 names deep, A2 101 (one past the cap), A3 exactly 100
    val book = named(
      5000,
      i => s"${lvl(i + 1)}+1",
      "RAND()+S!$B$1",
      "A1" -> s"${lvl(1)}+1",
      "A2" -> s"${lvl(4900)}+1",
      "A3" -> s"${lvl(4901)}+1"
    )
    for
      _ <- walksFinish(book)
      in <- write(book, "chain.xlsx")
      out = dir().resolve("chain-recalc.xlsx").toString
      audit <- CliHarness.run("-f", in, "--json", "audit")
      deep <- CliHarness.run("-f", in, "--json", "deps", "S!A1", "--direction", "precedents")
      near <- CliHarness.run("-f", in, "--json", "deps", "S!A3", "--direction", "precedents")
      recalc <- CliHarness.run("-f", in, "-o", out, "recalc")
      a3 <- cached(out, S, "A3")
    yield
      assertEquals(audit.exit, 0, audit.stderr)
      val report = data(audit)
      assertEquals(names(report("volatile")), Vector("S!A3"))
      assertEquals(names(report("unresolvedReaders")), Vector("S!A1", "S!A2"))
      assertEquals(deep.exit, 0, deep.stderr)
      assertEquals(refs(data(deep)("precedents")), Vector.empty)
      assertEquals(near.exit, 0, near.stderr)
      assertEquals(refs(data(near)("precedents")), Vector("S!B1"))
      // the two cells past the cap fail with the chain named — a RECALC_ERRORS warning, exit 0
      // without --strict — and the rest of the book recalculates
      assertEquals(recalc.exit, 0, recalc.stderr)
      assert(recalc.stderr.contains("RECALC_ERRORS"), recalc.stderr)
      assert(recalc.stdout.contains("2 errors"), recalc.stdout)
      assert(recalc.stdout.contains("Defined name chain too deep"), recalc.stdout)
      assert(
        a3.exists {
          case CellValue.Number(_) => true
          case _ => false
        },
        a3
      )
  }

  test("#691: a depth-30 diamond of names — recalc evaluates it, audit and deps follow it") {
    val book = named(30, i => s"${lvl(i + 1)}+${lvl(i + 1)}", "S!$B$1", "A1" -> s"${lvl(1)}+1")
    for
      _ <- walksFinish(book)
      in <- write(book, "diamond.xlsx")
      out = dir().resolve("diamond-recalc.xlsx").toString
      recalc <- CliHarness.run("-f", in, "-o", out, "recalc")
      audit <- CliHarness.run("-f", out, "--json", "audit")
      deps <- CliHarness.run("-f", out, "--json", "deps", "S!A1", "--direction", "precedents")
      a1 <- cached(out, S, "A1")
    yield
      assertEquals(recalc.exit, 0, recalc.stderr)
      assertEquals(a1, Some(CellValue.Number(BigDecimal(2).pow(29) + 1)))
      assertEquals(audit.exit, 0, audit.stderr)
      assertEquals(data(audit)("clean"), ujson.True)
      assertEquals(refs(data(deps)("precedents")), Vector("S!B1"))
  }

  test("#691: S's local X defined as T!X reaches T's own X — recalc, deps and audit agree") {
    val s = Sheet(S).put(at("A1"), formula("X+1")).put(at("A2"), formula("V+1"))
    val t = Sheet(T).put(at("B1"), CellValue.Number(5))
    val book = withNames(
      Workbook(Vector(s, t)),
      Vector(
        DefinedName("X", "T!X", localSheetId = Some(0)),
        DefinedName("X", "T!$B$1*2", localSheetId = Some(1)),
        DefinedName("V", "T!V", localSheetId = Some(0)),
        DefinedName("V", "RAND()", localSheetId = Some(1))
      )
    )
    for
      _ <- walksFinish(book)
      in <- write(book, "cross-scope.xlsx")
      out = dir().resolve("cross-scope-recalc.xlsx").toString
      recalc <- CliHarness.run("-f", in, "-o", out, "recalc")
      deps <- CliHarness.run("-f", out, "--json", "deps", "S!A1", "--direction", "precedents")
      audit <- CliHarness.run("-f", out, "--json", "audit")
      a1 <- cached(out, S, "A1")
    yield
      assertEquals(recalc.exit, 0, recalc.stderr)
      assertEquals(a1, Some(CellValue.Number(11)))
      assertEquals(refs(data(deps)("precedents")), Vector("T!B1"))
      val report = data(audit)
      assertEquals(names(report("volatile")), Vector("S!A2"))
      assertEquals(report("clean"), ujson.True)
  }
