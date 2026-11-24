package com.tjclp.xl.cli

import java.nio.file.{Files, Path}

import cats.effect.IO
import cats.effect.unsafe.implicits.global

import com.tjclp.xl.{*, given}
import com.tjclp.xl.cli.contract.{CliHarness, CliRun}
import com.tjclp.xl.io.ExcelIO

import munit.FunSuite

/**
 * GH-935 review 2: a streamed batch merge either writes exactly the file the in-memory batch writes
 * (values, formulas, merged areas, hyperlinks), or it refuses with no output file. Never a
 * different file. Plus the review's repros for blocked spills, unmerge re-spilling and legacy
 * arrays.
 */
class MergeStreamParitySpec extends FunSuite:

  override val munitTimeout = scala.concurrent.duration.Duration(180, "s")

  private def withDir[A](f: Path => A): A =
    val dir = Files.createTempDirectory("merge-stream-parity")
    try f(dir)
    finally
      Files.list(dir).forEach(p => Files.deleteIfExists(p))
      Files.deleteIfExists(dir)

  private def num(n: Int): CellValue = CellValue.Number(BigDecimal(n))
  private def text(s: String): CellValue = CellValue.Text(s)
  private def f(t: String, cache: Option[CellValue] = None): CellValue = CellValue.Formula(t, cache)
  private def at(a1: String): ARef = ARef.parse(a1).fold(e => fail(e), identity)
  private def rng(a1: String): CellRange = CellRange.parse(a1).fold(e => fail(e), identity)

  private def write(wb: Workbook, path: Path): Unit =
    ExcelIO.instance[IO].write(wb, path).unsafeRunSync()

  private def read(path: Path): Sheet =
    ExcelIO.instance[IO].read(path).unsafeRunSync().sheets.headOption.getOrElse(fail("no sheet"))

  private def run(args: String*): CliRun = CliHarness.run(args*).unsafeRunSync()

  private def runIn(stdin: String, args: String*): CliRun =
    CliHarness.run(args.toList, stdin).unsafeRunSync()

  /** What a reader of the file sees: every non-empty value or formula text, merges, links. */
  private def observed(s: Sheet): (Map[String, String], Set[String], Map[String, String]) =
    val values = s.cells.iterator
      .filter(_.value != CellValue.Empty)
      .map { c =>
        c.ref.toA1 -> (c.value match
          case CellValue.Formula(t, _, _) => s"=$t"
          case v => v.toString)
      }
      .toMap
    val linked = (s.cells.iterator.map(_.ref) ++ s.rangeHyperlinks.iterator.flatMap(_.range.cells))
      .flatMap(r => s.hyperlinkAt(r).map(l => r.toA1 -> l.toString))
      .toMap
    (values, s.mergedRanges.map(_.toA1), linked)

  private def sheet(cells: (String, CellValue)*): Sheet =
    cells.foldLeft(Sheet("Sheet1"))((s, kv) => s.put(at(kv._1), kv._2))

  /** name, the book, the batch, whether the stream must refuse, the in-memory expectations */
  private final case class Scenario(
    name: String,
    book: Sheet,
    ops: String,
    streamRefuses: Boolean,
    memoryRefuses: Boolean = false,
    check: Sheet => Unit = _ => ()
  )

  private val dyn = FormulaKind.dynamicArray(CellRange(ARef.from1(3, 1), ARef.from1(3, 3)))

  private val scenarios = Vector(
    Scenario(
      "a merge that grows over an existing merge moves a value (src.xlsx)",
      sheet("A2" -> text("keep"), "E1" -> f("A2", Some(text("keep"))))
        .copy(mergedRanges = Set(rng("B1:C3"))),
      """[{"op":"merge","range":"A1:B1"}]""",
      streamRefuses = true,
      check = s => {
        assertEquals(s(at("A1")).value, text("keep"))
        assertEquals(s.mergedRanges, Set(rng("A1:C3")))
        assertEquals(
          s(at("E1")).value match { case CellValue.Formula(t, _, _) => t; case o => o.toString },
          "A1"
        )
      }
    ),
    Scenario(
      "a merge that grows to a filled upper-left cell (src2.xlsx)",
      sheet("A2" -> text("top"), "A3" -> text("x")).copy(mergedRanges = Set(rng("B2:B3"))),
      """[{"op":"merge","range":"A3:B3"}]""",
      streamRefuses = false,
      check = s => {
        assertEquals(s(at("A2")).value, text("top"))
        assertEquals(s.cells.get(at("A3")).map(_.value).getOrElse(CellValue.Empty), CellValue.Empty)
        assertEquals(s.mergedRanges, Set(rng("A2:B3")))
      }
    ),
    Scenario(
      "a plain merge with a filled anchor",
      sheet("A1" -> num(1), "B1" -> num(2), "D1" -> f("SUM(A1:B1)", Some(num(3)))),
      """[{"op":"merge","range":"A1:B1"}]""",
      streamRefuses = false
    ),
    Scenario(
      "two merges in one batch, the second growing over the first",
      sheet("A1" -> num(1), "C1" -> num(3), "C2" -> num(4)),
      """[{"op":"merge","range":"A1:B1"},{"op":"merge","range":"B1:C2"}]""",
      streamRefuses = false,
      check = s => assertEquals(s.mergedRanges, Set(rng("A1:C2")))
    ),
    Scenario(
      "a merge elsewhere keeps the file's merges",
      sheet("E1" -> num(1)).copy(mergedRanges = Set(rng("B1:C2"))),
      """[{"op":"merge","range":"E1:F1"}]""",
      streamRefuses = false
    ),
    Scenario(
      "a merge over a dynamic spill",
      sheet(
        "C1" -> CellValue.Formula("SEQUENCE(3,1)", Some(num(1)), dyn),
        "C2" -> num(2),
        "C3" -> num(3)
      ),
      """[{"op":"merge","range":"B2:C3"}]""",
      streamRefuses = true
    ),
    Scenario(
      "a merge over a legacy array is refused by both",
      sheet(
        "A2" -> CellValue
          .Formula("SEQUENCE(2)", Some(num(1)), FormulaKind.ArrayFormula(rng("A2:A3"))),
        "A3" -> num(2)
      ),
      """[{"op":"merge","range":"A1:A3"}]""",
      streamRefuses = true,
      memoryRefuses = true
    ),
    Scenario(
      "a merge hiding a linked cell",
      sheet("A1" -> num(1))
        .put(Cell(at("B1"), text("x"), hyperlink = Some(Hyperlink("https://example.com")))),
      """[{"op":"merge","range":"A1:B1"}]""",
      streamRefuses = true
    )
  )

  for sc <- scenarios do
    test(s"GH-935 review 2 parity: ${sc.name}") {
      withDir { dir =>
        val in = dir.resolve("in.xlsx")
        val mem = dir.resolve("mem.xlsx")
        val str = dir.resolve("str.xlsx")
        write(Workbook(sc.book), in)
        val m = runIn(sc.ops, "-f", in.toString, "-o", mem.toString, "batch", "-")
        val s = runIn(sc.ops, "-f", in.toString, "-o", str.toString, "--stream", "batch", "-")
        if sc.memoryRefuses then
          assertNotEquals(m.exit, 0, m.stdout)
          assert(!Files.exists(mem))
        else
          assertEquals(m.exit, 0, m.stderr)
          sc.check(read(mem))
        if sc.streamRefuses then
          assertNotEquals(s.exit, 0, s.stdout)
          assert(!Files.exists(str), "a refused stream writes no file")
          assert(s.stderr.contains("drop --stream"), s.stderr)
        else
          assertEquals(s.exit, 0, s.stderr)
          assertEquals(observed(read(str)), observed(read(mem)))
      }
    }

  // ===== blocked spills (repro 2) =====

  private def formulaText(s: Sheet, a1: String): Option[String] =
    s.cells.get(at(a1)).map(_.value).collect { case CellValue.Formula(t, _, _) => t }

  private def cache(s: Sheet, a1: String): Option[CellValue] =
    s.cells.get(at(a1)).map(_.value).collect { case CellValue.Formula(_, c, _) => c }.flatten

  private def valueOf(s: Sheet, a1: String): CellValue =
    s.cells.get(at(a1)).map(_.value).getOrElse(CellValue.Empty)

  test("GH-935 review 2: a merge that blocks a spill empties the spill range; unmerge re-spills") {
    withDir { dir =>
      val base = dir.resolve("base.xlsx")
      write(Workbook(Sheet("Sheet1")), base)
      val sp = dir.resolve("sp.xlsx")
      assertEquals(
        run("-f", base.toString, "-o", sp.toString, "putf", "A1", "=SEQUENCE(3)", "--array").exit,
        0
      )
      assertEquals(run("-f", sp.toString, "-o", sp.toString, "putf", "C1", "=A3*10").exit, 0)
      assertEquals(cache(read(sp), "C1"), Some(num(30)))
      for range <- List("A2:B2", "A1:A2") do
        val merged = dir.resolve(s"merged-$range.xlsx")
        val r = run("-f", sp.toString, "-o", merged.toString, "merge", range)
        assertEquals(r.exit, 0, r.stderr)
        val s = read(merged)
        assertEquals(cache(s, "A1"), Some(CellValue.Error(CellError.Spill)), range)
        assertEquals(valueOf(s, "A2"), CellValue.Empty, range)
        assertEquals(valueOf(s, "A3"), CellValue.Empty, range)
        assertEquals(cache(s, "C1"), Some(num(0)), range)
        // a recalculation keeps it blocked, twice over
        val r1 = dir.resolve(s"recalc1-$range.xlsx")
        val r2 = dir.resolve(s"recalc2-$range.xlsx")
        assertEquals(run("-f", merged.toString, "-o", r1.toString, "recalc").exit, 0)
        assertEquals(run("-f", r1.toString, "-o", r2.toString, "recalc").exit, 0)
        assertEquals(cache(read(r2), "A1"), Some(CellValue.Error(CellError.Spill)), range)
        assertEquals(valueOf(read(r2), "A3"), CellValue.Empty, range)
        // unmerging lets it spill again, and its readers follow (verb and batch)
        val un = dir.resolve(s"unmerged-$range.xlsx")
        val unmergeRange = if range == "A2:B2" then "A2:B2" else "A1:A2"
        assertEquals(run("-f", merged.toString, "-o", un.toString, "unmerge", unmergeRange).exit, 0)
        val u = read(un)
        assertEquals(cache(u, "A1"), Some(num(1)), range)
        assertEquals(valueOf(u, "A2"), num(2), range)
        assertEquals(valueOf(u, "A3"), num(3), range)
        assertEquals(cache(u, "C1"), Some(num(30)), range)
        val unb = dir.resolve(s"unmerged-batch-$range.xlsx")
        val ops = s"""[{"op":"unmerge","range":"$unmergeRange"}]"""
        assertEquals(runIn(ops, "-f", merged.toString, "-o", unb.toString, "batch", "-").exit, 0)
        assertEquals(valueOf(read(unb), "A3"), num(3), range)
        assertEquals(cache(read(unb), "C1"), Some(num(30)), range)
    }
  }

  // ===== legacy arrays (repro 3) =====

  test("GH-935 review 2: merging into a legacy array formula is refused (verb and batch)") {
    withDir { dir =>
      val in = dir.resolve("cse.xlsx")
      val out = dir.resolve("out.xlsx")
      write(
        Workbook(
          sheet(
            "A2" -> CellValue
              .Formula("SEQUENCE(2)", Some(num(1)), FormulaKind.ArrayFormula(rng("A2:A3"))),
            "A3" -> num(2),
            "C1" -> f("SUM(A1:A3)", Some(num(3)))
          )
        ),
        in
      )
      val r = run("-f", in.toString, "-o", out.toString, "merge", "A1:A3")
      assertNotEquals(r.exit, 0)
      assert(r.stderr.contains("array formula at A2:A3"), r.stderr)
      assert(!Files.exists(out))
      val b = runIn(
        """[{"op":"merge","range":"A1:A3"}]""",
        "-f",
        in.toString,
        "-o",
        out.toString,
        "batch",
        "-"
      )
      assertNotEquals(b.exit, 0)
      assert(b.stderr.contains("array formula at A2:A3"), b.stderr)
    }
  }
