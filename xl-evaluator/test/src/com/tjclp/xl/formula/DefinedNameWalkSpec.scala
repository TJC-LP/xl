package com.tjclp.xl.formula

import java.util.concurrent.atomic.AtomicReference

import com.tjclp.xl.{*, given}
import com.tjclp.xl.addressing.{ARef, SheetName}
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.formula.eval.{NameWalk, WorkbookAudit}
import com.tjclp.xl.formula.graph.{DependencyGraph, NameChanges, QualifiedGraph, ReferenceScan}
import com.tjclp.xl.formula.graph.DependencyGraph.QualifiedRef
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.workbooks.{DefinedName, Workbook}
import munit.FunSuite

/**
 * #691: every defined-name walk — evaluation (`recalc`), the dependency graph (`deps`), dynamic
 * classification, unresolved readers, name-change readers, the edit-cone reach and the audit's
 * volatility check — on the three shapes that broke them, each run on a 1MB thread stack (the JVM
 * default on Linux x64) under a wall-clock budget:
 *
 *   - a 5,000-deep chain (`Lvl_1 = Lvl_2+1`, …) overflowed the stack. A chain resolves at most
 *     [[NameWalk.MaxDepth]] names deep; past that it is unresolvable — a per-cell error, an
 *     unresolved reader with no edges, never volatile or dynamic — and never a thrown error.
 *   - a depth-30 diamond (`Lvl_i = Lvl_{i+1}+Lvl_{i+1}`) has 2^29 paths, and every walk took each
 *     one: now each name is visited, and evaluated, once.
 *   - a sheet-local `X` on S defined as `T!X`, beside T's own `X`, read as a cycle because the
 *     guard was keyed on the name's text: it is keyed on the resolved definition now, so S's `X`
 *     reaches T's.
 */
class DefinedNameWalkSpec extends FunSuite:

  /** Generous: the shapes run in milliseconds; the walks this pins took hours or overflowed. */
  private val BudgetMs = 30000L

  /**
   * Run `body` on a fresh thread with a 1MB stack, the JVM default on Linux x64, rather than the
   * test runner's: a `StackOverflowError` fails the test by name, and so does a walk still running
   * when the budget runs out.
   */
  private def onSmallStack[A](body: => A): A =
    val out = new AtomicReference[Either[Throwable, A]](Left(new IllegalStateException("not run")))
    // catches StackOverflowError too (Try would not)
    val thread = Thread
      .ofPlatform()
      .name("gh-691-small-stack")
      .daemon(true)
      .stackSize(1024L * 1024L)
      .start(() =>
        out.set(
          try Right(body)
          catch case e: Throwable => Left(e)
        )
      )
    if !thread.join(java.time.Duration.ofMillis(BudgetMs)) then
      thread.interrupt()
      fail(s"the walks did not finish within ${BudgetMs}ms")
    out.get().fold(e => throw e, identity)

  private val S = SheetName.unsafe("S")
  private val T = SheetName.unsafe("T")

  private def at(a1: String): ARef = ARef.parse(a1).fold(err => fail(err), identity)
  private def q(sheet: SheetName, a1: String): QualifiedRef = QualifiedRef(sheet, at(a1))
  private def formula(text: String): CellValue = CellValue.Formula(text, None)
  private def num(n: BigDecimal): CellValue = CellValue.Number(n)

  private def withNames(wb: Workbook, names: Vector[DefinedName]): Workbook =
    wb.copy(metadata = wb.metadata.copy(definedNames = names))

  private def lvl(i: Int): String = s"Lvl_$i"

  /**
   * Names `Lvl_1 … Lvl_n`, `Lvl_i` defined as `body(i)` and `Lvl_n` as `leaf`; S!B1 = 1 and S
   * carries `cells`.
   */
  private def named(n: Int, body: Int => String, leaf: String, cells: (String, String)*): Workbook =
    val sheet = cells.foldLeft(Sheet(S).put(at("B1"), num(1))) { case (acc, (ref, text)) =>
      acc.put(at(ref), formula(text))
    }
    withNames(
      Workbook(sheet),
      (1 to n).toVector.map(i => DefinedName(lvl(i), if i == n then leaf else body(i)))
    )

  /** Each name the next plus one. */
  private def chain(n: Int, leaf: String, cells: (String, String)*): Workbook =
    named(n, i => s"${lvl(i + 1)}+1", leaf, cells*)

  /** Each name the next twice: 2^(n-1) paths from `Lvl_1` to the leaf. */
  private def diamond(n: Int, leaf: String, cells: (String, String)*): Workbook =
    named(n, i => s"${lvl(i + 1)}+${lvl(i + 1)}", leaf, cells*)

  private def cached(wb: Workbook, ref: QualifiedRef): Option[CellValue] =
    wb.sheets
      .find(_.name == ref.sheet)
      .flatMap(_.cells.get(ref.ref))
      .map(_.value)
      .collect { case CellValue.Formula(_, Some(value), _) => value }

  private def assertTooDeep[A](result: Either[XLError, A]): Unit =
    result match
      case Left(err) => assert(err.message.contains("too deep"), err.message)
      case Right(value) => fail(s"a chain past the cap must not evaluate, got $value")

  private def assertNumber(result: Either[XLError, CellValue]): BigDecimal =
    result match
      case Right(CellValue.Number(n)) => n
      case other => fail(s"expected a number, got $other")

  // A1 reads 5,000 names deep, A2 101 — one past the cap — and A3 exactly NameWalk.MaxDepth
  private val deepCells = Vector(
    "A1" -> s"${lvl(1)}+1",
    "A2" -> s"${lvl(4900)}+1",
    "A3" -> s"${lvl(4901)}+1"
  )

  test("#691: a 5,000-deep name chain is unresolvable past the cap in every walker, never thrown") {
    assertEquals(NameWalk.MaxDepth, 100)
    val wb = chain(5000, "RAND()+S!$B$1", deepCells*)
    val (a1, a2, a3, b1) = (q(S, "A1"), q(S, "A2"), q(S, "A3"), q(S, "B1"))
    onSmallStack {
      // audit: the 100-deep reader reaches RAND; the deeper ones are unresolved, never volatile
      val audit = WorkbookAudit.of(wb)
      assertEquals(audit.volatile, Vector(a3))
      assertEquals(audit.unresolvedReaders, Vector(a1, a2))
      assertEquals(audit.cycles, Vector.empty)
      assertEquals(audit.dynamic, Vector.empty)
      assertEquals(DependencyGraph.unresolvedReaders(wb), Set(a1, a2))

      // deps / the dependency graph: an edge to the leaf's cell within the cap, none past it
      val graph = QualifiedGraph.of(wb)
      assertEquals(graph.precedentsOf(a3), Set(b1))
      assertEquals(graph.precedentsOf(a2), Set.empty[QualifiedRef])
      assertEquals(graph.precedentsOf(a1), Set.empty[QualifiedRef])
      assertEquals(graph.dependents(b1, 0), Vector(Vector(a3)))
      val edges = DependencyGraph.fromWorkbook(wb)
      assertEquals(edges.get(a3), Some(Set(b1)))
      assertEquals(edges.get(a1), Some(Set.empty[QualifiedRef]))

      // recalc: past the cap is a per-cell error naming the chain, within it a value
      val result = wb.recalculate()
      assertEquals(result.errors.map(e => QualifiedRef(e.sheet, e.ref)).toSet, Set(a1, a2))
      result.errors.foreach(e => assertTooDeep(Left(e.error)))
      assert(
        cached(result.workbook, a3).exists {
          case CellValue.Number(_) => true
          case _ => false
        },
        result.errors
      )
      assertTooDeep(wb.evaluateFormula(s"=${lvl(4900)}", "S"))
      assertNumber(wb.evaluateFormula(s"=${lvl(4901)}", "S"))
      // a range slot follows the chain the same way
      assertTooDeep(wb.evaluateFormula(s"=SUM(${lvl(1)})", "S"))

      // redefining the leaf invalidates every reader (past the cap, conservatively)
      val redefined =
        withNames(
          wb,
          wb.metadata.definedNames.map(dn =>
            if dn.name == lvl(5000) then dn.copy(formula = "2") else dn
          )
        )
      assertEquals(NameChanges.readers(wb, redefined), Set(a1, a2, a3))
      // an unresolved reader's textual reach is unbounded past the cap
      assertEquals(ReferenceScan.reach(wb, S, s"=${lvl(1)}+1"), ReferenceScan.Reach.Unbounded)
    }
  }

  test("#691: a dynamic call 5,000 names deep is not dynamic past the cap; within it, it is") {
    val wb = chain(5000, "INDIRECT(\"S!B1\")", deepCells*)
    onSmallStack {
      assertEquals(DependencyGraph.dynamicCells(wb), Set(q(S, "A3")))
      assertEquals(WorkbookAudit.of(wb).dynamic, Vector(q(S, "A3")))
      val result = wb.recalculate()
      assertEquals(result.errors.map(_.ref).toSet, Set(at("A1"), at("A2")))
      assertEquals(cached(result.workbook, q(S, "A3")), Some(num(101)))
    }
  }

  test("#691: a depth-30 diamond of names costs one step per name in every walker") {
    // Lvl_30 = S!B1 = 1, so Lvl_1 = 2^29 and A1 = 2^29 + 1: 2^29 paths, each name read once
    val expected = BigDecimal(2).pow(29) + 1
    val steady = diamond(30, "S!$B$1", "A1" -> s"${lvl(1)}+1")
    val (a1, b1) = (q(S, "A1"), q(S, "B1"))
    onSmallStack {
      val audit = WorkbookAudit.of(steady)
      assertEquals(audit.volatile, Vector.empty)
      assertEquals(audit.unresolvedReaders, Vector.empty)
      assertEquals(audit.cycles, Vector.empty)
      assertEquals(DependencyGraph.unresolvedReaders(steady), Set.empty[QualifiedRef])
      assertEquals(DependencyGraph.dynamicCells(steady), Set.empty[QualifiedRef])
      assertEquals(QualifiedGraph.of(steady).precedentsOf(a1), Set(b1))
      assertEquals(DependencyGraph.fromWorkbook(steady).get(a1), Some(Set(b1)))
      val reach = ReferenceScan.reach(steady, S, s"=${lvl(1)}+1")
      assertNotEquals(reach, ReferenceScan.Reach.Unbounded)
      assert(reach.touches(Map(S -> ReferenceScan.SheetCells.of(Set(at("B1"))))), reach)

      val result = steady.recalculate()
      assertEquals(result.errors, Vector.empty)
      assertEquals(cached(result.workbook, a1), Some(num(expected)))
      assertEquals(assertNumber(steady.evaluateFormula(s"=${lvl(1)}+1", "S")), expected)

      // the same diamond over RAND is volatile, over INDIRECT dynamic — found once, not per path
      val volatile = diamond(30, "RAND()", "A1" -> s"${lvl(1)}+1")
      assertEquals(WorkbookAudit.of(volatile).volatile, Vector(a1))
      val dynamic = diamond(30, "INDIRECT(\"S!B1\")", "A1" -> s"${lvl(1)}+1")
      assertEquals(DependencyGraph.dynamicCells(dynamic), Set(a1))
      assertEquals(cached(dynamic.recalculate().workbook, a1), Some(num(expected)))
    }
  }

  test("#691: a chain of names each nested 63 calls deep is walked without recursing per name") {
    // the deepest nest the parser admits inside every name of a chain at the cap: the walks hold
    // one expression on the stack at a time, never one per name
    val nest = 63
    val wb = named(
      NameWalk.MaxDepth,
      i => ("SUM(" * nest) + lvl(i + 1) + (")" * nest),
      "INDIRECT(\"S!B1\")+S!$B$1",
      "A1" -> s"${lvl(1)}+1"
    )
    val (a1, b1) = (q(S, "A1"), q(S, "B1"))
    onSmallStack {
      assertEquals(QualifiedGraph.of(wb).precedentsOf(a1), Set(b1))
      assertEquals(DependencyGraph.fromWorkbook(wb).get(a1), Some(Set(b1)))
      val audit = WorkbookAudit.of(wb)
      assertEquals(audit.dynamic, Vector(a1))
      assertEquals(audit.unresolvedReaders, Vector.empty)
      assertEquals(DependencyGraph.dynamicCells(wb), Set(a1))
      val redefined = withNames(
        wb,
        wb.metadata.definedNames.map(dn =>
          if dn.name == lvl(NameWalk.MaxDepth) then dn.copy(formula = "2") else dn
        )
      )
      assertEquals(NameChanges.readers(wb, redefined), Set(a1))
      // evaluation recurses through values by nature: on this stack it may run out, and then the
      // cell carries the evaluator's typed error — never a thrown one
      val result = wb.recalculate()
      cached(result.workbook, a1) match
        case Some(value) => assertEquals(value, num(3))
        case None =>
          assert(result.errors.exists(_.error.message.contains("exhausted the stack")), result)
    }
  }

  /**
   * S (position 0) and T (position 1) each define their own `X`, `R`, `V`, `D` and `Y`: S's is the
   * same name on T (`T!X`), T's the real definition — except `Y`, which T defines as `S!Y`: a
   * genuine cycle across the two scopes.
   */
  private def crossScope: Workbook =
    val s = Sheet(S)
      .put(at("A1"), formula("X+1"))
      .put(at("A2"), formula("SUM(R)"))
      .put(at("A3"), formula("V+1"))
      .put(at("A4"), formula("D+1"))
      .put(at("A5"), formula("Y+1"))
    val t = Sheet(T).put(at("B1"), num(5)).put(at("B2"), num(7))
    def both(name: String, onT: String): Vector[DefinedName] =
      Vector(
        DefinedName(name, s"T!$name", localSheetId = Some(0)),
        DefinedName(name, onT, localSheetId = Some(1))
      )
    withNames(
      Workbook(Vector(s, t)),
      both("X", "T!$B$1*2") ++ both("R", "T!$B$1:$B$2") ++ both("V", "RAND()") ++
        both("D", "INDIRECT(\"T!B1\")") ++ both("Y", "S!Y")
    )

  test("#691: a sheet-local name defined as another sheet's same-named name reaches it") {
    val wb = crossScope
    val Seq(a1, a2, a3, a4, a5) = Seq("A1", "A2", "A3", "A4", "A5").map(q(S, _))
    onSmallStack {
      // evaluation: S's X is T!X, not a cycle — in a value position and in a range slot
      assertEquals(wb.evaluateFormula("=X+1", "S"), Right(num(11)))
      assertEquals(wb.evaluateFormula("=SUM(R)", "S"), Right(num(12)))
      val result = wb.recalculate()
      assertEquals(result.errors.map(_.ref), Vector(at("A5")), "only Y is a real cycle")
      assert(result.errors.forall(_.error.message.contains("cycle")), result.errors)
      assertEquals(cached(result.workbook, a1), Some(num(11)))
      assertEquals(cached(result.workbook, a2), Some(num(12)))
      assertEquals(cached(result.workbook, a4), Some(num(6)))

      // the graph follows S's name into T's: an edge to the cells T's definition reads
      val graph = QualifiedGraph.of(wb)
      assertEquals(graph.precedentsOf(a1), Set(q(T, "B1")))
      assertEquals(graph.precedentsOf(a2), Set(q(T, "B1"), q(T, "B2")))
      assertEquals(DependencyGraph.fromWorkbook(wb).get(a1), Some(Set(q(T, "B1"))))

      // the audit and dynamic classification see what T's definitions call
      val audit = WorkbookAudit.of(wb)
      assertEquals(audit.volatile, Vector(a3))
      assertEquals(audit.dynamic, Vector(a4))
      assertEquals(DependencyGraph.dynamicCells(wb), Set(a4))
      // the genuine cross-scope cycle is still caught, and unresolvable
      assertEquals(audit.unresolvedReaders, Vector(a5))
      assertEquals(DependencyGraph.unresolvedReaders(wb), Set(a5))
    }
  }

  test("#691: the name memo shares a value across paths, never a random draw") {
    // a pair of draws through one name stays two draws, as before the memo: Pair is not 0
    val draws = withNames(
      Workbook(Sheet(S)),
      Vector(DefinedName("Rnd", "RAND()"), DefinedName("Pair", "Rnd-Rnd"))
    )
    val pair = draws.evaluateFormula("=Pair", S, Clock.system, Rng.seeded(691L))
    assertNotEquals(assertNumber(pair), BigDecimal(0))

    // a name read twice through a named formula is its reading cell's: the memo never carries
    // one cell's value into another's
    val sheet = Sheet(S)
      .put(at("C1"), formula("Twice"))
      .put(at("C2"), formula("Twice"))
      .put(at("C3"), formula("Twice*1"))
    val per = withNames(
      Workbook(sheet),
      Vector(DefinedName("Here", "ROW()"), DefinedName("Twice", "Here+Here"))
    )
    val result = per.recalculate()
    assertEquals(result.errors, Vector.empty)
    assertEquals(
      Vector("C1", "C2", "C3").map(r => cached(result.workbook, q(S, r))),
      Vector(Some(num(2)), Some(num(4)), Some(num(6)))
    )
  }

  /** `prefix_1 … prefix_n`, each the next plus zero, `prefix_n` reading `last`. */
  private def links(prefix: String, n: Int, last: String): Vector[DefinedName] =
    (1 to n).toVector.map(i =>
      DefinedName(s"${prefix}_$i", if i == n then s"$last+0" else s"${prefix}_${i + 1}+0")
    )

  test("#691: a memoised name answers as evaluating it again would, on every path") {
    def book(names: Vector[DefinedName]): Workbook = withNames(Workbook(Sheet(S)), names)
    // a cycle the definitions themselves catch: Pee's IFERROR swallows Cee's refusal on the path
    // through Pee, but Cee read directly evaluates (Pee is 5, Cee 6) — never the refused value
    val swallowed = book(
      Vector(
        DefinedName("Rtop", "Pee+Cee"),
        DefinedName("Pee", "IFERROR(Cee,5)"),
        DefinedName("Cee", "Pee+1")
      )
    )
    assertEquals(swallowed.evaluateFormula("=Rtop", "S"), Right(num(11)))

    // Ex is cut by the cap on the 100-deep path IFERROR catches, then read 3 deep through Bee
    val cut = book(
      links("Lnk", 98, "Ex") ++ Vector(
        DefinedName("Ex", "Why+1"),
        DefinedName("Why", "1"),
        DefinedName("Bee", "Ex+0"),
        DefinedName("Top", "IFERROR(Lnk_1,Bee)")
      )
    )
    assertEquals(cut.evaluateFormula("=Top", "S"), Right(num(2)))

    // Ex evaluated 2 deep, then reached 100 deep where its own chain crosses the cap: refused
    // there, as evaluating it again would be, not served from the shallow visit
    val crossing = book(
      links("Deep", 98, "Ex") ++ Vector(
        DefinedName("Ex", "Why+1"),
        DefinedName("Why", "1"),
        DefinedName("Top", "Ex+Deep_1")
      )
    )
    assertTooDeep(crossing.evaluateFormula("=Top", "S"))
  }

  test("#691: the walks call a name too deep on any chain past the cap, as evaluation refuses it") {
    // Rtop3 reads Exx directly and through a 99-name chain: evaluating that chain refuses Exx,
    // the 101st name, so the static walks must not call Rtop3 resolved by its short path
    val wb = withNames(
      Workbook(Sheet(S).put(at("A1"), formula("Rtop3"))),
      links("Cn", 99, "Exx") ++ Vector(DefinedName("Exx", "7"), DefinedName("Rtop3", "Cn_1+Exx"))
    )
    onSmallStack {
      val result = wb.recalculate()
      assertEquals(result.errors.map(_.ref), Vector(at("A1")))
      result.errors.foreach(e => assertTooDeep(Left(e.error)))
      assertEquals(DependencyGraph.unresolvedReaders(wb), Set(q(S, "A1")))
      assertEquals(WorkbookAudit.of(wb).unresolvedReaders, Vector(q(S, "A1")))
    }
  }
