package com.tjclp.xl.formula.eval

import scala.collection.immutable.TreeSet

import com.tjclp.xl.addressing.{ARef, SheetName}
import com.tjclp.xl.cells.{Cell, CellValue}
import com.tjclp.xl.formula.{Clock, Rng}
import com.tjclp.xl.formula.eval.WorkbookEvaluator.*
import com.tjclp.xl.formula.graph.{CompactGraphGens, NaiveGraphs}
import com.tjclp.xl.formula.graph.DependencyGraph.QualifiedRef
import com.tjclp.xl.sheets.{CellStore, Sheet}
import com.tjclp.xl.workbooks.Workbook
import munit.ScalaCheckSuite
import org.scalacheck.{Gen, Prop}

/**
 * #767/#891: the recalculation order is canonical. A result depends only on the workbook's value:
 * the same whatever order its cells were inserted in (sequential with a seeded RAND, whose draws
 * follow evaluation order, and iterative under both schemes); a parallel pass equals the sequential
 * one; error vectors list in grid order. (Until #891 these laws compared the compact graph against
 * a second, per-cell engine, which is gone.) One law is independent of the pass's own graph: on
 * acyclic books, `recalculate` equals a fold that evaluates one formula at a time in the order a
 * `TreeSet` Kahn gives over the test-side per-cell dependency maps.
 */
class RecalcOrderLawSpec extends ScalaCheckSuite:

  override def scalaCheckTestParameters =
    super.scalaCheckTestParameters.withMinSuccessfulTests(80)

  private val clock = Clock.fixedDate(java.time.LocalDate.of(2026, 1, 2))

  private def same(a: RecalcResult, b: RecalcResult, label: String): Prop =
    Prop.all(
      Prop(a.errors == b.errors) :| s"$label errors\n${a.errors}\n${b.errors}",
      Prop(a.evaluated == b.evaluated) :| s"$label evaluated",
      Prop(a.workbook == b.workbook) :| s"$label workbook",
      Prop(a.errorReasons == b.errorReasons) :| s"$label reasons",
      Prop(a.converged == b.converged) :| s"$label converged",
      Prop(a.iterationsUsed == b.iterationsUsed) :| s"$label rounds",
      Prop(a.cycles == b.cycles) :| s"$label cycles"
    )

  /**
   * The same workbook with every sheet's cells inserted in a seeded shuffled order. (A `CellStore`
   * is canonical whatever the insertion order, so since #891 this pins that the shuffle is
   * invisible all the way through recalculation.)
   */
  private def shuffled(wb: Workbook, seed: Long): Workbook =
    val rnd = new scala.util.Random(seed)
    wb.copy(sheets = wb.sheets.map { s =>
      s.copy(cells = CellStore.from(rnd.shuffle(s.cells.iterator.toVector)))
    })

  /** The random graph books plus a few seeded RAND readers, so evaluation order is observable. */
  private val genBook: Gen[Workbook] =
    for
      wb <- CompactGraphGens.genWorkbook
      rands <- Gen.listOfN(4, Gen.zip(Gen.choose(0, 3), Gen.choose(1, 10)))
    yield rands.foldLeft(wb) { case (book, (c, r)) =>
      val name = SheetName.unsafe("Mid")
      book.sheets.find(_.name == name).fold(book) { sheet =>
        val ref = ARef.from0(c + 4, r - 1)
        book.put(sheet.put(ref, CellValue.Formula(s"RAND()+A$r")))
      }
    }

  private def run(options: RecalcOptions)(wb: Workbook): RecalcResult =
    recalculateCounted(wb, options, RecalcCounters.disabled)

  property("sequential with a seeded Rng: the result ignores the cells' insertion order") {
    Prop.forAll(genBook, Gen.choose(0L, 1000L)) { (wb, seed) =>
      def opts = RecalcOptions(clock = clock, rng = Rng.seeded(7L), iterative = IterativeMode.Off)
      same(run(opts)(wb), run(opts)(shuffled(wb, seed)), "sequential")
    }
  }

  property("parallel equals the sequential pass") {
    Prop.forAll(CompactGraphGens.genWorkbook) { wb =>
      val par = RecalcOptions(clock = clock, iterative = IterativeMode.Off, parallelism = 3)
      same(run(par)(wb), run(par.copy(parallelism = 1))(wb), "parallel vs sequential")
    }
  }

  /** Tall books: the depth-0 wave alone usually exceeds the parallel path's sequential cutoff. */
  private val genTall: Gen[Workbook] = CompactGraphGens.genWorkbookOf(40)

  /** The widest wave a parallel recalculation of `wb` would run (its cyclic core removed). */
  private def widestWave(wb: Workbook): Int =
    val graph = RecalcGraph.build(wb, RecalcCounters.disabled)
    val core = graph.cyclicCore
    val removed = core ++ graph.dependents(core, Set.empty)
    graph.order(removed).fold(_ => 0, o => o.waves(o.order).map(_.length).maxOption.getOrElse(0))

  test("the tall generator reaches the parallel path: waves wider than its cutoff") {
    val books =
      (1 to 20).flatMap(i => genTall(Gen.Parameters.default, org.scalacheck.rng.Seed(i.toLong)))
    assert(books.count(widestWave(_) >= 16) >= books.size / 2, books.map(widestWave).toString)
  }

  property("parallel over tall books equals the sequential pass") {
    Prop.forAll(genTall) { wb =>
      val par = RecalcOptions(clock = clock, iterative = IterativeMode.Off, parallelism = 4)
      same(run(par)(wb), run(par.copy(parallelism = 1))(wb), "parallel vs sequential")
    }
  }

  property("iterative, Gauss-Seidel and Jacobi: the result ignores the cells' insertion order") {
    Prop.forAll(
      genBook,
      Gen.oneOf(IterationScheme.GaussSeidel, IterationScheme.Jacobi),
      Gen.choose(0L, 1000L)
    ) { (wb, scheme, seed) =>
      def opts = RecalcOptions(
        clock = clock,
        rng = Rng.seeded(11L),
        iterative = IterativeMode.Force(IterativeCalc(25, BigDecimal("0.001"), scheme = scheme))
      )
      same(run(opts)(wb), run(opts)(shuffled(wb, seed)), s"$scheme")
    }
  }

  // ---- the order, against an independent fold --------------------------------------------------

  /**
   * A random acyclic three-sheet book of `rows` rows by four columns. Every slot draws a rank; a
   * formula reads constants, empty cells and formulas of a lower rank only, so the book has no
   * cycle and its dependency order is unrelated to its grid order. Formulas read single cells (on
   * their own sheet or another), small ranges and whole columns where every formula inside ranks
   * lower, and most draw RAND, so the order of evaluation shows in the values. Arithmetic is
   * addition, doubling, SUM and COUNT over numbers: no formula fails. `cache` is what every formula
   * holds before the pass.
   */
  private def dagBook(seed: Long, rows: Int, cache: Option[CellValue]): Workbook =
    val rnd = new scala.util.Random(seed)
    val names = CompactGraphGens.sheetNames
    val cols = Vector("A", "B", "C", "D")
    val slots = for s <- names; r <- 1 to rows; c <- cols.indices yield (s, c, r)
    val rank = rnd.shuffle(slots).zipWithIndex.toMap
    // 0: empty, 1: a constant, 2: a formula
    val kind = slots.map(slot => slot -> Vector(0, 1, 1, 2, 2, 2)(rnd.nextInt(6))).toMap
    def a1(slot: (String, Int, Int)): String = s"${cols(slot._2)}${slot._3}"
    def written(reader: (String, Int, Int), sheet: String, ref: String): String =
      if sheet == reader._1 && rnd.nextBoolean() then ref else s"'$sheet'!$ref"
    def formula(reader: (String, Int, Int)): String =
      def readable(slot: (String, Int, Int)): Boolean =
        slot != reader && (kind(slot) != 2 || rank(slot) < rank(reader))
      val points = slots.filter(readable)
      def point: String =
        if points.isEmpty then "1"
        else
          val slot = points(rnd.nextInt(points.size))
          written(reader, slot._1, a1(slot))
      def range: String =
        val sheet = names(rnd.nextInt(names.size))
        val (c0, r0) = (rnd.nextInt(cols.size), 1 + rnd.nextInt(rows))
        val whole = rnd.nextInt(6) == 0
        val (c1, r1) =
          if whole then (c0, rows)
          else (math.min(cols.size - 1, c0 + rnd.nextInt(2)), math.min(rows, r0 + rnd.nextInt(3)))
        val top = if whole then 1 else r0
        val inside = for c <- c0 to c1; r <- top to r1 yield (sheet, c, r)
        if !inside.forall(readable) then point
        else
          val text =
            if whole then s"${cols(c0)}:${cols(c0)}" else s"${cols(c0)}$r0:${cols(c1)}$r1"
          val call = if rnd.nextInt(4) == 0 then "COUNT" else "SUM"
          s"$call(${written(reader, sheet, text)})"
      def term: String = rnd.nextInt(8) match
        case 0 | 1 | 2 => point
        case 3 | 4 => range
        case 5 => s"$point*2"
        case 6 => "RAND()"
        case _ => "RANDBETWEEN(1,9)"
      (Vector.fill(1 + rnd.nextInt(3))(term) ++ Option.when(rnd.nextInt(3) > 0)("RAND()"))
        .mkString("+")
    // texts are drawn slot by slot in grid order, so one seed gives one book
    val cells = slots.flatMap { slot =>
      val ref = ARef.from0(slot._2, slot._3 - 1)
      kind(slot) match
        case 0 => None
        case 1 => Some(slot._1 -> Cell(ref, CellValue.Number(BigDecimal(rnd.nextInt(10)))))
        case _ => Some(slot._1 -> Cell(ref, CellValue.Formula(formula(slot), cache)))
    }
    Workbook(names.map { name =>
      Sheet(
        SheetName.unsafe(name),
        cells = com.tjclp.xl.sheets.CellStore.from(cells.collect { case (`name`, c) => c })
      )
    })

  /** Grid order, written out here: sheet name as plain text, then row, then column. */
  private val gridOrder: Ordering[QualifiedRef] =
    Ordering.by(q => (q.sheet.value, q.ref.row.index0, q.ref.col.index0))

  /**
   * The lexicographically smallest topological order in grid order: Kahn's algorithm over the
   * per-cell dependency maps of the test-side [[NaiveGraphs]], its ready set a `TreeSet`. None when
   * a cycle leaves formulas unordered.
   */
  private def kahnInGridOrder(wb: Workbook): Option[Vector[QualifiedRef]] =
    val (reads, readers) = NaiveGraphs.fromWorkbookFormulaGraph(wb)
    @scala.annotation.tailrec
    def go(
      ready: TreeSet[QualifiedRef],
      waiting: Map[QualifiedRef, Int],
      done: Vector[QualifiedRef]
    ): Vector[QualifiedRef] =
      ready.headOption match
        case None => done
        case Some(next) =>
          val (nowReady, stillWaiting) =
            readers.getOrElse(next, Set.empty).foldLeft((ready - next, waiting)) {
              case ((r, w), reader) =>
                val left = w.getOrElse(reader, 0) - 1
                (if left == 0 then r + reader else r, w.updated(reader, left))
            }
          go(nowReady, stillWaiting, done :+ next)
    val waiting = reads.map((q, read) => q -> read.size)
    val start = TreeSet.from(waiting.iterator.collect { case (q, 0) => q })(using gridOrder)
    val order = go(start, waiting, Vector.empty)
    Option.when(order.size == reads.size)(order)

  /**
   * The oracle: evaluate the formulas one at a time in `order`, each against the book as the ones
   * before it left it, and write its value back as its cache before the next. One seeded Rng serves
   * every formula, so a draw's position in the sequence is the formula's position in the order.
   * Left: a formula that failed.
   */
  private def foldInOrder(
    wb: Workbook,
    order: Vector[QualifiedRef],
    rng: Rng
  ): Either[String, Workbook] =
    val evaluator = Evaluator.instance(rng)
    order.foldLeft[Either[String, Workbook]](Right(wb)) { (acc, q) =>
      for
        book <- acc
        sheet <- book(q.sheet).left.map(_.message)
        outcome <- SheetEvaluator
          .evaluateCellOutcome(sheet, q.ref, evaluator, clock, Some(book))
          .left
          .map(e => s"$q: ${e.message}")
        text <- sheet.cells.get(q.ref).map(_.value) match
          case Some(formula: CellValue.Formula) => Right(formula)
          case other => Left(s"$q holds $other")
      yield book.put(sheet.put(q.ref, text.copy(cachedValue = Some(outcome.value))))
    }

  private def cellValues(wb: Workbook): Map[QualifiedRef, CellValue] =
    wb.sheets.iterator
      .flatMap(s => s.cells.iterator.map(cell => QualifiedRef(s.name, cell.ref) -> cell.value))
      .toMap

  /** The stale value every formula of a "cached" starting book holds. */
  private val stale: Option[CellValue] = Some(CellValue.Number(BigDecimal(-777)))

  property("recalculate equals a one-at-a-time fold in TreeSet-Kahn grid order") {
    Prop.forAll(Gen.choose(0L, Long.MaxValue), Gen.choose(3, 9), Gen.oneOf(None, stale)) {
      (seed, rows, cache) =>
        val wb = dagBook(seed, rows, cache)
        def opts = RecalcOptions(clock = clock, rng = Rng.seeded(7L), iterative = IterativeMode.Off)
        val result = run(opts)(wb)
        val order = kahnInGridOrder(wb)
        val folded =
          order.toRight("the book has a cycle").flatMap(foldInOrder(wb, _, Rng.seeded(7L)))
        val formulas = order.fold(0)(_.size)
        Prop.all(
          Prop(result.errors.isEmpty) :| s"errors ${result.errors}",
          Prop(folded.map(cellValues) == Right(cellValues(result.workbook))) :|
            s"seed $seed rows $rows cache $cache\nfold ${folded.map(cellValues)}\npass ${cellValues(result.workbook)}",
          Prop(result.evaluated.valuesIterator.map(_.size).sum == formulas) :| "evaluated count"
        )
    }
  }

  property("the uncached pass over a book with no caches equals the same fold") {
    Prop.forAll(Gen.choose(0L, Long.MaxValue), Gen.choose(3, 9)) { (seed, rows) =>
      val wb = dagBook(seed, rows, None)
      def opts = RecalcOptions(clock = clock, rng = Rng.seeded(7L), iterative = IterativeMode.Off)
      val result = recalculateUncachedCounted(wb, opts, RecalcCounters.disabled)
      val folded = kahnInGridOrder(wb)
        .toRight("the book has a cycle")
        .flatMap(foldInOrder(wb, _, Rng.seeded(7L)))
      Prop(folded.map(cellValues) == Right(cellValues(result.workbook))) :| s"seed $seed rows $rows"
    }
  }

  test("the fold law's books: several sheets, RAND readers, and an order that is not grid order") {
    val books = (1L to 40L).map(dagBook(_, 6, None))
    val orders = books.flatMap(kahnInGridOrder)
    assertEquals(orders.size, books.size, "every generated book is acyclic")
    // the dependency order differs from plain grid order in most books, and crosses sheets
    assert(orders.count(o => o != o.sorted(using gridOrder)) >= 30, orders.map(_.size).toString)
    assert(orders.forall(_.map(_.sheet).distinct.size == 3))
    val texts = books.flatMap(_.sheets.flatMap(_.cells.iterator.map(_.value))).collect {
      case CellValue.Formula(text, _, _) => text
    }
    assert(texts.count(_.contains("RAND")) > texts.size / 2)
    assert(texts.exists(_.contains("'!")) && texts.exists(_.contains("SUM(")))
    // a different order of the same formulas gives different values: the law can fail
    val wb = books.head
    val order = orders.head
    val inGridOrder = foldInOrder(wb, order.sorted(using gridOrder), Rng.seeded(7L))
    assertNotEquals(
      inGridOrder.map(cellValues),
      foldInOrder(wb, order, Rng.seeded(7L)).map(cellValues)
    )
  }

  test("the fold's grid order is the order the pass lists errors in") {
    val listed = Vector("b" -> "A10", "b" -> "A2", "B" -> "A3").map { (s, a1) =>
      QualifiedRef(SheetName.unsafe(s), ARef.parse(a1).getOrElse(ARef.from0(0, 0)))
    }
    assertEquals(
      listed.sorted(using gridOrder).map(q => s"${q.sheet.value}!${q.ref.toA1}"),
      Vector("B!A3", "b!A2", "b!A10")
    )
    assertEquals(listed.sorted(using gridOrder), listed.sortBy(QualifiedRef.gridKey))
  }

  // ---- the canonical order, observed --------------------------------------------------------

  private def sheet(name: String, cells: (String, CellValue)*): Sheet =
    Sheet(
      SheetName.unsafe(name),
      cells = CellStore.from(cells.map { (a1, v) =>
        Cell(ARef.parse(a1).getOrElse(ARef.from0(0, 0)), v)
      })
    )

  private def fx(text: String): CellValue = CellValue.Formula(text)

  test("cycle and blocked errors list in grid order: A2 before A10, sheet name first") {
    val wb = Workbook(
      Vector(
        sheet("b", "A10" -> fx("A10+1"), "A2" -> fx("A2+1"), "C1" -> fx("A10")),
        sheet("B", "A3" -> fx("A3*2"), "B1" -> fx("A3"))
      )
    )
    val opts = RecalcOptions(clock = clock, iterative = IterativeMode.Off)
    val errors = run(opts)(wb).errors.map(e => s"${e.sheet.value}!${e.ref.toA1}")
    // circular first, then blocked; each in (sheet name, row, column) order
    assertEquals(errors, Vector("B!A3", "b!A2", "b!A10", "B!B1", "b!C1"))
    // #891: the uncached pass lists them the same way
    val uncached = recalculateUncachedCounted(wb, opts, RecalcCounters.disabled).errors
      .map(e => s"${e.sheet.value}!${e.ref.toA1}")
    assertEquals(uncached, errors)
  }

  test("a seeded RAND draws in grid order whatever the cells' insertion order") {
    val cells = (1 to 12).reverse.map(r => s"A$r" -> fx("RAND()"))
    val wb = Workbook(Vector(sheet("S", cells*)))
    val opts = RecalcOptions(clock = clock, rng = Rng.seeded(3L), iterative = IterativeMode.Off)
    val result = run(opts)(wb)
    val rng = Rng.seeded(3L)
    val expected = (1 to 12).map(_ => rng.nextDouble())
    val actual = (1 to 12).map { r =>
      result.evaluated(SheetName.unsafe("S"))(ARef.from0(0, r - 1)) match
        case CellValue.Number(n) => n.toDouble
        case other => fail(s"A$r: $other")
    }
    assertEquals(actual.toVector, expected.toVector)
  }

  test("parallel waves re-list errors in the sequential order") {
    // 40 columns: row 3 fails reading a computed row-2 cell (depth 1), row 4 fails reading a
    // constant (depth 0); grid order puts row 3 first, the wave fold puts row 4 first
    val cols = 40
    val cells = (0 until cols).flatMap { c =>
      val col = ARef.from0(c, 0).toA1.takeWhile(_.isLetter)
      Vector(
        s"${col}1" -> CellValue.Number(BigDecimal(c)),
        s"${col}2" -> fx(s"${col}1*2"),
        s"${col}3" -> fx(s"NOSUCHFN(${col}2)"),
        s"${col}4" -> fx(s"NOSUCHFN(${col}1)")
      )
    }
    val wb = Workbook(Vector(sheet("P", cells*)))
    val seq = run(RecalcOptions(clock = clock, iterative = IterativeMode.Off))(wb)
    val par = run(RecalcOptions(clock = clock, iterative = IterativeMode.Off, parallelism = 4))(wb)
    assertEquals(par.errors, seq.errors)
    val listed = seq.errors.map(e => QualifiedRef(e.sheet, e.ref))
    assertEquals(listed, listed.sortBy(QualifiedRef.gridKey))
  }
