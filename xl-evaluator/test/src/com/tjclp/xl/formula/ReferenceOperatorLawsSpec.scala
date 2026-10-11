package com.tjclp.xl.formula

import com.tjclp.xl.{*, given}
import com.tjclp.xl.addressing.{ARef, SheetName}
import com.tjclp.xl.cells.{CellError, CellValue}
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.workbooks.Workbook
import munit.ScalaCheckSuite
import org.scalacheck.Gen
import org.scalacheck.Prop.*

/**
 * GH-669: the laws the union and intersection operators obey over static areas.
 *
 *   - Union desugaring: an aggregate over `(a,b,…)` equals the aggregate over `a,b,…`.
 *   - Intersection substitution: `a b` with a non-empty intersection `R` behaves as `R` in every
 *     context — a plain cell, array mode, SUM, arithmetic, INDEX, ROWS, `@`, AREAS; an empty one is
 *     `#NULL!`.
 *   - Dependencies: a union depends on its areas, an intersection on the cells it shares.
 *   - Shifting a union shifts each of its leaves.
 */
@SuppressWarnings(Array("org.wartremover.warts.AsInstanceOf"))
class ReferenceOperatorLawsSpec extends ScalaCheckSuite:

  override def scalaCheckTestParameters =
    super.scalaCheckTestParameters.withMinSuccessfulTests(60)

  /** A1:F10 of numbers, text and blanks (a deterministic mix), so every aggregate has cases. */
  private val sheet: Sheet =
    (for col <- 0 until 6; row <- 0 until 10 yield (col, row)).foldLeft(Sheet("S")) {
      case (acc, (col, row)) =>
        val at = ARef.from0(col, row)
        (col + row) % 7 match
          case 0 => acc
          case 3 => acc.put(at, CellValue.Text(s"t$col$row"))
          case _ => acc.put(at, CellValue.Number(BigDecimal(col * 10 + row + 1)))
    }
  private val book = Workbook(Vector(sheet))

  private val genArea: Gen[CellRange] =
    for
      c1 <- Gen.choose(0, 5)
      c2 <- Gen.choose(0, 5)
      r1 <- Gen.choose(0, 9)
      r2 <- Gen.choose(0, 9)
    yield CellRange(ARef.from0(c1 min c2, r1 min r2), ARef.from0(c1 max c2, r1 max r2))

  private val genAreas: Gen[List[CellRange]] = Gen.choose(2, 4).flatMap(Gen.listOfN(_, genArea))

  /** A plain cell's row: inside the grid's rows or past them (where intersections fail). */
  private val genHost: Gen[ARef] = Gen.choose(0, 14).map(row => ARef.from0(7, row))

  private def plain(at: ARef, formula: String): Either[String, CellValue] =
    val placed = sheet.put(at, CellValue.Formula(formula, None))
    placed.evaluateCell(at, Clock.system, Some(book.put(placed))).left.map(_.message)

  private def array(formula: String): Either[String, Any] =
    FormulaParser
      .parse(s"=$formula")
      .left
      .map(_.toString)
      .flatMap(expr =>
        Evaluator.arrayInstance
          .eval(expr.asInstanceOf[TExpr[Any]], sheet, workbook = Some(book))
          .left
          .map(e => EvalError.toErrorValue(e).fold(e.toString)(_.toExcel))
      )

  private def parsed(formula: String): TExpr[?] =
    FormulaParser.parse(formula).fold(err => fail(s"$formula: $err"), identity)

  private val aggregates =
    List(
      "SUM",
      "COUNT",
      "COUNTA",
      "AVERAGE",
      "MIN",
      "MAX",
      "MEDIAN",
      "STDEV",
      "STDEVP",
      "VAR",
      "VARP"
    )

  property("union desugaring: f((a,b,…)) == f(a,b,…) for every variadic aggregate") {
    forAll(genAreas, genHost, Gen.oneOf(aggregates)) { (areas, host, fn) =>
      val list = areas.map(_.toA1).mkString(",")
      assertEquals(plain(host, s"$fn(($list))"), plain(host, s"$fn($list)"), s"$fn(($list))")
      assertEquals(array(s"$fn(($list))"), array(s"$fn($list)"), s"evala $fn(($list))")
    }
  }

  property("intersection substitution: C[a b] == C[R]; an empty intersection is #NULL!") {
    val contexts: List[String => String] = List(
      x => x,
      x => s"SUM($x)",
      x => s"($x)*2",
      x => s"INDEX($x,1,1)",
      x => s"ROWS($x)",
      x => s"COLUMNS($x)",
      x => s"@($x)",
      x => s"AREAS($x)",
      x => s"COUNTA($x)"
    )
    forAll(genArea, genArea, genHost) { (a, b, host) =>
      val op = s"${a.toA1} ${b.toA1}"
      a.intersect(b) match
        case Some(shared) =>
          contexts.foreach { context =>
            val written = context(op)
            val substituted = context(shared.toA1)
            assertEquals(plain(host, written), plain(host, substituted), s"plain $written")
            assertEquals(array(written), array(substituted), s"evala $written")
          }
        case None =>
          assertEquals(plain(host, op), Right(CellValue.Error(CellError.Null)), op)
          assertEquals(plain(host, s"ISERROR($op)"), Right(CellValue.Bool(true)), op)
      true
    }
  }

  property("dependencies: a union's are its areas', an intersection's the shared cells'") {
    forAll(genAreas, genArea, genArea) { (areas, a, b) =>
      val list = areas.map(_.toA1).mkString(",")
      assertEquals(
        DependencyGraph.extractDependencies(parsed(s"=SUM(($list))")),
        DependencyGraph.extractDependencies(parsed(s"=SUM($list)"))
      )
      val shared = a.intersect(b).fold(Set.empty[ARef])(_.cells.toSet)
      assertEquals(
        DependencyGraph.extractDependencies(parsed(s"=SUM(${a.toA1} ${b.toA1})")),
        shared
      )
    }
  }

  // ===== GH-713: the range operator `:` =====

  private def hull(a: CellRange, b: CellRange): CellRange = a.expand(b.start).expand(b.end)

  /**
   * The ways an operand of `:` spells an area: as written, or as the reference IF, CHOOSE, OFFSET
   * or INDEX returns; a single cell also as one INDEX picks out of the whole grid.
   */
  private def spellings(area: CellRange): List[String] =
    val a = area.toA1
    val cell =
      if area.width == 1 && area.height == 1 then
        List(s"INDEX($$A$$1:$$F$$10,${area.start.row.index1},${area.start.col.index0 + 1})")
      else Nil
    List(a, s"IF(TRUE,$a,0)", s"CHOOSE(1,$a)", s"OFFSET($a,0,0)", s"INDEX($a,0,0)") ++ cell

  private val genSingleOrArea: Gen[CellRange] =
    Gen.frequency(
      3 -> genArea,
      1 -> genArea.map(area => CellRange(area.start, area.start))
    )

  private def genSpelled(area: CellRange): Gen[String] = Gen.oneOf(spellings(area))

  property("bounding substitution: C[a:b] == C[bbox(a,b)] whatever spells a and b") {
    val contexts: List[String => String] = List(
      x => x,
      x => s"SUM($x)",
      x => s"($x)*2",
      x => s"INDEX($x,1,1)",
      x => s"ROWS($x)",
      x => s"COLUMNS($x)",
      x => s"@($x)",
      x => s"AREAS($x)",
      x => s"COUNTA($x)"
    )
    val genCase =
      for
        a <- genSingleOrArea
        b <- genSingleOrArea
        sa <- genSpelled(a)
        sb <- genSpelled(b)
        host <- genHost
      yield (a, b, s"$sa:$sb", host)
    forAll(genCase) { (a, b, op, host) =>
      val box = hull(a, b).toA1
      contexts.foreach { context =>
        val written = context(op)
        val substituted = context(box)
        assertEquals(plain(host, written), plain(host, substituted), s"plain $written")
        assertEquals(array(written), array(substituted), s"evala $written")
      }
      true
    }
  }

  property("semilattice: a:b == b:a, (a:b):c == a:(b:c), a:a == a, a:(a b) == a") {
    val genCase =
      for
        a <- genArea
        b <- genArea
        c <- genArea
        sa <- genSpelled(a)
        sb <- genSpelled(b)
        sc <- genSpelled(c)
      yield (a, b, sa, sb, sc)
    forAll(genCase) { (a, b, sa, sb, sc) =>
      val host = ARef.from0(7, 0)
      def same(x: String, y: String): Unit =
        List((t: String) => s"SUM($t)", (t: String) => s"ROWS($t)*100+COLUMNS($t)").foreach { f =>
          assertEquals(plain(host, f(x)), plain(host, f(y)), s"${f(x)} vs ${f(y)}")
        }
      same(s"$sa:$sb", s"$sb:$sa")
      same(s"($sa:$sb):$sc", s"$sa:($sb:$sc)")
      same(s"$sa:$sa", a.toA1)
      if a.intersect(b).isDefined then same(s"$sa:(${a.toA1} ${b.toA1})", a.toA1)
      true
    }
  }

  property("edges: a computed range's edges hold every cell it reads, or it is dynamic") {
    val genCase =
      for
        a <- genSingleOrArea
        b <- genSingleOrArea
        sa <- genSpelled(a)
        sb <- genSpelled(b)
      yield (a, b, s"SUM($sa:$sb)")
    forAll(genCase) { (a, b, formula) =>
      val expr = parsed(s"=$formula")
      val read = hull(a, b).cells.toSet
      assert(
        DependencyGraph.containsDynamicReference(expr) ||
          read.subsetOf(DependencyGraph.extractDependencies(expr)),
        s"$formula reads $read, edges ${DependencyGraph.extractDependencies(expr)}"
      )
      // the pre-filter admits every formula the classifier flags
      if DependencyGraph.containsDynamicReference(expr) then
        assert(functions.ReferenceOperators.mayContainComputedRange(formula), formula)
      true
    }
  }

  property("static operands: a range's edges are exactly its bounding range") {
    forAll(genArea, genArea) { (a, b) =>
      assertEquals(
        DependencyGraph.extractDependencies(parsed(s"=SUM(${a.toA1}:${b.toA1})")),
        hull(a, b).cells.toSet
      )
    }
  }

  property("targeted recalculation after any one edit equals full recalculation") {
    val genCase =
      for
        formulas <- Gen.listOfN(
          4,
          for
            a <- genSingleOrArea
            b <- genSingleOrArea
            sa <- genSpelled(a)
            sb <- genSpelled(b)
          yield s"SUM($sa:$sb)"
        )
        col <- Gen.choose(0, 5)
        row <- Gen.choose(0, 9)
        value <- Gen.choose(-50, 50)
      yield (formulas, ARef.from0(col, row), value)
    forAll(genCase) { (formulas, edited, value) =>
      val withFormulas = formulas.zipWithIndex.foldLeft(sheet) { case (acc, (f, i)) =>
        acc.put(ARef.from0(7, i), CellValue.Formula(f, None))
      }
      val total = withFormulas.put(ARef.from0(8, 0), CellValue.Formula("SUM(H1:H4)", None))
      val settled = Workbook(Vector(total)).recalculate().workbook
      val name = SheetName.unsafe("S")
      val edit = settled(name)
        .fold(e => fail(e.message), identity)
        .put(edited, CellValue.Number(BigDecimal(value)))
      val book = settled.put(edit)
      val targeted = book.recalculateAfterEdit(name, Set(edited), RecalcOptions()).workbook
      val full = book.recalculate().workbook
      def caches(wb: Workbook): Map[ARef, Option[CellValue]] =
        wb(name).fold(e => fail(e.message), identity).cells.collect {
          case (at, cell) if cell.value.isInstanceOf[CellValue.Formula] =>
            at -> (cell.value match
              case CellValue.Formula(_, cached, _) => cached
              case _ => None)
        }
      assertEquals(
        caches(targeted),
        caches(full),
        s"${formulas.mkString("; ")} after ${edited.toA1}"
      )
      true
    }
  }

  test("drift: referenceFunctions is every Referencing spec, the selectors and the operators") {
    val referencing = functions.FunctionRegistry.all.collect {
      case spec: functions.FunctionSpec.Referencing[?, ?] => spec.name
    }.toSet
    assertEquals(
      Evaluator.referenceFunctions,
      referencing ++ Set("IF", "IFS", "CHOOSE", "SWITCH") ++
        Set(functions.ReferenceOperators.IntersectionName, functions.ReferenceOperators.RangeName)
    )
    // neither operator is registered, nor reachable by name
    List(
      functions.ReferenceOperators.UnionName,
      functions.ReferenceOperators.IntersectionName,
      functions.ReferenceOperators.RangeName
    ).foreach(op => assertEquals(functions.FunctionRegistry.lookup(op), None, op))
  }

  property("shifting a union shifts each leaf; an intersection shifts both operands") {
    forAll(genAreas, Gen.choose(0, 4), Gen.choose(0, 4)) { (areas, dc, dr) =>
      def shifted(formula: String): String =
        FormulaPrinter.print(FormulaShifter.shift(parsed(formula), dc, dr))
      val leaves = areas.map(area => shifted(s"=${area.toA1}").stripPrefix("="))
      assertEquals(
        shifted(s"=SUM((${areas.map(_.toA1).mkString(",")}))"),
        s"=SUM((${leaves.mkString(",")}))"
      )
      val (first, second) = (areas(0), areas(1))
      assertEquals(
        shifted(s"=${first.toA1} ${second.toA1}"),
        s"=${leaves(0)} ${leaves(1)}"
      )
    }
  }

  property("GH-713: shifting a range shifts both operands, a computed one inside its call") {
    forAll(genArea, genArea, Gen.choose(0, 4), Gen.choose(0, 4)) { (a, b, dc, dr) =>
      def shifted(formula: String): String =
        FormulaPrinter.printFileForm(FormulaShifter.shift(parsed(formula), dc, dr))
      val (la, lb) = (shifted(s"=${a.toA1}"), shifted(s"=${b.toA1}"))
      assertEquals(shifted(s"=SUM(${a.toA1}:INDEX(${b.toA1},1,1))"), s"SUM($la:INDEX($lb,1,1))")
      assertEquals(shifted(s"=SUM(INDEX(${a.toA1},1,1):${b.toA1})"), s"SUM(INDEX($la,1,1):$lb)")
    }
  }
