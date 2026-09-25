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
