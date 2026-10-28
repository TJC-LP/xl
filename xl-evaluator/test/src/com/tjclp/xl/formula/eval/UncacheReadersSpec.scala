package com.tjclp.xl.formula.eval

import munit.{FunSuite, ScalaCheckSuite}
import org.scalacheck.Gen
import org.scalacheck.Prop.forAllNoShrink

import com.tjclp.xl.{*, given}
import com.tjclp.xl.addressing.{ARef, CellRange}
import com.tjclp.xl.cells.CellValue

/**
 * GH-715: `--with`'s stale-reader search (`FormulaEvaluation.uncacheReaders`) indexes formulas by
 * the ranges they read, across sheets, and drops a reader from its candidate sets once it is stale.
 * On small random books it must find exactly what the naive search finds: expand every range, then
 * take the transitive closure from the overridden cells.
 */
class UncacheReadersSpec extends FunSuite with ScalaCheckSuite:

  override def scalaCheckTestParameters =
    super.scalaCheckTestParameters.withMinSuccessfulTests(300)

  private val names = Vector("S0", "S1", "S2")
  private val cols = 4
  private val rows = 6

  /** One read of a generated formula: on `sheet` (None: the formula's own, unqualified). */
  private final case class Read(sheet: Option[Int], range: CellRange)

  private val genRange: Gen[CellRange] =
    for
      c1 <- Gen.choose(0, cols - 1)
      c2 <- Gen.choose(0, cols - 1)
      r1 <- Gen.choose(0, rows - 1)
      r2 <- Gen.choose(0, rows - 1)
    yield CellRange(ARef.from0(c1 min c2, r1 min r2), ARef.from0(c1 max c2, r1 max r2))

  private val genRead: Gen[Read] =
    for
      sheet <- Gen.frequency(3 -> Gen.const(None), 2 -> Gen.choose(0, names.size - 1).map(Some(_)))
      range <- genRange
    yield Read(sheet, range)

  /** A cell: a cached SUM over one to three reads, a number, or nothing. */
  private val genCell: Gen[Option[Either[Int, List[Read]]]] =
    Gen.frequency(
      4 -> Gen.choose(1, 3).flatMap(Gen.listOfN(_, genRead)).map(r => Some(Right(r))),
      3 -> Gen.choose(1, 9).map(n => Some(Left(n))),
      3 -> Gen.const(None)
    )

  /** Per sheet, per cell: what it holds. */
  private type Model = Vector[Map[ARef, Either[Int, List[Read]]]]

  private val genModel: Gen[Model] =
    Gen
      .listOfN(
        names.size,
        Gen
          .listOfN(cols * rows, genCell)
          .map(_.zipWithIndex.flatMap((cell, i) => cell.map(ARef.from0(i % cols, i / cols) -> _)))
          .map(_.toMap)
      )
      .map(_.toVector)

  private val genChanged: Gen[Set[ARef]] =
    Gen
      .choose(1, 3)
      .flatMap(Gen.listOfN(_, Gen.zip(Gen.choose(0, cols - 1), Gen.choose(0, rows - 1))))
      .map(_.map((c, r) => ARef.from0(c, r)).toSet)

  private def render(read: Read): String =
    read.sheet.fold("")(s => s"${names(s)}!") + read.range.toA1

  /** The book the model describes, every formula cached, with the overridden cells set on S0. */
  private def book(model: Model, changed: Set[ARef]): Workbook =
    val sheets = model.zipWithIndex.map { (cells, s) =>
      val built = cells.foldLeft(Sheet(SheetName.unsafe(names(s)))) {
        case (sheet, (ref, Left(n))) => sheet.put(ref, CellValue.Number(BigDecimal(n)))
        case (sheet, (ref, Right(reads))) =>
          val text = reads.map(render).mkString("=SUM(", ",", ")")
          sheet.put(ref, CellValue.Formula(text, Some(CellValue.Number(BigDecimal(0)))))
      }
      if s == 0 then changed.foldLeft(built)((sh, ref) => sh.put(ref, CellValue.Number(1)))
      else built
    }
    Workbook(sheets)

  /** The naive answer: every range expanded, the stale formulas' closure from `changed`. */
  private def oracle(model: Model, changed: Set[ARef]): Set[(Int, ARef)] =
    val formulas: Vector[((Int, ARef), Set[(Int, ARef)])] =
      for
        (cells, s) <- model.zipWithIndex
        (ref, content) <- cells.toVector
        reads <- content.toOption.toList
        if !(s == 0 && changed.contains(ref))
      yield
        val cellsRead = for
          read <- reads
          cell <- read.range.cells.toList
        yield (read.sheet.getOrElse(s), cell)
        ((s, ref), cellsRead.toSet)
    val sources0: Set[(Int, ARef)] = changed.map(ref => (0, ref))
    @scala.annotation.tailrec
    def closure(stale: Set[(Int, ARef)]): Set[(Int, ARef)] =
      val sources = stale ++ sources0
      val next = formulas.collect {
        case (node, read) if read.exists(sources.contains) => node
      }.toSet
      if next == stale then stale else closure(next)
    closure(Set.empty)

  /** The formulas whose cache `uncacheReaders` dropped. */
  private def uncached(before: Workbook, after: Workbook): Set[(Int, ARef)] =
    before.sheets.zipWithIndex.flatMap { (sheet, s) =>
      sheet.cells.keys.collect {
        case ref if (sheet(ref).value, after.sheets(s)(ref).value) match
              case (CellValue.Formula(_, Some(_), _), CellValue.Formula(_, None, _)) => true
              case _ => false
            =>
          (s, ref)
      }
    }.toSet

  property("the stale set equals the naive expand-everything closure, across sheets") {
    forAllNoShrink(genModel, genChanged) { (model, changed) =>
      val before = book(model, changed)
      val after = FormulaEvaluation.uncacheReaders(before, SheetName.unsafe("S0"), changed)
      def named(cells: Set[(Int, ARef)]): List[String] =
        cells.toList.map((s, ref) => s"${names(s)}!${ref.toA1}").sorted
      val formulas = before.sheets.flatMap(sheet =>
        sheet.cells.toList.flatMap((ref, cell) =>
          cell.value match
            case CellValue.Formula(text, _, _) => List(s"${sheet.name.value}!${ref.toA1} $text")
            case _ => Nil
        )
      )
      assertEquals(
        named(uncached(before, after)),
        named(oracle(model, changed)),
        s"overridden on S0: ${changed.map(_.toA1)}\n${formulas.mkString("\n")}"
      )
    }
  }
