package com.tjclp.xl.formula.graph

import com.tjclp.xl.addressing.{ARef, SheetName}
import com.tjclp.xl.cells.{Cell, CellValue}
import com.tjclp.xl.formula.graph.DependencyGraph.QualifiedRef
import com.tjclp.xl.formula.graph.QualifiedGraph.{DeclaredRange, Node}
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.styles.units.StyleId
import com.tjclp.xl.workbooks.Workbook
import munit.ScalaCheckSuite
import org.scalacheck.Gen
import org.scalacheck.Prop.*

/**
 * The declared precedent view of [[QualifiedGraph]] — each range one node — is a regrouping of the
 * expanded view, never a different answer. On random two-sheet grids (values, style-only blanks,
 * formulas of one to three terms drawn from points, cross-sheet points, rectangles, whole columns
 * and rows, 1x1 range slots, anchored ranges and a sheet the book lacks, cycles allowed), for every
 * start cell and every depth:
 *
 *   - occupied: `occupiedIn(r)` is exactly the value-holding cells of `r` on its sheet, and every
 *     declared range has an entry;
 *   - single hop: the cells a start's declared nodes cover are its expanded precedents, plus only
 *     the empty cells a 1x1 range slot names (a cell node, occupied or not);
 *   - coverage: layer k's value-holding coverage (cell nodes and the occupied cells of range nodes,
 *     minus what earlier layers covered) equals the expanded layer k's value-holding cells, and
 *     every expanded cell is covered no later than its expanded layer;
 *   - uniqueness: no node repeats, no cell node is the start or covered by an earlier layer;
 *   - order and prefix: every layer is sorted by [[QualifiedGraph.nodeOrder]], and a bounded walk
 *     is a prefix of the unbounded one.
 */
class PrecedentNodesLawSpec extends ScalaCheckSuite:

  override def scalaCheckTestParameters =
    super.scalaCheckTestParameters.withMinSuccessfulTests(300)

  private val sheetNames = Vector(SheetName.unsafe("S1"), SheetName.unsafe("Q 2"))
  private val columns = Vector("A", "B", "C", "D")
  private val grid: Vector[ARef] =
    for
      col <- columns.indices.toVector
      row <- (0 until 5).toVector
    yield ARef.from0(col, row)

  private val gridRef: Gen[String] =
    for
      col <- Gen.oneOf(columns)
      row <- Gen.choose(1, 5)
    yield s"$col$row"
  private val rect: Gen[String] =
    for
      from <- gridRef
      to <- gridRef
    yield s"$from:$to"

  private val term: Gen[String] = Gen.oneOf(
    gridRef,
    gridRef.map(r => s"'Q 2'!$r"),
    gridRef.map(r => s"S1!$r"),
    rect.map(r => s"SUM($r)"),
    rect.map(r => s"SUM('Q 2'!$r)"),
    Gen.const("SUM(B:B)"),
    Gen.const("SUM('Q 2'!A:A)"),
    Gen.const("SUM(2:2)"),
    Gen.const("COUNTIF(C1,\">0\")"),
    Gen.const("SUM($A$1:$A$3)"),
    Gen.const("SUM(s1!A1:A3)"),
    Gen.const("SUM(Nope!A1:A2)")
  )

  private val formulaText: Gen[String] =
    Gen.choose(1, 3).flatMap(n => Gen.listOfN(n, term)).map(_.mkString("+"))

  private def cellGen(ref: ARef): Gen[Option[Cell]] = Gen.frequency(
    3 -> Gen.const(None),
    3 -> Gen.choose(1, 9).map(n => Some(Cell(ref, CellValue.Number(BigDecimal(n))))),
    1 -> Gen.const(Some(Cell(ref, CellValue.Empty, Some(StyleId(1))))),
    3 -> formulaText.map(f => Some(Cell(ref, CellValue.Formula(f, None))))
  )

  private def sheetGen(name: SheetName): Gen[Sheet] =
    grid.foldLeft(Gen.const(Sheet(name))) { (acc, ref) =>
      for
        sheet <- acc
        cell <- cellGen(ref)
      yield cell.fold(sheet)(sheet.put)
    }

  private val workbookGen: Gen[Workbook] =
    Gen.sequence[Vector[Sheet], Sheet](sheetNames.map(sheetGen)).map(Workbook(_))

  private def holds(value: CellValue): Boolean = value match
    case CellValue.Empty => false
    case _ => true

  private def holdsValue(wb: Workbook, c: QualifiedRef): Boolean =
    wb.sheets.find(_.name == c.sheet).flatMap(_.cells.get(c.ref)).exists(cell => holds(cell.value))

  private def cover(graph: QualifiedGraph, nodes: Iterable[Node]): Set[QualifiedRef] =
    nodes.iterator.flatMap {
      case Node.Cell(c) => Iterator.single(c)
      case Node.Range(r) => graph.occupiedIn(r).iterator
    }.toSet

  /** Each layer's coverage minus what the start and the earlier layers covered. */
  private def coverage(
    graph: QualifiedGraph,
    start: QualifiedRef,
    layers: Vector[Vector[Node]]
  ): Vector[Set[QualifiedRef]] =
    layers
      .foldLeft((Set(start), Vector.empty[Set[QualifiedRef]])) { case ((seen, acc), layer) =>
        val fresh = cover(graph, layer) -- seen
        (seen ++ fresh, acc :+ fresh)
      }
      ._2

  private val starts: Vector[QualifiedRef] =
    for
      sheet <- sheetNames
      ref <- grid
    yield QualifiedRef(sheet, ref)

  property("occupiedIn is exactly the value-holding cells of every declared range") {
    forAll(workbookGen) { wb =>
      val graph = QualifiedGraph.of(wb)
      val declaredRanges = graph.rangeReaders.iterator.flatMap { (sheet, entries) =>
        entries.iterator.map(e => DeclaredRange(sheet, e.range))
      }.toSet
      assertEquals(graph.rangeCells.keySet, declaredRanges)
      declaredRanges.foreach { r =>
        val brute = wb.sheets
          .find(_.name == r.sheet)
          .fold(Set.empty[QualifiedRef]) { sheet =>
            sheet.cells.iterator.collect {
              case (ref, cell) if r.range.contains(ref) && holds(cell.value) =>
                QualifiedRef(sheet.name, ref)
            }.toSet
          }
        assertEquals(graph.occupiedIn(r), brute, r.toString)
      }
    }
  }

  property("single hop: the declared nodes cover the expanded precedents, plus empty 1x1 slots") {
    forAll(workbookGen) { wb =>
      val graph = QualifiedGraph.of(wb)
      starts.foreach { start =>
        val nodes = graph.declaredPrecedentsOf(start)
        val covered = cover(graph, nodes)
        val expanded = graph.precedentsOf(start)
        assert(expanded.subsetOf(covered), s"$start: ${expanded -- covered} not covered")
        val extra = covered -- expanded
        assert(extra.forall(c => !holdsValue(wb, c)), s"$start: extra value-holding $extra")
        assertEquals(nodes, nodes.distinct.sorted(using QualifiedGraph.nodeOrder), start.toString)
      }
    }
  }

  property("coverage: each declared layer holds the expanded layer's value-holding cells") {
    forAll(workbookGen) { wb =>
      val graph = QualifiedGraph.of(wb)
      for
        start <- starts
        depth <- 0 to 4
      do
        val layers = graph.declaredPrecedents(start, depth)
        val covered = coverage(graph, start, layers)
        val expanded = graph.precedents(start, depth).map(_.toSet)
        (0 until math.max(covered.size, expanded.size)).foreach { k =>
          val declaredValues = covered.lift(k).getOrElse(Set.empty).filter(holdsValue(wb, _))
          val expandedValues = expanded.lift(k).getOrElse(Set.empty).filter(holdsValue(wb, _))
          assertEquals(declaredValues, expandedValues, s"$start depth $depth layer ${k + 1}")
        }
        expanded.zipWithIndex.foreach { (layer, k) =>
          val upTo = covered.take(k + 1).foldLeft(Set.empty[QualifiedRef])(_ ++ _)
          assert(layer.subsetOf(upTo), s"$start depth $depth: ${layer -- upTo} missing by ${k + 1}")
        }
    }
  }

  property(
    "uniqueness, order and prefix: every node once, each layer sorted, bounds are prefixes"
  ) {
    forAll(workbookGen) { wb =>
      val graph = QualifiedGraph.of(wb)
      starts.foreach { start =>
        val all = graph.declaredPrecedents(start, 0)
        val flat = all.flatten
        assertEquals(flat.distinct.size, flat.size, s"$start repeats a node")
        assert(all.forall(_.nonEmpty), s"$start has an empty layer")
        all.foreach(layer => assertEquals(layer, layer.sorted(using QualifiedGraph.nodeOrder)))
        all.foldLeft(Set(start)) { (seen, layer) =>
          layer.foreach {
            case Node.Cell(c) => assert(!seen.contains(c), s"$start relists $c")
            case Node.Range(_) => ()
          }
          seen ++ cover(graph, layer)
        }
        (1 to 4).foreach { d =>
          assertEquals(graph.declaredPrecedents(start, d), all.take(d), s"$start depth $d")
        }
      }
    }
  }
