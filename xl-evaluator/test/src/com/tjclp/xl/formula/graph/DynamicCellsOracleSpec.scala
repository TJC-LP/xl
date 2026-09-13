package com.tjclp.xl.formula.graph

import com.tjclp.xl.addressing.{ARef, SheetName}
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.formula.ast.TExpr
import com.tjclp.xl.formula.eval.Evaluator
import com.tjclp.xl.formula.graph.DependencyGraph.QualifiedRef
import com.tjclp.xl.formula.parser.FormulaParser
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.workbooks.{DefinedName, Workbook}
import munit.ScalaCheckSuite
import org.scalacheck.Gen
import org.scalacheck.Prop.*

/**
 * GH-537: `DependencyGraph.dynamicCells(workbook)` classifies sheet-independent names ONCE (not
 * once per sheet), dedups case-variant names by the resolution key, and parses each definition text
 * once. None of that may move the verdict: this spec keeps the pre-GH-537 algorithm — every (sheet
 * × name) pair classified independently, definitions parsed on every visit, no dedup — as a
 * brute-force oracle and checks the two agree on random small workbooks with workbook- and
 * sheet-scoped names, case-variant spellings, nested and sheet-qualified name references, and
 * sheets that do not exist.
 */
class DynamicCellsOracleSpec extends ScalaCheckSuite:

  private type Key = (SheetName, SheetName, String)

  /** The pre-GH-537 semantics, memo-free. */
  private def oracle(wb: Workbook): Set[QualifiedRef] =
    val positions: Map[SheetName, Int] =
      wb.sheets.zipWithIndex.reverseIterator.map((s, i) => s.name -> i).toMap

    def exprIsDynamic(expr: TExpr[?], current: SheetName, visiting: Set[Key]): Boolean =
      DependencyGraph.referencesMatching(
        expr,
        (name, scope) => nameIsDynamic(name, scope.getOrElse(current), current, visiting),
        includeDynamicCalls = true
      )

    def nameIsDynamic(
      name: String,
      lookupFrom: SheetName,
      fallback: SheetName,
      visiting: Set[Key]
    ): Boolean =
      val key = (lookupFrom, fallback, name.toUpperCase(java.util.Locale.ROOT))
      if visiting.contains(key) then false
      else
        (for
          dn <- Evaluator.lookupDefinedNameAt(wb, positions.get(lookupFrom), name)
          target <- FormulaParser.parse(dn.formula).toOption
        yield
          val defining = Evaluator.definedNameScope(wb, dn).map(_.name).getOrElse(fallback)
          exprIsDynamic(target, defining, visiting + key)
        ).getOrElse(false)

    wb.sheets.iterator.flatMap { sheet =>
      sheet.cells.iterator.flatMap { case (ref, cell) =>
        cell.value match
          case CellValue.Formula(text, _, _) =>
            FormulaParser.parse(text) match
              case Right(expr) if exprIsDynamic(expr, sheet.name, Set.empty) =>
                Some(QualifiedRef(sheet.name, ref))
              case _ => None
          case _ => None
      }
    }.toSet

  private val sheetNames = Vector("S1", "S2", "S3").map(SheetName.unsafe)

  // Definitions reference only LOWER-numbered names, so name chains never cycle: on a cyclic
  // chain the production memo's verdict for the inner name depends on which outer name was
  // classified first (a pre-existing quirk this spec does not legislate).
  private val leafDefs: Gen[String] =
    Gen.oneOf("42", "A1*2", "INDIRECT(\"A1\")", "OFFSET(A1,0,0)", "SUM(A1:A3)", "=1+")
  private def definitionFor(level: Int): Gen[String] = level match
    case 1 => leafDefs
    case 2 => Gen.oneOf(leafDefs, Gen.oneOf("n1", "N1*2", "S1!n1", "SUM(n1)", "S3!N1"))
    case _ => Gen.oneOf(leafDefs, Gen.oneOf("n1", "n2", "N2+n1", "S2!n2", "SUM(n2)", "S3!n1"))

  private val nameGen: Gen[DefinedName] =
    for
      level <- Gen.choose(1, 3)
      spelling <- Gen.oneOf(s"n$level", s"N$level")
      scope <- Gen.option(Gen.choose(0, 2))
      definition <- definitionFor(level)
    yield DefinedName(spelling, definition, scope)

  private val cellFormulas: Vector[String] = Vector(
    "=n1",
    "=N2*2",
    "=n3",
    "=SUM(n2)",
    "=S2!n1",
    "=S1!N3+1",
    "=INDIRECT(\"B1\")",
    "=A1+1",
    "=n1+n3",
    "=OFFSET(n2,0,0)"
  )

  private val workbookGen: Gen[Workbook] =
    for
      sheetCount <- Gen.choose(1, 3)
      nameCount <- Gen.choose(0, 6)
      names <- Gen.listOfN(nameCount, nameGen)
      cells <- Gen.listOf(
        for
          sheet <- Gen.choose(0, sheetCount - 1)
          col <- Gen.choose(0, 4)
          text <- Gen.oneOf(cellFormulas)
        yield (sheet, col, text)
      )
    yield
      val sheets = (0 until sheetCount).map { i =>
        val base = Sheet(sheetNames(i))
          .put(ARef.from0(0, 0), CellValue.Number(BigDecimal(1)))
          .put(ARef.from0(1, 0), CellValue.Number(BigDecimal(2)))
        cells.filter(_._1 == i).foldLeft(base) { case (s, (_, col, text)) =>
          s.put(ARef.from0(col, 1), CellValue.Formula(text, None))
        }
      }
      val wb = Workbook(sheets.toVector)
      wb.copy(metadata = wb.metadata.copy(definedNames = names.toVector))

  property("GH-537: dynamicCells(workbook) equals the per-sheet brute-force oracle") {
    forAll(workbookGen) { wb =>
      val got = DependencyGraph.dynamicCells(wb)
      val want = oracle(wb)
      (got == want) :|
        s"names=${wb.metadata.definedNames.map(d => s"${d.name}:${d.formula}@${d.localSheetId}")} got=$got want=$want"
    }
  }
