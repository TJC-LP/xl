package com.tjclp.xl.formula.graph

import com.tjclp.xl.addressing.{ARef, SheetName}
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.formula.ast.TExpr
import com.tjclp.xl.formula.eval.Evaluator
import com.tjclp.xl.formula.functions.FunctionRegistry
import com.tjclp.xl.formula.graph.DependencyGraph.QualifiedRef
import com.tjclp.xl.formula.parser.FormulaParser
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.workbooks.{DefinedName, Workbook}
import munit.ScalaCheckSuite
import org.scalacheck.Gen
import org.scalacheck.Prop.*
import org.scalacheck.rng.Seed

/**
 * GH-537: `DependencyGraph.dynamicCells(workbook)` classifies sheet-independent names ONCE (not
 * once per sheet), dedups case-variant names by the resolution key, and parses each definition text
 * once. None of that may move the verdict: this spec keeps the pre-GH-537 algorithm — every (sheet
 * × name) pair classified independently, definitions parsed on every visit, no dedup, and every
 * declared spelling's `toUpperCase` form as a substring pre-filter token — as a brute-force oracle
 * and checks the two agree on random small workbooks with workbook- and sheet-scoped names,
 * case-variant spellings, nested and sheet-qualified name references, and sheets that do not exist.
 *
 * Spellings are name-shaped (`nm_1`, never `n1`, which the parser reads as cell N1), so every
 * generated name reference really is a `NameRef`. A Kelvin-sign spelling (U+212A) sits beside its
 * ASCII twin `k`: it upper-cases to ITSELF yet resolves as `k` under `equalsIgnoreCase`, the one
 * place where the token relation and the resolution relation part.
 *
 * Two shapes separate the algorithms and are too rare to leave to chance, so dedicated generators
 * mix them in beside the random workbooks: a workbook-scoped name nesting an unqualified name that
 * a later sheet shadows with a differently-classified local (the sheet-independent shortcut must
 * not take it), and the Kelvin/`k` twins declared in either order with a reader of each spelling
 * (every declared spelling must contribute its own pre-filter token).
 */
class DynamicCellsOracleSpec extends ScalaCheckSuite:

  // A fixed seed over a large sample: the shapes that separate the two algorithms (a Kelvin-sign
  // spelling declared before its ASCII twin, a nested name shadowed on a later sheet) are rare, and
  // an equivalence claim must not flake.
  override def scalaCheckTestParameters =
    super.scalaCheckTestParameters.withMinSuccessfulTests(400).withInitialSeed(Seed(0x537L))

  private type Key = (SheetName, SheetName, String)

  private def upper(text: String): String = text.toUpperCase(java.util.Locale.ROOT)

  /** The pre-GH-537 semantics, memo-free, substring pre-filter included. */
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
      val key = (lookupFrom, fallback, upper(name))
      if visiting.contains(key) then false
      else
        (for
          dn <- Evaluator.lookupDefinedNameAt(wb, positions.get(lookupFrom), name)
          target <- FormulaParser.parse(dn.formula).toOption
        yield
          val defining = Evaluator.definedNameScope(wb, dn).map(_.name).getOrElse(fallback)
          exprIsDynamic(target, defining, visiting + key)
        ).getOrElse(false)

    // Every declared spelling classified from every sheet contributes its own token.
    val tokens: Set[String] =
      FunctionRegistry.dynamicFunctionNames.toSet ++
        wb.sheets.iterator.flatMap { sheet =>
          wb.metadata.definedNames.iterator
            .map(_.name)
            .filter(name => nameIsDynamic(name, sheet.name, sheet.name, Set.empty))
            .map(upper)
        }

    wb.sheets.iterator.flatMap { sheet =>
      sheet.cells.iterator.flatMap { case (ref, cell) =>
        cell.value match
          case CellValue.Formula(text, _, _) if tokens.exists(upper(text).contains) =>
            FormulaParser.parse(text) match
              case Right(expr) if exprIsDynamic(expr, sheet.name, Set.empty) =>
                Some(QualifiedRef(sheet.name, ref))
              case _ => None
          case _ => None
      }
    }.toSet

  private val sheetNames = Vector("S1", "S2", "S3").map(SheetName.unsafe)

  private val kelvin = "K"

  // Level 0 is the Kelvin-sign pair; levels 1–3 are ASCII case variants of a name-shaped spelling.
  private def spellings(level: Int): Gen[String] = level match
    case 0 => Gen.oneOf(kelvin, "k")
    case n => Gen.oneOf(s"nm_$n", s"NM_$n")

  // Definitions reference only LOWER levels, so name chains never cycle: on a cyclic chain the
  // production memo's verdict for the inner name depends on which outer name was classified first
  // (a pre-existing quirk this spec does not legislate).
  private val leafDefs: Gen[String] =
    Gen.oneOf("42", "A1*2", "INDIRECT(\"A1\")", "OFFSET(A1,0,0)", "SUM(A1:A3)", "=1+")
  private def definitionFor(level: Int): Gen[String] = level match
    case 0 => leafDefs
    case 1 => Gen.frequency(1 -> leafDefs, 2 -> Gen.oneOf("k", s"$kelvin*2", "S1!k", "SUM(k)"))
    case 2 =>
      Gen.frequency(
        1 -> leafDefs,
        2 -> Gen.oneOf("nm_1", "NM_1*2", "S1!nm_1", "SUM(nm_1)", "S3!NM_1", "k+nm_1")
      )
    case _ =>
      Gen.frequency(
        1 -> leafDefs,
        2 -> Gen.oneOf("nm_1", "nm_2", "NM_2+nm_1", "S2!nm_2", "SUM(nm_2)", "S3!nm_1", kelvin)
      )

  private val nameGen: Gen[DefinedName] =
    for
      level <- Gen.choose(0, 3)
      spelling <- spellings(level)
      scope <- Gen.option(Gen.choose(0, 2))
      definition <- definitionFor(level)
    yield DefinedName(spelling, definition, scope)

  // A reader for every spelling, the Kelvin sign's and `k`'s included.
  private val cellFormulas: Vector[String] = Vector(
    "=nm_1",
    "=NM_2*2",
    "=nm_3",
    "=SUM(nm_2)",
    "=S2!nm_1",
    "=S1!NM_3+1",
    "=INDIRECT(\"B1\")",
    "=A1+1",
    "=nm_1+nm_3",
    "=OFFSET(nm_2,0,0)",
    "=k*2",
    s"=$kelvin+1",
    "=S1!k",
    "=SUM(k)"
  )

  private def assemble(
    sheetCount: Int,
    names: Vector[DefinedName],
    cells: List[(Int, Int, String)]
  ): Workbook =
    val sheets = (0 until sheetCount).map { i =>
      val base = Sheet(sheetNames(i))
        .put(ARef.from0(0, 0), CellValue.Number(BigDecimal(1)))
        .put(ARef.from0(1, 0), CellValue.Number(BigDecimal(2)))
      cells.filter(_._1 == i).foldLeft(base) { case (s, (_, col, text)) =>
        s.put(ARef.from0(col, 1), CellValue.Formula(text, None))
      }
    }
    val wb = Workbook(sheets.toVector)
    wb.copy(metadata = wb.metadata.copy(definedNames = names))

  private def cellsGen(sheetCount: Int): Gen[List[(Int, Int, String)]] =
    Gen.listOf(
      for
        sheet <- Gen.choose(0, sheetCount - 1)
        col <- Gen.choose(0, 4)
        text <- Gen.oneOf(cellFormulas)
      yield (sheet, col, text)
    )

  /** Declaration order decides which entry a lookup returns, so it is randomised too. */
  private def shuffled[A](xs: Vector[A]): Gen[Vector[A]] =
    Gen.long.map(seed => new scala.util.Random(seed).shuffle(xs))

  private val randomWorkbook: Gen[Workbook] =
    for
      sheetCount <- Gen.choose(1, 3)
      nameCount <- Gen.choose(0, 8)
      names <- Gen.listOfN(nameCount, nameGen)
      cells <- cellsGen(sheetCount)
    yield assemble(sheetCount, names.toVector, cells)

  // `nm_2` (workbook-scoped, nesting an unqualified `nm_1`) is read on every sheet; `nm_1` has a
  // local on a later sheet and maybe a global, so `nm_2`'s verdict may differ between the first
  // sheet and the shadowing one.
  private val shadowedChain: Gen[Workbook] =
    for
      sheetCount <- Gen.choose(2, 3)
      shadowSheet <- Gen.choose(1, sheetCount - 1)
      outerSpelling <- Gen.oneOf("nm_2", "NM_2")
      outerDef <- Gen.oneOf("nm_1", "NM_1*2", "SUM(nm_1)", "k+nm_1")
      localSpelling <- Gen.oneOf("nm_1", "NM_1")
      localDef <- Gen.oneOf("INDIRECT(\"A1\")", "OFFSET(A1,0,0)", "42")
      globalName <- Gen.option(
        for
          spelling <- Gen.oneOf("nm_1", "NM_1")
          definition <- Gen.oneOf("42", "A1*2", "INDIRECT(\"A1\")")
        yield DefinedName(spelling, definition)
      )
      extras <- Gen.listOfN(2, nameGen).map(_.filterNot(_.name.equalsIgnoreCase("nm_2")))
      names <- shuffled(
        Vector(
          DefinedName(outerSpelling, outerDef),
          DefinedName(localSpelling, localDef, Some(shadowSheet))
        ) ++ globalName ++ extras
      )
      cells <- cellsGen(sheetCount)
    yield
      val outerReaders =
        (0 until sheetCount).toList.flatMap(i => List((i, 2, "=NM_2*2"), (i, 3, "=SUM(nm_2)")))
      assemble(sheetCount, names, outerReaders ++ cells)

  // The Kelvin sign and `k`, declared in either order, with a reader of each spelling everywhere.
  private val kelvinTwins: Gen[Workbook] =
    for
      sheetCount <- Gen.choose(1, 3)
      kelvinDef <- Gen.oneOf("INDIRECT(\"A1\")", "OFFSET(A1,0,0)", "42")
      kDef <- Gen.oneOf("INDIRECT(\"A1\")", "42")
      kelvinScope <- Gen.option(Gen.choose(0, sheetCount - 1))
      kScope <- Gen.option(Gen.choose(0, sheetCount - 1))
      extras <- Gen.listOfN(2, nameGen)
      names <- shuffled(
        Vector(DefinedName(kelvin, kelvinDef, kelvinScope), DefinedName("k", kDef, kScope)) ++
          extras
      )
      cells <- cellsGen(sheetCount)
    yield
      val twinReaders = (0 until sheetCount).toList.flatMap(i =>
        List((i, 2, "=k*2"), (i, 3, s"=$kelvin+1"), (i, 4, "=S1!k"))
      )
      assemble(sheetCount, names, twinReaders ++ cells)

  private val workbookGen: Gen[Workbook] =
    Gen.frequency(3 -> randomWorkbook, 1 -> shadowedChain, 1 -> kelvinTwins)

  property("GH-537: dynamicCells(workbook) equals the per-sheet brute-force oracle") {
    forAll(workbookGen) { wb =>
      val got = DependencyGraph.dynamicCells(wb)
      val want = oracle(wb)
      (got == want) :|
        s"names=${wb.metadata.definedNames.map(d => s"${d.name}:${d.formula}@${d.localSheetId}")} got=$got want=$want"
    }
  }

  test("GH-537: the generator's spellings are names to the parser, never cell references") {
    // Guards the generator itself: `n1` would be cell N1, and a property over cell refs exercises
    // none of the name classification.
    val nameShaped = Vector("=nm_1", "=NM_2*2", "=k*2", s"=$kelvin+1", "=S1!k", "=SUM(nm_2)")
    nameShaped.foreach { text =>
      val expr = FormulaParser.parse(text).toOption.getOrElse(fail(s"$text must parse"))
      assert(
        DependencyGraph.referencesMatching(expr, (_, _) => true, includeDynamicCalls = false),
        s"$text must carry a name reference"
      )
    }
  }
