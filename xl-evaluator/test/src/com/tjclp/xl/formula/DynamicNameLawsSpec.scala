package com.tjclp.xl.formula

import com.tjclp.xl.{*, given}
import com.tjclp.xl.addressing.{ARef, SheetName}
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.formula.ast.TExpr
import com.tjclp.xl.formula.eval.Evaluator
import com.tjclp.xl.formula.functions.CriteriaRangeResize
import com.tjclp.xl.formula.parser.FormulaParser
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.workbooks.Workbook
import munit.ScalaCheckSuite
import org.scalacheck.Gen
import org.scalacheck.Prop.forAllNoShrink

/** GH-748: text discovery and LET's two binding representations must preserve semantics. */
class DynamicNameLawsSpec extends ScalaCheckSuite:
  // Generic tuple/list shrinking can create row 0 or an empty grid, outside this grammar.
  // Preserve the valid generated witness so a mutation fails for its semantics, not a parse error.
  override def scalaCheckTestParameters =
    super.scalaCheckTestParameters.withMinSuccessfulTests(60)

  private def parsed(formula: String): TExpr[?] =
    FormulaParser.parse(formula).fold(error => fail(s"$formula: $error"), identity)

  // The generated grammar has exactly one resizing call, optionally under LET. Inspect the
  // parsed slots, independently of the scanner, rather than discarding unparseable/non-dynamic
  // samples through an implication. Every sample must exercise the antecedent of the law.
  private def dynamicResize(expr: TExpr[?]): Boolean = expr match
    case TExpr.Let(_, body) => dynamicResize(body)
    case call: TExpr.Call[?] =>
      CriteriaRangeResize.hasDynamicResize(
        call.spec.name,
        call.spec.argSpec.toValues(call.args),
        _ => true
      )
    case _ => false

  property("GH-748: parsed dynamic name resizes never escape the text pre-filter") {
    val gen = for
      suffix <- Gen.choose(10, 9999)
      row <- Gen.choose(2, 30)
      whitespace <- Gen.oneOf("", " ", "\t", "\u2003")
    yield (suffix, row, whitespace)
    forAllNoShrink(gen) { (suffix, row, whitespace) =>
      // Visit every lexical family on every trial; chance cannot omit Unicode, escaped quotes,
      // literal decoys or a colon corner. Name predicates model computed definitions, not static
      // names: any parsed Name slot is a possible computed reference in a workbook.
      val names = List(
        s"amount$suffix",
        s"_amt$suffix",
        s"\\amt$suffix",
        s"Sales.Total$suffix",
        s"ümsatz$suffix",
        "XFE1"
      )
      val functions = List("SUMIF", "averageif", "ſUMIF", "SUMıF", "AVERAGEıF")
      val criteria = List("1", "\"name:decoy\"", "\"a\"\"b\"", "#N/A", "#NAME?")
      names.foreach { name =>
        functions.foreach { function =>
          val f = function + whitespace
          val formulas = criteria.flatMap { criterion =>
            List(
              s"$f(A1:A$row,$criterion,$name)",
              s"$f($name,$criterion,C1)",
              s"$f('Q 1''s'!$name,$criterion,'Q 1''s'!C1)",
              s"LET(local,C1,$f(A1:A$row,$criterion,local))"
            )
            // Computed colon ranges cannot yet occupy a range-typed slot (#723); put them
            // in the criterion expression while keeping a genuine dynamic range slot.
          } ++ List(
            s"$f(A1:A$row,SUM(A1:corner$suffix),$name)",
            s"$f(A1:A$row,SUM(corner$suffix:A$row),$name)",
            s"$f(A1:A$row,SUM(left$suffix:right$suffix),$name)",
            s"_xlfn.LET(_xlpm.local,C1,$f(A1:A$row,1,_xlpm.local))"
          )
          formulas.foreach { formula =>
            assert(dynamicResize(parsed(formula)), s"generator missed dynamic resize: $formula")
            assert(CriteriaRangeResize.mayResizeThroughName(formula), formula)
            assert(
              CriteriaRangeResize
                .mayResizeThroughName(formula, formula.toUpperCase(java.util.Locale.ROOT)),
              formula
            )
          }
        }
      }
      true
    }
  }

  private val sheetName = SheetName.unsafe("Data")
  private def number(n: Int): CellValue = CellValue.Number(BigDecimal(n))

  property("GH-748: array LET reference binding is neutral for unrelated array values") {
    val gen = for
      size <- Gen.choose(3, 8)
      values <- Gen.listOfN(size, Gen.choose(-9, 9))
      factor <- Gen.choose(2, 5)
      threshold <- Gen.choose(-4, 4)
    yield (values, factor, threshold)
    forAllNoShrink(gen) { (values, factor, threshold) =>
      val size = values.size
      val range = s"A1:A$size"
      val sheet = values.zipWithIndex.foldLeft(Sheet(sheetName)) { case (s, (n, row)) =>
        s.put(ARef.from0(0, row), number(n)).put(ARef.from0(1, row), number(n + 10))
      }
      val book = Workbook(sheet).withDefinedName("dyn", s"OFFSET(Data!A1,0,0,$size,1)")
      def spill(formula: String): (Int, Int, List[CellValue]) =
        sheet
          .evaluateArrayFormula(s"=$formula", ref"J1", Clock.system, Some(book))
          .fold(
            error => fail(s"$formula: ${error.message}"),
            (result, area) =>
              (area.height, area.width, area.cells.map(at => result(at).value).toList)
          )
      val count = values.count(_ > threshold)
      val triggers = List(
        (s"COUNTIF(r,\">$threshold\")", s"COUNTIF(dyn,\">$threshold\")", count),
        ("SUM(OFFSET(r,0,0,1,1))", "A1", values.headOption.getOrElse(0))
      )
      // IF and INDEX denote references, so s changes representation with the whole LET's flag.
      // Multiplication denotes an array value: it must stay an array even when r needs a reference.
      val bindings =
        List(range, "dyn", s"IF(TRUE,$range,B1:B$size)", s"INDEX(A1:B$size,0,1)", s"$range*$factor")
      bindings.foreach { bound =>
        val multiplier = if bound == s"$range*$factor" then factor else 1
        triggers.foreach { (referenceRead, valueRead, contribution) =>
          val reads = List(
            ("SUMPRODUCT(s)", List(number(contribution + values.sum * multiplier))),
            (s"s*$factor", values.map(n => number(contribution + n * multiplier * factor)))
          )
          reads.foreach { (read, expected) =>
            val together = s"LET(r,dyn,s,$bound,$referenceRead+$read)"
            val valuePath = s"LET(r,dyn,s,$bound,$valueRead+$read)"
            val referenceExpr = parsed(together)
            val valueExpr = parsed(valuePath)
            assert(Evaluator.readsBindingAsRange(referenceExpr), together)
            assert(!Evaluator.readsBindingAsRange(valueExpr), valuePath)
            assert(!Evaluator.inspectsReference(valueExpr), valuePath)
            // Pin success and actual numbers as well as equality: two identical errors cannot pass.
            assertEquals(spill(valuePath), (expected.size, 1, expected), valuePath)
            assertEquals(spill(together), (expected.size, 1, expected), together)
          }
        }
      }
      true
    }
  }
