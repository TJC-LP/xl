package com.tjclp.xl.formula.graph

import com.tjclp.xl.{*, given}
import com.tjclp.xl.addressing.ARef
import com.tjclp.xl.formula.ast.TExpr
import com.tjclp.xl.formula.parser.FormulaParser
import munit.FunSuite

class ValueDependenciesSpec extends FunSuite:
  private def reads(formula: String): Set[ARef] =
    val expr = FormulaParser.parse(formula).fold(error => fail(error.toString), identity)
    ValueDependencies.extract(expr, Some(ref"A1:F10"))

  List(
    "LET(chosen,D5,ROW(chosen))",
    "LET(chosen,INDEX(D5:D6,1),ROWS(chosen))",
    "LET(chosen,D5:D6,COLUMNS(chosen))",
    "LET(outer,D5,LET(inner,outer,COLUMN(inner)))",
    "LET(chosen,D5,LET(chosen,A1,ROWS(chosen)))",
    "CELL(\"address\",D5)",
    "CELL(\"filename\",D5)"
  ).foreach { formula =>
    test(s"$formula inspects LET reference identity without reading referenced values") {
      assertEquals(reads(formula), Set.empty[ARef])
    }
  }

  List(
    "LET(chosen,D5,chosen+1)" -> Set(ref"D5"),
    "LET(chosen,D5,ROW(chosen)+chosen)" -> Set(ref"D5"),
    "LET(chosen,INDEX(A1:A10,D5),ROWS(chosen))" -> Set(ref"D5"),
    "LET(chosen,OFFSET(A1,D5,0),ROWS(chosen))" -> Set(ref"D5"),
    "LET(chosen,IF(D5,A1,A2),ROWS(chosen))" -> Set(ref"D5"),
    "LET(chosen,CHOOSE(D5,A1,A2),ROWS(chosen))" -> Set(ref"D5"),
    "LET(unused,D5+1,ROW(A1))" -> Set(ref"D5"),
    "LET(unused,OFFSET(A1,D5,0),ROW(A1))" -> Set(ref"D5"),
    "LET(outer,D5,LET(inner,outer,inner+1))" -> Set(ref"D5"),
    "LET(chosen,D5,LET(chosen,chosen+1,chosen))" -> Set(ref"D5"),
    "ROWS(XLOOKUP(9,A1:A2,B1:B2,D5))" -> Set(ref"A1", ref"A2", ref"D5"),
    "CELL(D5,D5)" -> Set(ref"D5")
  ).foreach { (formula, expected) =>
    test(s"$formula retains scalar, selector and fallback value reads") {
      assertEquals(reads(formula), expected)
    }
  }

  test("a range-slot lexical name follows its binding's value reads") {
    // The parser substitutes literal ranges today; range-slot bindings have the same identifier
    // representation as this unqualified location, so the analysis needs no parser rewrite.
    val source = TExpr.Aggregate("SUM", TExpr.RangeLocation.Name("named", None))
    val formula = TExpr.Let(List("named" -> TExpr.PolyRef(ref"D5")), source)
    assertEquals(ValueDependencies.extract(formula, Some(ref"A1:F10")), Set(ref"D5"))
  }
