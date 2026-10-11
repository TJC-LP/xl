package com.tjclp.xl.formula.graph

import com.tjclp.xl.{*, given}
import com.tjclp.xl.addressing.{ARef, SheetName}
import com.tjclp.xl.formula.ast.TExpr
import com.tjclp.xl.formula.eval.Evaluator
import com.tjclp.xl.formula.functions.{FunctionRegistry, FunctionSpec, ReferenceOperators}
import com.tjclp.xl.formula.parser.FormulaParser
import munit.FunSuite

class ValueDependenciesSpec extends FunSuite:
  private def parsed(formula: String): TExpr[?] =
    FormulaParser.parse(formula).fold(error => fail(error.toString), identity)

  private def reads(formula: String): Set[ARef] =
    ValueDependencies.extract(parsed(formula), Some(ref"A1:F10"), None)

  /** The value reads of `formula` placed on the sheet named Test. */
  private def readsOnTest(formula: String): Set[ARef] =
    ValueDependencies.extract(parsed(formula), Some(ref"A1:F10"), Some(SheetName.unsafe("Test")))

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
    assertEquals(ValueDependencies.extract(formula, Some(ref"A1:F10"), None), Set(ref"D5"))
  }

  // --- a reference qualified with the formula's own sheet is a read on that sheet --------------

  List(
    "Test!D5+1" -> Set(ref"D5"),
    "test!D5" -> Set(ref"D5"),
    "'Test'!D5*2" -> Set(ref"D5"),
    "SUM(Test!D4:D5)" -> Set(ref"D4", ref"D5"),
    "Test!D4:D5" -> Set(ref"D4", ref"D5"),
    "SUM(Test!D4:E9 Test!D5:D6)" -> Set(ref"D5", ref"D6"),
    "LET(chosen,Test!D5,chosen+1)" -> Set(ref"D5"),
    "ROW(Test!D5)+ROWS(Test!D4:D5)" -> Set.empty[ARef],
    "LET(chosen,Test!D5,ROW(chosen))" -> Set.empty[ARef],
    "Other!D5+SUM(Other!D4:D5)" -> Set.empty[ARef]
  ).foreach { (formula, expected) =>
    test(s"$formula on sheet Test reads its own sheet's qualified references as local") {
      assertEquals(readsOnTest(formula), expected)
    }
  }

  test("without a home sheet every qualified reference is another sheet's") {
    assertEquals(reads("Test!D5+SUM(Test!D4:D5)"), Set.empty[ARef])
  }

  // --- the reference-passing table covers every function that can return a reference ----------

  test("every reference-returning function names the arguments its reference passes through") {
    val referencing = FunctionRegistry.all.collect { case spec: FunctionSpec.Referencing[?, ?] =>
      spec.name
    }
    // the walk itself must find the known Referencing specs, or the assertion below is vacuous
    assert(Set("INDEX", "OFFSET", "INDIRECT").subsetOf(referencing.toSet), referencing.toString)
    val missing = (referencing.toSet ++ Evaluator.referenceFunctions + ReferenceOperators.UnionName)
      .filterNot(ValueDependencies.referencePassing.contains)
    assertEquals(missing, Set.empty[String], "add them to ValueDependencies.referencePassing")
  }
