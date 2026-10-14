package com.tjclp.xl.formula.graph

import com.tjclp.xl.{*, given}
import com.tjclp.xl.addressing.{ARef, SheetName}
import com.tjclp.xl.cells.{CellError, CellValue}
import com.tjclp.xl.formula.Arity
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.formula.ast.TExpr
import com.tjclp.xl.formula.eval.Evaluator
import com.tjclp.xl.formula.functions.{
  FunctionRegistry,
  FunctionSpec,
  FunctionSpecs,
  ReferenceOperators
}
import com.tjclp.xl.formula.parser.FormulaParser
import com.tjclp.xl.workbooks.{DefinedName, Workbook}
import munit.FunSuite

class ValueDependenciesSpec extends FunSuite:
  private def parsed(formula: String): TExpr[?] =
    FormulaParser.parse(formula).fold(error => fail(error.toString), identity)

  private def reads(formula: String): Set[ARef] =
    ValueDependencies.Reader(Some(ref"A1:F10"), None).reads(parsed(formula))

  /** The value reads of `formula` placed on the sheet named Test. */
  private def readsOnTest(formula: String): Set[ARef] =
    ValueDependencies
      .Reader(Some(ref"A1:F10"), Some(SheetName.unsafe("Test")))
      .reads(parsed(formula))

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
    assertEquals(ValueDependencies.Reader(Some(ref"A1:F10"), None).reads(formula), Set(ref"D5"))
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
    "Other!D5+SUM(Other!D4:D5)" -> Set.empty[ARef],
    // unary wrappers keep both the own-sheet qualifier and reference-only inspection
    "-Test!D5" -> Set(ref"D5"),
    "Test!D5%" -> Set(ref"D5"),
    "+Test!D5" -> Set(ref"D5"),
    "-ROW(D5)" -> Set.empty[ARef],
    "-ROW(Test!D5)%" -> Set.empty[ARef]
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

  test("every reference operator is in the table, and the table names no other operator") {
    val operators = ValueDependenciesSpec.referenceOperators.map(_.name).toSet
    // the walk must see the operators that exist, or the checks below are vacuous
    assert(
      Set(
        ReferenceOperators.UnionName,
        ReferenceOperators.IntersectionName,
        ReferenceOperators.RangeName
      ).subsetOf(operators),
      operators.toString
    )
    assertEquals(
      operators.filterNot(ValueDependencies.referencePassing.contains),
      Set.empty[String]
    )
    val stray = ValueDependencies.referencePassing.keySet.filter(_.startsWith("(")) -- operators
    assertEquals(stray, Set.empty[String])
  }

  // --- the address-only rules hold only while those functions never read a value ---------------

  /** `formula` evaluated on a sheet whose A1 holds `a1`. */
  private def evaluatedOver(a1: CellValue, formula: String): Either[String, CellValue] =
    Sheet("Test").put(ref"A1", a1).evaluateFormula(formula).left.map(_.toString)

  private val valuesOfA1 = List(
    CellValue.Number(BigDecimal(42)),
    CellValue.Text("text"),
    CellValue.Error(CellError.Div0)
  )

  /** Whether `formula` answers the same whatever A1 holds, without an error. */
  private def addressOnly(formula: String): Boolean =
    val results = valuesOfA1.map(evaluatedOver(_, formula))
    results.distinct.sizeIs == 1 && results.forall {
      case Right(CellValue.Error(_)) | Left(_) => false
      case Right(_) => true
    }

  test("every inspector is a registered function whose result ignores the referenced value") {
    ValueDependencies.inspectors.foreach { name =>
      assert(FunctionRegistry.lookup(name).isDefined, s"$name is not a registered function")
      assert(addressOnly(s"=$name(A1)"), s"$name(A1) depends on A1's value")
    }
  }

  test("CELL's supported info types are exactly the address-only arms the CELL rule assumes") {
    // Every info type Excel documents. If CELL learns one that reads the referenced value
    // ("contents", "type"), the supported set grows and this fails: revisit the CELL rule in
    // ValueDependencies.Reader before updating the expected set.
    val excel = List(
      "address",
      "col",
      "color",
      "contents",
      "filename",
      "format",
      "parentheses",
      "prefix",
      "protect",
      "row",
      "type",
      "width"
    )
    val unsupported = Right(CellValue.Error(CellError.Value))
    val supported = excel.filter { info =>
      evaluatedOver(CellValue.Number(BigDecimal(42)), s"=CELL(\"$info\",A1)") != unsupported
    }
    assertEquals(supported.toSet, Set("address", "col", "filename", "row"))
    supported.foreach(info =>
      assert(addressOnly(s"=CELL(\"$info\",A1)"), s"CELL(\"$info\") depends on A1's value")
    )
  }

  test("computed range inspection reads endpoint selectors but not the referenced cells") {
    // Build GH-713's operator shape directly: this branch can merge before its parser support.
    // The same ArgSpec carries the endpoints of all two-operand reference operators.
    type Operands = (ReferenceOperators.Operand, ReferenceOperators.Operand)
    val range = FunctionSpec.simple[Any, Operands]("(range)", Arity.two)((_, _) => Right(0))(using
      ReferenceOperators.intersection.argSpec
    )
    val endpoint = ReferenceOperators.referenceOperand
      .parse(List(parsed("INDEX(D5:D6,A1)")), 0, "(range)")
      .fold(error => fail(error.toString), _._1)
    val computed = TExpr.Call(
      range,
      (Left(TExpr.RangeLocation.Local(ref"D5:D5")), endpoint)
    )
    val inspected = TExpr.Call(FunctionSpecs.rows, computed)
    val reader = ValueDependencies.Reader(Some(ref"A1:F10"), None)
    assertEquals(reader.reads(inspected), Set(ref"A1"))
    // Reading the range as a value must still retain those cell dependencies.
    assertEquals(reader.reads(computed), Set(ref"A1", ref"D5", ref"D6"))
  }

  /**
   * GH-715: Data and Other, with rate = Data!$B$1, base = rate, ubase = $B$1 (unqualified), loc =
   * Data!$A$1 for the workbook and loc = Data!$B$1 scoped to Other, ping and pong naming each
   * other.
   */
  private val named: Workbook =
    val book = Workbook(Vector(Sheet("Data"), Sheet("Other")))
    book.copy(metadata =
      book.metadata.copy(definedNames =
        Vector(
          DefinedName("rate", "Data!$B$1"),
          DefinedName("base", "rate"),
          DefinedName("ubase", "$B$1"),
          DefinedName("loc", "Data!$A$1"),
          DefinedName("loc", "Data!$B$1", localSheetId = Some(1)),
          DefinedName("ping", "pong+1"),
          DefinedName("pong", "ping+Data!$C$1"),
          DefinedName("block", "Data!$B$1:$C$2")
        )
      )
    )

  /** What `formula`, sitting on `from`, reads on Data through `named`'s names. */
  private def readsOnData(
    formula: String,
    from: String,
    names: ValueDependencies.Names = ValueDependencies.Names(named)
  ): Set[ARef] =
    ValueDependencies
      .Reader(
        Some(ref"A1:F10"),
        Some(SheetName.unsafe("Data")),
        localReads = from == "Data",
        names = Some(names),
        current = Some(SheetName.unsafe(from))
      )
      .readRanges(parsed(formula))
      .flatMap(_.cells)

  test("GH-715: a defined name read as a value reads its definition's reads, through names") {
    assertEquals(readsOnData("=rate*2", "Other"), Set(ref"B1"))
    assertEquals(readsOnData("=base*2", "Other"), Set(ref"B1"))
    // an unqualified body resolves from the referencing sheet: from Other it reads Other!B1
    assertEquals(readsOnData("=ubase", "Other"), Set.empty[ARef])
    assertEquals(readsOnData("=ubase", "Data"), Set(ref"B1"))
    // a sheet-scoped name shadows the workbook's on its own sheet only
    assertEquals(readsOnData("=loc", "Other"), Set(ref"B1"))
    assertEquals(readsOnData("=loc", "Data"), Set(ref"A1"))
    // inspecting a name is not a read; an unknown name reads nothing
    assertEquals(readsOnData("=ROWS(rate)", "Other"), Set.empty[ARef])
    assertEquals(readsOnData("=nosuchname", "Other"), Set.empty[ARef])
  }

  test("GH-715: names in a cycle terminate and read what their chain reaches") {
    assertEquals(readsOnData("=ping", "Other"), Set(ref"C1"))
  }

  test("GH-761: a defined name in a range slot reads its definition's reads, through names") {
    // SUMIF's first slot is a range slot: the name is a RangeLocation.Name, not a NameRef
    assertEquals(readsOnData("=SUMIF(block,\">0\")", "Other"), ref"B1:C2".cells.toSet)
    assertEquals(readsOnData("=COUNTIF(rate,\">5\")", "Other"), Set(ref"B1"))
    assertEquals(readsOnData("=SUMIF(base,\">0\")", "Other"), Set(ref"B1"))
    // the same resolution rules as a value-position name: an unqualified body resolves from the
    // referencing sheet, a sheet-scoped name shadows the workbook's on its own sheet, and a
    // qualifier names the sheet the name resolves from
    assertEquals(readsOnData("=SUMIF(ubase,\">0\")", "Other"), Set.empty[ARef])
    assertEquals(readsOnData("=SUMIF(ubase,\">0\")", "Data"), Set(ref"B1"))
    assertEquals(readsOnData("=COUNTIF(loc,\">5\")", "Other"), Set(ref"B1"))
    assertEquals(readsOnData("=COUNTIF(loc,\">5\")", "Data"), Set(ref"A1"))
    assertEquals(readsOnData("=COUNTIF(Other!loc,\">5\")", "Data"), Set(ref"B1"))
    // inspecting a name is not a read; an unknown name reads nothing
    assertEquals(readsOnData("=ROWS(block)", "Other"), Set.empty[ARef])
    assertEquals(readsOnData("=SUMIF(nosuchname,\">0\")", "Other"), Set.empty[ARef])
  }

  test("GH-761: a range-slot name reads nothing without names, as a value-position name does") {
    val reader = ValueDependencies.Reader(Some(ref"A1:F10"), Some(SheetName.unsafe("Data")))
    assertEquals(reader.reads(parsed("=SUMIF(block,\">0\")")), Set.empty[ARef])
    assertEquals(reader.reads(parsed("=block*2")), Set.empty[ARef])
    // nor does it size SUMIF's sum range, or get sized: the written ranges read as written
    assertEquals(reader.reads(parsed("=SUMIF(block,\">0\",Data!$D$1)")), Set(ref"D1"))
    assertEquals(reader.reads(parsed("=SUMIF(Data!D1:D2,\">0\",block)")), ref"D1:D2".cells.toSet)
  }

  test("GH-761: a name sizes SUMIF's and AVERAGEIF's sum range, or is sized, through names") {
    // block is B1:C2: a sum range from D1 is summed over D1:E2, as evaluation sizes it (#733)
    val blockAndD1E2 = ref"B1:C2".cells.toSet ++ ref"D1:E2".cells.toSet
    assertEquals(readsOnData("=SUMIF(block,\">0\",Data!$D$1)", "Other"), blockAndD1E2)
    assertEquals(readsOnData("=AVERAGEIF(block,\">0\",Data!$D$1)", "Other"), blockAndD1E2)
    // a named sum range takes the criteria range's shape from its upper-left cell, through a
    // chain of names too: rate and base name B1, summed over B1:B3 beside A1:A3
    val a1a3AndB1B3 = ref"A1:A3".cells.toSet ++ ref"B1:B3".cells.toSet
    assertEquals(readsOnData("=SUMIF(Data!A1:A3,\">0\",rate)", "Other"), a1a3AndB1B3)
    assertEquals(readsOnData("=SUMIF(Data!A1:A3,\">0\",base)", "Other"), a1a3AndB1B3)
    // SUMIFS sizes nothing: its sum range is read as written
    assertEquals(
      readsOnData("=SUMIFS(Data!$D$1,block,\">0\")", "Other"),
      ref"B1:C2".cells.toSet + ref"D1"
    )
  }

  test("GH-761: one Names resolves a sum range's name per qualifier and referencing sheet") {
    // loc names Data!$A$1 for the workbook and Data!$B$1 on Other; one Names serves every reader
    // of a --with search, so its memo must keep the three apart
    val names = ValueDependencies.Names(named)
    val sized = "=SUMIF(Data!D1:D2,\">0\",loc)"
    assertEquals(
      readsOnData(sized, "Data", names),
      ref"D1:D2".cells.toSet ++ ref"A1:A2".cells.toSet
    )
    assertEquals(
      readsOnData(sized, "Other", names),
      ref"D1:D2".cells.toSet ++ ref"B1:B2".cells.toSet
    )
    assertEquals(
      readsOnData("=SUMIF(Data!D1:D2,\">0\",Other!loc)", "Data", names),
      ref"D1:D2".cells.toSet ++ ref"B1:B2".cells.toSet
    )
  }

object ValueDependenciesSpec:
  /**
   * Every operator spec ReferenceOperators defines, found by reflection over its accessors (the
   * operators live outside FunctionRegistry), so a new operator needs no list here.
   */
  def referenceOperators: List[FunctionSpec[?]] =
    ReferenceOperators.getClass.getMethods.toList
      .filter(m =>
        m.getParameterCount == 0 && classOf[FunctionSpec[?]].isAssignableFrom(m.getReturnType)
      )
      .map(_.invoke(ReferenceOperators))
      .collect { case spec: FunctionSpec[?] => spec }
