package com.tjclp.xl.formula.graph

import com.tjclp.xl.{*, given}
import com.tjclp.xl.formula.graph.ReferenceScan.Reach
import com.tjclp.xl.workbooks.DefinedName

import munit.ScalaCheckSuite
import org.scalacheck.Gen
import org.scalacheck.Prop.forAll
import org.scalacheck.rng.Seed

/**
 * GH-606: the textual over-approximation of what a formula the parser rejects may read. Every token
 * class either contributes references, is provably not a reference, or makes the reach `Unbounded`;
 * these cases pin each class and the name-resolution rules.
 */
class ReferenceScanSpec extends ScalaCheckSuite:
  // A fixed seed over a large sample: the law below is a soundness claim and must not flake
  override def scalaCheckTestParameters =
    super.scalaCheckTestParameters.withMinSuccessfulTests(600).withInitialSeed(Seed(0x606606L))

  private val sheet1 = SheetName.unsafe("Sheet1")
  private val sheet2 = SheetName.unsafe("Sheet2")

  private val workbook: Workbook =
    val base = Workbook(Sheet("Sheet1"), Sheet("Sheet2"), Sheet("Sheet3"), Sheet("Q1 Data"))
      .withDefinedName("Rate", "0.08")
      .withDefinedName("Inputs", "Sheet2!$B$2:$B$5")
      .withDefinedName("Multi", "Sheet3!$A$1:$A$3,Sheet3!$A$8:$A$10")
      .withDefinedName("Alias", "Multi")
      .withDefinedName("Dyn", "INDIRECT(\"A1\")")
      .withDefinedName("Loop", "Loop")
      .withDefinedName("Bare", "$A$1:$A$3,$A$8:$A$10")
      .withDefinedName("Chain", "Inputs")
      .withDefinedName("MyFunc", "_xlfn.LAMBDA(_xlpm.x,_xlpm.x+Sheet2!$A$1)")
      .withDefinedName("Twice", "_xlfn.LAMBDA(Sheet2!$A$1*2)")
      .withDefinedName("Point", "$A$1")
    // Sheet1 shadows the workbook-scoped Multi with a parseable local definition; Sheet2 carries
    // a sheet-scoped copy of the unparseable union so its unqualified areas have a home.
    base.copy(metadata =
      base.metadata.copy(definedNames =
        base.metadata.definedNames :+
          DefinedName("Multi", "Sheet1!$C$1", localSheetId = Some(0)) :+
          DefinedName("Local", "$A$1:$A$3,$A$8:$A$10", localSheetId = Some(1))
      )
    )

  private def area(sheet: String, range: String): (SheetName, CellRange) =
    val parsed = CellRange.parse(range).fold(error => fail(error), identity)
    (SheetName.unsafe(sheet), CellRange(parsed.start, parsed.end))

  private def areas(pairs: (String, String)*): Reach = Reach.Areas(pairs.map(area).toSet)

  private def reach(text: String, from: SheetName = sheet1): Reach =
    ReferenceScan.reach(workbook, from, text)

  test("a cell reference reaches only that cell on the reader's sheet") {
    assertEquals(reach("=ZZZNOTAFUNC(C3)"), areas("Sheet1" -> "C3"))
    assertEquals(reach("ZZZNOTAFUNC(C3)"), areas("Sheet1" -> "C3"), "the leading '=' is optional")
  }

  test("anchors are stripped and a range is its bounding rectangle") {
    assertEquals(reach("=ZZZNOTAFUNC($C$3,A$1:$B2)"), areas("Sheet1" -> "C3", "Sheet1" -> "A1:B2"))
    assertEquals(reach("=ZZZNOTAFUNC(B20:A1)"), areas("Sheet1" -> "A1:B20"))
  }

  test("a sheet qualifier applies to the immediately following reference only") {
    assertEquals(
      reach("=ZZZNOTAFUNC('Q1 Data'!B16,C3)"),
      areas("Q1 Data" -> "B16", "Sheet1" -> "C3")
    )
    assertEquals(reach("=ZZZNOTAFUNC('Q1 Data'!A1:B2)"), areas("Q1 Data" -> "A1:B2"))
  }

  test("an unquoted qualifier resolves case-insensitively to the workbook's spelling") {
    assertEquals(reach("=ZZZNOTAFUNC(sheet2!A1:A5)"), areas("Sheet2" -> "A1:A5"))
  }

  test("a qualifier naming no sheet reads #REF! and contributes nothing") {
    assertEquals(reach("=ZZZNOTAFUNC(Missing!A1,'Also Missing'!B2)"), Reach.Areas(Set.empty))
  }

  test("a 3-D span covers every sheet between its ends, in either order") {
    val expected = areas("Sheet1" -> "A1", "Sheet2" -> "A1", "Sheet3" -> "A1")
    assertEquals(reach("=ZZZNOTAFUNC(Sheet1:Sheet3!A1)"), expected)
    assertEquals(reach("=ZZZNOTAFUNC('Sheet3:Sheet1'!A1)"), expected)
    assertEquals(reach("=ZZZNOTAFUNC(Sheet1:Missing!A1)"), Reach.Unbounded)
  }

  test("string literals are not references") {
    assertEquals(reach("=ZZZNOTAFUNC(\"B16\")"), Reach.Areas(Set.empty))
    assertEquals(reach("=ZZZNOTAFUNC(\"say \"\"B16\"\" twice\")"), Reach.Areas(Set.empty))
    assertEquals(reach("=ZZZNOTAFUNC(\"unterminated"), Reach.Unbounded)
  }

  test("the CellError literals are dead operands; any other '#' token is unknown") {
    val literals = CellError.values.map(_.toExcel).mkString(",")
    assertEquals(reach(s"=ZZZNOTAFUNC($literals)"), Reach.Areas(Set.empty))
    assertEquals(
      reach("=ZZZNOTAFUNC(#N/A,#NAME?,#NUM!,#NULL!)"),
      Reach.Areas(Set.empty),
      "prefixes"
    )
    assertEquals(reach("=ZZZNOTAFUNC(#WHAT)"), Reach.Unbounded)
    assertEquals(reach("=ZZZNOTAFUNC(#SPILL!)"), Reach.Unbounded, "not modelled by CellError")
  }

  test("a spill reference A1# is unbounded, never read as A1") {
    assertEquals(reach("=ZZZNOTAFUNC(A1#)"), Reach.Unbounded)
    assertEquals(reach("=ZZZNOTAFUNC(Sheet2!A1#)"), Reach.Unbounded)
    assertEquals(reach("=SUM(A1#)"), Reach.Unbounded)
  }

  test("LET and LAMBDA parameter names do not resolve and are unbounded") {
    assertEquals(reach("=ZZZNOTAFUNC(LET(x,1,x+A1))"), Reach.Unbounded)
    assertEquals(reach("=ZZZNOTAFUNC(_xlfn.LET(_xlpm.x,1,_xlpm.x+A1))"), Reach.Unbounded)
    assertEquals(reach("=ZZZNOTAFUNC(LAMBDA(x,x+A1)(1))"), Reach.Unbounded)
  }

  test(
    "a parseable workbook-scoped definition with an unqualified reference is read from the reader's sheet"
  ) {
    // The parsed branch follows the evaluator: `Point = $A$1` resolves against the sheet the
    // reader is on (the textual branch, exercised by `Bare`, has no such home and is unbounded)
    assertEquals(reach("=ZZZNOTAFUNC(Point)", sheet1), areas("Sheet1" -> "A1"))
    assertEquals(reach("=ZZZNOTAFUNC(Point)", sheet2), areas("Sheet2" -> "A1"))
    assertEquals(
      reach("=ZZZNOTAFUNC(Sheet2!Point)", sheet1),
      areas("Sheet1" -> "A1"),
      "the qualifier picks the name, the reader picks the home"
    )
  }

  test("a name chain deeper than the guard is unbounded") {
    // `Link1` and not `N1`: a one-to-three-letter name followed by digits IS a cell reference
    val deep = (1 to 120)
      .foldLeft(workbook) { (wb, i) => wb.withDefinedName(s"Link$i", s"Link${i + 1}") }
      .withDefinedName("Link121", "Sheet2!$A$1")
    assertEquals(ReferenceScan.reach(deep, sheet1, "=ZZZNOTAFUNC(Link1)"), Reach.Unbounded)
    assertEquals(
      ReferenceScan.reach(deep, sheet1, "=ZZZNOTAFUNC(Link100)"),
      areas("Sheet2" -> "A1"),
      "22 hops are within the guard"
    )
  }

  test("numbers, operators, array constants and booleans carry no references") {
    assertEquals(
      reach("=ZZZNOTAFUNC(1.5E+3+2%-.5*4/5^6&\"x\"<>7<=8>=9,{1,2;3,4},TRUE,FALSE,@1)"),
      Reach.Areas(Set.empty)
    )
  }

  test("a colon between two integers is a whole-row range") {
    assertEquals(reach("=ZZZNOTAFUNC(3:3,$5:$7)"), areas("Sheet1" -> "3:3", "Sheet1" -> "5:7"))
    assertEquals(reach("=ZZZNOTAFUNC(Sheet2!10:3)"), areas("Sheet2" -> "3:10"))
  }

  test("a colon between two column letters is a whole-column range") {
    assertEquals(reach("=ZZZNOTAFUNC(B:B,$A:$C)"), areas("Sheet1" -> "B:B", "Sheet1" -> "A:C"))
    assertEquals(reach("=ZZZNOTAFUNC(Sheet3!XFD:XFC)"), areas("Sheet3" -> "XFC:XFD"))
  }

  test("function names are not references; a dynamic function is unbounded") {
    assertEquals(
      reach("=ZZZNOTAFUNC(SUM(A1),_xlfn.XLOOKUP(1,B1:B2,C1:C2))"),
      areas("Sheet1" -> "A1", "Sheet1" -> "B1:B2", "Sheet1" -> "C1:C2")
    )
    assertEquals(reach("=ZZZNOTAFUNC(INDIRECT(\"A1\"))"), Reach.Unbounded)
    assertEquals(reach("=ZZZNOTAFUNC(_xlfn.indirect(\"A1\"))"), Reach.Unbounded)
    assertEquals(reach("=ZZZNOTAFUNC(OFFSET(A1,1,1))"), Reach.Unbounded)
  }

  test("a call to a name the registry does not know reads through the definition") {
    // Excel 365 stores a LAMBDA as a defined name and its caller as `MyFunc(1)`, which the parser
    // rejects; the parametrised body is unbounded through `_xlpm.x`, a parameterless one bounded
    assertEquals(reach("=MyFunc(1)"), Reach.Unbounded)
    assertEquals(reach("=ZZZNOTAFUNC(Twice())"), areas("Sheet2" -> "A1"))
    assertEquals(
      reach("=ZZZNOTAFUNC(SINGLE(A1),RRI(3,B1,B2))"),
      areas("Sheet1" -> "A1", "Sheet1" -> "B1", "Sheet1" -> "B2"),
      "an unknown function that is no name reads only its arguments"
    )
  }

  test("ANCHORARRAY, SUMIF and AVERAGEIF read beyond their argument text") {
    assertEquals(reach("=ZZZNOTAFUNC(_xlfn.ANCHORARRAY(A1))"), Reach.Unbounded)
    assertEquals(reach("=ZZZNOTAFUNC(SUMIF(A1:A10,\">0\",C1))"), Reach.Unbounded)
    assertEquals(reach("=ZZZNOTAFUNC(averageif(A1:A10,\">0\",C1))"), Reach.Unbounded)
    assertEquals(
      reach("=ZZZNOTAFUNC(SUMIFS(C1:C10,A1:A10,\">0\"))"),
      areas("Sheet1" -> "C1:C10", "Sheet1" -> "A1:A10"),
      "SUMIFS demands matching shapes and stays bounded"
    )
  }

  test("a defined name reads what its parseable definition reads") {
    assertEquals(reach("=ZZZNOTAFUNC(Rate)"), Reach.Areas(Set.empty))
    assertEquals(reach("=ZZZNOTAFUNC(Inputs)"), areas("Sheet2" -> "B2:B5"))
    assertEquals(reach("=ZZZNOTAFUNC(Chain)"), areas("Sheet2" -> "B2:B5"), "name -> name chains")
  }

  test("an unparseable definition is scanned textually, through aliases too") {
    val expected = areas("Sheet3" -> "A1:A3", "Sheet3" -> "A8:A10")
    assertEquals(reach("=ZZZNOTAFUNC(Multi)", sheet2), expected)
    assertEquals(reach("=ZZZNOTAFUNC(Alias)", sheet2), expected)
    assertEquals(reach("=ZZZNOTAFUNC(multi)", sheet2), expected, "name lookup is case-insensitive")
  }

  test("a local name shadows the workbook-scoped one, from its own sheet or by qualifier") {
    assertEquals(reach("=ZZZNOTAFUNC(Multi)", sheet1), areas("Sheet1" -> "C1"))
    assertEquals(reach("=ZZZNOTAFUNC(Sheet1!Multi)", sheet2), areas("Sheet1" -> "C1"))
    assertEquals(
      reach("=ZZZNOTAFUNC(Sheet2!Multi)", sheet1),
      areas("Sheet3" -> "A1:A3", "Sheet3" -> "A8:A10")
    )
  }

  test("an unqualified reference inside an unparseable definition needs a sheet scope") {
    assertEquals(reach("=ZZZNOTAFUNC(Bare)"), Reach.Unbounded, "workbook-scoped: no home sheet")
    assertEquals(
      reach("=ZZZNOTAFUNC(Local)", sheet2),
      areas("Sheet2" -> "A1:A3", "Sheet2" -> "A8:A10"),
      "sheet-scoped: the scope sheet"
    )
  }

  test("a missing, dynamic or self-referential name is unbounded") {
    assertEquals(reach("=ZZZNOTAFUNC(NoSuchName)"), Reach.Unbounded)
    assertEquals(reach("=ZZZNOTAFUNC(Dyn)"), Reach.Unbounded)
    assertEquals(reach("=ZZZNOTAFUNC(Loop)"), Reach.Unbounded)
    assertEquals(reach("=ZZZNOTAFUNC(Missing!Rate)"), Reach.Unbounded)
  }

  test("structured and external references are unbounded") {
    assertEquals(reach("=ZZZNOTAFUNC(Table1[Col])"), Reach.Unbounded)
    assertEquals(reach("=ZZZNOTAFUNC([1]Sheet1!A1)"), Reach.Unbounded)
    assertEquals(reach("=ZZZNOTAFUNC('[1]Sheet1'!A1)"), Reach.Unbounded)
  }

  test("a colon with a name or a call on either side is unbounded") {
    assertEquals(reach("=ZZZNOTAFUNC(A1:INDEX(A1:A10,3))"), Reach.Unbounded)
    assertEquals(reach("=ZZZNOTAFUNC(Inputs:A1)"), Reach.Unbounded)
    assertEquals(reach("=ZZZNOTAFUNC(A1:Inputs)"), Reach.Unbounded)
    assertEquals(reach("=ZZZNOTAFUNC(OFFSET(A1,1,1):B5)"), Reach.Unbounded)
    assertEquals(reach("=ZZZNOTAFUNC(A1:)"), Reach.Unbounded)
  }

  test("any other character is unknown") {
    assertEquals(reach("=ZZZNOTAFUNC(A1~B2)"), Reach.Unbounded)
    assertEquals(reach("=ZZZNOTAFUNC('lonely quoted')"), Reach.Unbounded)
    assertEquals(reach("=ZZZNOTAFUNC($)"), Reach.Unbounded)
  }

  test("whitespace, the intersection operator and line breaks are harmless") {
    assertEquals(
      reach("=ZZZNOTAFUNC( A1:B2 B2:C3 ,\n D4 )"),
      areas("Sheet1" -> "A1:B2", "Sheet1" -> "B2:C3", "Sheet1" -> "D4")
    )
  }

  test("a reference just past the grid is a name, not a cell") {
    assertEquals(reach("=ZZZNOTAFUNC(XFE1)"), Reach.Unbounded)
    assertEquals(reach("=ZZZNOTAFUNC(A1048577)"), Reach.Unbounded)
    assertEquals(reach("=ZZZNOTAFUNC(XFD1048576)"), areas("Sheet1" -> "XFD1048576"))
  }

  test("a parseable definition's reach covers every reference node shape") {
    // The parsed branch reads a definition through `extractQualifiedDependencies` with `cellsFor`
    // as a visitor (see `definitionReach`); each definition below is one reference-bearing TExpr
    // node, so a node the extraction stopped visiting fails here rather than under-approximating.
    val base = Workbook(Sheet("Sheet1"), Sheet("Sheet2"))
      .withDefinedName("RefDef", "Sheet2!$A$1")
      .withDefinedName("RangeDef", "Sheet2!$A$1:$B$2")
      .withDefinedName("AggDef", "SUM(Sheet2!A1:A3)")
      .withDefinedName("LookupDef", "VLOOKUP(1,Sheet2!A1:D5,3,FALSE)")
      .withDefinedName("ArithDef", "Sheet2!$A$1*2+Sheet2!$B$1^2-Sheet2!$C$1/Sheet2!$D$1")
      .withDefinedName("CompareDef", "IF(Sheet2!A1>Sheet2!B1,Sheet2!C1,Sheet2!D1)")
      .withDefinedName("TextDef", "Sheet2!$A$1&LEN(Sheet2!$B$1)")
      .withDefinedName("LetDef", "LET(x,Sheet2!$A$1,x*Sheet2!$B$1)")
      .withDefinedName("CoercedDef", "ROUND(Sheet2!$A$1,0)+Sheet2!$B$1%")
      .withDefinedName("CountDef", "COUNTIF(Sheet2!A1:A9,\">0\")")
    val wb = base.copy(metadata =
      base.metadata.copy(definedNames =
        base.metadata.definedNames :+
          DefinedName("LocalRef", "$A$1", localSheetId = Some(1)) :+
          DefinedName("LocalRange", "$A$1:$B$2", localSheetId = Some(1)) :+
          DefinedName("LocalAgg", "SUM(A1:A3)", localSheetId = Some(1))
      )
    )
    def of(name: String): Reach = ReferenceScan.reach(wb, sheet1, s"=ZZZNOTAFUNC($name)")
    assertEquals(of("RefDef"), areas("Sheet2" -> "A1"))
    assertEquals(of("RangeDef"), areas("Sheet2" -> "A1:B2"))
    assertEquals(of("AggDef"), areas("Sheet2" -> "A1:A3"))
    assertEquals(
      of("LookupDef"),
      areas("Sheet2" -> "A1:D5"),
      "the full table, never the two strips"
    )
    assertEquals(
      of("ArithDef"),
      areas("Sheet2" -> "A1", "Sheet2" -> "B1", "Sheet2" -> "C1", "Sheet2" -> "D1")
    )
    assertEquals(
      of("CompareDef"),
      areas("Sheet2" -> "A1", "Sheet2" -> "B1", "Sheet2" -> "C1", "Sheet2" -> "D1")
    )
    assertEquals(of("TextDef"), areas("Sheet2" -> "A1", "Sheet2" -> "B1"))
    assertEquals(of("LetDef"), areas("Sheet2" -> "A1", "Sheet2" -> "B1"))
    assertEquals(of("CoercedDef"), areas("Sheet2" -> "A1", "Sheet2" -> "B1"))
    assertEquals(of("CountDef"), areas("Sheet2" -> "A1:A9"))
    assertEquals(of("Sheet2!LocalRef"), areas("Sheet2" -> "A1"))
    assertEquals(of("Sheet2!LocalRange"), areas("Sheet2" -> "A1:B2"))
    assertEquals(of("Sheet2!LocalAgg"), areas("Sheet2" -> "A1:A3"))
  }

  // ----- the soundness law: textual reach ⊇ parsed reach -----

  private val lawWorkbook: Workbook =
    val base = Workbook(Sheet("Sheet1"), Sheet("Sheet2"), Sheet("Q1 Data"), Sheet("It's"))
      .withDefinedName("Rate", "0.08")
      .withDefinedName("Inputs", "Sheet2!$B$2:$B$5")
      .withDefinedName("Cell1", "Sheet2!$A$1")
      .withDefinedName("Chain", "Inputs")
      .withDefinedName("Bare", "$A$1")
      .withDefinedName("Multi", "Sheet2!$A$1:$A$3,Sheet2!$A$8:$A$10")
      .withDefinedName("Dyn", "INDIRECT(\"A1\")")
    base.copy(metadata =
      base.metadata.copy(definedNames =
        base.metadata.definedNames :+ DefinedName("Local", "$C$1:$C$3", localSheetId = Some(0))
      )
    )

  private val genA1: Gen[String] =
    for
      col <- Gen.choose(0, 30)
      row <- Gen.choose(0, 40)
      anchor <- Gen.oneOf("", "$", "row", "col")
    yield
      val ref = ARef.from0(col, row)
      anchor match
        case "$" => s"$$${ref.col.toLetter}$$${ref.row.index1}"
        case "row" => s"${ref.col.toLetter}$$${ref.row.index1}"
        case "col" => s"$$${ref.col.toLetter}${ref.row.index1}"
        case _ => ref.toA1

  private val genRangeText: Gen[String] =
    for
      a <- genA1
      w <- Gen.choose(0, 5)
      h <- Gen.choose(0, 5)
      start = ARef.parse(a.replace("$", "")).fold(fail(_), identity)
    yield s"$a:${ARef.from0(start.col.index0 + w, start.row.index0 + h).toA1}"

  private val genQualifier: Gen[String] = Gen.oneOf(
    "Sheet1!",
    "Sheet2!",
    "sheet2!",
    "'Sheet2'!",
    "'Q1 Data'!",
    "'It''s'!",
    "Missing!",
    "'Also Missing'!",
    "Sheet1:Sheet2!",
    "'Sheet1:Q1 Data'!",
    ""
  )

  private def genFormulaText(depth: Int): Gen[String] =
    val leaf: Gen[String] = Gen.oneOf(
      genA1,
      for q <- genQualifier; a <- genA1 yield s"$q$a",
      for q <- genQualifier; r <- genRangeText yield s"SUM($q$r)",
      for q <- genQualifier; r <- genRangeText yield s"MAX($q$r,1)",
      Gen.oneOf("A:A", "$B:$D", "Sheet2!C:C", "'Q1 Data'!A:B").map(c => s"SUM($c)"),
      Gen.oneOf("1:1", "$3:$5", "Sheet2!7:9").map(r => s"SUM($r)"),
      Gen.choose(-999, 999).map(_.toString),
      Gen.oneOf("1.5", ".5", "2.", "1E5", "1.5E-3", "TRUE", "FALSE", "#N/A", "#REF!", "#DIV/0!"),
      Gen.alphaNumStr.map(t => "\"" + t + "\""),
      Gen.const("\"Sheet2!A1:B2 \"\"quoted\"\"\""),
      Gen.oneOf("Rate", "Inputs", "Cell1", "Chain", "Bare", "Local", "Multi", "Dyn", "NoSuchName"),
      Gen.oneOf("Sheet1!Local", "Sheet2!Rate", "Sheet2!Bare", "sheet1!Cell1", "Missing!Rate"),
      Gen.oneOf("Inputs", "Multi", "Local", "Chain").map(n => s"SUM($n)"),
      for r <- genRangeText yield s"VLOOKUP(1,Sheet2!$r,2,FALSE)",
      for r <- genRangeText yield s"COUNTIF($r,\">0\")",
      for r <- genRangeText yield s"SUMIF($r,\">0\",C1)",
      for r <- genRangeText yield s"SUMIFS(C1:C5,$r,\">0\")",
      genA1.map(a => s"INDIRECT(\"$a\")"),
      genA1.map(a => s"OFFSET($a,1,1)"),
      genA1.map(a => s"LET(x,$a,x*2)"),
      genA1.map(a => s"_xlfn.XLOOKUP(1,$a:$a,$a:$a)"),
      genA1.map(a => s"@$a"),
      genA1.map(a => s"$a#"),
      genA1.map(a => s"LEN(\"$a\")"),
      genA1.map(a => s"ROUND($a,0)")
    )
    if depth >= 2 then leaf
    else
      Gen.frequency(
        3 -> leaf,
        1 -> Gen.lzy(for
          x <- genFormulaText(depth + 1)
          y <- genFormulaText(depth + 1)
          op <- Gen.oneOf("+", "-", "*", "/", "^", "&", "=", "<>", "<", "<=", ">", ">=")
        yield s"($x$op$y)"),
        1 -> Gen.lzy(genFormulaText(depth + 1).map(x => s"-$x")),
        1 -> Gen.lzy(genFormulaText(depth + 1).map(x => s"$x%")),
        1 -> Gen.lzy(for
          c <- genFormulaText(depth + 1)
          a <- genFormulaText(depth + 1)
          b <- genFormulaText(depth + 1)
        yield s"IF($c>0,$a,$b)")
      )

  /**
   * Whether `range` on `sheet` lies inside the reach — inside one of its rectangles, or unbounded.
   */
  private def covers(reach: Reach, sheet: SheetName, range: CellRange): Boolean = reach match
    case Reach.Unbounded => true
    case Reach.Areas(areas) =>
      areas.exists((s, r) => s == sheet && r.contains(range.start) && r.contains(range.end))

  property(
    "law: textual reach ⊇ parsed reach — every dependency the parser extracts lies inside the scanner's areas, or the reach is unbounded"
  ) {
    val existing = lawWorkbook.sheets.map(_.name).toSet
    val canonical = DependencyGraph.sheetCanonicaliser(lawWorkbook)
    forAll(genFormulaText(0), Gen.oneOf(sheet1, sheet2)) { (text, from) =>
      // The law speaks only about formulas the parser accepts; the scanner is exercised on every
      // draw regardless (it must be total), and the 3-D and error-literal leaves stay in the pool
      // so the law binds the day the parser learns them.
      val reach = ReferenceScan.reach(lawWorkbook, from, text)
      FormulaParser.parse(s"=$text").foreach { expr =>
        val ranges = scala.collection.mutable.HashSet.empty[(SheetName, CellRange)]
        val points = DependencyGraph.extractQualifiedDependencies(
          expr,
          from,
          cellsFor = (sheet, range) =>
            ranges += ((sheet, range))
            Set.empty
          ,
          workbook = Some(lawWorkbook),
          canonicalSheet = canonical
        )
        val parsed = ranges.toSet ++ points.map(q => (q.sheet, CellRange(q.ref, q.ref)))
        // A dependency on a sheet the workbook lacks is #REF! whatever any cell holds; both the
        // graph and the scanner treat it as reading nothing that exists.
        parsed.filter((sheet, _) => existing.contains(sheet)).foreach { (sheet, range) =>
          assert(
            covers(reach, sheet, range),
            s"'$text' read from ${from.value}: parsed dependency ${sheet.value}!${range.toA1} escapes the textual reach $reach"
          )
        }
      }
    }
  }

  test("reaches scans every reader once against its own sheet") {
    val wb = Workbook(
      Sheet("Sheet1").put(ref"A1", CellValue.Formula("ZZZNOTAFUNC(C3)")),
      Sheet("Sheet2").put(ref"A1", CellValue.Formula("ZZZNOTAFUNC(Table1[Col])"))
    )
    val readers = Set(
      DependencyGraph.QualifiedRef(sheet1, ref"A1"),
      DependencyGraph.QualifiedRef(sheet2, ref"A1")
    )
    assertEquals(
      ReferenceScan.reaches(wb, readers),
      Map(
        DependencyGraph.QualifiedRef(sheet1, ref"A1") -> areas("Sheet1" -> "C3"),
        DependencyGraph.QualifiedRef(sheet2, ref"A1") -> Reach.Unbounded
      )
    )
  }
