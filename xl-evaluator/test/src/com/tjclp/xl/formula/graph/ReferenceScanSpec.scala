package com.tjclp.xl.formula.graph

import com.tjclp.xl.{*, given}
import com.tjclp.xl.formula.graph.ReferenceScan.Reach
import com.tjclp.xl.workbooks.DefinedName

import munit.FunSuite

/**
 * GH-606: the textual over-approximation of what a formula the parser rejects may read. Every token
 * class either contributes references, is provably not a reference, or makes the reach `Unbounded`;
 * these cases pin each class and the name-resolution rules.
 */
class ReferenceScanSpec extends FunSuite:
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

  test("error literals are dead operands; any other '#' token is unknown") {
    val literals = "#REF!,#N/A,#DIV/0!,#NAME?,#VALUE!,#NUM!,#NULL!,#SPILL!,#CALC!,#GETTING_DATA"
    assertEquals(reach(s"=ZZZNOTAFUNC($literals)"), Reach.Areas(Set.empty))
    assertEquals(reach("=ZZZNOTAFUNC(#WHAT)"), Reach.Unbounded)
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
