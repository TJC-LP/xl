package com.tjclp.xl.formula

import com.tjclp.xl.{*, given}
import com.tjclp.xl.addressing.{ARef, SheetName}
import com.tjclp.xl.cells.{CellError, CellValue}
import com.tjclp.xl.formula.ast.TExpr
import com.tjclp.xl.formula.eval.{EvalError, Evaluator, RangeOperand}
import com.tjclp.xl.formula.parser.FormulaParser
import com.tjclp.xl.formula.printer.FormulaPrinter
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.workbooks.Workbook
import munit.FunSuite

/**
 * GH-710: a dynamic named range — a name whose formula computes a reference (OFFSET, INDIRECT,
 * INDEX, IF or CHOOSE over references) — and a LET name bound to a reference are references in
 * every slot that takes a range: the criteria and lookup functions, INDEX, MATCH, LARGE, RANK, NPV,
 * TRANSPOSE. They used to be `#VALUE!` ("does not refer to a range") or a parse error.
 *
 * The first table is LibreOffice 24.2's recalculation of the same plain cells (Empty!A1, A2, … in
 * order). LibreOffice 24.2 has neither LET, MAXIFS, MINIFS, XMATCH nor XLOOKUP: those cases are
 * pinned to the value of the same formula written with the reference itself, which is what Excel's
 * LET (it binds the reference) and those functions compute.
 */
class DynamicNameRangeSlotSpec extends FunSuite:

  private def N(n: String): CellValue = CellValue.Number(BigDecimal(n))
  private def T(s: String): CellValue = CellValue.Text(s)
  private def E(error: CellError): CellValue = CellValue.Error(error)

  private val a = Vector("1", "-2", "3.5", "0", "10", "7", "-4.25", "100", "2", "5")
  private val c = Vector(
    "apple",
    "banana",
    "cherry",
    "date",
    "elder",
    "fig",
    "grape",
    "honeydew",
    "iceberg",
    "jam"
  )
  private val d = Vector(1, 2, 3, 1, 2, 3, 1, 2, 3, 1)

  private val sheet1Name = SheetName.unsafe("Sheet1")
  private val emptyName = SheetName.unsafe("Empty")

  /** Sheet1: A1:A10 numbers, B1:B10 = 2, C1:C10 text, D1:D10 a key column, A12:E12 = 1..5. */
  private val sheet1: Sheet =
    val columns = (0 until 10).foldLeft(Sheet(sheet1Name)) { (s, i) =>
      s.put(ARef.from0(0, i), N(a(i)))
        .put(ARef.from0(1, i), N("2"))
        .put(ARef.from0(2, i), T(c(i)))
        .put(ARef.from0(3, i), CellValue.Number(BigDecimal(d(i))))
    }
    (0 until 5).foldLeft(columns) { (s, j) =>
      s.put(ARef.from0(j, 11), CellValue.Number(BigDecimal(j + 1)))
    }

  private val other: Sheet = (0 until 10).foldLeft(Sheet(SheetName.unsafe("Other"))) { (s, i) =>
    s.put(ARef.from0(0, i), CellValue.Number(BigDecimal((i + 1) * 10)))
  }

  private def withNames(wb: Workbook): Workbook =
    wb.withDefinedName("dyn", "OFFSET(Sheet1!$A$1,0,0,10,1)")
      .withDefinedName("dynTab", "OFFSET(Sheet1!$A$1,0,0,10,4)")
      .withDefinedName("dynRow", "OFFSET(Sheet1!$A$12,0,0,1,5)")
      .withDefinedName("dynKey", "OFFSET(Sheet1!$D$1,0,0,10,1)")
      .withDefinedName("dynTop", "OFFSET(Sheet1!$A$1,0,0,1,1)")
      .withDefinedName("dynOther", "OFFSET(Other!$A$1,0,0,10,1)")
      .withDefinedName("dynInd", "INDIRECT(\"Sheet1!A1:A10\")")
      .withDefinedName("dynIdx", "INDEX(Sheet1!$A$1:$D$10,0,1)")
      .withDefinedName("dynIf", "IF(TRUE,Sheet1!$A$1:$A$10,Sheet1!$B$1:$B$10)")
      .withDefinedName("dynCh", "CHOOSE(2,Sheet1!$B$1:$B$10,Sheet1!$A$1:$A$10)")
      .withDefinedName("dynBad", "INDIRECT(\"nope!A1:A3\")")
      .withDefinedName("dynRef", "OFFSET(Sheet1!$A$1,-1,0,10,1)")
      .withDefinedName("alias", "dyn")
      .withDefinedName("konst", "5")

  private val book: Workbook = withNames(Workbook(Vector(sheet1, other, Sheet(emptyName))))

  /** `formula` as a plain formula cell at `sheet!at`, evaluated at its own position. */
  private def plain(sheet: String, at: String, formula: String): Either[String, CellValue] =
    val ref = ARef.parse(at).fold(msg => fail(msg), identity)
    book.sheets.find(_.name.value == sheet) match
      case None => Left(s"no sheet $sheet")
      case Some(target) =>
        val placed = target.put(ref, CellValue.Formula(formula, None))
        placed.evaluateCell(ref, Clock.system, Some(book.put(placed))).left.map(_.message)

  /** `formula` evaluated without a cell position — an array context, as `xl eval` evaluates. */
  private def unplaced(formula: String): Either[String, CellValue] =
    book.evaluateFormula(s"=$formula", sheet1Name).left.map(_.message)

  private def agree(obtained: Either[String, CellValue], expected: CellValue): Unit =
    (obtained, expected) match
      case (Right(CellValue.Number(got)), CellValue.Number(want)) =>
        assert((got - want).abs < BigDecimal("1e-9"), s"$got is not $want")
      case _ => assertEquals(obtained, Right(expected))

  private val libreOffice: List[(String, CellValue)] = List(
    "COUNTIF(dyn,\">2\")" -> N("5"),
    "INDEX(dyn,3)" -> N("3.5"),
    "MATCH(7,dyn,0)" -> N("6"),
    "VLOOKUP(7,dyn,1,FALSE)" -> N("7"),
    "VLOOKUP(7,dynTab,3,FALSE)" -> T("fig"),
    "HLOOKUP(3,dynRow,1,FALSE)" -> N("3"),
    "INDEX(dynTab,6,3)" -> T("fig"),
    "SUMIF(dyn,\">2\")" -> N("125.5"),
    "SUMIF(dynKey,2,dyn)" -> N("108"),
    "AVERAGEIF(dyn,\">2\")" -> N("25.1"),
    "AVERAGEIF(dynKey,1,dyn)" -> N("0.4375"),
    "COUNTIFS(dyn,\">0\",dynKey,1)" -> N("2"),
    "SUMIFS(dyn,dynKey,1)" -> N("1.75"),
    "AVERAGEIFS(dyn,dynKey,1)" -> N("0.4375"),
    "LARGE(dyn,2)" -> N("10"),
    "SMALL(dyn,1)" -> N("-4.25"),
    "PERCENTILE(dyn,0.5)" -> N("2.75"),
    "QUARTILE(dyn,1)" -> N("0.25"),
    "RANK(7,dyn)" -> N("3"),
    "RANK(-4.25,dyn,1)" -> N("1"),
    "NPV(0.1,dyn)" -> N("59.29205859357761"),
    "SUM(TRANSPOSE(dyn))" -> N("122.25"),
    "COUNTIF(dynInd,\">2\")" -> N("5"),
    "COUNTIF(dynIdx,\">2\")" -> N("5"),
    "COUNTIF(dynIf,\">2\")" -> N("5"),
    "COUNTIF(dynCh,\">2\")" -> N("5"),
    "MATCH(7,dynInd,0)" -> N("6"),
    "INDEX(dynIf,3)" -> N("3.5"),
    "COUNTIF(dynOther,\">20\")" -> N("8"),
    "MATCH(30,dynOther,0)" -> N("3"),
    "MATCH(3,dynRow,0)" -> N("3"),
    "SUMIF(Sheet1!A1:A10,\">2\",dyn)" -> N("125.5"),
    // the sum range is resized to the criteria range's shape, from the reference's top-left cell
    "SUMIF(Sheet1!D1:D10,1,dynTop)" -> N("1.75"),
    "COUNTIF(dynTab,\">2\")" -> N("8"),
    "COUNTIF(dynKey,1)" -> N("4"),
    "HLOOKUP(1,dyn,3,FALSE)" -> N("3.5"),
    "COUNTIF(alias,\">2\")" -> N("5"),
    "NPV(0.1,alias)" -> N("59.29205859357761"),
    // a name that computes no reference is still #VALUE!
    "COUNTIF(konst,\">2\")" -> E(CellError.Value),
    // the error the name's formula evaluates to is the slot's error
    "COUNTIF(dynBad,1)" -> E(CellError.Ref)
  )

  libreOffice.zipWithIndex.foreach { case ((formula, expected), i) =>
    val at = s"A${i + 1}"
    test(s"=$formula at Empty!$at is $expected") {
      agree(plain("Empty", at, formula), expected)
    }
  }

  test("INDEX's column of a dynamic name in a plain cell's value position is intersected") {
    // LibreOffice: Empty!A8 reads row 8 of the column INDEX returns
    agree(plain("Empty", "A8", "INDEX(dyn,0,1)"), N("100"))
    agree(plain("Empty", "A9", "SUM(INDEX(dyn,0,1))"), N("122.25"))
  }

  test("a name computing an off-grid reference is #REF! in a range slot, as OFFSET is") {
    // Excel's OFFSET off the grid is #REF! (LibreOffice says #VALUE! for OFFSET itself, and so
    // for COUNTIF over it): the slot carries the error the reference evaluated to
    agree(plain("Empty", "A1", "COUNTIF(dynRef,\">2\")"), E(CellError.Ref))
    agree(plain("Empty", "A2", "INDEX(dynRef,3)"), E(CellError.Ref))
  }

  test("MAXIFS, MINIFS, XMATCH and XLOOKUP take a dynamic name like a written range") {
    agree(plain("Empty", "A1", "MAXIFS(dyn,dynKey,1)"), N("5"))
    agree(plain("Empty", "A2", "MINIFS(dyn,dynKey,1)"), N("-4.25"))
    agree(plain("Empty", "A3", "MAXIFS(Sheet1!A1:A10,Sheet1!D1:D10,1)"), N("5"))
    agree(plain("Empty", "A4", "XMATCH(7,dyn)"), N("6"))
    agree(plain("Empty", "A5", "XLOOKUP(7,dyn,dynKey)"), N("3"))
  }

  test("a LET name bound to a dynamic name is a reference in a range slot") {
    agree(plain("Empty", "A1", "LET(r,dyn,COUNTIF(r,\">2\"))"), N("5"))
    agree(plain("Empty", "A2", "LET(r,dyn,INDEX(r,3))"), N("3.5"))
    agree(plain("Empty", "A3", "LET(r,dyn,MATCH(7,r,0))"), N("6"))
    agree(plain("Empty", "A4", "LET(t,dynTab,VLOOKUP(7,t,3,FALSE))"), T("fig"))
    agree(plain("Empty", "A5", "LET(r,dyn,k,dynKey,SUMIFS(r,k,1))"), N("1.75"))
  }

  test("a LET name bound to OFFSET, INDIRECT, INDEX or a written range is a reference too") {
    agree(plain("Sheet1", "H1", "LET(r,OFFSET(A1,0,0,10,1),COUNTIF(r,\">2\"))"), N("5"))
    agree(plain("Sheet1", "H2", "LET(r,INDIRECT(\"A1:A10\"),LARGE(r,2))"), N("10"))
    agree(plain("Sheet1", "H3", "LET(r,INDEX(A1:D10,0,1),SUMIF(r,\">2\"))"), N("125.5"))
    agree(plain("Sheet1", "H4", "LET(r,A1:A10,COUNTIF(r,\">2\"))"), N("5"))
    agree(plain("Sheet1", "H5", "LET(x,A1,COUNTIF(x,1))"), N("1"))
    // nested LETs: the inner name is bound to the outer one
    agree(plain("Sheet1", "H6", "LET(r,dyn,LET(s,r,COUNTIF(s,\">2\")))"), N("5"))
  }

  test("INDEX over a LET name bound to an array indexes the array") {
    agree(plain("Empty", "A1", "LET(x,SEQUENCE(3),INDEX(x,2))"), N("2"))
    agree(unplaced("LET(x,SEQUENCE(3,2),INDEX(x,3,2))"), N("6"))
    agree(unplaced("LET(r,dyn,INDEX(r,3))"), N("3.5"))
    agree(unplaced("LET(r,dyn,AREAS(r))"), N("1"))
  }

  test("a LET name bound to a value is #VALUE! in a range slot, as a name bound to one is") {
    agree(plain("Empty", "A1", "LET(r,5,COUNTIF(r,\">2\"))"), E(CellError.Value))
    agree(plain("Empty", "A2", "LET(r,konst,COUNTIF(r,\">2\"))"), E(CellError.Value))
  }

  test("without a cell position (an array context) dynamic names and LET names resolve too") {
    agree(unplaced("COUNTIF(dyn,\">2\")"), N("5"))
    agree(unplaced("INDEX(dyn,3)"), N("3.5"))
    agree(unplaced("VLOOKUP(7,dynTab,3,FALSE)"), T("fig"))
    agree(unplaced("LET(r,dyn,COUNTIF(r,\">2\"))"), N("5"))
    agree(unplaced("LET(r,OFFSET(A1,0,0,10,1),COUNTIF(r,\">2\")+SUM(r))"), N("127.25"))
    agree(unplaced("SUMPRODUCT(LET(r,dyn,COUNTIF(r,r)))"), N("10"))
  }

  test("a LET name in a range slot prints and re-parses as written") {
    val text = "=LET(r,dyn,COUNTIF(r,\">2\"))"
    val parsed = FormulaParser.parse(text).fold(err => fail(err.toString), identity)
    val printed = FormulaPrinter.printFileForm(parsed)
    assertEquals(printed, "LET(r,dyn,COUNTIF(r,\">2\"))")
    assertEquals(FormulaParser.parse(s"=$printed"), Right(parsed))
  }

  test("a name whose reference reads itself through a range slot is a clean name-cycle error") {
    // Excel flags the circular reference; xl's name-cycle guard (#691) stops the walk the second
    // time `loop` is entered and the cell reports that cycle, uncached — never a stack overflow
    val looped = book.withDefinedName("loop", "OFFSET(Sheet1!$A$1,0,0,COUNTIF(loop,\">2\"),1)")
    val placed = Sheet(emptyName).put(ref"A1", CellValue.Formula("COUNTIF(loop,\">2\")", None))
    val result = placed.evaluateCell(ref"A1", Clock.system, Some(looped.put(placed)))
    assertEquals(
      result.left.map(_.message),
      Left(
        "Formula error in 'COUNTIF(loop,\">2\")': Evaluation failed: Defined name cycle " +
          "detected while resolving 'loop'. (COUNTIF(loop,\">2\"))"
      )
    )
  }

  // ----- the static resolution decides whether a name is evaluated (review of #717) -----

  private val unread: (String, Option[SheetName]) => Either[EvalError, Any] =
    (name, _) => fail(s"$name must not be evaluated: its static resolution already failed")

  private def slot(
    wb: Workbook,
    name: String,
    evalName: (String, Option[SheetName]) => Either[EvalError, Any] = unread
  ): Either[EvalError, (Sheet, CellRange)] =
    Evaluator.resolveRangeSlot(
      TExpr.RangeLocation.Name(name, None),
      sheet1,
      Some(wb),
      Evaluator.NameTrail.empty,
      Map.empty,
      evalName
    )

  test("a name that fails statically is never evaluated, and keeps its own diagnostic") {
    val wb = book
      .withDefinedName("badValue", "#VALUE!")
      .withDefinedName("badNa", "#N/A")
      .withDefinedName("cycA", "cycB")
      .withDefinedName("cycB", "cycA")
    assertEquals(
      slot(wb, "badValue"),
      Left(EvalError.ErrorValue(CellError.Value, Some("Defined name 'badValue' is #VALUE!")))
    )
    assertEquals(
      slot(wb, "badNa"),
      Left(EvalError.ErrorValue(CellError.NA, Some("Defined name 'badNa' is #N/A")))
    )
    assertEquals(
      slot(wb, "cycA"),
      Left(EvalError.EvalFailed("Defined name cycle detected while resolving 'cycA'.", None))
    )
    assertEquals(
      slot(wb, "nope"),
      Left(EvalError.EvalFailed("Name 'nope' is not defined in this workbook.", None))
    )
    // a written range and a name chain ending at one resolve without evaluation too
    assertEquals(
      slot(wb.withDefinedName("written", "Sheet1!$A$1:$A$3"), "written").map(_._2.toA1),
      Right("A1:A3")
    )
  }

  test("a name bound to a constant is #VALUE! statically, never evaluated") {
    val wb = book
      .withDefinedName("negative", "-0.05")
      .withDefinedName("pct", "5%")
      .withDefinedName("label", "\"x\"")
    List("konst" -> "5", "negative" -> "-0.05", "pct" -> "5%", "label" -> "\"x\"").foreach {
      (name, refersTo) =>
        assertEquals(
          slot(wb, name),
          Left(
            EvalError.ErrorValue(
              CellError.Value,
              Some(s"name '$name' does not refer to a range (refersTo: $refersTo)")
            )
          )
        )
    }
  }

  test("a name bound to a formula is evaluated; a value it computes keeps the static diagnostic") {
    val wb = book.withDefinedName("calc", "SUM(Sheet1!$A$1:$A$3)")
    val notARange =
      EvalError.ErrorValue(
        CellError.Value,
        Some("name 'calc' does not refer to a range (refersTo: SUM(Sheet1!$A$1:$A$3))")
      )
    // only a value the evaluation computes is relabelled "does not refer to a range"
    assertEquals(slot(wb, "calc", (_, _) => Right(BigDecimal(5))), Left(notARange))
    // an error the evaluation itself produces is the slot's as it is: a bare #VALUE! (from
    // OFFSET(A1,"x",0), say) is not blamed on the name's shape
    val bare = EvalError.ErrorValue(CellError.Value, None)
    assertEquals(slot(wb, "calc", (_, _) => Left(bare)), Left(bare))
    val failed = EvalError.EvalFailed("boom", None)
    assertEquals(slot(wb, "calc", (_, _) => Left(failed)), Left(failed))
    assertEquals(
      slot(wb, "calc", (_, _) => Right(CellValue.Error(CellError.Ref))),
      Left(EvalError.ErrorValue(CellError.Ref))
    )
    // the evaluation is handed the name itself, unqualified as written
    assertEquals(
      slot(
        book,
        "dyn",
        (name, scope) =>
          assertEquals((name, scope), ("dyn", None))
          Right(RangeOperand(other, CellRange(ref"A1", ref"A2")))
      ),
      Right((other, CellRange(ref"A1", ref"A2")))
    )
  }

  test("a name whose OFFSET fails carries that error, not the name's shape") {
    // OFFSET(A1,"x",0) is #VALUE! on its text row offset (#709); the slot keeps that error
    // rather than failing because the name "does not refer to a range"
    val wb = book.withDefinedName("badOffset", "OFFSET(Sheet1!$A$1,\"x\",0)")
    val placed = Sheet(emptyName).put(ref"A1", CellValue.Formula("COUNTIF(badOffset,1)", None))
    assertEquals(
      placed.evaluateCell(ref"A1", Clock.system, Some(wb.put(placed))),
      Right(CellValue.Error(CellError.Value))
    )
  }

  test("a hand-built LET name outside its LET is a clean failure, never an exception") {
    // the parser never emits it: a Binding exists only inside the LET that binds it
    assertEquals(
      Evaluator.resolveRangeSlot(
        TExpr.RangeLocation.Binding("nope"),
        sheet1,
        Some(book),
        Evaluator.NameTrail.empty,
        Map.empty,
        unread
      ),
      Left(EvalError.EvalFailed("LET name 'nope' is not in scope", None))
    )
  }

  // ----- the typed Aggregate node resolves its slot like a function's range slot -----

  test("a typed Aggregate over a dynamic name or a LET name reads the reference it denotes") {
    val sumDyn = TExpr.Aggregate("SUM", TExpr.RangeLocation.Name("dyn", None))
    val sumLet = TExpr.Let(
      List("r" -> TExpr.NameRef("dyn")),
      TExpr.Aggregate("SUM", TExpr.RangeLocation.Binding("r"))
    )
    List(Evaluator.instance, Evaluator.instance.withArrayResults).foreach { evaluator =>
      assertEquals(
        evaluator.eval(sumDyn, sheet1, workbook = Some(book)),
        Right(BigDecimal("122.25"))
      )
      assertEquals(
        evaluator.eval(sumLet, sheet1, workbook = Some(book)),
        Right(BigDecimal("122.25"))
      )
    }
    // a name computing a value is still #VALUE!, triaged as SUM triages an error argument
    assertEquals(
      Evaluator.instance
        .eval(
          TExpr.Aggregate("COUNT", TExpr.RangeLocation.Name("konst", None)),
          sheet1,
          workbook = Some(book)
        ),
      Right(BigDecimal(0))
    )
  }

  // ----- every wrapper between a LET and its range slot keeps the binding a reference -----

  List(
    "-LET(r,dyn,COUNTIF(r,\">2\"))" -> N("-5"),
    "LET(r,dyn,-COUNTIF(r,\">2\"))" -> N("-5"),
    "LET(r,dyn,+COUNTIF(r,\">2\"))" -> N("5"),
    "LET(r,dyn,COUNTIF(r,\">2\")%)" -> N("0.05"),
    "LET(r,dyn,COUNTIF(r,\">2\")&\"\")" -> T("5"),
    "LET(r,dyn,COUNTIF(r,\">2\")=5)" -> CellValue.Bool(true),
    "LET(r,dyn,LEFT(\"abcdef\",COUNTIF(r,\">2\")))" -> T("abcde"),
    "LET(r,dyn,ROUND(COUNTIF(r,\">2\")/3,1))" -> N("1.7"),
    "LET(r,dyn,DATE(2000,1,COUNTIF(r,\">2\"))+0)" -> N("36530"),
    "LET(r,dyn,IF(TRUE,COUNTIF(r,\">2\"),0))" -> N("5"),
    "SUMPRODUCT(-LET(r,dyn,COUNTIF(r,\">2\")))" -> N("-5")
  ).foreach { case (formula, expected) =>
    test(s"=$formula reads its LET name as a reference in array mode too") {
      agree(unplaced(formula), expected)
      agree(plain("Empty", "A1", formula), expected)
    }
  }

  test("a range-slot read of one LET name leaves the other bindings' array values unchanged") {
    // r is read by COUNTIF's range slot, so the array-mode LET binds references; s, an INDEX read
    // by FILTER as an array, must give the same values as when nothing makes the LET bind them
    val withSlot =
      "LET(r,IF(TRUE,A1:A3,B1:B3),s,INDEX(A1:B3,0,1),COUNTIF(r,\">0\")+SUM(FILTER(s,s>1)))"
    val withoutSlot = "LET(r,IF(TRUE,A1:A3,B1:B3),s,INDEX(A1:B3,0,1),SUM(FILTER(s,s>1)))"
    // A1:A3 is 1, -2, 3.5: two are positive, and 3.5 is the only one above 1
    agree(unplaced(withoutSlot), N("3.5"))
    agree(unplaced("COUNTIF(A1:A3,\">0\")"), N("2"))
    agree(unplaced(withSlot), N("5.5"))
    agree(plain("Sheet1", "H1", withSlot), N("5.5"))
  }

  test("OFFSET's reference binding leaves another binding's value read unchanged") {
    // OFFSET makes the array-mode LET bind references; x, an array value, must still read as one.
    // Array mode: x is {1,-2,3.5}, SUM 2.5, plus B6's 2. A plain cell is Excel's legacy formula:
    // A1:A3*1 is intersected at row 1 (1), plus 2 — with or without the OFFSET, in both modes.
    val withOffset = "LET(c,B5,x,A1:A3*1,OFFSET(c,1,0)+SUM(x))"
    val withoutOffset = "LET(x,A1:A3*1,SUM(x))+Sheet1!B6"
    agree(unplaced(withOffset), N("4.5"))
    agree(unplaced(withoutOffset), N("4.5"))
    agree(plain("Sheet1", "H1", withOffset), N("3"))
    agree(plain("Sheet1", "H1", withoutOffset), N("3"))
  }

  // ----- a reference-reading position sees a LET name as its reference, whatever else the LET reads -----

  // Each pair: the LET with an unrelated range-slot read of another LET name, and the same LET
  // alone plus that read outside it. Excel's LET always binds the reference, so both forms agree
  // in array mode too (B5:B6 hold 2, COUNTIF(dyn,">2") is 5). OFFSET's base was #VALUE! in array
  // mode without the unrelated read before the LET bound references for it.
  List(
    ("OFFSET(c,1,0)", "B5", N("7")),
    ("SUM(OFFSET(c,0,0,2,1))", "B5", N("9")),
    ("SUM(OFFSET(IF(TRUE,c,c),1,0))", "B5", N("7")),
    ("ROWS(c)", "B5:B6", N("7")),
    ("COLUMNS(c)", "B5:B6", N("6")),
    ("AREAS(c)", "B5", N("6")),
    ("INDEX(c,1)", "B5", N("7")),
    ("@c", "B5", N("7"))
  ).foreach { case (read, bound, expected) =>
    val together = s"LET(c,$bound,r,dyn,$read+COUNTIF(r,\">2\"))"
    val apart = s"LET(c,$bound,$read)+COUNTIF(dyn,\">2\")"
    test(s"$read over a LET name agrees with and without an unrelated range-slot read") {
      agree(unplaced(together), expected)
      agree(unplaced(apart), expected)
      agree(plain("Sheet1", "H1", together), expected)
      agree(plain("Sheet1", "H1", apart), expected)
    }
  }

  // ----- a LET name through the reference operators' operand slots (union, intersection, AREAS) -----

  // A LET name in an operand slot of `,`, ` ` or AREAS parses as a range location (Binding) and
  // resolves to the reference it is bound to. LibreOffice 24.2 has no LET: each value is its
  // recalculation of the same formula with the bound reference written in place (`~` is its union).
  List(
    "LET(a,OFFSET(A1,0,0,3,1),AREAS((a,B1)))" -> N("2"),
    "LET(a,OFFSET(A1,0,0,3,1),SUM((a,B1)))" -> N("4.5"),
    "LET(a,OFFSET(A1,0,0,3,1),COUNT((a,B1)))" -> N("4"),
    "LET(a,dynTab,AREAS((a,B1,C1)))" -> N("3"),
    "LET(a,dynTab,SUM(a B1:B3))" -> N("6"),
    "LET(a,dynTab,COUNT(a B1:B3))" -> N("3"),
    "LET(a,dynTab,ROWS(a B1:B3))" -> N("3"),
    "LET(a,dyn,SUM(a A2:B3))" -> N("1.5"),
    "LET(a,dyn,AREAS(a A2:B3))" -> N("1")
  ).foreach { case (formula, expected) =>
    test(s"=$formula reads the LET name's reference through the operator, in both modes") {
      agree(plain("Sheet1", "H1", formula), expected)
      agree(unplaced(formula), expected)
    }
  }

  // ----- an unrelated OFFSET leaves a multi-cell binding's array value unchanged -----

  test("an unrelated OFFSET in an array LET leaves a multi-cell binding's values as before") {
    // OFFSET makes the array-mode LET bind x as its reference; x*2 must still be the array
    // {2,-4,7}. LibreOffice 24.2 (no LET), element by element via SUMPRODUCT of the written form
    // OFFSET(B1,1,0)+INDEX(A1:B3,0,1)*2: 4, -2, 9.
    def spill(formula: String): List[CellValue] =
      sheet1
        .evaluateArrayFormula(s"=$formula", ref"J1", Clock.system, Some(book))
        .fold(
          error => fail(error.message),
          (spilled, range) => range.cells.map(at => spilled(at).value).toList
        )
    val expected = List(N("4"), N("-2"), N("9"))
    assertEquals(spill("LET(x,INDEX(A1:B3,0,1),OFFSET(B1,1,0)+x*2)"), expected)
    assertEquals(spill("LET(x,INDEX(A1:B3,0,1),2+x*2)"), expected)
    assertEquals(spill("OFFSET(B1,1,0)+INDEX(A1:B3,0,1)*2"), expected)
  }

  // ----- one smoke case per family whose range slot now resolves dynamically -----

  /** H: B1:B2 holidays, C1:C3 cash flows, D1:D3 their dates, E1:E3 keys, F1:F3 results. */
  private val familyBook: Workbook =
    val hName = SheetName.unsafe("H")
    def num(n: String) = CellValue.Number(BigDecimal(n))
    val h = Sheet(hName)
      .put(ref"B1", num("45293"))
      .put(ref"B2", num("45294"))
      .put(ref"C1", num("-1000"))
      .put(ref"C2", num("600"))
      .put(ref"C3", num("600"))
      .put(ref"D1", num("45292"))
      .put(ref"D2", num("45474"))
      .put(ref"D3", num("45658"))
      .put(ref"E1", num("1"))
      .put(ref"E2", num("3"))
      .put(ref"E3", num("5"))
      .put(ref"F1", T("a"))
      .put(ref"F2", T("b"))
      .put(ref"F3", T("c"))
    withNames(Workbook(Vector(sheet1, other, Sheet(emptyName), h)))
      .withDefinedName("holsDyn", "OFFSET(H!$B$1,0,0,2,1)")
      .withDefinedName("valsDyn", "OFFSET(H!$C$1,0,0,3,1)")
      .withDefinedName("datesDyn", "OFFSET(H!$D$1,0,0,3,1)")
      .withDefinedName("keysDyn", "OFFSET(H!$E$1,0,0,3,1)")
      .withDefinedName("resDyn", "OFFSET(H!$F$1,0,0,3,1)")

  private def family(formula: String): Either[String, CellValue] =
    val placed = Sheet(emptyName).put(ref"A1", CellValue.Formula(formula, None))
    placed.evaluateCell(ref"A1", Clock.system, Some(familyBook.put(placed))).left.map(_.message)

  private def near(obtained: Either[String, CellValue], want: String, tolerance: String): Unit =
    obtained match
      case Right(CellValue.Number(got)) =>
        assert((got - BigDecimal(want)).abs < BigDecimal(tolerance), s"$got is not $want")
      case other => fail(s"$other is not $want")

  test("SORT and UNIQUE take a dynamic name or a LET name like the written range") {
    // a plain cell shows the sorted array's top-left element
    agree(family("SORT(dyn)"), N("-4.25"))
    agree(family("SORT(dyn,1,-1)"), N("100"))
    agree(family("LET(r,dyn,SORT(r,1,-1))"), N("100"))
    agree(family("SORT(Sheet1!A1:A10,1,-1)"), N("100"))
    agree(family("SUM(UNIQUE(dynKey))"), N("6"))
    agree(family("LET(k,dynKey,SUM(UNIQUE(k)))"), N("6"))
    agree(family("SUM(UNIQUE(Sheet1!D1:D10))"), N("6"))
  }

  // LibreOffice 24.2 values for the same formulas with the OFFSET written in place of the name
  test("NETWORKDAYS and WORKDAY take dynamic holidays") {
    agree(family("NETWORKDAYS(DATE(2024,1,1),DATE(2024,1,12),holsDyn)"), N("8"))
    agree(family("LET(h,holsDyn,NETWORKDAYS(DATE(2024,1,1),DATE(2024,1,12),h))"), N("8"))
    agree(family("WORKDAY(DATE(2024,1,1),5,holsDyn)+0"), N("45301"))
  }

  test("XNPV, XIRR and IRR take dynamic values and dates") {
    near(family("XNPV(0.1,valsDyn,datesDyn)"), "117.464382846325", "1e-9")
    near(family("LET(v,valsDyn,d,datesDyn,XNPV(0.1,v,d))"), "117.464382846325", "1e-9")
    near(family("XIRR(valsDyn,datesDyn)"), "0.278158944317214", "1e-7")
    near(family("IRR(valsDyn)"), "0.130662386291808", "1e-9")
  }

  test("LOOKUP takes a dynamic lookup vector and result vector") {
    agree(family("LOOKUP(4,keysDyn,resDyn)"), T("b"))
    agree(family("LOOKUP(5,keysDyn,resDyn)"), T("c"))
    agree(family("LET(k,keysDyn,r,resDyn,LOOKUP(4,k,r))"), T("b"))
  }

  // ----- name scope through a range slot -----

  test("a sheet-scoped name shadows the workbook name in a range slot, per reading sheet") {
    // `amt` is Other!A1:A3 for the workbook and Sheet1!A1:A3 on Sheet1: SUMIF on each sheet reads
    // its own, and a sheet-qualified read from Empty reads Sheet1's
    val scoped = book
      .withDefinedName("amt", "Other!$A$1:$A$3")
      .withDefinedName("amt", "Sheet1!$A$1:$A$3", sheet1Name)
      .fold(error => fail(error.message), identity)
    def at(sheet: String, formula: String): Either[String, CellValue] =
      val ref = ref"H1"
      scoped.sheets.find(_.name.value == sheet) match
        case None => Left(s"no sheet $sheet")
        case Some(target) =>
          val placed = target.put(ref, CellValue.Formula(formula, None))
          placed.evaluateCell(ref, Clock.system, Some(scoped.put(placed))).left.map(_.message)
    // Sheet1!A1:A3 is 1, -2, 3.5; Other!A1:A3 is 10, 20, 30
    agree(at("Sheet1", "SUMIF(amt,\">0\")"), N("4.5"))
    agree(at("Other", "SUMIF(amt,\">0\")"), N("60"))
    agree(at("Empty", "SUMIF(amt,\">0\")"), N("60"))
    agree(at("Empty", "SUMIF(Sheet1!amt,\">0\")"), N("4.5"))
    // the resized form reads the same scope: A1:A3's shape from B1 on the reading sheet's name
    agree(at("Sheet1", "SUMIF(amt,\">0\",B1)"), N("4"))

    // the graph's name memos key on the reading sheet: each reader gets its own sheet's cells
    val readers = scoped
      .put(sheet1.put(ref"H1", CellValue.Formula("SUMIF(amt,\">0\",B1)", None)))
      .put(other.put(ref"H1", CellValue.Formula("SUMIF(amt,\">0\",B1)", None)))
    val graph = com.tjclp.xl.formula.graph.DependencyGraph.fromWorkbook(readers)
    def deps(sheet: SheetName) = graph.getOrElse(
      com.tjclp.xl.formula.graph.DependencyGraph.QualifiedRef(sheet, ref"H1"),
      Set.empty
    )
    assert(deps(sheet1Name).forall(_.sheet == sheet1Name), deps(sheet1Name).toString)
    assert(
      deps(SheetName.unsafe("Other")).forall(_.sheet == SheetName.unsafe("Other")),
      deps(SheetName.unsafe("Other")).toString
    )
    assertEquals(
      com.tjclp.xl.formula.graph.DependencyGraph.dynamicCells(readers),
      Set.empty[com.tjclp.xl.formula.graph.DependencyGraph.QualifiedRef]
    )
    val result = readers.recalculate()
    assertEquals(result.errors, Vector.empty)
    assertEquals(result.evaluated(sheet1Name)(ref"H1"), N("4"))
    // Other's B column is empty: A1:A3's positives pair with nothing there
    assertEquals(result.evaluated(SheetName.unsafe("Other"))(ref"H1"), N("0"))
  }

  test("a sheet-qualified computed name in a range slot evaluates in its sheet's scope") {
    // a sheet-scoped `dynLocal` on Other computes an OFFSET from Other!A1; read qualified from Empty
    val scoped = book
      .withDefinedName("dynLocal", "OFFSET($A$1,0,0,10,1)", SheetName.unsafe("Other"))
      .fold(error => fail(error.message), identity)
    val placed = Sheet(emptyName)
      .put(ref"A1", CellValue.Formula("COUNTIF(Other!dynLocal,\">20\")", None))
      .put(ref"A2", CellValue.Formula("MATCH(30,Other!dynLocal,0)", None))
    val wb = scoped.put(placed)
    agree(placed.evaluateCell(ref"A1", Clock.system, Some(wb)).left.map(_.message), N("8"))
    agree(placed.evaluateCell(ref"A2", Clock.system, Some(wb)).left.map(_.message), N("3"))
    agree(
      scoped.evaluateFormula("=COUNTIF(Other!dynLocal,\">20\")", emptyName).left.map(_.message),
      N("8")
    )
  }

  // ----- a LET name as an operand of the reference operators -----

  List(
    "LET(r,dyn,SUM((r,Sheet1!B1)))" -> N("124.25"),
    "LET(r,dyn,COUNT((r,Sheet1!B1)))" -> N("11"),
    "LET(r,dyn,AREAS((r,Sheet1!B1)))" -> N("2"),
    "LET(r,OFFSET(Sheet1!A1,0,0,3,1),AREAS(r))" -> N("1"),
    "LET(r,OFFSET(Sheet1!A1,0,0,3,1),SUM(r Sheet1!A2:B2))" -> N("-2"),
    "LET(r,dyn,SUM(r Sheet1!A3:B3))" -> N("3.5"),
    "LET(r,dyn,COUNTIF(r,\">2\")+AREAS((r,Sheet1!B1)))" -> N("7")
  ).foreach { case (formula, expected) =>
    test(s"=$formula unions and intersects the LET name's reference in both modes") {
      agree(plain("Empty", "A1", formula), expected)
      agree(unplaced(formula), expected)
    }
  }

  test("recalculation reads a dynamic range's precedent formulas before the range slot") {
    // A2 is an uncached formula inside the dynamic range: the readers must see its value (20)
    val withFormula = sheet1
      .put(ref"A2", CellValue.Formula("B1*10", None))
      .put(ref"H1", CellValue.Formula("COUNTIF(dyn,\">2\")", None))
      .put(ref"H2", CellValue.Formula("LET(r,OFFSET(A1,0,0,10,1),COUNTIF(r,\">2\"))", None))
      .put(ref"H3", CellValue.Formula("MATCH(20,dyn,0)", None))
    val result = withNames(Workbook(Vector(withFormula, other, Sheet(emptyName)))).recalculate()
    assert(result.isClean, result.errors.map(_.render).mkString("; "))
    assertEquals(result.evaluated(sheet1Name)(ref"H1"), N("6"))
    assertEquals(result.evaluated(sheet1Name)(ref"H2"), N("6"))
    assertEquals(result.evaluated(sheet1Name)(ref"H3"), N("2"))
  }
