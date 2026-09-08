package com.tjclp.xl.formula

import com.tjclp.xl.{*, given}
import com.tjclp.xl.formula.eval.{DependentRecalculation, StructuralEditor}

import munit.FunSuite

class CacheInvalidationIntegritySpec extends FunSuite:
  private def num(value: Int): CellValue = CellValue.Number(BigDecimal(value))
  private def formula(text: String, cached: Int): CellValue =
    CellValue.Formula(text, Some(num(cached)))

  private def cache(wb: Workbook, sheet: String, ref: ARef): Option[CellValue] =
    wb(SheetName.unsafe(sheet)).fold(error => fail(error.message), identity)(ref).value match
      case CellValue.Formula(_, value, _) => value
      case other => fail(s"Expected formula at $sheet!${ref.toA1}, got $other")

  private def failingChain: Workbook = Workbook(
    Sheet("S")
      .put(ref"A1", formula("UNSUPPORTED(1)", 6))
      .put(ref"A2", formula("A1*2", 12))
      .put(ref"A3", formula("SUM(A1:A2)", 18))
      .put(ref"B1", formula("10+1", 11))
  )

  test("GH-563: failed formulas and their dependent caches are withdrawn and reported") {
    val before = failingChain
    val result = before.recalculate()
    List(ref"A1", ref"A2", ref"A3").foreach { ref =>
      assertEquals(cache(result.workbook, "S", ref), None)
      assert(result.errors.exists(_.ref == ref), s"Missing diagnostic for ${ref.toA1}")
      assert(!result.evaluated.getOrElse(SheetName.unsafe("S"), Map.empty).contains(ref))
    }
    assertEquals(cache(result.workbook, "S", ref"B1"), Some(num(11)))
    assertEquals(cache(before, "S", ref"A1"), Some(num(6)))
    assert(!result.certified)
  }

  test("GH-563: failures propagate across sheets and into financial range readers") {
    val wb = failingChain.put(
      Sheet("Other")
        .put(ref"C1", formula("S!A2+1", 13))
        .put(ref"C2", formula("NPV(0,S!A1:A3)", 36))
    )
    val result = wb.recalculate()
    assertEquals(cache(result.workbook, "Other", ref"C1"), None)
    assertEquals(cache(result.workbook, "Other", ref"C2"), None)
    assert(result.errors.exists(e => e.sheet.value == "Other" && e.ref == ref"C1"))
    assert(result.errors.exists(e => e.sheet.value == "Other" && e.ref == ref"C2"))
  }

  test("GH-563: a guard cannot certify a fallback banked behind a host failure") {
    val wb = failingChain.put(
      failingChain.sheets.head
        .put(ref"C1", formula("IFERROR(A1,7)", 6))
    )
    val result = wb.recalculate()
    assertEquals(cache(result.workbook, "S", ref"C1"), None)
    assert(result.errors.exists(_.ref == ref"C1"))
  }

  test("GH-563: genuine Excel error values and their guards still evaluate normally") {
    val wb = Workbook(
      Sheet("S")
        .put(ref"A1", formula("1/0", 99))
        .put(ref"A2", formula("IFERROR(A1,7)", 99))
    )
    val result = wb.recalculate()
    assertEquals(result.errors, Vector.empty)
    assertEquals(cache(result.workbook, "S", ref"A1"), Some(CellValue.Error(CellError.Div0)))
    assertEquals(cache(result.workbook, "S", ref"A2"), Some(num(7)))
  }

  test("GH-563: a dynamic reader cannot fall back to a failed precedent's loaded cache") {
    val wb = failingChain.put(
      failingChain.sheets.head
        .put(ref"C1", formula("INDIRECT(\"A1\")*2", 12))
    )
    assertEquals(cache(wb.recalculate().workbook, "S", ref"C1"), None)
  }

  test("GH-563: non-iterated cycles and dynamic readers of them lose their stale caches") {
    val wb = Workbook(
      Sheet("S")
        .put(ref"A1", formula("A2", 1))
        .put(ref"A2", formula("A1", 1))
        .put(ref"B1", formula("INDIRECT(\"A1\")+1", 2))
    )
    val result = wb.recalculate()
    List(ref"A1", ref"A2", ref"B1").foreach(ref =>
      assertEquals(cache(result.workbook, "S", ref), None)
    )
  }

  test("GH-563: sequential and parallel failure results agree") {
    val wb = (1 to 200).foldLeft(failingChain) { (book, i) =>
      book.put(book.sheets.head.put(ARef.from0(i, 4), formula("A2+1", 13)))
    }
    assertEquals(wb.recalculateParallel(4), wb.recalculate())
  }

  private def namedWorkbook: Workbook =
    val data = (1 to 10).foldLeft(Sheet("Data")) { (sheet, row) =>
      sheet.put(ARef.from0(0, row - 1), num(row))
    }
    Workbook(
      data,
      Sheet("Other")
        .put(ref"B2", formula("SUM(Multi)", 33))
        .put(ref"B3", formula("B2*2", 66))
        .put(ref"C1", formula("1+1", 2))
    )
      .withDefinedName("Multi", "Data!$A$1:$A$3,Data!$A$8:$A$10")

  test("GH-507: structural edits invalidate unknown named dependencies and their consumers") {
    val edited = StructuralEditor
      .deleteRowsChecked(namedWorkbook, SheetName.unsafe("Data"), 1, 1)
      .fold(error => fail(error.message), identity)
    assertEquals(cache(edited, "Other", ref"B2"), None)
    assertEquals(cache(edited, "Other", ref"B3"), None)
    assertEquals(cache(edited, "Other", ref"C1"), Some(num(2)))
  }

  test("GH-507: aliases of unsupported definitions are equally unknown") {
    val before = namedWorkbook.withDefinedName("Alias", "Multi")
    val other = before.sheets(1).put(ref"B2", formula("SUM(Alias)", 33))
    val edited = StructuralEditor
      .deleteRowsChecked(before.put(other), SheetName.unsafe("Data"), 1, 1)
      .fold(error => fail(error.message), identity)
    assertEquals(cache(edited, "Other", ref"B2"), None)
    assertEquals(cache(edited, "Other", ref"B3"), None)
  }

  test("GH-507: name-like string literals are not dependencies") {
    val before = namedWorkbook.put(
      namedWorkbook
        .sheets(1)
        .put(ref"D1", formula("LEN(\"Multi\")", 5))
    )
    val edited = StructuralEditor
      .deleteRowsChecked(before, SheetName.unsafe("Data"), 1, 1)
      .fold(error => fail(error.message), identity)
    assertEquals(cache(edited, "Other", ref"D1"), Some(num(5)))
  }

  test("GH-507: a local valid name shadows an unsupported global definition") {
    val source = namedWorkbook
    val local = com.tjclp.xl.workbooks.DefinedName("Multi", "Other!$C$1", localSheetId = Some(1))
    val before = source
      .copy(metadata =
        source.metadata.copy(
          definedNames = source.metadata.definedNames :+ local
        )
      )
      .put(
        source
          .sheets(1)
          .put(ref"B2", formula("SUM(Multi)", 2))
          .put(ref"B3", formula("B2*2", 4))
      )
    val edited = StructuralEditor
      .deleteRowsChecked(before, SheetName.unsafe("Data"), 1, 1)
      .fold(error => fail(error.message), identity)
    assertEquals(cache(edited, "Other", ref"B2"), Some(num(2)))
  }

  test("GH-507: an unsafe intersection-name rewrite is refused before mutating the workbook") {
    val before = namedWorkbook.withDefinedName("Multi", "Data!$A$1:$A$8 Data!$A$5:$A$10")
    val result = StructuralEditor.deleteRowsChecked(before, SheetName.unsafe("Data"), 1, 1)
    assert(result.isLeft)
    assert(result.left.toOption.exists(_.message.contains("Multi")))
    assertEquals(cache(before, "Other", ref"B2"), Some(num(33)))
  }

  test(
    "GH-507: unsupported definitions mentioning other sheets or quoted text do not block an edit"
  ) {
    val before = namedWorkbook
      .withDefinedName("Unrelated", "UNSUPPORTED(OtherData!A1)")
      .withDefinedName("Quoted", "UNSUPPORTED(\"Data!A1\")")
      .withDefinedName("External", "UNSUPPORTED('[1]Data'!A1)")
    assert(StructuralEditor.deleteRowsChecked(before, SheetName.unsafe("Data"), 1, 1).isRight)
  }

  test("GH-563: pinned external caches survive an unrelated formula failure") {
    val before = failingChain.put(
      failingChain.sheets.head
        .put(ref"D1", formula("'[1]External'!A1", 42))
    )
    assertEquals(cache(before.recalculate().workbook, "S", ref"D1"), Some(num(42)))
  }

  test("GH-563: recalculation is stable after withdrawing failed caches") {
    val first = failingChain.recalculate()
    assertEquals(first.workbook.recalculate(), first)
  }

  // ===== GH-606: blind readers are dirty only when the edit lies inside their textual reach =====

  private def afterEdit(wb: Workbook, sheet: String, refs: ARef*) =
    DependentRecalculation.recalculateAfterEdit(
      wb,
      SheetName.unsafe(sheet),
      refs.toSet,
      Clock.system
    )

  /**
   * `Data!C3` is a parseable formula over `Data!A1`; every `Calc` formula is a blind reader whose
   * text names only cells the scanner can bound (`ZZZNOTAFUNC` is a stable parse failure).
   */
  private def boundedBlindWorkbook: Workbook = Workbook(
    Sheet("Data")
      .put(ref"A1", num(2))
      .put(ref"C3", formula("A1*2", 4)),
    Sheet("Calc")
      .put(ref"C1", num(1))
      .put(ref"C3", num(5))
      .put(ref"B2", formula("ZZZNOTAFUNC(C3)", 7))
      .put(ref"B3", formula("B2*2", 14))
      .put(ref"D1", formula("ZZZNOTAFUNC(\"B16\")", 1))
      .put(ref"D2", formula("ZZZNOTAFUNC(#REF!)", 2))
      .put(ref"D3", formula("ZZZNOTAFUNC(PI())", 3))
      .put(ref"F1", formula("ZZZNOTAFUNC(Data!C3)", 8))
      .put(ref"G1", formula("ZZZNOTAFUNC('Q1 Data'!B16)", 9))
      .put(ref"G2", formula("ZZZNOTAFUNC(3:3)", 10))
      .put(ref"G3", formula("ZZZNOTAFUNC(H:H)", 11)),
    Sheet("Q1 Data").put(ref"B16", num(16))
  )

  private val boundedCalcRefs =
    List(ref"B2", ref"B3", ref"D1", ref"D2", ref"D3", ref"F1", ref"G1", ref"G2", ref"G3")

  test(
    "GH-606: an edit outside every blind reader's reach keeps their caches and reports nothing"
  ) {
    val before = boundedBlindWorkbook
    val result = afterEdit(before, "Q1 Data", ref"B15")
    assertEquals(result.errors, Vector.empty)
    boundedCalcRefs.foreach { ref =>
      assertEquals(cache(result.workbook, "Calc", ref), cache(before, "Calc", ref), ref.toA1)
    }
    assertEquals(result.workbook, before)
  }

  test(
    "GH-606: an edit inside a blind reader's reach withdraws its cache and reports the failure"
  ) {
    val result = afterEdit(boundedBlindWorkbook, "Calc", ref"C3")
    assertEquals(cache(result.workbook, "Calc", ref"B2"), None)
    assert(result.errors.exists(e => e.sheet.value == "Calc" && e.ref == ref"B2"))
    // a parseable reader of the blind cell cannot trust a value computed from a failed precedent
    assertEquals(cache(result.workbook, "Calc", ref"B3"), None)
    assert(result.errors.exists(e => e.sheet.value == "Calc" && e.ref == ref"B3"))
    // C3 lies in row 3, which the whole-row reader names
    assertEquals(cache(result.workbook, "Calc", ref"G2"), None)
    // literals, error literals and function names are not dependencies
    assertEquals(cache(result.workbook, "Calc", ref"D1"), Some(num(1)))
    assertEquals(cache(result.workbook, "Calc", ref"D2"), Some(num(2)))
    assertEquals(cache(result.workbook, "Calc", ref"D3"), Some(num(3)))
    assertEquals(result.errors.map(_.ref).toSet, Set(ref"B2", ref"B3", ref"G2"))
  }

  test("GH-606: a blind reader of a formula inside the cone is withdrawn transitively") {
    // Data!A1 -> Data!C3 (parseable) -> Calc!F1 (blind, names Data!C3 textually)
    val result = afterEdit(boundedBlindWorkbook, "Data", ref"A1")
    assertEquals(cache(result.workbook, "Data", ref"C3"), Some(num(4)))
    assertEquals(cache(result.workbook, "Calc", ref"F1"), None)
    assert(result.errors.exists(e => e.sheet.value == "Calc" && e.ref == ref"F1"))
    assertEquals(cache(result.workbook, "Calc", ref"B2"), Some(num(7)))
    assertEquals(result.errors.size, 1)
  }

  test("GH-606: a quoted cross-sheet qualifier bounds the reach to that sheet") {
    assertEquals(
      cache(afterEdit(boundedBlindWorkbook, "Q1 Data", ref"B16").workbook, "Calc", ref"G1"),
      None
    )
    assertEquals(
      cache(afterEdit(boundedBlindWorkbook, "Calc", ref"B16").workbook, "Calc", ref"G1"),
      Some(num(9)),
      "B16 on the reader's own sheet is not the qualified B16"
    )
  }

  test("GH-606: whole-row and whole-column reaches") {
    assertEquals(
      cache(afterEdit(boundedBlindWorkbook, "Calc", ref"Z3").workbook, "Calc", ref"G2"),
      None
    )
    assertEquals(
      cache(afterEdit(boundedBlindWorkbook, "Calc", ref"A4").workbook, "Calc", ref"G2"),
      Some(num(10))
    )
    assertEquals(
      cache(afterEdit(boundedBlindWorkbook, "Calc", ref"H100").workbook, "Calc", ref"G3"),
      None
    )
    assertEquals(
      cache(afterEdit(boundedBlindWorkbook, "Calc", ref"I1").workbook, "Calc", ref"G3"),
      Some(num(11))
    )
    assertEquals(
      cache(afterEdit(boundedBlindWorkbook, "Data", ref"H1").workbook, "Calc", ref"G3"),
      Some(num(11))
    )
  }

  test("GH-606: a 3-D span reaches every sheet between its ends") {
    val wb = boundedBlindWorkbook.put(
      boundedBlindWorkbook.sheets(1).put(ref"J1", formula("ZZZNOTAFUNC(Data:Calc!A1)", 12))
    )
    assertEquals(cache(afterEdit(wb, "Calc", ref"A1").workbook, "Calc", ref"J1"), None)
    assertEquals(cache(afterEdit(wb, "Q1 Data", ref"A1").workbook, "Calc", ref"J1"), Some(num(12)))
  }

  test("GH-606: unbounded blind readers are withdrawn on any edit") {
    val wb = Workbook(
      Sheet("Data").put(ref"A1", num(2)),
      Sheet("Calc")
        .put(ref"E1", formula("ZZZNOTAFUNC(INDIRECT(\"A1\"))", 1))
        .put(ref"E2", formula("ZZZNOTAFUNC(Table1[Col])", 2))
        .put(ref"E3", formula("ZZZNOTAFUNC(NoSuchName)", 3))
        .put(ref"E4", formula("ZZZNOTAFUNC(A1:INDEX(A1:A9,3))", 4))
    )
    val result = afterEdit(wb, "Data", ref"A1")
    List(ref"E1", ref"E2", ref"E3", ref"E4").foreach { ref =>
      assertEquals(cache(result.workbook, "Calc", ref), None, ref.toA1)
      assert(result.errors.exists(e => e.sheet.value == "Calc" && e.ref == ref), ref.toA1)
    }
  }

  test("GH-606: name resolution honours sheet-scoped shadowing") {
    val data = (1 to 10).foldLeft(Sheet("Data")) { (sheet, row) =>
      sheet.put(ARef.from0(0, row - 1), num(row))
    }
    val base = Workbook(
      data,
      Sheet("Calc").put(ref"C1", num(1)).put(ref"B5", formula("ZZZNOTAFUNC(Multi)", 2)),
      Sheet("Other").put(ref"B5", formula("ZZZNOTAFUNC(Multi)", 33))
    ).withDefinedName("Multi", "Data!$A$1:$A$3,Data!$A$8:$A$10")
    val local = com.tjclp.xl.workbooks.DefinedName("Multi", "Calc!$C$1", localSheetId = Some(1))
    val wb =
      base.copy(metadata = base.metadata.copy(definedNames = base.metadata.definedNames :+ local))
    // Data row 2 is inside the global union: Other's reader is dirty, Calc's shadowed one is not
    val dataEdit = afterEdit(wb, "Data", ref"A2")
    assertEquals(cache(dataEdit.workbook, "Other", ref"B5"), None)
    assertEquals(cache(dataEdit.workbook, "Calc", ref"B5"), Some(num(2)))
    // Calc!C1 is what the local definition reads
    val calcEdit = afterEdit(wb, "Calc", ref"C1")
    assertEquals(cache(calcEdit.workbook, "Calc", ref"B5"), None)
    assertEquals(cache(calcEdit.workbook, "Other", ref"B5"), Some(num(33)))
    // Data row 5 is outside both areas of the union
    val outside = afterEdit(wb, "Data", ref"A5")
    assertEquals(outside.errors, Vector.empty)
    assertEquals(cache(outside.workbook, "Other", ref"B5"), Some(num(33)))
  }
