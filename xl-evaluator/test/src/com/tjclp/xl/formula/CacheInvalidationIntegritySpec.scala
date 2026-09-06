package com.tjclp.xl.formula

import com.tjclp.xl.{*, given}
import com.tjclp.xl.formula.eval.StructuralEditor

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
