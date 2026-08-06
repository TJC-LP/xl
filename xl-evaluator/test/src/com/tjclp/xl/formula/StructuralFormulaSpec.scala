package com.tjclp.xl.formula

import com.tjclp.xl.{*, given}
import com.tjclp.xl.addressing.SheetName
import com.tjclp.xl.cells.{CellError, CellValue}
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.workbooks.Workbook
import com.tjclp.xl.formula.eval.StructuralEditor
import munit.FunSuite

/**
 * GH-128 / GH-129: structural editing WITH formula rewriting (the xl-evaluator layer over the pure
 * xl-core cell shift). Covers conditional shifting, #REF! on deletion, range shrink, the
 * insert↔delete identity law, and cross-sheet reference rewriting.
 */
@SuppressWarnings(Array("org.wartremover.warts.OptionPartial"))
class StructuralFormulaSpec extends FunSuite:

  private val S = SheetName.unsafe("S")

  private def sheetNamed(wb: Workbook, n: String): Sheet =
    wb.sheets.find(_.name == SheetName.unsafe(n)).get

  private def formulaCell(s: String): CellValue = CellValue.Formula(s, None)

  test("insert rows shifts local refs at/after the insertion point") {
    val s = new Sheet(name = S)
      .put(ref"B1", formulaCell("=A5"))
      .put(ref"B2", formulaCell("=A1"))
    val r = StructuralEditor.insertRows(Workbook(Vector(s)), S, at = 2, count = 1)
    val s2 = sheetNamed(r, "S")
    assertEquals(s2(ref"B1").value, formulaCell("A6")) // A5 (row 4 >= 2) -> A6
    assertEquals(s2(ref"B2").value, formulaCell("A1")) // A1 (row 0 < 2) unchanged
  }

  test("insert rows moves the formula cell itself and still rewrites refs") {
    val s = new Sheet(name = S).put(ref"B5", formulaCell("=A1"))
    val r = StructuralEditor.insertRows(Workbook(Vector(s)), S, at = 2, count = 1)
    val s2 = sheetNamed(r, "S")
    assertEquals(s2(ref"B6").value, formulaCell("A1")) // cell B5 (row 4) -> B6; ref A1 unchanged
    assert(!s2.contains(ref"B5"))
  }

  test("delete rows: ref into the deleted band becomes #REF!") {
    val s = new Sheet(name = S).put(ref"B1", formulaCell("=A3"))
    val r = StructuralEditor.deleteRows(Workbook(Vector(s)), S, at = 2, count = 1)
    // A3 = row index 2 = the deleted row
    assertEquals(sheetNamed(r, "S")(ref"B1").value, CellValue.Error(CellError.Ref))
  }

  test("delete rows: refs after the band shift up") {
    val s = new Sheet(name = S).put(ref"B1", formulaCell("=A5"))
    val r = StructuralEditor.deleteRows(Workbook(Vector(s)), S, at = 2, count = 1)
    assertEquals(sheetNamed(r, "S")(ref"B1").value, formulaCell("A4")) // A5 (row 4) -> A4
  }

  test("delete rows: a range straddling the deletion shrinks") {
    val s = new Sheet(name = S).put(ref"B1", formulaCell("=SUM(A1:A10)"))
    val r = StructuralEditor.deleteRows(Workbook(Vector(s)), S, at = 2, count = 1)
    assertEquals(sheetNamed(r, "S")(ref"B1").value, formulaCell("SUM(A1:A9)"))
  }

  test("insert then delete the same band is identity on formula refs") {
    val s = new Sheet(name = S).put(ref"B1", formulaCell("=A5"))
    val inserted = StructuralEditor.insertRows(Workbook(Vector(s)), S, at = 2, count = 2)
    val restored = StructuralEditor.deleteRows(inserted, S, at = 2, count = 2)
    assertEquals(sheetNamed(restored, "S")(ref"B1").value, formulaCell("A5"))
  }

  test("insert columns shifts column refs at/after the insertion point") {
    val s = new Sheet(name = S).put(ref"A1", formulaCell("=C1"))
    val r = StructuralEditor.insertColumns(Workbook(Vector(s)), S, at = 1, count = 1)
    // C1 (col 2 >= 1) -> D1; cell A1 (col 0 < 1) stays put
    assertEquals(sheetNamed(r, "S")(ref"A1").value, formulaCell("D1"))
  }

  test("delete columns: ref into the deleted band becomes #REF!") {
    val s = new Sheet(name = S).put(ref"A1", formulaCell("=C1"))
    val r = StructuralEditor.deleteColumns(Workbook(Vector(s)), S, at = 2, count = 1)
    // C1 = col index 2 = the deleted column
    assertEquals(sheetNamed(r, "S")(ref"A1").value, CellValue.Error(CellError.Ref))
  }

  test("cross-sheet references to the edited sheet are rewritten") {
    val data = new Sheet(name = SheetName.unsafe("Data")).put(ref"A5", CellValue.Number(99))
    val report =
      new Sheet(name = SheetName.unsafe("Report")).put(ref"B1", formulaCell("=Data!A5"))
    val r =
      StructuralEditor.insertRows(Workbook(Vector(data, report)), SheetName.unsafe("Data"), 2, 1)
    assertEquals(sheetNamed(r, "Report")(ref"B1").value, formulaCell("Data!A6"))
  }

  test("references to OTHER sheets are left untouched") {
    val data = new Sheet(name = SheetName.unsafe("Data"))
    val report = new Sheet(name = SheetName.unsafe("Report")).put(ref"B1", formulaCell("=Data!A5"))
    // Edit Report (not Data); the cross-ref to Data must not move.
    val r =
      StructuralEditor.insertRows(Workbook(Vector(data, report)), SheetName.unsafe("Report"), 2, 1)
    assertEquals(sheetNamed(r, "Report")(ref"B1").value, formulaCell("Data!A5"))
  }

  // ===== GH-428: insert-side clamp at the sheet edge =====

  test("GH-428: a formula ref pushed past the last row degrades to #REF!") {
    val s = new Sheet(name = S)
      .put(ref"B1", formulaCell("A1048576"))
      .put(ref"C1", formulaCell("XFD1"))
    val rows = StructuralEditor.insertRows(Workbook(Vector(s)), S, at = 0, count = 1)
    // the formula cell itself moved to B2; its ref had no home past the edge
    assertEquals(sheetNamed(rows, "S")(ref"B2").value, CellValue.Error(CellError.Ref))
    val cols = StructuralEditor.insertColumns(Workbook(Vector(s)), S, at = 0, count = 1)
    assertEquals(sheetNamed(cols, "S")(ref"D1").value, CellValue.Error(CellError.Ref))
  }

  test("GH-428: a range end clamps at the edge; a range start pushed past it voids the formula") {
    val s = new Sheet(name = S)
      .put(ref"B1", formulaCell("SUM(A10:A1048576)"))
      .put(ref"C1", formulaCell("SUM(A1048570:A1048576)"))
    val r = StructuralEditor.insertRows(Workbook(Vector(s)), S, at = 2, count = 2)
    val s2 = sheetNamed(r, "S")
    assertEquals(s2(ref"B1").value, formulaCell("SUM(A12:A1048576)")) // end pinned at the max
    val voided = StructuralEditor.insertRows(Workbook(Vector(s)), S, at = 0, count = 10)
    // the formula cell itself moved to C11; its range START passed the edge -> #REF!
    assertEquals(sheetNamed(voided, "S")(ref"C11").value, CellValue.Error(CellError.Ref))
  }

  // ===== GH-427: equals-free rewrite + structural cache invalidation =====

  test("GH-427: rewrite prints the model's equals-free form (no '=' lands in <f>)") {
    val s = new Sheet(name = S).put(ref"B1", formulaCell("A5*2")) // file-canonical form
    val r = StructuralEditor.insertRows(Workbook(Vector(s)), S, at = 2, count = 1)
    assertEquals(sheetNamed(r, "S")(ref"B1").value, formulaCell("A6*2"))
  }

  test("structural edits invalidate caches even when every formula reference survives") {
    val cached = Some(CellValue.Number(BigDecimal(4)))
    val s = new Sheet(name = S)
      .put(ref"B1", CellValue.Formula("A5*2", cached))
      .put(ref"B2", CellValue.Formula("A1*3", cached)) // refs above the cut: position unchanged
    val r = StructuralEditor.insertRows(Workbook(Vector(s)), S, at = 2, count = 1)
    val s2 = sheetNamed(r, "S")
    assertEquals(s2(ref"B1").value, CellValue.Formula("A6*2", None))
    // Default (unflagged) behaviour is unchanged: B2's cache drops even though nothing it reads
    // moved. `preserveUntouchedCaches` is what opts into the GH-503 narrowing.
    assertEquals(s2(ref"B2").value, CellValue.Formula("A1*3", None))
  }

  // ===== GH-503: opt-in cache preservation for edits the formula cannot have noticed =====

  test("GH-503: preserveUntouchedCaches keeps a cache nothing the edit touched can change") {
    val cached = Some(CellValue.Number(BigDecimal(4)))
    val s = new Sheet(name = S)
      .put(ref"A1", CellValue.Number(BigDecimal(2)))
      .put(ref"B2", CellValue.Formula("A1*3", cached)) // reads A1, above the cut
    val r = StructuralEditor
      .insertRows(Workbook(Vector(s)), S, at = 20, count = 1, preserveUntouchedCaches = true)
    assertEquals(sheetNamed(r, "S")(ref"B2").value, CellValue.Formula("A1*3", cached))
  }

  test("GH-503: a RELOCATED formula drops its cache even when its text is identical") {
    // The position-sensitive case: =ROW() keeps its text and changes its answer. `ref` is a
    // POST-shift address while the cone is keyed on PRE-edit ones, so relocation is tested
    // against the post-edit cut rather than by cone membership.
    val cached = Some(CellValue.Number(BigDecimal(30)))
    val s = new Sheet(name = S).put(ref"C30", CellValue.Formula("ROW()", cached))
    val r = StructuralEditor
      .insertRows(Workbook(Vector(s)), S, at = 19, count = 1, preserveUntouchedCaches = true)
    assertEquals(sheetNamed(r, "S")(ref"C31").value, CellValue.Formula("ROW()", None))
  }

  test("GH-503: a formula IN the edit's cone drops its cache under the flag too") {
    val cached = Some(CellValue.Number(BigDecimal(777)))
    val s = new Sheet(name = S)
      .put(ref"A25", CellValue.Number(BigDecimal(5)))
      .put(ref"C1", CellValue.Formula("A25*2", cached)) // reads a cell the insert MOVES
    val r = StructuralEditor
      .insertRows(Workbook(Vector(s)), S, at = 20, count = 1, preserveUntouchedCaches = true)
    // A25 moved to A26, so the reference is rewritten — a changed question, no cache carried.
    assertEquals(sheetNamed(r, "S")(ref"C1").value, CellValue.Formula("A26*2", None))
  }

  test("GH-503: a shrinking SUM drops its cache under the flag (text changed)") {
    val cached = Some(CellValue.Number(BigDecimal(6)))
    val s = new Sheet(name = S)
      .put(ref"A1", CellValue.Number(BigDecimal(1)))
      .put(ref"A2", CellValue.Number(BigDecimal(2)))
      .put(ref"A3", CellValue.Number(BigDecimal(3)))
      .put(ref"C1", CellValue.Formula("SUM(A1:A3)", cached))
    val r = StructuralEditor
      .deleteRows(Workbook(Vector(s)), S, at = 1, count = 1, preserveUntouchedCaches = true)
    assertEquals(sheetNamed(r, "S")(ref"C1").value, CellValue.Formula("SUM(A1:A2)", None))
  }

  test("deleting a row invalidates the old cached result of a shrinking SUM range") {
    val s = new Sheet(name = S)
      .put(ref"A1", CellValue.Number(BigDecimal(1)))
      .put(ref"A2", CellValue.Number(BigDecimal(2)))
      .put(ref"A3", CellValue.Number(BigDecimal(3)))
      .put(
        ref"C1",
        CellValue.Formula("SUM(A1:A3)", Some(CellValue.Number(BigDecimal(6))))
      )
    val r = StructuralEditor.deleteRows(Workbook(Vector(s)), S, at = 1, count = 1)
    assertEquals(
      sheetNamed(r, "S")(ref"C1").value,
      CellValue.Formula("SUM(A1:A2)", None)
    )
  }

  test("moving a position-sensitive formula invalidates its cached result") {
    val s = new Sheet(name = S)
      .put(ref"B2", CellValue.Formula("ROW()", Some(CellValue.Number(BigDecimal(2)))))
    val r = StructuralEditor.insertRows(Workbook(Vector(s)), S, at = 0, count = 1)
    assertEquals(sheetNamed(r, "S")(ref"B3").value, CellValue.Formula("ROW()", None))
  }

  test("GH-427: a fully-deleted reference still degrades to #REF! (cache dropped)") {
    val s = new Sheet(name = S)
      .put(ref"B1", CellValue.Formula("A3", Some(CellValue.Number(BigDecimal(7)))))
    val r = StructuralEditor.deleteRows(Workbook(Vector(s)), S, at = 2, count = 1)
    assertEquals(sheetNamed(r, "S")(ref"B1").value, CellValue.Error(CellError.Ref))
  }

  // ===== GH-274: INDIRECT and structural edits =====

  test("GH-274: insert row freezes INDIRECT text but shifts INDIRECT ref arguments") {
    val s = new Sheet(name = S)
      .put(ref"B1", formulaCell("=INDIRECT(\"A5\")")) // text is data — frozen (Excel parity)
      .put(ref"C1", formulaCell("=INDIRECT(A5)")) // the A5 ARGUMENT is a real ref — shifts
    val r = StructuralEditor.insertRows(Workbook(Vector(s)), S, at = 2, count = 1)
    val s2 = sheetNamed(r, "S")
    assertEquals(s2(ref"B1").value, formulaCell("INDIRECT(\"A5\")"))
    assertEquals(s2(ref"C1").value, formulaCell("INDIRECT(A6)"))
  }

  // ===== GH-455: grouping parens survive the rewrite; untouched sheets ride byte-identical =====

  private def num(i: Int): CellValue = CellValue.Number(BigDecimal(i))

  test("GH-455: insert rows keeps grouping parens and the evaluated value (=E13/(E12*E14))") {
    val s = new Sheet(name = S)
      .put(ref"E12", num(2))
      .put(ref"E13", num(4))
      .put(ref"E14", num(2))
      .put(ref"G1", formulaCell("=E13/(E12*E14)"))
    val before = Workbook(Vector(s)).evaluateFormula("=E13/(E12*E14)", "S")
    assertEquals(before, Right(num(1)))
    val r = StructuralEditor.insertRows(Workbook(Vector(s)), S, at = 2, count = 2)
    val s2 = sheetNamed(r, "S")
    assertEquals(s2(ref"G1").value, formulaCell("E15/(E14*E16)"))
    assertEquals(r.evaluateFormula("=E15/(E14*E16)", "S"), before)
  }

  test("GH-455: insert rows keeps the nested grouped divisor (=E12/((E13+E14)/2))") {
    val s = new Sheet(name = S)
      .put(ref"E12", num(6))
      .put(ref"E13", num(2))
      .put(ref"E14", num(4))
      .put(ref"G1", formulaCell("=E12/((E13+E14)/2)"))
    val before = Workbook(Vector(s)).evaluateFormula("=E12/((E13+E14)/2)", "S")
    assertEquals(before, Right(num(2)))
    val r = StructuralEditor.insertRows(Workbook(Vector(s)), S, at = 2, count = 2)
    val s2 = sheetNamed(r, "S")
    assertEquals(s2(ref"G1").value, formulaCell("E14/((E15+E16)/2)"))
    assertEquals(r.evaluateFormula("=E14/((E15+E16)/2)", "S"), before)
  }

  test("GH-455: insert rows keeps the grouped sum inside IF/ROUND (tie-out stays 'Yes')") {
    val s = new Sheet(name = S)
      .put(ref"E12", num(6))
      .put(ref"E13", num(2))
      .put(ref"E14", num(4))
      .put(ref"G1", formulaCell("=IF(ROUND(E12-(E13+E14),0)=0,\"Yes\",\"No\")"))
    val r = StructuralEditor.insertRows(Workbook(Vector(s)), S, at = 2, count = 2)
    val s2 = sheetNamed(r, "S")
    // GH-484: the rewrite reprints in Excel's bare-comma file form
    assertEquals(
      s2(ref"G1").value,
      formulaCell("IF(ROUND(E14-(E15+E16),0)=0,\"Yes\",\"No\")")
    )
    assertEquals(
      r.evaluateFormula("=IF(ROUND(E14-(E15+E16), 0)=0, \"Yes\", \"No\")", "S"),
      Right(CellValue.Text("Yes"))
    )
  }

  test("GH-455: cross-sheet grouped divisor on a non-edited sheet keeps its parens") {
    val a = new Sheet(name = SheetName.unsafe("A"))
      .put(ref"E12", num(2))
      .put(ref"E13", num(4))
      .put(ref"E14", num(2))
    val b = new Sheet(name = SheetName.unsafe("B"))
      .put(ref"B1", formulaCell("=A!E13/(A!E12*A!E14)"))
    val r = StructuralEditor.insertRows(Workbook(Vector(a, b)), SheetName.unsafe("A"), 2, 2)
    assertEquals(sheetNamed(r, "B")(ref"B1").value, formulaCell("A!E15/(A!E14*A!E16)"))
    assertEquals(r.evaluateFormula("=A!E15/(A!E14*A!E16)", "B"), Right(num(1)))
  }

  test("GH-455: formulas on a non-participating sheet are byte-identical after insert-rows") {
    val cached = Some(num(7))
    val data = new Sheet(name = SheetName.unsafe("Data")).put(ref"A1", num(1))
    val other = new Sheet(name = SheetName.unsafe("Other")).put(ref"A1", num(7))
    val report = new Sheet(name = SheetName.unsafe("Report"))
      .put(ref"B1", CellValue.Formula("C1/(A1*D1)", cached)) // local refs only
      .put(ref"B2", CellValue.Formula("SUM(Other!A1:A5)", cached)) // refs a DIFFERENT sheet
      .put(
        ref"B3",
        CellValue.Formula("IF(A1,1,0)", cached)
      ) // would canonicalize to ", " if reprinted
    val r = StructuralEditor.insertRows(
      Workbook(Vector(data, other, report)),
      SheetName.unsafe("Data"),
      at = 0,
      count = 3
    )
    val rep = sheetNamed(r, "Report")
    // Text AND cached values ride untouched: the sheet does not participate in the edit.
    assertEquals(rep(ref"B1").value, CellValue.Formula("C1/(A1*D1)", cached))
    assertEquals(rep(ref"B2").value, CellValue.Formula("SUM(Other!A1:A5)", cached))
    assertEquals(rep(ref"B3").value, CellValue.Formula("IF(A1,1,0)", cached))
  }

  test("GH-455: a two-hop cross-sheet dependent keeps its text but drops its stale cache") {
    val alpha = new Sheet(name = SheetName.unsafe("Alpha")).put(ref"A1", num(10))
    val beta = new Sheet(name = SheetName.unsafe("Beta"))
      .put(ref"A1", num(3))
      .put(ref"X1", CellValue.Formula("Alpha!A1", Some(num(10)))) // hop 1: names Alpha
      .put(ref"Y1", CellValue.Formula("X1*2", Some(num(20)))) // hop 2: never names Alpha
      .put(ref"Z1", CellValue.Formula("A1*3", Some(num(9)))) // independent control
    val r = StructuralEditor.insertRows(
      Workbook(Vector(alpha, beta)),
      SheetName.unsafe("Alpha"),
      at = 0,
      count = 1
    )
    val out = sheetNamed(r, "Beta")
    // Hop 1 participates: rewritten and uncached (existing behavior).
    assertEquals(out(ref"X1").value, CellValue.Formula("Alpha!A2", None))
    // Hop 2 rides byte-identical in TEXT, but its cache transitively read Alpha: stale, dropped.
    assertEquals(out(ref"Y1").value, CellValue.Formula("X1*2", None))
    // Genuinely independent formula keeps text AND cache untouched.
    assertEquals(out(ref"Z1").value, CellValue.Formula("A1*3", Some(num(9))))
  }

  test("GH-455: a three-hop chain across three sheets is cache-invalidated end to end") {
    val alpha = new Sheet(name = SheetName.unsafe("Alpha")).put(ref"A1", num(10))
    val beta = new Sheet(name = SheetName.unsafe("Beta"))
      .put(ref"B1", CellValue.Formula("Alpha!A1", Some(num(10))))
    val gamma = new Sheet(name = SheetName.unsafe("Gamma"))
      .put(ref"C1", CellValue.Formula("Beta!B1*2", Some(num(20)))) // hop 2: never names Alpha
      .put(ref"D1", CellValue.Formula("C1+1", Some(num(21)))) // hop 3: local ref only
      .put(ref"C2", num(5))
      .put(ref"E1", CellValue.Formula("C2*2", Some(num(10)))) // independent control
    val r = StructuralEditor.insertRows(
      Workbook(Vector(alpha, beta, gamma)),
      SheetName.unsafe("Alpha"),
      at = 0,
      count = 2
    )
    val out = sheetNamed(r, "Gamma")
    assertEquals(sheetNamed(r, "Beta")(ref"B1").value, CellValue.Formula("Alpha!A3", None))
    assertEquals(out(ref"C1").value, CellValue.Formula("Beta!B1*2", None))
    assertEquals(out(ref"D1").value, CellValue.Formula("C1+1", None))
    assertEquals(out(ref"E1").value, CellValue.Formula("C2*2", Some(num(10))))
  }

  test("GH-455: cross-sheet dynamic refs and their dependents drop stale caches") {
    val alpha = new Sheet(name = SheetName.unsafe("Alpha")).put(ref"A1", num(10))
    val beta = new Sheet(name = SheetName.unsafe("Beta"))
      .put(ref"X1", CellValue.Formula("INDIRECT(\"Alpha!A1\")", Some(num(10))))
      .put(ref"Y1", CellValue.Formula("X1*2", Some(num(20))))
      .put(ref"Z1", CellValue.Formula("1+2", Some(num(3))))
    val r = StructuralEditor.insertRows(
      Workbook(Vector(alpha, beta)),
      SheetName.unsafe("Alpha"),
      at = 0,
      count = 1
    )
    val out = sheetNamed(r, "Beta")
    // INDIRECT text is data and stays byte-identical, but now resolves to the newly empty A1.
    assertEquals(out(ref"X1").value, CellValue.Formula("INDIRECT(\"Alpha!A1\")", None))
    assertEquals(out(ref"Y1").value, CellValue.Formula("X1*2", None))
    assertEquals(out(ref"Z1").value, CellValue.Formula("1+2", Some(num(3))))
  }

  test("GH-455: a non-edited sheet MENTIONING the edited sheet still rewrites (and only then)") {
    val data = new Sheet(name = SheetName.unsafe("Data"))
    val report = new Sheet(name = SheetName.unsafe("Report"))
      .put(ref"B1", CellValue.Formula("Data!A5*2", Some(num(7)))) // participates: rewritten
      .put(ref"B2", CellValue.Formula("\"Data\"&C1", Some(CellValue.Text("Datax")))) // text only
    val r = StructuralEditor.insertRows(
      Workbook(Vector(data, report)),
      SheetName.unsafe("Data"),
      at = 2,
      count = 1
    )
    val rep = sheetNamed(r, "Report")
    assertEquals(rep(ref"B1").value, CellValue.Formula("Data!A6*2", None))
    // Mentions "Data" only inside a string literal: the AST gate proves non-participation.
    assertEquals(
      rep(ref"B2").value,
      CellValue.Formula("\"Data\"&C1", Some(CellValue.Text("Datax")))
    )
  }

  test("GH-484: structural rewrites keep Excel's bare-comma file form") {
    val s = new Sheet(name = S)
      .put(ref"B1", formulaCell("IF(A5=1,1,0)"))
      .put(ref"B2", formulaCell("SUMIF(A1:A10,\">2\",C1:C10)"))
    val r = StructuralEditor.insertRows(Workbook(Vector(s)), S, at = 2, count = 1)
    val s2 = sheetNamed(r, "S")
    assertEquals(s2(ref"B1").value, formulaCell("IF(A6=1,1,0)"))
    assertEquals(s2(ref"B2").value, formulaCell("SUMIF(A1:A11,\">2\",C1:C11)"))
  }
