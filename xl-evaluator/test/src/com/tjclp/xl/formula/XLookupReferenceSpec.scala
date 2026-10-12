package com.tjclp.xl.formula

import com.tjclp.xl.{*, given}
import com.tjclp.xl.addressing.{ARef, SheetName}
import com.tjclp.xl.cells.{CellError, CellValue}
import com.tjclp.xl.formula.eval.ArrayResult
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.workbooks.Workbook
import munit.FunSuite

/**
 * GH-713: XLOOKUP returns a reference — the matched row (or column) of `return_array` — so a 2-D
 * return array yields a whole row, a reference position (SUM, ROWS, the `:` operator) reads it
 * whole, and a plain cell reads it through the implicit intersection, as Excel 365's `=@XLOOKUP(…)`
 * does.
 *
 * LibreOffice 24.2 has no XLOOKUP: the values are Excel's documented semantics. The INDEX lifting
 * pins are LibreOffice 24.2.7's.
 */
@SuppressWarnings(Array("org.wartremover.warts.AsInstanceOf"))
class XLookupReferenceSpec extends FunSuite:

  private def N(n: String): CellValue = CellValue.Number(BigDecimal(n))
  private def E(code: String): CellValue =
    CellValue.Error(CellError.parse(code).fold(msg => fail(msg), identity))

  /** Sheet1: A1:A3 = 1,2,3, B1:B3 = 10,20,30, C1:C3 = 100,200,300, D1 blank, row 5 = keys. */
  private val book: Workbook =
    val grid = (1 to 3).foldLeft(Sheet(SheetName.unsafe("Sheet1"))) { (s, i) =>
      s.put(ARef.from1(1, i), N(i.toString))
        .put(ARef.from1(2, i), N((i * 10).toString))
        .put(ARef.from1(3, i), N((i * 100).toString))
    }
    // a horizontal key row: A5:C5 = 1,2,3
    val keyed = (1 to 3).foldLeft(grid)((s, i) => s.put(ARef.from1(i, 5), N(i.toString)))
    Workbook(Vector(keyed))

  private def sheet1: Sheet = book.sheets.headOption.getOrElse(fail("no sheet"))

  private def plain(at: String, formula: String): Either[String, CellValue] =
    val ref = ARef.parse(at).fold(msg => fail(msg), identity)
    val placed = sheet1.put(ref, CellValue.Formula(formula, None))
    placed.evaluateCell(ref, Clock.system, Some(book.put(placed))).left.map(_.message)

  private def array(formula: String): Either[EvalError, Any] =
    Evaluator.arrayInstance.eval(
      FormulaParser
        .parse(formula)
        .fold(err => fail(s"$formula: $err"), identity)
        .asInstanceOf[TExpr[Any]],
      sheet1,
      workbook = Some(book)
    )

  private def row(values: String*): ArrayResult = ArrayResult(Vector(values.map(N).toVector))

  test("GH-713: a 2-D return array yields the whole matched row") {
    assertEquals(array("=XLOOKUP(2,A1:A3,B1:C3)"), Right(row("20", "200")))
    assertEquals(plain("H1", "ROWS(XLOOKUP(2,A1:A3,B1:C3))"), Right(N("1")))
    assertEquals(plain("H1", "COLUMNS(XLOOKUP(2,A1:A3,B1:C3))"), Right(N("2")))
    assertEquals(plain("H1", "SUM(XLOOKUP(2,A1:A3,B1:C3))"), Right(N("220")))
    assertEquals(array("=SUM(XLOOKUP(2,A1:A3,B1:C3))"), Right(BigDecimal(220)))
  }

  test("GH-713: a plain cell reads the matched row in its own column (=@XLOOKUP)") {
    assertEquals(plain("B9", "XLOOKUP(2,A1:A3,B1:C3)"), Right(N("20")))
    assertEquals(plain("C9", "XLOOKUP(2,A1:A3,B1:C3)"), Right(N("200")))
    assertEquals(plain("E9", "XLOOKUP(2,A1:A3,B1:C3)"), Right(E("#VALUE!")))
    // a one-cell result is the same value in every column
    for at <- List("B9", "E9", "Z42") do
      assertEquals(plain(at, "XLOOKUP(2,A1:A3,B1:B3)"), Right(N("20")), at)
      assertEquals(plain(at, "XLOOKUP(2,A1:A3,B1:B3)*2"), Right(N("40")), at)
  }

  test("GH-713: a horizontal lookup returns a column; other shapes are #VALUE!") {
    // keys across A5:C5, return rows 1..3 of A:C: the matched column, top to bottom
    assertEquals(
      array("=XLOOKUP(2,A5:C5,A1:C3)"),
      Right(ArrayResult(Vector(Vector(N("10")), Vector(N("20")), Vector(N("30")))))
    )
    assertEquals(plain("H1", "SUM(XLOOKUP(3,A5:C5,A1:C3))"), Right(N("600")))
    // a 2-D lookup array, and a return array of another length
    plain("H1", "XLOOKUP(2,A1:B3,B1:C3)") match
      case Right(CellValue.Error(CellError.Value)) => ()
      case other => fail(s"2-D lookup_array: $other")
    val message = array("=XLOOKUP(2,A1:B3,B1:C3)") match
      case Left(EvalError.ErrorValue(CellError.Value, Some(msg))) => msg
      case other => fail(s"2-D lookup_array: $other")
    assert(message.contains("dimension"), message)
    assertEquals(plain("H1", "XLOOKUP(2,A1:A3,B1:B2)"), Right(E("#VALUE!")))
    // a 1×1 lookup array with a one-row return array: the row (the vertical rule wins)
    assertEquals(array("=XLOOKUP(1,A1,A1:C1)"), Right(row("1", "10", "100")))
  }

  test("GH-713: a miss is #N/A without if_not_found, its value with one") {
    assertEquals(plain("H1", "XLOOKUP(9,A1:A3,B1:C3)"), Right(E("#N/A")))
    assertEquals(plain("H1", "ISNA(XLOOKUP(9,A1:A3,B1:C3))"), Right(CellValue.Bool(true)))
    assertEquals(plain("H1", "IFNA(XLOOKUP(9,A1:A3,B1:B3),-1)"), Right(N("-1")))
    assertEquals(plain("H1", "XLOOKUP(9,A1:A3,B1:B3,\"none\")"), Right(CellValue.Text("none")))
    assertEquals(plain("H1", "SUM(XLOOKUP(9,A1:A3,B1:C3,0))"), Right(N("0")))
    // a matched blank reads as INDEX over a blank does
    val blankRow = sheet1.put(ref"A7", N("7"))
    val placed = blankRow.put(ref"H1", CellValue.Formula("XLOOKUP(7,A1:A7,D1:D7)", None))
    val index = blankRow.put(ref"H1", CellValue.Formula("INDEX(D1:D7,7)", None))
    assertEquals(
      placed.evaluateCell(ref"H1", Clock.system, Some(book.put(placed))),
      index.evaluateCell(ref"H1", Clock.system, Some(book.put(index)))
    )
  }

  test("GH-713: XLOOKUP is an operand of ':'") {
    assertEquals(plain("H1", "SUM(B1:XLOOKUP(2,A1:A3,B1:B3))"), Right(N("30")))
    assertEquals(
      plain("H1", "SUM(XLOOKUP(1,A1:A3,B1:B3):XLOOKUP(3,A1:A3,C1:C3))"),
      Right(N("660"))
    )
    assertEquals(plain("H1", "SUM(B1:XLOOKUP(9,A1:A3,B1:B3))"), Right(E("#N/A")))
  }

  test("GH-713: a year's row — SUM(XLOOKUP(key,ids,B:M)) — costs the data") {
    val wide = (1 to 3).foldLeft(sheet1) { (s, r) =>
      (2 to 13).foldLeft(s)((acc, c) => acc.put(ARef.from1(c, r), N((r * c).toString)))
    }
    val placed = wide.put(ref"O1", CellValue.Formula("SUM(XLOOKUP(2,A:A,B:M))", None))
    // row 2: 2*(2+…+13) = 2*90
    assertEquals(
      placed.evaluateCell(ref"O1", Clock.system, Some(book.put(placed))),
      Right(N("180"))
    )
  }

  /**
   * A tall sheet's lookup: A1:A3 = 1,2,3 and B1:B3 = 10,20,30, nothing below row 3, so every
   * position of `A:A` past row 3 is blank. `ROWS(B1:XLOOKUP(…))` is the matched row.
   */
  private def tall(formula: String): Either[String, CellValue] =
    val data = (1 to 3).foldLeft(Sheet(SheetName.unsafe("Tall"))) { (s, i) =>
      s.put(ARef.from1(1, i), N(i.toString)).put(ARef.from1(2, i), N((i * 10).toString))
    }
    val placed = data.put(ref"H1", CellValue.Formula(formula, None))
    placed.evaluateCell(ref"H1", Clock.system, Some(Workbook(Vector(placed)))).left.map(_.message)

  /**
   * The same data, A1:A3 = 1,2,3 and B1:B3 = 10,20,30, on a sheet whose used range runs to row 20
   * (Z20 holds a value), so a lookup over A1:A20 reads every one of its blank cells — the full read
   * the whole-column clamp must agree with.
   */
  private def short(formula: String): Either[String, CellValue] =
    val data = (1 to 3).foldLeft(Sheet(SheetName.unsafe("Short"))) { (s, i) =>
      s.put(ARef.from1(1, i), N(i.toString)).put(ARef.from1(2, i), N((i * 10).toString))
    }
    val placed = data.put(ref"Z20", N("0")).put(ref"H1", CellValue.Formula(formula, None))
    placed.evaluateCell(ref"H1", Clock.system, Some(Workbook(Vector(placed)))).left.map(_.message)

  test("GH-713: every match × search mode lands where a full read does, blank tail included") {
    // A:A on Tall reads rows 1-3 and lets one blank stand for rows 4-1048576; A1:A20 on Short reads
    // all 20 rows. A hit on Short's first blank (row 4) is Tall's too; a hit on Short's last row
    // (20, a reverse search's first blank) is Tall's last row, 1048576.
    val cases =
      for
        key <- List("Z9", "\"\"", "0", "0.5", "-1", "1", "2.5", "3", "5", "\"a\"")
        matchMode <- List("-1", "1", "0", "2")
        searchMode <- List("1", "-1", "2", "-2")
      yield (key, matchMode, searchMode)
    def fullRead(k: String, m: String, s: String): Either[String, CellValue] =
      short(s"ROWS(B1:XLOOKUP($k,A1:A20,B1:B20,,$m,$s))") match
        case Right(CellValue.Number(n)) if n == BigDecimal(20) => Right(N("1048576"))
        case other => other
    val expected = cases.map((k, m, s) => (s"$k,$m,$s", fullRead(k, m, s)))
    val clamped =
      cases.map((k, m, s) => (s"$k,$m,$s", tall(s"ROWS(B1:XLOOKUP($k,A:A,B:B,,$m,$s))")))
    assertEquals(clamped, expected)
    // Today no mode lands on a blank — exact, wildcard and approximate matching all skip blank
    // cells, whatever the key — so the full read and the clamp agree because the stand-in never
    // wins. Pinned: if a change ever lets a blank match, this fails first, and the comparison
    // above then checks the stand-in's position (it relies on first-tie-wins, see blankTail).
    assert(
      !expected.exists((_, r) => r == Right(N("4")) || r == Right(N("1048576"))),
      expected.mkString("\n")
    )
    // the matrix does reach both ends of the data in every search direction
    assert(
      expected.count((_, r) => r == Right(N("1"))) >= 8 && expected.count((_, r) =>
        r == Right(N("3"))
      ) >= 8
    )
  }

  test("GH-713: a whole-column lookup reads the data, and matches past it as before") {
    // the data
    assertEquals(tall("XLOOKUP(2,A:A,B:B)"), Right(N("20")))
    assertEquals(tall("ROWS(B1:XLOOKUP(3,A:A,B:B,,0,-1))"), Right(N("3")))
    // keys that could meet the blanks past the data (a blank cell, "", 0, wildcards), forward
    // and in both reverse modes, exact and wildcard. None matches a blank — a wildcard matches
    // the numbers as text, so it lands on the first or the last of the data. These are the
    // answers the search gave when it read every cell of A:A.
    val blankLike =
      for
        key <- List("Z9", "\"\"", "0", "\"*\"", "\"?*\"")
        mode <- List("", ",,0,-1", ",,0,-2", ",,2", ",,2,-1", ",,2,-2")
      yield (key, mode)
    def expected(key: String, mode: String): Either[String, CellValue] =
      (key.contains('*'), mode) match
        case (true, ",,2") => Right(N("1"))
        case (true, ",,2,-1" | ",,2,-2") => Right(N("3"))
        case _ => Right(E("#N/A"))
    assertEquals(
      blankLike.map((k, m) => (k + m, tall(s"ROWS(B1:XLOOKUP($k,A:A,B:B$m))"))),
      blankLike.map((k, m) => (k + m, expected(k, m)))
    )
    // the approximate modes over data then blanks: next smaller from either end, next larger
    assertEquals(tall("ROWS(B1:XLOOKUP(5,A:A,B:B,,-1))"), Right(N("3")))
    assertEquals(tall("ROWS(B1:XLOOKUP(5,A:A,B:B,,-1,-1))"), Right(N("3")))
    assertEquals(tall("ROWS(B1:XLOOKUP(0,A:A,B:B,,1))"), Right(N("1")))
    assertEquals(tall("XLOOKUP(9,A:A,B:B,,1)"), Right(E("#N/A")))
    // a finite lookup array past the data, and a horizontal one over a whole row
    assertEquals(tall("ROWS(B1:XLOOKUP(\"*\",A1:A10,B1:B10,,2,-1))"), Right(N("3")))
    assertEquals(tall("ROWS(B1:XLOOKUP(5,A1:A10,B1:B10,,-1,-1))"), Right(N("3")))
    assertEquals(tall("COLUMNS(A2:XLOOKUP(\"*\",2:2,2:2,,2,-1))"), Right(N("2")))
    assertEquals(tall("XLOOKUP(20,2:2,2:2)"), Right(N("20")))
    assertEquals(tall("XLOOKUP(Z9,2:2,2:2,,0,-1)"), Right(E("#N/A")))
  }

  test("GH-713: an array key still lifts under an aggregate (no reference collapse)") {
    assertEquals(plain("H1", "SUM(XLOOKUP({1;2},A1:A3,B1:B3))"), Right(N("30")))
    assertEquals(plain("H1", "SUMPRODUCT(XLOOKUP({1;2},A1:A3,B1:B3))"), Right(N("30")))
    // LibreOffice 24.2.7: INDEX lifts over an array position under an aggregate
    assertEquals(plain("H1", "SUM(INDEX(B1:C3,{1;2},1))"), Right(N("30")))
    assertEquals(plain("H1", "SUM(INDEX(B1:B3,{1;2}))"), Right(N("30")))
  }

  test("GH-713: an array key over a 2-D return array keeps each match's first cell") {
    // Excel lifts XLOOKUP over the array key, and an element that is itself an array (the matched
    // row {10,100}) is not a value of the lifted array: Excel keeps its first element, as
    // SEQUENCE({2;3}) is {1;1}. So the result is {10;20} and the sum 30, not 330.
    // Excel and LibreOffice differ here. LibreOffice 24.2.7 cannot answer (no XLOOKUP: #NAME?),
    // and where it lifts a row-returning call — the INDEX analogue SUM(INDEX(B1:C3,{1;2},0)) — it
    // keeps every cell of each selected row (330), where Excel keeps one element per key.
    assertEquals(plain("H1", "SUM(XLOOKUP({1;2},A1:A3,B1:C3))"), Right(N("30")))
    assertEquals(plain("H1", "SUMPRODUCT(XLOOKUP({1;2},A1:A3,B1:C3))"), Right(N("30")))
    assertEquals(
      array("=XLOOKUP({1;2},A1:A3,B1:C3)"),
      Right(ArrayResult(Vector(Vector(N("10")), Vector(N("20")))))
    )
  }

  test("GH-713: the matched strip is total — a position outside the return array is #REF!") {
    val column = CellRange.parse("A1:A3").fold(msg => fail(msg), identity)
    val block = CellRange.parse("B1:C3").fold(msg => fail(msg), identity)
    val strip =
      functions.FunctionSpecs.xlookupStrip(column, block).fold(e => fail(e.toString), identity)
    assertEquals(strip(1).map(_.toA1), Right("B2:C2"))
    List(-1, 3).foreach { i =>
      strip(i) match
        case Left(EvalError.ErrorValue(CellError.Ref, _)) => ()
        case other => fail(s"row $i: expected #REF!, got $other")
    }
    val keys = CellRange.parse("A5:C5").fold(msg => fail(msg), identity)
    val grid = CellRange.parse("A1:C3").fold(msg => fail(msg), identity)
    val horizontal =
      functions.FunctionSpecs.xlookupStrip(keys, grid).fold(e => fail(e.toString), identity)
    assertEquals(horizontal(2).map(_.toA1), Right("C1:C3"))
    assert(horizontal(3).isLeft)
  }

  test("GH-713: an if_not_found value is not a reference, so a `:` over it is #VALUE!") {
    // the miss returns if_not_found, a value: the range operator refuses it as an operand; an
    // error value there keeps its own code
    assertEquals(plain("H1", "SUM(B1:XLOOKUP(9,A1:A3,B1:B3,\"x\"))"), Right(E("#VALUE!")))
    assertEquals(plain("H1", "SUM(B1:XLOOKUP(9,A1:A3,B1:B3,5))"), Right(E("#VALUE!")))
    assertEquals(plain("H1", "SUM(B1:XLOOKUP(9,A1:A3,B1:B3,#N/A))"), Right(E("#N/A")))
    // a hit is a reference: B1:B2
    assertEquals(plain("H1", "SUM(B1:XLOOKUP(2,A1:A3,B1:B3,\"x\"))"), Right(N("30")))
  }

  test("GH-713: a 1×1 lookup array beside an N×1 return array returns the whole column") {
    // Today's answer, pinned so it cannot drift silently: a one-cell lookup array is also a one-row
    // array of the return array's width (1), so a match returns the matched column — all of
    // B1:B3. Excel's answer is unconfirmed (LibreOffice 24.2 has no XLOOKUP to ask); reading the
    // lookup array as a column instead would make the heights differ, `#VALUE!`.
    assertEquals(
      array("=XLOOKUP(1,A1,B1:B3)"),
      Right(ArrayResult(Vector(Vector(N("10")), Vector(N("20")), Vector(N("30")))))
    )
    assertEquals(plain("H1", "SUM(XLOOKUP(1,A1,B1:B3))"), Right(N("60")))
    assertEquals(plain("H1", "ROWS(XLOOKUP(1,A1,B1:B3))"), Right(N("3")))
  }

  test("GH-713: the result keeps round-tripping as _xlfn.XLOOKUP") {
    val expr =
      FormulaParser.parse("=SUM(B1:XLOOKUP(2,A1:A3,B1:B3))").fold(e => fail(e.toString), identity)
    assertEquals(FormulaPrinter.printFileForm(expr), "SUM(B1:XLOOKUP(2,A1:A3,B1:B3))")
    assertEquals(
      com.tjclp.xl.ooxml.FormulaStorage.toStored(FormulaPrinter.printFileForm(expr)),
      "SUM(B1:_xlfn.XLOOKUP(2,A1:A3,B1:B3))"
    )
  }

  /** Sheet1: C1:E3 = i*100, i*10, i; G1:G3 = 1,2,3; columns A and B empty. */
  private val offsetBook: Workbook =
    val grid = (1 to 3).foldLeft(Sheet(SheetName.unsafe("Sheet1"))) { (s, i) =>
      s.put(ARef.from1(3, i), N((i * 100).toString))
        .put(ARef.from1(4, i), N((i * 10).toString))
        .put(ARef.from1(5, i), N(i.toString))
        .put(ARef.from1(7, i), N(i.toString))
    }
    Workbook(Vector(grid))

  private def offsetArray(formula: String): Either[EvalError, Any] =
    Evaluator.arrayInstance.eval(
      FormulaParser
        .parse(formula)
        .fold(err => fail(s"$formula: $err"), identity)
        .asInstanceOf[TExpr[Any]],
      offsetBook.sheets.headOption.getOrElse(fail("no sheet")),
      workbook = Some(offsetBook)
    )

  private def offsetPlain(formula: String): Either[String, CellValue] =
    val target = offsetBook.sheets.headOption.getOrElse(fail("no sheet"))
    val placed = target.put(ref"J1", CellValue.Formula(formula, None))
    placed.evaluateCell(ref"J1", Clock.system, Some(offsetBook.put(placed))).left.map(_.message)

  test("GH-713: a returned row keeps the reference's shape — leading blanks stay in place") {
    val expected = ArrayResult(
      Vector(Vector(CellValue.Empty, CellValue.Empty, N("200"), N("20"), N("2")))
    )
    for call <- List("XLOOKUP(2,G1:G3,A1:E3)", "INDEX(A1:E3,2,0)") do
      assertEquals(offsetArray(s"=$call"), Right(expected), call)
      assertEquals(offsetPlain(s"COLUMNS($call)"), Right(N("5")), call)
      assertEquals(offsetPlain(s"SUMPRODUCT($call*{1,2,3,4,5})"), Right(N("690")), call)
      assertEquals(offsetPlain(s"SUMPRODUCT($call,{1,2,3,4,5})"), Right(N("690")), call)
  }

  test("GH-713: a returned row past the used range keeps its trailing blanks") {
    val months = "{1,2,3,4,5,6,7,8,9,10,11,12}"
    array("=XLOOKUP(2,A1:A3,B1:M3)") match
      case Right(r: ArrayResult) =>
        assertEquals((r.rows, r.cols), (1, 12))
        assertEquals(r.values.headOption.map(_.take(2)), Some(Vector(N("20"), N("200"))))
        assert(r.values.flatten.drop(2).forall(_ == CellValue.Empty), r.toString)
      case other => fail(s"XLOOKUP(2,A1:A3,B1:M3): $other")
    assertEquals(plain("H1", s"SUMPRODUCT(XLOOKUP(2,A1:A3,B1:M3)*$months)"), Right(N("420")))
    assertEquals(plain("H1", s"SUMPRODUCT(INDEX(B1:M3,2,0)*$months)"), Right(N("420")))
  }

  test("GH-713: a whole-column result is still bounded to the used rows, from its start") {
    // used rows end at row 5 (the key row): A:A materializes rows 1..5, never a million blanks
    array("=INDEX(A:A,0)") match
      case Right(r: ArrayResult) => assertEquals((r.rows, r.cols), (5, 1))
      case other => fail(s"INDEX(A:A,0): $other")
    // a column reaching the sheet's last row past the data: anchored at its start, one blank
    array("=INDEX(A9:A1048576,0)") match
      case Right(r: ArrayResult) =>
        assertEquals(r, ArrayResult(Vector(Vector(CellValue.Empty))))
      case other => fail(s"INDEX(A9:A1048576,0): $other")
    // a whole-row result is bounded to the used columns
    array("=INDEX(2:2,0)") match
      case Right(r: ArrayResult) => assertEquals((r.rows, r.cols), (1, 3))
      case other => fail(s"INDEX(2:2,0): $other")
  }
