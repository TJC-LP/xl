package com.tjclp.xl.sheets

import munit.ScalaCheckSuite
import org.scalacheck.Prop.forAll

import com.tjclp.xl.{*, given}
import com.tjclp.xl.Generators
import com.tjclp.xl.cells.{Cell, CellError, CellValue, FormulaKind}
import com.tjclp.xl.error.XLError

/**
 * GH-419 data-table authoring semantics: V1-V8 validation (exact XLError case + steering message,
 * deterministic row-major first-offender selection), corner-only materialization, the
 * consume-and-normalize rule for existing records, the absorb rule for plain corner scalars, seeds,
 * and the authored record dialect (dt2D/dtr per shape, single 1-D input riding r1, ca=1).
 *
 * The API is exercised through the public `com.tjclp.xl.{*, given}` export path — the same hop
 * scripts use — so this spec also gates the dataTableSyntax export wiring.
 */
class SheetDataTableSpec extends ScalaCheckSuite:

  private def range(a1: String): CellRange = CellRange.parse(a1).fold(fail(_), identity)
  private def aref(a1: String): ARef = ARef.parse(a1).fold(fail(_), identity)
  private def num(n: Int): CellValue = CellValue.Number(BigDecimal(n))

  private def expectLeft(result: XLResult[Sheet]): XLError =
    result.fold(identity, sheet => fail(s"expected Left, got authored sheet: $sheet"))

  private def expectRight(result: XLResult[Sheet]): Sheet =
    result.fold(err => fail(s"expected authored sheet, got $err"), identity)

  /** Interior D5:F6 -> corner C4, row axis D4:F4, column axis C5:C6; inputs B1/B2. */
  private def base2D: Sheet =
    Sheet("S")
      .put(ref"B1", num(10))
      .put(ref"B2", num(20))
      .put(ref"C4", CellValue.Formula("B1*B2"))
      .put(ref"D4", num(1))
      .put(ref"E4", num(2))
      .put(ref"F4", num(3))
      .put(ref"C5", num(100))
      .put(ref"C6", num(200))

  private val interior2D = range("D5:F6")

  private def kind2D(interior: CellRange, r1: String, r2: String): FormulaKind.DataTable =
    FormulaKind.DataTable(
      ref = interior,
      dt2D = true,
      dtr = true,
      r1 = Some(aref(r1)),
      r2 = Some(aref(r2)),
      del1 = false,
      del2 = false,
      ca = true
    )

  // ===== Happy paths =====

  test("2-D: corner-only record with the Excel dialect (dt2D=1, dtr=1, r1=row, r2=col, ca=1)") {
    val authored = expectRight(base2D.dataTable(interior2D, ref"B1", ref"B2"))
    assertEquals(
      authored(ref"D5").value,
      CellValue.dataTable(kind2D(interior2D, "B1", "B2"), None)
    )
    // Every other interior cell stays untouched (here: absent).
    List(ref"E5", ref"F5", ref"D6", ref"E6", ref"F6").foreach { r =>
      assert(authored.cells.get(r).isEmpty, s"non-corner interior ${r.toA1} must stay untouched")
    }
    // Corner formula and axes untouched.
    assertEquals(authored(ref"C4").value, CellValue.Formula("B1*B2"))
    assertEquals(authored(ref"D4").value, num(1))
  }

  test("2-D seeded: corner cache = seeds(0)(0), non-corner interiors become plain seed values") {
    val seeds = Seq(
      Seq(num(1000), num(2000), num(3000)),
      Seq(num(2000), num(4000), num(6000))
    )
    val authored = expectRight(base2D.dataTable(interior2D, ref"B1", ref"B2", seeds))
    assertEquals(
      authored(ref"D5").value,
      CellValue.dataTable(kind2D(interior2D, "B1", "B2"), Some(num(1000)))
    )
    assertEquals(authored(ref"E5").value, num(2000))
    assertEquals(authored(ref"F5").value, num(3000))
    assertEquals(authored(ref"D6").value, num(2000))
    assertEquals(authored(ref"E6").value, num(4000))
    assertEquals(authored(ref"F6").value, num(6000))
  }

  test("2-D: an Error seed caches as CellValue.Error (#NUM! interiors are legitimate data)") {
    val numErr = CellValue.Error(CellError.Num)
    val seeds = Seq(Seq(numErr, num(2), num(3)), Seq(num(4), num(5), num(6)))
    val authored = expectRight(base2D.dataTable(interior2D, ref"B1", ref"B2", seeds))
    assertEquals(
      authored(ref"D5").value,
      CellValue.dataTable(kind2D(interior2D, "B1", "B2"), Some(numErr))
    )
  }

  test("1-D row-oriented: dt2D=0, dtr=1, the input rides r1") {
    // Interior B21:D21 -> axis row B20:D20, source formula column A21 (one per interior row).
    val sheet = Sheet("S")
      .put(ref"A1", num(2))
      .put(ref"B20", num(7))
      .put(ref"C20", num(8))
      .put(ref"D20", num(9))
      .put(ref"A21", CellValue.Formula("A1+1"))
    val authored = expectRight(sheet.dataTableRow(range("B21:D21"), ref"A1"))
    val expectedKind: FormulaKind.DataTable = FormulaKind.DataTable(
      ref = range("B21:D21"),
      dt2D = false,
      dtr = true,
      r1 = Some(aref("A1")),
      r2 = None,
      del1 = false,
      del2 = false,
      ca = true
    )
    assertEquals(authored(ref"B21").value, CellValue.dataTable(expectedKind, None))
    assertEquals(
      authored(ref"B21").value match
        case CellValue.Formula(expression, _, _) => expression
        case other => fail(s"expected formula, got $other")
      ,
      "TABLE(A1,)"
    )
    assert(authored.cells.get(ref"C21").isEmpty)
  }

  test("1-D column-oriented: dt2D=0, dtr=0, the single input STILL rides r1") {
    // Interior F10:F12 -> source formula F9 (above the column), axis column E10:E12.
    val sheet = Sheet("S")
      .put(ref"A2", num(3))
      .put(ref"F9", CellValue.Formula("A2*100"))
      .put(ref"E10", num(1))
      .put(ref"E11", num(2))
      .put(ref"E12", num(3))
    val authored = expectRight(sheet.dataTableCol(range("F10:F12"), ref"A2"))
    val expectedKind: FormulaKind.DataTable = FormulaKind.DataTable(
      ref = range("F10:F12"),
      dt2D = false,
      dtr = false,
      r1 = Some(aref("A2")),
      r2 = None,
      del1 = false,
      del2 = false,
      ca = true
    )
    assertEquals(authored(ref"F10").value, CellValue.dataTable(expectedKind, None))
    assertEquals(
      authored(ref"F10").value match
        case CellValue.Formula(expression, _, _) => expression
        case other => fail(s"expected formula, got $other")
      ,
      "TABLE(,A2)"
    )
  }

  test("multi-result-column 1-D column table: a source formula above EVERY interior column") {
    // Interior T2:U4 (two result columns) -> formulas T1 and U1, axis S2:S4.
    val sheet = Sheet("S")
      .put(ref"A2", num(3))
      .put(ref"T1", CellValue.Formula("A2*100"))
      .put(ref"U1", CellValue.Formula("A2+7"))
      .put(ref"S2", num(1))
      .put(ref"S3", num(2))
      .put(ref"S4", num(3))
    val authored = expectRight(sheet.dataTableCol(range("T2:U4"), ref"A2"))
    authored(ref"T2").value match
      case CellValue.Formula(_, _, dt: FormulaKind.DataTable) =>
        assertEquals(dt.ref, range("T2:U4"))
        assertEquals(dt.dt2D, false)
        assertEquals(dt.dtr, false)
      case other => fail(s"expected record corner, got $other")
    // Missing one of the two source formulas is refused.
    val short = sheet.remove(ref"U1")
    expectLeft(short.dataTableCol(range("T2:U4"), ref"A2")) match
      case XLError.FormulaError(expression, reason) =>
        assertEquals(expression, "U1")
        assert(reason.contains("source formula at U1"), reason)
      case other => fail(s"expected FormulaError, got $other")
  }

  test("1x1 interiors are legal") {
    val sheet = Sheet("S")
      .put(ref"A1", num(2))
      .put(ref"A2", num(3))
      .put(ref"P1", CellValue.Formula("A1*A2"))
      .put(ref"Q1", num(10))
      .put(ref"P2", num(100))
    val authored = expectRight(sheet.dataTable(range("Q2:Q2"), ref"A1", ref"A2"))
    assertEquals(
      authored(ref"Q2").value,
      CellValue.dataTable(kind2D(range("Q2:Q2"), "A1", "A2"), None)
    )
  }

  test("string overloads: happy path parses, junk refs are Left (never thrown)") {
    val authored = expectRight(base2D.dataTable("D5:F6", "B1", "B2"))
    assertEquals(
      authored(ref"D5").value,
      CellValue.dataTable(kind2D(interior2D, "B1", "B2"), None)
    )
    expectLeft(base2D.dataTable("NOT A RANGE", "B1", "B2")) match
      case XLError.InvalidRange(range, _) => assertEquals(range, "NOT A RANGE")
      case other => fail(s"expected InvalidRange, got $other")
    expectLeft(base2D.dataTable("D5:F6", "NOPE!", "B2")) match
      case XLError.InvalidCellRef(ref, _) => assertEquals(ref, "NOPE!")
      case other => fail(s"expected InvalidCellRef, got $other")
    val rowAuthored = Sheet("S")
      .put(ref"A1", num(2))
      .put(ref"B20", num(7))
      .put(ref"A21", CellValue.Formula("A1+1"))
      .dataTableRow("B21:B21", "A1")
    assert(rowAuthored.isRight, s"dataTableRow string overload failed: $rowAuthored")
    val colAuthored = Sheet("S")
      .put(ref"A2", num(3))
      .put(ref"F9", CellValue.Formula("A2*100"))
      .put(ref"E10", num(1))
      .dataTableCol("F10:F10", "A2")
    assert(colAuthored.isRight, s"dataTableCol string overload failed: $colAuthored")
  }

  // ===== V1: geometry =====

  test("V1: an interior starting in row 1 or column A has no room for corner + axes") {
    expectLeft(base2D.dataTable(range("A5:B6"), ref"H1", ref"H2")) match
      case XLError.InvalidRange(range, reason) =>
        assertEquals(range, "A5:B6")
        assert(reason.contains("corner formula one-up-one-left"), reason)
      case other => fail(s"expected InvalidRange, got $other")
    expectLeft(base2D.dataTable(range("D1:F2"), ref"H1", ref"H2")) match
      case XLError.InvalidRange(_, reason) =>
        assert(reason.contains("cannot start in row 1 or column A"), reason)
      case other => fail(s"expected InvalidRange, got $other")
  }

  // ===== V2: distinct inputs =====

  test("V2: a 2-D table needs two different input cells") {
    expectLeft(base2D.dataTable(interior2D, ref"B1", ref"B1")) match
      case XLError.InvalidReference(reason) =>
        assert(reason.contains("must be two different cells"), reason)
      case other => fail(s"expected InvalidReference, got $other")
  }

  // ===== V3: inputs outside the block =====

  test("V3: an input anywhere in the table block (corner, axes, interior) is refused") {
    // C4 = the corner cell itself; D4 = axis; E6 = interior.
    List(ref"C4", ref"D4", ref"E6").foreach { input =>
      expectLeft(base2D.dataTable(interior2D, input, ref"B2")) match
        case XLError.InvalidCellRef(bad, reason) =>
          assertEquals(bad, input.toA1)
          assert(
            reason.contains(
              "input cell reference is not valid — Excel refuses inputs anywhere in the table block"
            ),
            reason
          )
        case other => fail(s"expected InvalidCellRef, got $other")
    }
    // 1-D forms enforce the same rule on their single input.
    val rowSheet = Sheet("S")
      .put(ref"B20", num(7))
      .put(ref"A21", CellValue.Formula("A1+1"))
    expectLeft(rowSheet.dataTableRow(range("B21:D21"), ref"A21")) match
      case XLError.InvalidCellRef(bad, _) => assertEquals(bad, "A21")
      case other => fail(s"expected InvalidCellRef, got $other")
  }

  // ===== V4: source formulas present =====

  test("V4: a missing 2-D corner formula steers to put-the-corner-formula-first") {
    val noCorner = base2D.remove(ref"C4")
    expectLeft(noCorner.dataTable(interior2D, ref"B1", ref"B2")) match
      case XLError.FormulaError(expression, reason) =>
        assertEquals(expression, "C4")
        assert(reason.contains("put the corner formula first"), reason)
        assert(reason.contains("sheet.put(ref\"C4\", fx\"=B14\")"), reason)
        assert(reason.contains("CellValue.dataTable + put"), reason)
      case other => fail(s"expected FormulaError, got $other")
  }

  test("V4: a plain scalar or a DataTable record at the corner does not qualify") {
    val scalarCorner = base2D.put(ref"C4", num(42))
    expectLeft(scalarCorner.dataTable(interior2D, ref"B1", ref"B2")) match
      case XLError.FormulaError(expression, _) => assertEquals(expression, "C4")
      case other => fail(s"expected FormulaError, got $other")
    val recordCorner = base2D.put(
      ref"C4",
      CellValue.dataTable(kind2D(range("C4:C4"), "H1", "H2"), Some(num(1)))
    )
    expectLeft(recordCorner.dataTable(interior2D, ref"B1", ref"B2")) match
      case XLError.FormulaError(expression, _) => assertEquals(expression, "C4")
      case other => fail(s"expected FormulaError, got $other")
  }

  test("V4: a plain `=$B$14`-style ref corner qualifies (real corners are often bare refs)") {
    val bareRefCorner = base2D.put(ref"C4", CellValue.Formula("$B$1"))
    assert(bareRefCorner.dataTable(interior2D, ref"B1", ref"B2").isRight)
  }

  test("V4: 1-D forms demand a source formula per interior row/column and name the gap") {
    val rowSheet = Sheet("S")
      .put(ref"A1", num(2))
      .put(ref"B20", num(7))
      .put(ref"A21", CellValue.Formula("A1+1"))
    // Interior spans rows 21..22 but only A21 has a formula.
    expectLeft(rowSheet.dataTableRow(range("B21:B22"), ref"A1")) match
      case XLError.FormulaError(expression, reason) =>
        assertEquals(expression, "A22")
        assert(reason.contains("source formula at A22"), reason)
      case other => fail(s"expected FormulaError, got $other")
  }

  // ===== V5: overlap hygiene =====

  test("V5: partial overlap from outside tears the existing table and is refused") {
    // Existing record covers F6:G7 — intersects interior D5:F6 at F6 but sticks out to G7.
    val foreign = kind2D(range("F6:G7"), "H1", "H2")
    val torn = base2D.put(ref"F6", CellValue.dataTable(foreign, Some(num(5))))
    expectLeft(torn.dataTable(interior2D, ref"B1", ref"B2")) match
      case XLError.InvalidRange(range, reason) =>
        assertEquals(range, "F6:G7")
        assert(reason.contains("would tear data table at F6:G7"), reason)
        assert(reason.contains("Excel: cannot change part of a data table"), reason)
      case other => fail(s"expected InvalidRange, got $other")
  }

  test("V5 determinism: the row-major-first record cell reports the tear") {
    // Two foreign records both partially overlap; E5 precedes F6 in row-major order.
    val first = kind2D(range("E5:G5"), "H1", "H2")
    val second = kind2D(range("F6:G7"), "H1", "H2")
    val torn = base2D
      .put(ref"F6", CellValue.dataTable(second, Some(num(6))))
      .put(ref"E5", CellValue.dataTable(first, Some(num(5))))
    expectLeft(torn.dataTable(interior2D, ref"B1", ref"B2")) match
      case XLError.InvalidRange(range, _) => assertEquals(range, "E5:G5")
      case other => fail(s"expected InvalidRange, got $other")
  }

  test("V5/D6: fully-contained records are CONSUMED — every-interior books normalize") {
    // Foreign every-interior dialect (like the Sentinel books): all six interior cells carry
    // the record; one is cache-less.
    val foreign: FormulaKind.DataTable = FormulaKind.DataTable(
      ref = interior2D,
      dt2D = true,
      dtr = false, // foreign dtr=0 dialect
      r1 = Some(aref("B1")),
      r2 = Some(aref("B2"))
    )
    val everyInterior = interior2D.cellsRowMajor.zipWithIndex
      .foldLeft(base2D) { case (s, (r, i)) =>
        val cache = if r == ref"F6" then None else Some(num(1000 + i))
        s.put(r, CellValue.dataTable(foreign, cache))
      }
    val authored = expectRight(everyInterior.dataTable(interior2D, ref"B1", ref"B2"))
    // Corner: re-authored record in the authored dialect, keeping the Excel-written cache.
    assertEquals(
      authored(ref"D5").value,
      CellValue.dataTable(kind2D(interior2D, "B1", "B2"), Some(num(1000)))
    )
    // Non-corner records reverted to their caches as plain values; the cache-less one dropped.
    assertEquals(authored(ref"E5").value, num(1001))
    assertEquals(authored(ref"F5").value, num(1002))
    assertEquals(authored(ref"D6").value, num(1003))
    assertEquals(authored(ref"E6").value, num(1004))
    assert(authored.cells.get(ref"F6").isEmpty, "cache-less consumed record must drop")
  }

  test("re-author over the same interior is idempotent (dataTable . dataTable = dataTable)") {
    val once = expectRight(base2D.dataTable(interior2D, ref"B1", ref"B2"))
    val twice = expectRight(once.dataTable(interior2D, ref"B1", ref"B2"))
    assertEquals(twice, once)
    // And a seeded re-author keeps the corner cache through the unseeded round.
    val seeded = expectRight(
      base2D.dataTable(
        interior2D,
        ref"B1",
        ref"B2",
        Seq(Seq(num(1), num(2), num(3)), Seq(num(4), num(5), num(6)))
      )
    )
    val reauthored = expectRight(seeded.dataTable(interior2D, ref"B1", ref"B2"))
    assertEquals(reauthored, seeded)
  }

  // ===== V6: no silent formula destruction =====

  test("V6: a real formula inside the interior is refused with clear-first steering") {
    val withFormula = base2D.put(ref"E6", CellValue.Formula("SUM(A1:A3)"))
    expectLeft(withFormula.dataTable(interior2D, ref"B1", ref"B2")) match
      case XLError.FormulaError(expression, reason) =>
        assertEquals(expression, "SUM(A1:A3)")
        assert(reason.contains("clear E6 first (removeRange) or bake values"), reason)
      case other => fail(s"expected FormulaError, got $other")
  }

  test("V6 determinism: the row-major-first formula cell reports, not Map iteration order") {
    // Insert in anti-row-major order; E5 must still be selected over D6/F6.
    val adversarial = base2D
      .put(ref"F6", CellValue.Formula("F1*2"))
      .put(ref"D6", CellValue.Formula("D1*2"))
      .put(ref"E5", CellValue.Formula("E1*2"))
    expectLeft(adversarial.dataTable(interior2D, ref"B1", ref"B2")) match
      case XLError.FormulaError(expression, reason) =>
        assertEquals(expression, "E1*2")
        assert(reason.contains("clear E5 first"), reason)
      case other => fail(s"expected FormulaError, got $other")
  }

  // ===== V7: merges =====

  test("V7: a merged range intersecting the interior is refused (unmerge first)") {
    val merged = base2D.merge(range("E6:G8"))
    expectLeft(merged.dataTable(interior2D, ref"B1", ref"B2")) match
      case XLError.InvalidRange(range, reason) =>
        assertEquals(range, "E6:G8")
        assert(reason.contains("unmerge first"), reason)
      case other => fail(s"expected InvalidRange, got $other")
    // A merge fully outside the block is fine.
    val outside = base2D.merge(range("H1:J2"))
    assert(outside.dataTable(interior2D, ref"B1", ref"B2").isRight)
  }

  // ===== V8: seeds =====

  test("V8: seeds must be exactly interior-shaped (rows, then row width)") {
    expectLeft(
      base2D.dataTable(interior2D, ref"B1", ref"B2", Seq(Seq(num(1), num(2), num(3))))
    ) match
      case XLError.ValueCountMismatch(expected, actual, context) =>
        assertEquals(expected, 2)
        assertEquals(actual, 1)
        assertEquals(context, "data table interior D5:F6")
      case other => fail(s"expected ValueCountMismatch, got $other")
    expectLeft(
      base2D.dataTable(interior2D, ref"B1", ref"B2", Seq(Seq(num(1), num(2), num(3)), Seq(num(4))))
    ) match
      case XLError.ValueCountMismatch(expected, actual, context) =>
        assertEquals(expected, 3)
        assertEquals(actual, 1)
        assertEquals(context, "data table interior D5:F6")
      case other => fail(s"expected ValueCountMismatch, got $other")
  }

  test("V8: a Formula seed is refused before construction (require can never throw)") {
    val seeds = Seq(
      Seq(num(1), num(2), num(3)),
      Seq(num(4), CellValue.Formula("A1*2"), num(6))
    )
    expectLeft(base2D.dataTable(interior2D, ref"B1", ref"B2", seeds)) match
      case XLError.FormulaError(expression, reason) =>
        assertEquals(expression, "A1*2")
        assert(reason.contains("seeds must be plain cached values"), reason)
        assert(reason.contains("row 1 column 1"), reason)
      case other => fail(s"expected FormulaError, got $other")
  }

  // ===== Absorb rule + composition =====

  test("absorb rule: fillBy-then-dataTable keeps the corner scalar as the record cache") {
    val filled = base2D.fillBy(interior2D)((_, _) => num(777))
    val authored = expectRight(filled.dataTable(interior2D, ref"B1", ref"B2"))
    assertEquals(
      authored(ref"D5").value,
      CellValue.dataTable(kind2D(interior2D, "B1", "B2"), Some(num(777)))
    )
    // Non-corner fillBy values stay untouched as the grid's plain caches.
    assertEquals(authored(ref"E6").value, num(777))
  }

  test("authoring commutes with puts disjoint from the block") {
    val putThenAuthor = expectRight(
      base2D.put(ref"Z9", num(9)).dataTable(interior2D, ref"B1", ref"B2")
    )
    val authorThenPut = expectRight(base2D.dataTable(interior2D, ref"B1", ref"B2"))
      .put(ref"Z9", num(9))
    assertEquals(putThenAuthor, authorThenPut)
  }

  // ===== Generative law =====

  property("authored geometry law: corner record carries the D2 dialect, interiors untouched") {
    forAll(Generators.genAuthoredDataTable) { case (interior, orientation, input1, input2) =>
      val startCol = interior.start.col.index0
      val startRow = interior.start.row.index0
      val corner = ARef.from0(startCol - 1, startRow - 1)
      val sourceRefs: Vector[ARef] = orientation match
        case 0 => Vector(corner)
        case 1 => (startRow to interior.end.row.index0).toVector.map(ARef.from0(startCol - 1, _))
        case _ => (startCol to interior.end.col.index0).toVector.map(ARef.from0(_, startRow - 1))
      val sheet = sourceRefs.foldLeft(Sheet("G")) { (s, r) =>
        s.put(r, CellValue.Formula("A1*2"))
      }
      val authored = orientation match
        case 0 => sheet.dataTable(interior, input1, input2)
        case 1 => sheet.dataTableRow(interior, input1)
        case _ => sheet.dataTableCol(interior, input1)
      authored match
        case Left(err) => fail(s"valid-geometry author failed: $err")
        case Right(out) =>
          val expectedKind: FormulaKind.DataTable = orientation match
            case 0 =>
              FormulaKind.DataTable(
                interior,
                true,
                true,
                Some(input1),
                Some(input2),
                false,
                false,
                true
              )
            case 1 =>
              FormulaKind.DataTable(interior, false, true, Some(input1), None, false, false, true)
            case _ =>
              FormulaKind.DataTable(interior, false, false, Some(input1), None, false, false, true)
          assertEquals(out(interior.start).value, CellValue.dataTable(expectedKind, None))
          interior.cellsRowMajor.filterNot(_ == interior.start).foreach { r =>
            assert(out.cells.get(r).isEmpty, s"interior ${r.toA1} must stay untouched")
          }
          true
    }
  }
