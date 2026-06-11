package com.tjclp.xl.formula

import com.tjclp.xl.{*, given}
import com.tjclp.xl.addressing.{ARef, SheetName}
import com.tjclp.xl.cells.{CellError, CellValue}
import com.tjclp.xl.workbooks.Workbook
import munit.ScalaCheckSuite
import org.scalacheck.Gen
import org.scalacheck.Prop.*

/**
 * GH-274: INDIRECT(ref_text, [a1]) — dynamic text-to-reference resolution.
 *
 * INDIRECT resolves A1-style text to a cell/range at evaluation time and rides the same
 * ArrayResult mechanism as OFFSET (GH-122): aggregates flatten it, array formulas spill it,
 * scalar contexts collapse 1×1 results. Unresolvable text is the #REF! VALUE (total, never
 * throws); R1C1 mode (a1=FALSE) is a documented-unsupported eval error.
 */
class IndirectFunctionSpec extends ScalaCheckSuite:

  private def num(i: Int): CellValue = CellValue.Number(BigDecimal(i))
  private val refErr: CellValue = CellValue.Error(CellError.Ref)

  private val s = new Sheet(name = SheetName.unsafe("S"))
  private val base = s.put("A1" -> 10).put("B2" -> 42)
  private val col = s.put("A1" -> 10).put("A2" -> 20).put("A3" -> 30)

  // ===== 1. Quote laws: INDIRECT("X") ≡ X =====

  test("INDIRECT(\"B2\") returns B2's value (quote law)") {
    assertEquals(base.evaluateFormula("=INDIRECT(\"B2\")"), base.evaluateFormula("=B2"))
    assertEquals(base.evaluateFormula("=INDIRECT(\"B2\")"), Right(num(42)))
  }

  test("INDIRECT of an empty cell preserves emptiness (OFFSET parity)") {
    // Direct `=Z9` resolves Empty -> 0 (decodeResolvedValue); the array path preserves
    // Empty so aggregate semantics stay correct (e.g. COUNTA over an empty target is 0).
    // Both render as zero-valued; pinned deliberately.
    assertEquals(base.evaluateFormula("=INDIRECT(\"Z9\")"), Right(CellValue.Empty))
    assertEquals(base.evaluateFormula("=SUM(INDIRECT(\"Z9\"))"), Right(num(0)))
  }

  test("INDIRECT(\"$B$2\") accepts anchored text") {
    assertEquals(base.evaluateFormula("=INDIRECT(\"$B$2\")"), Right(num(42)))
  }

  test("INDIRECT(\"b2\") is case-insensitive") {
    assertEquals(base.evaluateFormula("=INDIRECT(\"b2\")"), Right(num(42)))
  }

  test("INDIRECT over a cached formula target uses the cache") {
    val sheet = base.put(ref"C1", CellValue.Formula("=B2*2", Some(num(84))))
    assertEquals(sheet.evaluateFormula("=INDIRECT(\"C1\")"), Right(num(84)))
  }

  test("INDIRECT over an UNCACHED formula target evaluates it (eval-aware extractor)") {
    val sheet = base.put(ref"C1", CellValue.Formula("=B2*2", None))
    assertEquals(sheet.evaluateFormula("=INDIRECT(\"C1\")"), Right(num(84)))
  }

  test("SUM(INDIRECT(range)) evaluates uncached formulas inside the target range") {
    val sheet = s
      .put("A1" -> 1)
      .put(ref"A2", CellValue.Formula("=A1+1", None))
      .put(ref"A3", CellValue.Formula("=A1+2", Some(num(3))))
    assertEquals(sheet.evaluateFormula("=SUM(INDIRECT(\"A1:A3\"))"), Right(num(6)))
  }

  test("INDIRECT trims surrounding whitespace in ref_text") {
    assertEquals(base.evaluateFormula("=INDIRECT(\" B2 \")"), Right(num(42)))
  }

  // ===== 2. Dynamic text =====

  test("INDIRECT(A1) where A1 holds 'B1' dereferences the contained text") {
    val sheet = s.put("A1" -> "B1").put("B1" -> 7)
    assertEquals(sheet.evaluateFormula("=INDIRECT(A1)"), Right(num(7)))
  }

  test("INDIRECT(\"B\"&\"2\") accepts concatenated text") {
    assertEquals(base.evaluateFormula("=INDIRECT(\"B\"&\"2\")"), Right(num(42)))
  }

  test("INDIRECT(\"B\"&C1) builds text from a numeric cell") {
    val sheet = base.put("C1" -> 2)
    assertEquals(sheet.evaluateFormula("=INDIRECT(\"B\"&C1)"), Right(num(42)))
  }

  test("INDIRECT(ADDRESS(2,2)) composes with ADDRESS") {
    assertEquals(base.evaluateFormula("=INDIRECT(ADDRESS(2,2))"), Right(num(42)))
  }

  test("INDIRECT of a numeric cell yields #REF! (text '5' is not a reference)") {
    val sheet = s.put("A1" -> 5)
    assertEquals(sheet.evaluateFormula("=INDIRECT(A1)"), Right(refErr))
  }

  // ===== 3. Aggregates (the headline) =====

  test("SUM/AVERAGE/MAX/COUNT over INDIRECT(\"A1:A3\")") {
    assertEquals(col.evaluateFormula("=SUM(INDIRECT(\"A1:A3\"))"), Right(num(60)))
    assertEquals(col.evaluateFormula("=AVERAGE(INDIRECT(\"A1:A3\"))"), Right(num(20)))
    assertEquals(col.evaluateFormula("=MAX(INDIRECT(\"A1:A3\"))"), Right(num(30)))
    assertEquals(col.evaluateFormula("=COUNT(INDIRECT(\"A1:A3\"))"), Right(num(3)))
  }

  test("SUM(INDIRECT(\"A1:A\"&B1)) with a dynamic row bound") {
    val sheet = col.put("B1" -> 2)
    assertEquals(sheet.evaluateFormula("=SUM(INDIRECT(\"A1:A\"&B1))"), Right(num(30)))
  }

  test("full-column text bounds to the used range: SUM(INDIRECT(\"A:A\")) == SUM(A:A)") {
    assertEquals(
      col.evaluateFormula("=SUM(INDIRECT(\"A:A\"))"),
      col.evaluateFormula("=SUM(A:A)")
    )
    assertEquals(col.evaluateFormula("=SUM(INDIRECT(\"A:A\"))"), Right(num(60)))
  }

  test("full column over an empty sheet collapses to the empty array (SUM = 0)") {
    assertEquals(s.evaluateFormula("=SUM(INDIRECT(\"A:A\"))"), Right(num(0)))
  }

  test("whole-sheet text (full column AND full row) bounds to the used range") {
    assertEquals(col.evaluateFormula("=SUM(INDIRECT(\"A1:XFD1048576\"))"), Right(num(60)))
  }

  test("oversize resolved range is capped (Left), not materialized") {
    // B2:XFD1048576 is neither a full column (rows 2..max) nor a full row (cols B..XFD),
    // so no used-range bounding applies — the 1,048,576-cell cap must reject it quickly.
    col.evaluateFormula("=INDIRECT(\"B2:XFD1048576\")") match
      case Left(err) => assert(err.message.contains("1,048,576"), err.message)
      case Right(v) => fail(s"expected Left(cap), got $v")
  }

  // ===== 4. Cross-sheet =====

  test("INDIRECT(\"S2!B2\") resolves cross-sheet with workbook context") {
    val s1 = new Sheet(name = SheetName.unsafe("S1"))
    val s2 = (new Sheet(name = SheetName.unsafe("S2"))).put("B2" -> 99)
    val wb = Workbook(s1, s2)
    assertEquals(wb.evaluateFormula("=INDIRECT(\"S2!B2\")", "S1"), Right(num(99)))
  }

  test("INDIRECT(\"'My Sheet'!A1:B2\") quoted sheet range composes with SUM") {
    val s1 = new Sheet(name = SheetName.unsafe("S1"))
    val s2 = (new Sheet(name = SheetName.unsafe("My Sheet")))
      .put("A1" -> 1)
      .put("B1" -> 2)
      .put("A2" -> 3)
      .put("B2" -> 4)
    val wb = Workbook(s1, s2)
    assertEquals(wb.evaluateFormula("=SUM(INDIRECT(\"'My Sheet'!A1:B2\"))", "S1"), Right(num(10)))
  }

  test("GH-263: quoted cell-ref-shaped sheet name in ref_text — INDIRECT(\"'Q1'!A1\")") {
    val s1 = new Sheet(name = SheetName.unsafe("S1"))
    val q1 = (new Sheet(name = SheetName.unsafe("Q1"))).put("A1" -> 5)
    val wb = Workbook(s1, q1)
    assertEquals(wb.evaluateFormula("=INDIRECT(\"'Q1'!A1\")", "S1"), Right(num(5)))
  }

  test("missing sheet named in ref_text yields #REF! (the sheet name is data)") {
    val s1 = new Sheet(name = SheetName.unsafe("S1"))
    val wb = Workbook(s1)
    assertEquals(wb.evaluateFormula("=INDIRECT(\"Nope!A1\")", "S1"), Right(refErr))
  }

  test("sheet-qualified text without workbook context is a config error (Left)") {
    val res = s.evaluateFormula("=INDIRECT(\"S2!A1\")")
    res match
      case Left(err) => assert(err.message.contains("workbook"), err.message)
      case Right(v) => fail(s"expected Left(missing workbook), got $v")
  }

  // ===== 5. Error table (D11: total, no throw) =====

  test("unresolvable / non-reference text yields the #REF! value") {
    val sheet = s.put("A1" -> 1)
    val cases = List(
      "\"\"", // empty
      "\"hello\"", // bare name
      "\"5\"", // number literal
      "\"A0\"", // row 0 out of grid
      "\"XFE1\"", // column beyond XFD
      "\"A1048577\"", // row beyond 1048576
      "\"A1+B2\"", // expression, not a reference
      "\"[Book1.xlsx]Sheet1!A1\"", // external workbook
      "\"Table1[Col]\"", // structured reference
      "\"R1C1\"" // R1C1-style text in A1 mode
    )
    cases.foreach { lit =>
      assertEquals(sheet.evaluateFormula(s"=INDIRECT($lit)"), Right(refErr), lit)
    }
  }

  test("INDIRECT(\"A1\", TRUE) works; a1=FALSE is a documented-unsupported Left") {
    val sheet = s.put("A1" -> 11)
    assertEquals(sheet.evaluateFormula("=INDIRECT(\"A1\", TRUE)"), Right(num(11)))
    sheet.evaluateFormula("=INDIRECT(\"A1\", FALSE)") match
      case Left(err) => assert(err.message.contains("R1C1"), err.message)
      case Right(v) => fail(s"expected Left(R1C1 unsupported), got $v")
  }

  // ===== 6. Scalar argument positions (GH-302: implicit-intersection collapse) =====

  test("IFERROR(INDIRECT(\"bad\"),0) returns the fallback") {
    assertEquals(s.evaluateFormula("=IFERROR(INDIRECT(\"bad\"),0)"), Right(num(0)))
  }

  test("GH-302: IFERROR(INDIRECT(\"B2\"),0) returns B2's VALUE for a valid ref") {
    // Scalar argument positions collapse a 1×1 ArrayResult to its value (implicit
    // intersection) instead of rejecting it — previously the rejection made IFERROR
    // return its fallback even for valid references.
    assertEquals(base.evaluateFormula("=IFERROR(INDIRECT(\"B2\"),0)"), Right(num(42)))
  }

  test("GH-302: LEFT(INDIRECT(\"B2\"), 1) collapses to text position") {
    assertEquals(
      base.evaluateFormula("=LEFT(INDIRECT(\"B2\"), 1)"),
      Right(CellValue.Text("4"))
    )
  }

  test("GH-302: ABS(INDIRECT(\"B2\")) collapses 1x1 to numeric position") {
    assertEquals(base.evaluateFormula("=ABS(INDIRECT(\"B2\"))"), Right(num(42)))
  }

  test("GH-302: multi-cell INDIRECT in a scalar position collapses to top-left") {
    assertEquals(col.evaluateFormula("=ABS(INDIRECT(\"A1:A3\"))"), Right(num(10)))
    assertEquals(col.evaluateFormula("=IFERROR(INDIRECT(\"A1:A3\"),0)"), Right(num(10)))
  }

  test("GH-302: truly invalid ref still fires the IFERROR fallback") {
    // The collapsed value of INDIRECT("bad") is the #REF! VALUE — IFERROR must keep
    // catching it after the collapse change.
    assertEquals(s.evaluateFormula("=IFERROR(INDIRECT(\"nope\"),7)"), Right(num(7)))
    assertEquals(s.evaluateFormula("=ISERROR(INDIRECT(\"nope\"))"), Right(CellValue.Bool(true)))
  }

  test("GH-302: ISNUMBER(INDIRECT(\"B2\")) sees the collapsed number") {
    assertEquals(base.evaluateFormula("=ISNUMBER(INDIRECT(\"B2\"))"), Right(CellValue.Bool(true)))
  }

  // ===== 6b. Scalar-mode OPERATOR positions collapse like argument positions (GH-302) =====

  test("GH-302: INDIRECT in arithmetic collapses in scalar mode") {
    // Operator positions collapse 1×1 ArrayResults exactly like scalar argument
    // positions (=ABS(INDIRECT("B2")) already worked; =INDIRECT("B2")+1 must too).
    assertEquals(base.evaluateFormula("=INDIRECT(\"B2\")+1"), Right(num(43)))
    assertEquals(base.evaluateFormula("=INDIRECT(\"A1\")+INDIRECT(\"B2\")"), Right(num(52)))
    assertEquals(base.evaluateFormula("=INDIRECT(\"B2\")*10"), Right(num(420)))
  }

  test("GH-302: multi-cell INDIRECT in arithmetic collapses to top-left") {
    // col A1:A3 = 10,20,30 — same top-left convention as the standalone collapse.
    assertEquals(col.evaluateFormula("=INDIRECT(\"A1:A3\")*10"), Right(num(100)))
  }

  test("GH-302: INDIRECT in comparison collapses in scalar mode") {
    assertEquals(base.evaluateFormula("=INDIRECT(\"B2\")>15"), Right(CellValue.Bool(true)))
    assertEquals(base.evaluateFormula("=INDIRECT(\"A1\")>15"), Right(CellValue.Bool(false)))
  }

  test("GH-302: IF over an INDIRECT comparison works in scalar mode") {
    assertEquals(
      base.evaluateFormula("=IF(INDIRECT(\"B2\")>15,\"big\",\"small\")"),
      Right(CellValue.Text("big"))
    )
    assertEquals(
      base.evaluateFormula("=IF(INDIRECT(\"A1\")>15,\"big\",\"small\")"),
      Right(CellValue.Text("small"))
    )
  }

  test("GH-302: INDIRECT equality collapses in scalar mode (all three operand shapes)") {
    // array=scalar, scalar=array, array=array — the three equality branches
    assertEquals(base.evaluateFormula("=INDIRECT(\"B2\")=42"), Right(CellValue.Bool(true)))
    assertEquals(base.evaluateFormula("=42=INDIRECT(\"B2\")"), Right(CellValue.Bool(true)))
    assertEquals(
      base.evaluateFormula("=INDIRECT(\"A1\")=INDIRECT(\"B2\")"),
      Right(CellValue.Bool(false))
    )
    assertEquals(base.evaluateFormula("=INDIRECT(\"B2\")<>42"), Right(CellValue.Bool(false)))
  }

  // ===== 7. Spill / collapse / broadcast =====

  test("INDIRECT(\"A1:B2\") spills 2x2 via evaluateArrayFormula") {
    val sheet = s.put("A1" -> 1).put("B1" -> 2).put("A2" -> 3).put("B2" -> 4)
    val (out, range) =
      sheet.evaluateArrayFormula("=INDIRECT(\"A1:B2\")", ref"E1").fold(e => fail(e.message), identity)
    assertEquals(range.height, 2)
    assertEquals(range.width, 2)
    assertEquals(out(ref"E1").value, num(1))
    assertEquals(out(ref"F1").value, num(2))
    assertEquals(out(ref"E2").value, num(3))
    assertEquals(out(ref"F2").value, num(4))
  }

  test("standalone multi-cell INDIRECT collapses to top-left in scalar eval") {
    assertEquals(col.evaluateFormula("=INDIRECT(\"A1:A3\")"), Right(num(10)))
  }

  test("INDIRECT(\"A1:A3\")*10 broadcasts as an array formula") {
    val (out, range) =
      col.evaluateArrayFormula("=INDIRECT(\"A1:A3\")*10", ref"E1").fold(e => fail(e.message), identity)
    assertEquals(range.height, 3)
    assertEquals(out(ref"E1").value, num(100))
    assertEquals(out(ref"E2").value, num(200))
    assertEquals(out(ref"E3").value, num(300))
  }

  // ===== 8. Properties =====

  private val genRef: Gen[ARef] =
    for
      c <- Gen.choose(0, 30)
      r <- Gen.choose(0, 30)
    yield ARef.from0(c, r)

  property("INDIRECT(ref.toA1) ≡ direct deref for value-bearing cells") {
    forAll(genRef) { (target: ARef) =>
      val sheet = s.put(target, num(7))
      val direct = sheet.evaluateFormula("=" + target.toA1)
      val indirect = sheet.evaluateFormula(s"=INDIRECT(\"${target.toA1}\")")
      indirect == direct && indirect == Right(num(7))
    }
  }

  property("SUM(INDIRECT(r)) == SUM(r) over mixed values and formulas") {
    val genRange: Gen[CellRange] =
      for
        c1 <- Gen.choose(0, 5)
        r1 <- Gen.choose(0, 5)
        c2 <- Gen.choose(0, 5)
        r2 <- Gen.choose(0, 5)
      yield CellRange(
        ARef.from0(math.min(c1, c2), math.min(r1, r2)),
        ARef.from0(math.max(c1, c2), math.max(r1, r2))
      )
    val seeded = s
      .put("A1" -> 1)
      .put("B2" -> 2)
      .put("C3" -> 3)
      .put(ref"D4", CellValue.Formula("=A1+1", None)) // uncached formula
      .put(ref"E5", CellValue.Formula("=B2*2", Some(num(4)))) // cached formula
    forAll(genRange) { (r: CellRange) =>
      seeded.evaluateFormula(s"=SUM(INDIRECT(\"${r.toA1}\"))") ==
        seeded.evaluateFormula(s"=SUM(${r.toA1})")
    }
  }

  property("INDIRECT is total over arbitrary text (never throws)") {
    forAll { (text: String) =>
      val sheet = s.put("A1" -> text)
      sheet.evaluateFormula("=INDIRECT(A1)") match
        case Right(_) => true
        case Left(_) => true // e.g. sheet-qualified text without workbook context — still total
    }
  }

  property("non-reference text yields the #REF! value") {
    val genNonRef = Gen.oneOf(
      "hello world",
      "12.5",
      "=SUM",
      "A+B",
      "!!",
      "Sheet1!",
      ":",
      "A1:B",
      "ZZZZZ99999999"
    )
    forAll(genNonRef) { (text: String) =>
      val sheet = s.put("A1" -> text)
      sheet.evaluateFormula("=INDIRECT(A1)") == Right(refErr)
    }
  }
