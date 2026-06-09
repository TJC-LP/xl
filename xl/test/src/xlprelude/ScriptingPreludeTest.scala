// Deliberately OUTSIDE com.tjclp.xl: scripts import the prelude from a fresh namespace,
// so this test must not see the package-level exports that files under com.tjclp.xl inherit.
package xlprelude

import java.time.{LocalDate, LocalDateTime}

import munit.FunSuite

import com.tjclp.xl.scripting.{*, given}

/**
 * Gate test for the scripting prelude: every public surface a script touches must resolve
 * through `import com.tjclp.xl.scripting.{*, given}` alone — macros (transparent inline through
 * an export hop), given instances, DSL operators, evaluator extensions, sync IO, and the unsafe
 * boundary. Compile success is most of the test; runtime assertions confirm semantics.
 */
class ScriptingPreludeTest extends FunSuite:

  test("compile-time ref macros resolve through the prelude"):
    val cell: ARef = ref"A1"
    val range: CellRange = ref"A1:B10"
    assertEquals(cell.toA1, "A1")
    assertEquals(range.start.toA1, "A1")
    assertEquals(range.end.toA1, "B10")

  test("formatted literal macros resolve through the prelude"):
    val m: Formatted = money"$$1,234.56"
    val p: Formatted = percent"12.5%"
    val d: Formatted = date"2025-11-24"
    val f: CellValue = fx"=SUM(A1:B10)"
    assert(m.numFmt != NumFmt.General)
    assert(p.numFmt != NumFmt.General)
    assert(d.numFmt != NumFmt.General)
    f match
      case CellValue.Formula(expr, _) => assert(expr.nonEmpty)
      case other => fail(s"fx literal should produce CellValue.Formula, got $other")

  test("Sheet construction and literal-string puts are infallible"):
    val sheet = Sheet("Demo")
      .put("A1", "Title")
      .put("B1", 42)
    assertEquals(sheet.cells.size, 2)

  test("CellCodec givens and conversions resolve: typed put and readTyped"):
    val sheet = Sheet("Typed")
      .put(ref"A1", "text")
      .put(ref"B1", 42)
      .put(ref"C1", BigDecimal("1000.50"))
      .put(ref"D1", LocalDate.of(2025, 1, 15))
      .put(ref"E1", LocalDateTime.of(2025, 1, 15, 10, 30))
    assertEquals(sheet.readTyped[Int](ref"B1"), Right(Some(42)): Either[CodecError, Option[Int]])
    assertEquals(
      sheet.readTyped[String](ref"A1"),
      Right(Some("text")): Either[CodecError, Option[String]]
    )
    assertEquals(
      sheet.readTyped[BigDecimal](ref"C1"),
      Right(Some(BigDecimal("1000.50"))): Either[CodecError, Option[BigDecimal]]
    )

  test("Patch DSL composes: :=, ++, styled, merge"):
    val patch =
      (ref"A1" := "Report") ++
        ref"A1".styled(CellStyle.default.bold) ++
        ref"A1:C1".merge ++
        (ref"B2" := 19.99)
    val sheet = Sheet("Patched").put(patch)
    assertEquals(sheet.cells.size, 2)
    assertEquals(sheet.mergedRanges.size, 1)

  test("range fill := and ARef navigation resolve through the prelude"):
    val sheet = Sheet("Fill").put(ref"A1:B2" := 0)
    assertEquals(sheet.cells.size, 4)
    assertEquals(ref"A1".down(2).right(1), ref"B3")

  test("runtime ref interpolation returns Either and := works on RefType"):
    val row = "5"
    val patch = (for refType <- ref"A$row"
    yield refType := "dynamic").getOrElse(Patch.empty)
    val sheet = Sheet("Dyn").put(patch)
    assertEquals(sheet.cells.size, 1)

  test("rich text extensions resolve"):
    val rich = "Status: ".bold + "ACTIVE".green
    val sheet = Sheet("Rich").put(ref"A1", rich)
    assertEquals(sheet.cells.size, 1)

  test("unsafe boundary resolves: .unsafe unwraps XLResult"):
    val runtimeRef = "A1"
    val sheet = Sheet("Unsafe").put(runtimeRef, "value").unsafe
    assertEquals(sheet.cells.size, 1)
    // Literal invalid refs fail at compile time (verified: the macro rejects them through the
    // prelude hop); a runtime string is needed to exercise the XLException path.
    val invalidRef = "NOT A REF!!!"
    intercept[XLException]:
      Sheet("Unsafe").put(invalidRef, "boom").unsafe

  test("formula evaluation extensions resolve through the prelude"):
    val sheet = Sheet("Calc")
      .put(ref"A1", 2)
      .put(ref"A2", 3)
    val result = sheet.evaluateFormula("=SUM(A1:A2)")
    assert(result.exists {
      case CellValue.Number(n) => n == BigDecimal(5)
      case _ => false
    })
    assert(sheet.evaluateWithDependencyCheck().isRight)
    assert(FormulaParser.parse("=SUM(A1:A10)").isRight)
    val clock = Clock.system
    assert(clock != null || true)

  test("display interpolator resolves with given Sheet"):
    val sheet = Sheet("Disp").put(ref"A1", 100)
    given Sheet = sheet
    assertEquals(excel"${ref"A1"}", "100")

  test("upsert, wb.evaluateFormula, and readTypedOr resolve through the prelude"):
    val wb = Workbook(Sheet("Sales").put(ref"A1", 10).put(ref"A2", 20))
      .upsert("Summary", _.put(ref"A1", 5))
    assertEquals(
      wb.evaluateFormula("=SUM(Sales!A1:A2) + A1", "Summary"),
      Right(CellValue.Number(BigDecimal(35))): XLResult[CellValue]
    )
    val sheet = wb.sheets.headOption.getOrElse(fail("missing sheet"))
    assertEquals(sheet.readTypedOr[Int](ref"A1", 0), 10)
    assertEquals(sheet.readTypedOpt[Int](ref"Z9"), None)

  test("workbook construction, withCachedFormulas, and sync Excel round-trip"):
    val sheet = Sheet("Data")
      .put(ref"A1", 10)
      .put(ref"A2", 20)
      .put(ref"A3", fx"=SUM(A1:A2)")
    val wb = Workbook(sheet).withCachedFormulas()
    val path = java.nio.file.Files.createTempDirectory("xl-prelude").resolve("roundtrip.xlsx")
    Excel.write(wb, path.toString)
    val loaded = Excel.read(path.toString)
    assertEquals(loaded.sheets.size, 1)
    assertEquals(
      loaded.sheets.headOption.map(_.cells.size),
      Some(3)
    )

  test("ExcelIO escape hatch is reachable"):
    val io = ExcelIO
    assert(io != null || true)
