package com.tjclp.xl.cli.batch

import munit.FunSuite

import cats.effect.unsafe
import com.tjclp.xl.{Sheet, Workbook}
import com.tjclp.xl.addressing.ARef
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.cli.helpers.{AppearanceOps, BatchParser}
import com.tjclp.xl.extensions.style
import com.tjclp.xl.sheets.syntax.*
import com.tjclp.xl.styles.CellStyle
import com.tjclp.xl.styles.font.Font
import com.tjclp.xl.styles.numfmt.NumFmt

/**
 * GH-560 (ADR-017 invariant 4): a hint on a value is Inferred or Explicit, and only an Explicit
 * `format` overrides a non-General numFmt. The inferred hint (a detected date or currency string
 * with no `format` key) keeps `Sheet.put`'s General-only merge rule.
 */
@SuppressWarnings(Array("org.wartremover.warts.OptionPartial"))
class BatchPutFormatSpec extends FunSuite:

  given unsafe.IORuntime = unsafe.IORuntime.global

  private val a1 = ARef.from0(0, 0)
  private val a2 = ARef.from0(0, 1)

  /** The issue's fixture: a percent code with negatives in parentheses. */
  private val percentCode = """0.0%_);\(0.0%\)"""

  /** The issue's target: an integer code with negatives in parentheses. */
  private val integerCode = """#,##0_);\(#,##0\)"""

  private def styled(ref: ARef, numFmt: NumFmt): Sheet =
    val style = CellStyle.default.withFont(Font.default.withBold()).withNumFmt(numFmt)
    Sheet("Data").put(ref, CellValue.Number(BigDecimal("0.5"))).style(ref, style)

  private def run(sheet: Sheet, json: String): Sheet =
    val wb = Workbook(Vector(sheet))
    BatchParser
      .parseBatchOperations(json)
      .flatMap(result => BatchParser.applyScoped(wb, Some(sheet), result.scoped))
      .unsafeRunSync()
      .sheets
      .headOption
      .getOrElse(fail("no sheet"))

  private def numFmtOf(sheet: Sheet, ref: ARef): Option[NumFmt] =
    sheet.getCellStyle(ref).map(_.numFmt)

  test(
    "GH-560: an explicit format replaces a custom number format and keeps the rest of the style"
  ) {
    val before = styled(a1, NumFmt.Custom(percentCode))
    // JSON needs the backslashes doubled; the parsed code is the issue's exact text
    val json = """[{"op":"put","ref":"A1","value":1234,"format":"#,##0_);\\(#,##0\\)"}]"""
    val after = run(before, json)
    assertEquals(numFmtOf(after, a1), Some(NumFmt.Custom(integerCode)))
    assertEquals(after.cells.get(a1).map(_.value), Some(CellValue.Number(BigDecimal(1234))))
    assert(after.getCellStyle(a1).exists(_.font.bold), "the explicit format touches only numFmt")
  }

  test("GH-560: an explicit named format replaces a Currency numFmt") {
    val after = run(
      styled(a1, NumFmt.Currency),
      """[{"op":"put","ref":"A1","value":0.25,"format":"percent"}]"""
    )
    assertEquals(numFmtOf(after, a1), Some(NumFmt.Percent))
  }

  test("an inferred date hint does not override a Currency numFmt") {
    val after =
      run(styled(a1, NumFmt.Currency), """[{"op":"put","ref":"A1","value":"2025-01-15"}]""")
    assertEquals(numFmtOf(after, a1), Some(NumFmt.Currency))
    assert(after.getCellStyle(a1).exists(_.font.bold))
    after.cells.get(a1).map(_.value) match
      case Some(CellValue.DateTime(_)) | Some(CellValue.Number(_)) => ()
      case other => fail(s"expected the detected date value, got $other")
  }

  test("an inferred currency hint does not override a Custom numFmt") {
    val after = run(
      styled(a1, NumFmt.Custom(percentCode)),
      """[{"op":"put","ref":"A1","value":"$1,234.56"}]"""
    )
    assertEquals(numFmtOf(after, a1), Some(NumFmt.Custom(percentCode)))
  }

  test("an inferred hint still applies onto a General cell (the rule's positive half)") {
    val plain = Sheet("Data").put(a1, CellValue.Text("x"))
    val after = run(plain, """[{"op":"put","ref":"A1","value":"2025-01-15"}]""")
    assertEquals(numFmtOf(after, a1), Some(NumFmt.Date))
  }

  test("an explicit format applies onto an unstyled cell") {
    val plain = Sheet("Data")
    val after = run(plain, """[{"op":"put","ref":"A1","value":0.5,"format":"percent"}]""")
    assertEquals(numFmtOf(after, a1), Some(NumFmt.Percent))
  }

  test(
    "an explicit format on a values array replaces the numFmt of every target (GH-416 x GH-560)"
  ) {
    val before = styled(a1, NumFmt.Custom(percentCode))
      .style(a2, CellStyle.default.withNumFmt(NumFmt.Currency))
    val after = run(before, """[{"op":"put","ref":"A1:A2","values":[1,2],"format":"percent"}]""")
    assertEquals(numFmtOf(after, a1), Some(NumFmt.Percent))
    assertEquals(numFmtOf(after, a2), Some(NumFmt.Percent))
    assert(after.getCellStyle(a1).exists(_.font.bold))
  }

  test("inferred hints inside a values array keep the General-only rule") {
    val before = styled(a1, NumFmt.Custom(percentCode))
    val after = run(before, """[{"op":"put","ref":"A1:A2","values":["$5","2025-01-15"]}]""")
    assertEquals(numFmtOf(after, a1), Some(NumFmt.Custom(percentCode)))
    assertEquals(numFmtOf(after, a2), Some(NumFmt.Date), "A2 was General, so the hint applies")
  }

  test("the numFormat alias on put is the same explicit format") {
    val after = run(
      styled(a1, NumFmt.Currency),
      """[{"op":"put","ref":"A1","value":1,"numFormat":"percent"}]"""
    )
    assertEquals(numFmtOf(after, a1), Some(NumFmt.Percent))
  }

  test("GH-463: page-setup accepts fitToHeight 0 (automatic) through batch") {
    val after = run(Sheet("Data"), """[{"op":"page-setup","fitToWidth":1,"fitToHeight":0}]""")
    assertEquals(after.pageSetup.flatMap(_.fitToHeight), Some(0))
    assertEquals(after.pageSetup.flatMap(_.fitToWidth), Some(1))
  }

  test("GH-463: the shared page-setup applier accepts 0 and still rejects negatives") {
    assert(AppearanceOps.applyPageSetup(Sheet("Data"), None, None, Some(1), Some(0), None).isRight)
    assert(AppearanceOps.applyPageSetup(Sheet("Data"), None, None, Some(0), Some(1), None).isRight)
    assert(AppearanceOps.applyPageSetup(Sheet("Data"), None, None, Some(1), Some(-1), None).isLeft)
    assert(AppearanceOps.applyPageSetup(Sheet("Data"), None, None, Some(-1), Some(1), None).isLeft)
  }
