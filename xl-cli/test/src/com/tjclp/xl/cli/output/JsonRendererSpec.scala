package com.tjclp.xl.cli.output

import cats.effect.IO
import munit.CatsEffectSuite

import com.tjclp.xl.{*, given}
import com.tjclp.xl.addressing.CellRange
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.cli.ViewFormat
import com.tjclp.xl.cli.commands.ReadCommands
import com.tjclp.xl.macros.ref

/**
 * GH-357: JSON output carries a dedicated "formula" field for formula cells.
 *
 * Formula cells always render as {ref, type: "formula", formula, value, formatted}: `formula` holds
 * the "=…" expression while value/formatted hold the computed (evalFormulas) or cached value —
 * null/"" when uncached. --formulas no longer swaps the JSON payload (it only affects non-JSON
 * display formats); non-formula cells keep the historical {ref, type, value, formatted} shape
 * byte-for-byte.
 */
class JsonRendererSpec extends CatsEffectSuite:

  private val sheet = Sheet("Data")
    .put(ref"A1", CellValue.Number(BigDecimal(42)))
    .put(ref"B1", CellValue.Text("Revenue"))
    .put(ref"C1", CellValue.Formula("=A1*2", Some(CellValue.Number(BigDecimal(84)))))
    .put(ref"D1", CellValue.Formula("=A1+1")) // uncached: no value until eval
    .put(ref"E1", CellValue.Bool(true))

  private val row1 = CellRange(ref"A1", ref"F1") // F1 unset → empty cell

  private def countOccurrences(haystack: String, needle: String): Int =
    haystack.split(java.util.regex.Pattern.quote(needle), -1).length - 1

  // ========== rows path: formula cells ==========

  test("rows: cached formula carries formula, cached value, and formatted value") {
    val out = JsonRenderer.renderRange(sheet, row1)
    assert(
      out.contains(
        """{"ref": "C1", "type": "formula", "formula": "=A1*2", "value": 84, "formatted": "84"}"""
      ),
      s"missing formula cell shape:\n$out"
    )
  }

  test("rows: uncached formula renders null value and blank formatted") {
    val out = JsonRenderer.renderRange(sheet, row1)
    assert(
      out.contains(
        """{"ref": "D1", "type": "formula", "formula": "=A1+1", "value": null, "formatted": ""}"""
      ),
      s"missing uncached formula shape:\n$out"
    )
  }

  test("rows: cache rewrite (--eval pre-pass style) is reflected in value") {
    val rewritten =
      sheet.put(ref"C1", CellValue.Formula("=A1*2", Some(CellValue.Number(BigDecimal(999)))))
    val out = JsonRenderer.renderRange(rewritten, row1)
    assert(
      out.contains(
        """{"ref": "C1", "type": "formula", "formula": "=A1*2", "value": 999, "formatted": "999"}"""
      ),
      s"cache rewrite not reflected:\n$out"
    )
  }

  test("rows: text-typed cached value renders as a JSON string") {
    val textCached =
      Sheet("Data").put(ref"A1", CellValue.Formula("=B1", Some(CellValue.Text("Revenue"))))
    val out = JsonRenderer.renderRange(textCached, CellRange(ref"A1", ref"A1"))
    assert(
      out.contains(
        """{"ref": "A1", "type": "formula", "formula": "=B1", "value": "Revenue", "formatted": "Revenue"}"""
      ),
      s"typed cached value not rendered:\n$out"
    )
  }

  test("rows: renderer-level evalFormulas computes values and keeps the formula field") {
    val out = JsonRenderer.renderRange(sheet, row1, evalFormulas = true)
    assert(
      out.contains(
        """{"ref": "D1", "type": "formula", "formula": "=A1+1", "value": 43, "formatted": "43"}"""
      ),
      s"evaluated uncached formula not rendered:\n$out"
    )
    assert(
      out.contains(
        """{"ref": "C1", "type": "formula", "formula": "=A1*2", "value": 84, "formatted": "84"}"""
      ),
      s"evaluated cached formula not rendered:\n$out"
    )
  }

  test("rows: evalFormulas failure renders the error token as value and formatted") {
    val errSheet = Sheet("Data").put(ref"A1", CellValue.Formula("=NOSUCHFN(1)"))
    val out =
      JsonRenderer.renderRange(errSheet, CellRange(ref"A1", ref"A1"), evalFormulas = true)
    assert(out.contains(""""type": "formula""""), s"formula type lost on error:\n$out")
    assert(out.contains(""""formula": "=NOSUCHFN(1)""""), s"formula field lost on error:\n$out")
    assert(out.contains(""""value": "#"""), s"error token missing from value:\n$out")
    assert(out.contains(""""formatted": "#"""), s"error token missing from formatted:\n$out")
  }

  test("rows: expression stored without leading = is normalized in the formula field") {
    val bare =
      Sheet("Data").put(ref"A1", CellValue.Formula("1+1", Some(CellValue.Number(BigDecimal(2)))))
    val out = JsonRenderer.renderRange(bare, CellRange(ref"A1", ref"A1"))
    assert(out.contains(""""formula": "=1+1""""), s"expression not normalized:\n$out")
  }

  test("rows: --formulas flag does not change JSON output") {
    val withFlag = JsonRenderer.renderRange(sheet, row1, showFormulas = true)
    val withoutFlag = JsonRenderer.renderRange(sheet, row1, showFormulas = false)
    assertEquals(withFlag, withoutFlag, "JSON must be identical regardless of --formulas")
  }

  // ========== rows path: non-formula cells unchanged ==========

  test("rows: non-formula cells keep the historical {ref, type, value, formatted} shape") {
    val out = JsonRenderer.renderRange(sheet, row1)
    assert(
      out.contains("""{"ref": "A1", "type": "number", "value": 42, "formatted": "42"}"""),
      s"number cell changed:\n$out"
    )
    assert(
      out.contains(
        """{"ref": "B1", "type": "text", "value": "Revenue", "formatted": "Revenue"}"""
      ),
      s"text cell changed:\n$out"
    )
    assert(
      out.contains("""{"ref": "E1", "type": "boolean", "value": true, "formatted": "TRUE"}"""),
      s"boolean cell changed:\n$out"
    )
    assert(
      out.contains("""{"ref": "F1", "type": "empty", "value": null, "formatted": ""}"""),
      s"empty cell changed:\n$out"
    )
    assertEquals(
      countOccurrences(out, "\"formula\":"),
      2,
      s"only the two formula cells may carry a formula field:\n$out"
    )
  }

  // ========== records path (--header-row) ==========

  private val recordsSheet = Sheet("Data")
    .put(ref"A1", CellValue.Text("Qty"))
    .put(ref"B1", CellValue.Text("Double"))
    .put(ref"C1", CellValue.Text("Plus"))
    .put(ref"A2", CellValue.Number(BigDecimal(42)))
    .put(ref"B2", CellValue.Formula("=A2*2", Some(CellValue.Number(BigDecimal(84)))))
    .put(ref"C2", CellValue.Formula("=A2+1")) // uncached

  private val recordsRange = CellRange(ref"A1", ref"C2")

  test("records: formula cells project the cached value, null when uncached") {
    val out = JsonRenderer.renderRange(recordsSheet, recordsRange, headerRow = Some(1))
    assert(
      out.contains("""{"Qty": 42, "Double": 84, "Plus": null}"""),
      s"records must carry cached values, not expressions:\n$out"
    )
  }

  test("records: evalFormulas computes values for uncached formulas") {
    val out = JsonRenderer.renderRange(
      recordsSheet,
      recordsRange,
      headerRow = Some(1),
      evalFormulas = true
    )
    assert(
      out.contains("""{"Qty": 42, "Double": 84, "Plus": 43}"""),
      s"records must carry evaluated values:\n$out"
    )
  }

  test("records: --formulas flag does not change JSON output") {
    val withFlag =
      JsonRenderer.renderRange(recordsSheet, recordsRange, showFormulas = true, headerRow = Some(1))
    val withoutFlag =
      JsonRenderer.renderRange(
        recordsSheet,
        recordsRange,
        showFormulas = false,
        headerRow = Some(1)
      )
    assertEquals(withFlag, withoutFlag, "records JSON must be identical regardless of --formulas")
  }

  // ========== end-to-end: view --format json / markdown ==========

  private val wb = Workbook(Vector(sheet))

  private def runView(
    range: String,
    format: ViewFormat,
    showFormulas: Boolean = false,
    evalFormulas: Boolean = false
  ): IO[String] =
    ReadCommands.view(
      wb,
      wb.sheets.headOption,
      range,
      showFormulas = showFormulas,
      evalFormulas = evalFormulas,
      strict = false,
      limit = 0,
      format = format,
      printScale = false,
      showGridlines = false,
      showLabels = false,
      dpi = 96,
      quality = 90,
      rasterOutput = None,
      skipEmpty = false,
      headerRow = None
    )

  test("view json: formula cells carry formula and cached value by default") {
    runView("A1:F1", ViewFormat.Json).map { out =>
      assert(
        out.contains(
          """{"ref": "C1", "type": "formula", "formula": "=A1*2", "value": 84, "formatted": "84"}"""
        ),
        s"view json missing formula cell shape:\n$out"
      )
    }
  }

  test("view json: --formulas yields byte-identical output") {
    for
      withFlag <- runView("A1:F1", ViewFormat.Json, showFormulas = true)
      withoutFlag <- runView("A1:F1", ViewFormat.Json)
    yield assertEquals(withFlag, withoutFlag, "view json must ignore --formulas")
  }

  test("view json: --eval surfaces computed values") {
    // ReadCommands pre-evaluates via evaluateSheetFormulas, which replaces Formula cells
    // with their computed plain values before the renderer runs — so under --eval the
    // formula field is absent today (type reflects the computed value). Pinned here to
    // document the boundary of the GH-357 renderer fix.
    runView("A1:F1", ViewFormat.Json, evalFormulas = true).map { out =>
      assert(
        out.contains(""""value": 43"""),
        s"--eval must surface the computed value of D1 (=A1+1):\n$out"
      )
      assert(
        out.contains(""""value": 84"""),
        s"--eval must surface the computed value of C1 (=A1*2):\n$out"
      )
    }
  }

  test("view markdown: --formulas still shows expressions (non-JSON formats unchanged)") {
    for
      withFlag <- runView("A1:D1", ViewFormat.Markdown, showFormulas = true)
      withoutFlag <- runView("A1:D1", ViewFormat.Markdown)
    yield
      assert(
        withFlag.contains("=A1*2"),
        s"markdown --formulas must show the expression:\n$withFlag"
      )
      assert(
        !withoutFlag.contains("=A1*2"),
        s"markdown without --formulas must show the cached value:\n$withoutFlag"
      )
      assert(withoutFlag.contains("84"), s"markdown must show the cached value:\n$withoutFlag")
  }
