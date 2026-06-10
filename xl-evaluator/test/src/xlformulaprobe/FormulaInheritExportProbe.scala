// Deliberately OUTSIDE com.tjclp.xl: export forwarders for extension methods can break at
// external use sites while everything inside the package compiles (see the landmine notes in
// exports.scala). This probe pins GH-184's opt-in surface as consumers see it.
package xlformulaprobe

import munit.FunSuite

import com.tjclp.xl.{*, given}

class FormulaInheritExportProbe extends FunSuite:

  private val currency = CellStyle.default.withNumFmt(NumFmt.Currency)

  test("putFormulaInheriting resolves through com.tjclp.xl exports (sheet overload)") {
    val sheet = Sheet("Probe").put(ref"B2", 100, currency)
    val result: XLResult[Sheet] = sheet.putFormulaInheriting(ref"B4", "=B2*2")
    result match
      case Right(updated) =>
        val fmt = updated.cells.get(ref"B4").flatMap(_.styleId).flatMap(updated.styleRegistry.get)
        assertEquals(fmt.map(_.numFmt), Some(NumFmt.Currency))
      case Left(err) => fail(s"putFormulaInheriting failed: $err")
  }

  test("putFormulaInheriting workbook overload resolves through com.tjclp.xl exports") {
    val data = Sheet("Data").put(ref"B2", 250, currency)
    val summary = Sheet("Summary")
    val wb = Workbook(data, summary)
    val result: XLResult[Sheet] = summary.putFormulaInheriting(ref"A1", "=Data!B2*2", wb)
    result match
      case Right(updated) =>
        val fmt = updated.cells.get(ref"A1").flatMap(_.styleId).flatMap(updated.styleRegistry.get)
        assertEquals(fmt.map(_.numFmt), Some(NumFmt.Currency))
      case Left(err) => fail(s"putFormulaInheriting failed: $err")
  }

  test("FormulaFormatting object (default-param method) resolves through com.tjclp.xl exports") {
    val sheet = Sheet("Probe").put(ref"B2", 100, currency)
    FormulaParser.parse("=B2*2") match
      case Right(expr) =>
        // Default `workbook` param omitted: plain object methods tolerate defaults through exports
        assertEquals(
          FormulaFormatting.inferFormatFromReferences(expr, sheet),
          Some(NumFmt.Currency)
        )
      case Left(err) => fail(s"parse failed: $err")
  }
