package com.tjclp.xl.cli

import java.nio.file.{Files, Path}

import cats.effect.IO
import munit.CatsEffectSuite

import com.tjclp.xl.{*, given}
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.cli.commands.ReadCommands
import com.tjclp.xl.io.ExcelIO
import com.tjclp.xl.macros.ref

/**
 * Tests for eval command, particularly --with override propagation (TJC-698).
 *
 * The eval command with --with flag should propagate overrides through formula chains. For
 * example, if A1=100, B1=A1*2, C1=B1+50, then `eval "=C1" --with "A1=200"` should return 450 (not
 * 250).
 */
@SuppressWarnings(
  Array(
    "org.wartremover.warts.OptionPartial"
  )
)
class EvalCommandSpec extends CatsEffectSuite:

  // Create a test workbook with formula chain: A1=100, B1=A1*2, C1=B1+50
  private def createFormulaChainWorkbook: IO[Workbook] =
    IO {
      val sheet = Sheet("Test")
        .put(ref"A1", CellValue.Number(BigDecimal(100)))
        .put(ref"B1", CellValue.Formula("=A1*2"))
        .put(ref"C1", CellValue.Formula("=B1+50"))

      Workbook(Vector(sheet))
    }

  test("eval: formula without overrides returns correct value") {
    for
      wb <- createFormulaChainWorkbook
      sheet = wb.sheets.head
      result <- ReadCommands.eval(wb, Some(sheet), "=C1", Nil)
    yield
      // C1 = B1+50 = (A1*2)+50 = 100*2+50 = 250
      assert(result.contains("250"), s"Expected 250, got: $result")
  }

  test("eval: override propagates through single formula (TJC-698)") {
    for
      wb <- createFormulaChainWorkbook
      sheet = wb.sheets.head
      result <- ReadCommands.eval(wb, Some(sheet), "=B1", List("A1=200"))
    yield
      // B1 = A1*2 = 200*2 = 400
      assert(result.contains("400"), s"Expected 400 (A1=200 → B1=400), got: $result")
  }

  test("eval: override propagates through formula chain (TJC-698)") {
    for
      wb <- createFormulaChainWorkbook
      sheet = wb.sheets.head
      result <- ReadCommands.eval(wb, Some(sheet), "=C1", List("A1=200"))
    yield
      // C1 = B1+50 = (A1*2)+50 = 200*2+50 = 450
      // Bug before fix: returned 250 (B1 wasn't recalculated with new A1 value)
      assert(result.contains("450"), s"Expected 450 (A1=200 → B1=400 → C1=450), got: $result")
  }

  test("eval: multiple overrides propagate correctly") {
    for
      wb <- IO {
        val sheet = Sheet("Test")
          .put(ref"A1", CellValue.Number(BigDecimal(10)))
          .put(ref"A2", CellValue.Number(BigDecimal(20)))
          .put(ref"B1", CellValue.Formula("=A1+A2"))

        Workbook(Vector(sheet))
      }
      sheet = wb.sheets.head
      result <- ReadCommands.eval(wb, Some(sheet), "=B1", List("A1=100", "A2=200"))
    yield
      // B1 = A1+A2 = 100+200 = 300
      assert(result.contains("300"), s"Expected 300, got: $result")
  }

  test("eval: override with no dependent formulas works") {
    for
      wb <- createFormulaChainWorkbook
      sheet = wb.sheets.head
      result <- ReadCommands.eval(wb, Some(sheet), "=A1*10", List("A1=50"))
    yield
      // Direct formula using overridden value: 50*10 = 500
      assert(result.contains("500"), s"Expected 500, got: $result")
  }

  test("eval: constant formula without overrides doesn't need file") {
    for result <- ReadCommands.eval(Workbook.empty, None, "=1+1", Nil)
    yield assert(result.contains("2"), s"Expected 2, got: $result")
  }
