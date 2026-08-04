package com.tjclp.xl.formula

import com.tjclp.xl.{*, given}
import com.tjclp.xl.addressing.SheetName
import com.tjclp.xl.formula.eval.StructuralEditor
import com.tjclp.xl.sheets.{DataValidation, DvKind, DvOperator, Sheet}
import com.tjclp.xl.workbooks.Workbook
import munit.FunSuite

/**
 * GH-429: structural edits rewrite TYPED data-validation formulas through the formula engine,
 * mirroring the conditional-format behavior — absolute list sources track the edit, inline list
 * literals ride verbatim, fully-deleted sources degrade the formula text to "#REF!" (entry kept),
 * and Preserved payload formulas are never rewritten (their sqref envelope is shifted by the pure
 * core layer instead).
 */
@SuppressWarnings(Array("org.wartremover.warts.OptionPartial"))
class StructuralDvSpec extends FunSuite:

  private val S = SheetName.unsafe("S")

  private def sheetNamed(wb: Workbook, n: String): Sheet =
    wb.sheets.find(_.name == SheetName.unsafe(n)).get

  private def dvKinds(s: Sheet): Vector[DvKind] =
    s.typedDataValidations.map(_.kind)

  test("an absolute list source shifts with the rows it points at") {
    val s = new Sheet(name = S)
      .withDataValidation(ref"B2:B10", DataValidation.list("$Z$5:$Z$7"))
    val r = StructuralEditor.insertRows(Workbook(Vector(s)), S, at = 0, count = 2)
    val s2 = sheetNamed(r, "S")
    assertEquals(dvKinds(s2), Vector(DvKind.List("$Z$7:$Z$9")))
    // the sqref envelope moved too (core layer)
    assertEquals(s2.typedDataValidations.flatMap(_.ranges.map(_.toA1)), Vector("B4:B12"))
  }

  test("an inline list literal rides verbatim (string literals are data, not refs)") {
    val s = new Sheet(name = S)
      .withDataValidation(ref"H5:H10", DataValidation.list("\"yes,no\""))
    val r = StructuralEditor.insertRows(Workbook(Vector(s)), S, at = 1, count = 2)
    assertEquals(dvKinds(sheetNamed(r, "S")), Vector(DvKind.List("\"yes,no\"")))
    assertEquals(
      sheetNamed(r, "S").typedDataValidations.flatMap(_.ranges.map(_.toA1)),
      Vector("H7:H12")
    )
  }

  test("a fully-deleted list source degrades the formula text to #REF! (entry kept)") {
    val s = new Sheet(name = S)
      .withDataValidation(ref"B20:B30", DataValidation.list("$Z$5:$Z$7"))
    val r = StructuralEditor.deleteRows(Workbook(Vector(s)), S, at = 4, count = 3) // rows 5-7
    assertEquals(dvKinds(sheetNamed(r, "S")), Vector(DvKind.List("#REF!")))
  }

  test("a cross-sheet list source shifts when the SOURCE sheet is edited") {
    val lists = new Sheet(name = SheetName.unsafe("Lists"))
    val data = new Sheet(name = SheetName.unsafe("Data"))
      .withDataValidation(ref"C2:C9", DataValidation.list("Lists!$A$5:$A$8"))
    val r = StructuralEditor
      .insertRows(Workbook(Vector(lists, data)), SheetName.unsafe("Lists"), at = 0, count = 3)
    assertEquals(dvKinds(sheetNamed(r, "Data")), Vector(DvKind.List("Lists!$A$8:$A$11")))
    // the envelope on Data did NOT move (the edit was on Lists)
    assertEquals(
      sheetNamed(r, "Data").typedDataValidations.flatMap(_.ranges.map(_.toA1)),
      Vector("C2:C9")
    )
  }

  test("bounded formulas shift; custom formulas shift; Preserved payload formulas never do") {
    val base = new Sheet(name = S)
      .withDataValidation(
        ref"D1:D9",
        DataValidation.whole(DvOperator.Between, "$F$5", Some("$F$6"))
      )
      .withDataValidation(ref"E1:E9", DataValidation.custom("E1<$F$5"))
    val s = base.copy(dataValidations =
      base.dataValidations :+ DataValidation.Preserved(
        """<dataValidation type="custom" sqref="G1"><formula1>$F$5</formula1></dataValidation>"""
      )
    )
    val r = StructuralEditor.insertRows(Workbook(Vector(s)), S, at = 0, count = 1)
    val s2 = sheetNamed(r, "S")
    assertEquals(
      dvKinds(s2),
      Vector(
        DvKind.Bounded(
          com.tjclp.xl.sheets.DvBoundedType.Whole,
          DvOperator.Between,
          "$F$6",
          Some("$F$7")
        ),
        DvKind.Custom("E2<$F$6")
      )
    )
    // Preserved: envelope shifted textually by the core layer, formula text untouched
    assertEquals(
      s2.dataValidations.collect { case DataValidation.Preserved(x) => x },
      Vector(
        """<dataValidation type="custom" sqref="G2"><formula1>$F$5</formula1></dataValidation>"""
      )
    )
  }

  test("GH-455: a DV formula on a non-participating sheet rides byte-identical") {
    val data = new Sheet(name = SheetName.unsafe("Data"))
    val report = new Sheet(name = SheetName.unsafe("Report"))
      // Non-canonical spelling (lowercase call, extra spaces): a reprint would rewrite it.
      .withDataValidation(ref"C2:C9", DataValidation.custom("C2 < sum($F$5:$F$9)"))
    val r = StructuralEditor
      .insertRows(Workbook(Vector(data, report)), SheetName.unsafe("Data"), at = 0, count = 3)
    assertEquals(dvKinds(sheetNamed(r, "Report")), Vector(DvKind.Custom("C2 < sum($F$5:$F$9)")))
  }
