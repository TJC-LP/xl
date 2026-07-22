package com.tjclp.xl.sheets

import com.tjclp.xl.{*, given}
import com.tjclp.xl.addressing.CellRange
import munit.FunSuite

/**
 * GH-375: typed data-validation model — factories, Sheet authoring API, and structural shifts (the
 * conditionalFormats clamp/split/drop algebra: without it, dropdowns detach from their cells on
 * row/column insert or delete).
 */
class DataValidationSpec extends FunSuite:

  // ===== factories =====

  test("list stores the formula verbatim (inline quoted list or range reference)") {
    assertEquals(
      DataValidation.list("\"Low,Med,High\""),
      DataValidation.Rules(Vector.empty, DvKind.List("\"Low,Med,High\""))
    )
    assertEquals(
      DataValidation.list("$Z$1:$Z$3", allowBlank = false, showDropdown = false),
      DataValidation.Rules(
        Vector.empty,
        DvKind.List("$Z$1:$Z$3"),
        allowBlank = false,
        showDropdown = false
      )
    )
  }

  test("listOf quotes inline values and doubles embedded quotes") {
    assertEquals(
      DataValidation.listOf("Yes", "No").kind,
      DvKind.List("\"Yes,No\"")
    )
    assertEquals(
      DataValidation.listOf("say \"hi\"").kind,
      DvKind.List("\"say \"\"hi\"\"\"")
    )
  }

  test("GH-429: bounded/custom/anyValue factories and the message combinators") {
    assertEquals(
      DataValidation.whole(DvOperator.Between, "1", Some("9")).kind,
      DvKind.Bounded(DvBoundedType.Whole, DvOperator.Between, "1", Some("9"))
    )
    assertEquals(
      DataValidation.textLength(DvOperator.LessThanOrEqual, "10").kind,
      DvKind.Bounded(DvBoundedType.TextLength, DvOperator.LessThanOrEqual, "10", None)
    )
    assertEquals(DataValidation.custom("ISNUMBER(A1)").kind, DvKind.Custom("ISNUMBER(A1)"))
    assertEquals(DataValidation.anyValue.kind, DvKind.AnyValue)
    // attach-implies-show (Excel-UI parity)
    val prompted = DataValidation.listOf("a", "b").withPrompt("Pick", "Choose one")
    assertEquals(prompted.messages.showInputMessage, true)
    assertEquals(prompted.messages.promptTitle, Some("Pick"))
    assertEquals(prompted.messages.prompt, Some("Choose one"))
    val errored = DataValidation.anyValue.withError("No", "Try again", DvErrorStyle.Warning)
    assertEquals(errored.messages.showErrorMessage, true)
    assertEquals(errored.messages.errorStyle, DvErrorStyle.Warning)
    assertEquals(errored.messages.errorTitle, Some("No"))
    assertEquals(errored.messages.error, Some("Try again"))
    // defaults equal the OOXML schema defaults
    assertEquals(DataValidation.listOf("x").messages, DvMessages.default)
  }

  // ===== Sheet authoring =====

  test("withDataValidation appends one entry with the given range") {
    val s = Sheet("DV")
      .withDataValidation(ref"B2:B10", DataValidation.listOf("1", "2", "3"))
      .withDataValidation(ref"C2:C10", DataValidation.list("$Z$1:$Z$3"))
    assertEquals(s.dataValidations.size, 2)
    assertEquals(
      s.dataValidations(0),
      DataValidation.Rules(Vector(ref"B2:B10": CellRange), DvKind.List("\"1,2,3\""))
    )
    assertEquals(
      s.dataValidations(1),
      DataValidation.Rules(Vector(ref"C2:C10": CellRange), DvKind.List("$Z$1:$Z$3"))
    )
  }

  test("withDataValidation multi-range variant keeps one entry; empty ranges are identity") {
    val s = Sheet("DV").withDataValidation(
      Vector(ref"A1:A5": CellRange, ref"C1:C5": CellRange),
      DataValidation.listOf("a", "b")
    )
    assertEquals(s.dataValidations.size, 1)
    assertEquals(s.typedDataValidations.map(_.ranges.map(_.toA1)), Vector(Vector("A1:A5", "C1:C5")))
    assertEquals(
      Sheet("DV").withDataValidation(Vector.empty, DataValidation.listOf("a")),
      Sheet("DV")
    )
  }

  test("removeDataValidation mirrors removeConditionalFormat (identity out of range)") {
    val s = Sheet("DV")
      .withDataValidation(ref"A1:A5", DataValidation.listOf("x"))
      .withDataValidation(ref"B1:B5", DataValidation.listOf("y"))
    assertEquals(s.removeDataValidation(0).dataValidations.size, 1)
    assertEquals(
      s.removeDataValidation(0).dataValidations(0),
      s.dataValidations(1)
    )
    assertEquals(s.removeDataValidation(7), s)
    assertEquals(s.removeDataValidation(-1), s)
  }

  test("typedDataValidations excludes Preserved entries") {
    val s = Sheet("DV")
      .copy(dataValidations = Vector(DataValidation.Preserved("<dataValidation/>")))
      .withDataValidation(ref"A1:A2", DataValidation.listOf("x"))
    assertEquals(s.dataValidations.size, 2)
    assertEquals(s.typedDataValidations.size, 1)
  }

  // ===== structural shifts (the mergedRanges/cf clamp/split/drop algebra) =====

  private def dvRanges(s: Sheet): Vector[Vector[String]] =
    s.typedDataValidations.map(_.ranges.map(_.toA1))

  test("insertRows shifts validation ranges below the cut (dropdowns stay attached)") {
    val s = Sheet("DV").withDataValidation(ref"B2:B4", DataValidation.listOf("1", "2"))
    assertEquals(dvRanges(s.insertRows(at = 0, count = 2)), Vector(Vector("B4:B6")))
  }

  test("deleteRows clamps a validation range spanning the cut") {
    val s = Sheet("DV").withDataValidation(ref"A1:A4", DataValidation.listOf("x"))
    assertEquals(dvRanges(s.deleteRows(at = 1, count = 2)), Vector(Vector("A1:A2")))
  }

  test("deleteColumns drops a fully-deleted range and removes an emptied entry") {
    val s = Sheet("DV")
      .withDataValidation(ref"B1:B5", DataValidation.listOf("x"))
      .withDataValidation(
        Vector(ref"B2:B3": CellRange, ref"D2:D3": CellRange),
        DataValidation.listOf("y")
      )
    val r = s.deleteColumns(at = 1, count = 1) // delete column B
    // first entry: only range fully deleted -> whole entry removed
    // second entry: B2:B3 dropped, D2:D3 -> C2:C3
    assertEquals(dvRanges(r), Vector(Vector("C2:C3")))
  }

  test("GH-428: insertColumns clamps a full-width validation at XFD (never XFE)") {
    val s = Sheet("DV").withDataValidation(ref"A1:XFD1", DataValidation.listOf("x"))
    assertEquals(dvRanges(s.insertColumns(at = 1, count = 2)), Vector(Vector("A1:XFD1")))
  }

  test("GH-429: structural shifts move Preserved entries' sqref (textual envelope rewrite)") {
    val preserved = DataValidation.Preserved("""<dataValidation type="whole" sqref="Z9"/>""")
    val s = Sheet("DV")
      .copy(dataValidations = Vector(preserved))
      .withDataValidation(ref"A2:A4", DataValidation.listOf("x"))
    val r = s.insertRows(at = 0, count = 1)
    assertEquals(
      r.dataValidations(0),
      DataValidation.Preserved("""<dataValidation type="whole" sqref="Z10"/>""")
    )
    assertEquals(dvRanges(r), Vector(Vector("A3:A5")))
  }
