package com.tjclp.xl.workbooks

import com.tjclp.xl.Workbook
import com.tjclp.xl.addressing.SheetName
import com.tjclp.xl.error.XLError
import com.tjclp.xl.sheets.Sheet
import munit.FunSuite

class WorkbookSheetStateSpec extends FunSuite:

  private val sheet1 = Sheet("Sheet1")
  private val sheet2 = Sheet("Sheet2")
  private val sheet3 = Sheet("Sheet3")
  private val workbook = Workbook(Vector(sheet1, sheet2, sheet3))

  private def unwrap[A](result: Either[XLError, A]): A =
    result.fold(err => fail(s"Expected Right but got: $err"), identity)

  // ========== State Validation ==========

  test("setSheetState accepts hidden state") {
    val wb = unwrap(workbook.setSheetState(SheetName.unsafe("Sheet2"), Some("hidden")))
    assertEquals(wb.getSheetState(SheetName.unsafe("Sheet2")), Some("hidden"))
  }

  test("setSheetState accepts veryHidden state") {
    val wb = unwrap(workbook.setSheetState(SheetName.unsafe("Sheet2"), Some("veryHidden")))
    assertEquals(wb.getSheetState(SheetName.unsafe("Sheet2")), Some("veryHidden"))
  }

  test("setSheetState accepts None (visible) state") {
    // First hide, then show
    val hidden = unwrap(workbook.setSheetState(SheetName.unsafe("Sheet2"), Some("hidden")))
    val wb = unwrap(hidden.setSheetState(SheetName.unsafe("Sheet2"), None))
    assertEquals(wb.getSheetState(SheetName.unsafe("Sheet2")), None)
  }

  test("setSheetState rejects invalid state") {
    val result = workbook.setSheetState(SheetName.unsafe("Sheet2"), Some("invalid"))
    assert(result.isLeft, "Expected Left for invalid state")
    result match
      case Left(XLError.InvalidWorkbook(msg)) =>
        assert(msg.contains("invalid"), s"Error message should mention invalid state: $msg")
        assert(
          msg.contains("hidden") || msg.contains("veryHidden"),
          s"Error message should mention valid options: $msg"
        )
      case other =>
        fail(s"Expected InvalidWorkbook error but got: $other")
  }

  test("setSheetState rejects empty string state") {
    val result = workbook.setSheetState(SheetName.unsafe("Sheet2"), Some(""))
    assert(result.isLeft, "Expected Left for empty string state")
    result match
      case Left(XLError.InvalidWorkbook(_)) => () // expected
      case other => fail(s"Expected InvalidWorkbook error but got: $other")
  }

  test("setSheetState rejects misspelled state") {
    val result = workbook.setSheetState(SheetName.unsafe("Sheet2"), Some("Hidden"))
    assert(result.isLeft, "Expected Left for case-sensitive misspelling")
    result match
      case Left(XLError.InvalidWorkbook(_)) => () // expected
      case other => fail(s"Expected InvalidWorkbook error but got: $other")
  }

  // ========== Last Visible Sheet Protection ==========

  test("cannot hide the last visible sheet") {
    // Hide all but one
    val wb1 = unwrap(workbook.setSheetState(SheetName.unsafe("Sheet1"), Some("hidden")))
    val wb2 = unwrap(wb1.setSheetState(SheetName.unsafe("Sheet2"), Some("hidden")))

    // Try to hide the last one
    val result = wb2.setSheetState(SheetName.unsafe("Sheet3"), Some("hidden"))
    assert(result.isLeft, "Expected Left when trying to hide last visible sheet")
    result match
      case Left(XLError.InvalidWorkbook(msg)) =>
        assert(msg.contains("last visible"), s"Error should mention last visible: $msg")
      case other =>
        fail(s"Expected InvalidWorkbook error but got: $other")
  }

  test("cannot veryHide the last visible sheet") {
    // Hide all but one
    val wb1 = unwrap(workbook.setSheetState(SheetName.unsafe("Sheet1"), Some("veryHidden")))
    val wb2 = unwrap(wb1.setSheetState(SheetName.unsafe("Sheet2"), Some("veryHidden")))

    // Try to veryHide the last one
    val result = wb2.setSheetState(SheetName.unsafe("Sheet3"), Some("veryHidden"))
    assert(result.isLeft, "Expected Left when trying to veryHide last visible sheet")
    result match
      case Left(XLError.InvalidWorkbook(msg)) =>
        assert(msg.contains("last visible"), s"Error should mention last visible: $msg")
      case other =>
        fail(s"Expected InvalidWorkbook error but got: $other")
  }

  // ========== Unhiding ==========

  test("can unhide hidden sheet") {
    val hidden = unwrap(workbook.setSheetState(SheetName.unsafe("Sheet2"), Some("hidden")))
    val wb = unwrap(hidden.setSheetState(SheetName.unsafe("Sheet2"), None))
    assertEquals(wb.getSheetState(SheetName.unsafe("Sheet2")), None)
  }

  test("can unhide veryHidden sheet") {
    val hidden = unwrap(workbook.setSheetState(SheetName.unsafe("Sheet2"), Some("veryHidden")))
    val wb = unwrap(hidden.setSheetState(SheetName.unsafe("Sheet2"), None))
    assertEquals(wb.getSheetState(SheetName.unsafe("Sheet2")), None)
  }

  test("unhiding already visible sheet is a no-op") {
    val wb = unwrap(workbook.setSheetState(SheetName.unsafe("Sheet2"), None))
    assertEquals(wb.getSheetState(SheetName.unsafe("Sheet2")), None)
  }

  // ========== Error Cases ==========

  test("setSheetState returns error for non-existent sheet") {
    val result = workbook.setSheetState(SheetName.unsafe("NonExistent"), Some("hidden"))
    assert(result.isLeft, "Expected Left for non-existent sheet")
    result match
      case Left(XLError.SheetNotFound(name)) =>
        assertEquals(name, "NonExistent")
      case other =>
        fail(s"Expected SheetNotFound error but got: $other")
  }

  // ========== State Transitions ==========

  test("can change from hidden to veryHidden") {
    val hidden = unwrap(workbook.setSheetState(SheetName.unsafe("Sheet2"), Some("hidden")))
    val wb = unwrap(hidden.setSheetState(SheetName.unsafe("Sheet2"), Some("veryHidden")))
    assertEquals(wb.getSheetState(SheetName.unsafe("Sheet2")), Some("veryHidden"))
  }

  test("can change from veryHidden to hidden") {
    val veryHidden = unwrap(workbook.setSheetState(SheetName.unsafe("Sheet2"), Some("veryHidden")))
    val wb = unwrap(veryHidden.setSheetState(SheetName.unsafe("Sheet2"), Some("hidden")))
    assertEquals(wb.getSheetState(SheetName.unsafe("Sheet2")), Some("hidden"))
  }

  test("hiding already hidden sheet is idempotent") {
    val hidden = unwrap(workbook.setSheetState(SheetName.unsafe("Sheet2"), Some("hidden")))
    val wb = unwrap(hidden.setSheetState(SheetName.unsafe("Sheet2"), Some("hidden")))
    assertEquals(wb.getSheetState(SheetName.unsafe("Sheet2")), Some("hidden"))
  }
