package com.tjclp.xl.sheets

import com.tjclp.xl.{*, given}
import munit.FunSuite

/**
 * GH-420: `Sheet.named` is THE documented factory for dynamic (runtime-`String`) sheet names.
 *
 * `Sheet(name)` is `transparent inline`: a string literal validates at compile time and specializes
 * to `Sheet`, while a dynamic `String` silently specializes to `XLResult[Sheet]` — sound, but
 * nothing at the call site hints the return type changed (it cost a field build three edit cycles).
 * `Sheet.named` spells the `XLResult` in a plain non-inline signature so dynamic-name callers see
 * the `.map`/`.unsafe` step coming.
 */
class SheetNamedSpec extends FunSuite:

  test("named with a valid dynamic name yields Right, usable in a put chain") {
    // Deliberately a runtime value (the FinAgent field repro held the name in a val).
    val dynamic: String = List("Acquisitions").mkString
    val result: XLResult[Sheet] = Sheet.named(dynamic).map(_.put(ref"A1", 42).put(ref"B1", "ok"))
    result match
      case Right(sheet) =>
        assertEquals(sheet.name.value, "Acquisitions")
        assertEquals(sheet.cells.get(ref"A1").map(_.value), Some(CellValue.Number(BigDecimal(42))))
        assertEquals(sheet.cells.get(ref"B1").map(_.value), Some(CellValue.Text("ok")))
      case Left(err) => fail(s"expected Right, got Left($err)")
  }

  test("named with an invalid name yields Left(InvalidSheetName)") {
    val cases = List(
      "Q1:Q2", // invalid character ':'
      "", // empty
      "X" * 32, // over the 31-char maximum
      "History" // reserved by Excel
    )
    cases.foreach { bad =>
      Sheet.named(bad) match
        case Left(XLError.InvalidSheetName(name, _)) => assertEquals(name, bad)
        case other => fail(s"expected Left(InvalidSheetName) for '$bad', got $other")
    }
  }

  test("named agrees with the dynamic branch of the inline apply") {
    val names = List("Sales", "Q1:Q2", "", "X" * 32, "History", "Income Statement")
    names.foreach { s =>
      val dynamic: String = List(s).mkString // defeat constant folding: force the runtime branch
      assertEquals(Sheet.named(s), Sheet(dynamic): XLResult[Sheet])
    }
  }
