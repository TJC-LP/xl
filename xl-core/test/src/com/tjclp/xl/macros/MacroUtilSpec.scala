package com.tjclp.xl.macros

import munit.FunSuite

class MacroUtilSpec extends FunSuite:

  // ===== String Reconstruction Tests =====

  test("reconstructString: empty literals (no interpolation)") {
    val parts = Seq("A1")
    val literals = Seq.empty
    assertEquals(MacroUtil.reconstructString(parts, literals), "A1")
  }

  test("reconstructString: single literal") {
    val parts = Seq("", "!A1")
    val literals = Seq("Sales")
    assertEquals(MacroUtil.reconstructString(parts, literals), "Sales!A1")
  }

  test("reconstructString: multiple literals") {
    val parts = Seq("", "!", "")
    val literals = Seq("Sales", "A1")
    assertEquals(MacroUtil.reconstructString(parts, literals), "Sales!A1")
  }

  test("reconstructString: literals with numbers") {
    val parts = Seq("", "", "")
    val literals = Seq("B", 42)
    assertEquals(MacroUtil.reconstructString(parts, literals), "B42")
  }

  test("reconstructString: complex interpolation") {
    val parts = Seq("", "!", ":", "")
    val literals = Seq("Sheet1", "A1", "B10")
    assertEquals(MacroUtil.reconstructString(parts, literals), "Sheet1!A1:B10")
  }

  test("reconstructString: with prefix") {
    val parts = Seq("Sales!", "")
    val literals = Seq("A1")
    assertEquals(MacroUtil.reconstructString(parts, literals), "Sales!A1")
  }

  test("reconstructString: with suffix") {
    val parts = Seq("", ":B10")
    val literals = Seq("A1")
    assertEquals(MacroUtil.reconstructString(parts, literals), "A1:B10")
  }

  test("reconstructString: all empty parts") {
    val parts = Seq("", "")
    val literals = Seq("A1")
    assertEquals(MacroUtil.reconstructString(parts, literals), "A1")
  }

  test("reconstructString: invariant violation fails") {
    val parts = Seq("A", "B")
    val literals = Seq("1", "2") // parts.length != literals.length + 1
    intercept[IllegalArgumentException] {
      MacroUtil.reconstructString(parts, literals)
    }
  }

  // ===== Error Formatting Tests =====

  test("formatCompileError: includes all components") {
    val err = MacroUtil.formatCompileError("ref", "INVALID!@#$", "Invalid characters in sheets name")
    assert(err.contains("Invalid ref literal"))
    assert(err.contains("INVALID!@#$"))
    assert(err.contains("Invalid characters"))
    assert(err.contains("Hint"))
  }

  test("formatCompileError: includes macro name") {
    val err = MacroUtil.formatCompileError("money", "$ABC", "non-numeric")
    assert(err.contains("Invalid money literal"))
    assert(err.contains("$ABC"))
  }
