package com.tjclp.xl.macros

import com.tjclp.xl.*
import com.tjclp.xl.addressing.{ARef, CellRange, RefType, SheetName}
import munit.{FunSuite, ScalaCheckSuite}
import org.scalacheck.Prop.*

class RefCompileTimeOptimizationSpec extends ScalaCheckSuite:

  // ===== Compile-Time Optimization Tests (Phase 2) =====

  test("Optimized: ref\"$sheet!$cell\" with all literals") {
    inline val sheet = "Sales" // inline makes it compile-time constant
    inline val cell = "A1"
    val result = ref"$sheet!$cell"

    // Should return unwrapped RefType (compile-time optimized, NOT Either)
    result match
      case RefType.QualifiedCell(name, cellRef) =>
        assertEquals(name.value, "Sales")
        assertEquals(cellRef.toA1, "A1")
      case other => fail(s"Expected optimized QualifiedCell, got $other (should not be Either)")
  }

  test("Optimized: ref\"$col$row\" with number literals") {
    inline val col = "B"
    inline val row = 42
    val result = ref"$col$row"

    // Optimization works - type checks as ARef directly
    val aref: ARef = result
    assertEquals(aref.toA1, "B42")
  }

  test("Optimized: ref\"$start:$end\" range with literals") {
    inline val start = "A1"
    inline val end = "B10"
    val result = ref"$start:$end"

    // Optimization works - type checks as CellRange directly
    val range: CellRange = result
    assertEquals(range.toA1, "A1:B10")
  }

  test("Optimized: qualified range with all literals") {
    inline val sheet = "Sales"
    inline val start = "A1"
    inline val end = "B10"
    val result = ref"$sheet!$start:$end"

    result match
      case RefType.QualifiedRange(name, range) =>
        assertEquals(name.value, "Sales")
        assertEquals(range.toA1, "A1:B10")
      case other => fail(s"Expected optimized QualifiedRange, got $other")
  }

  test("Optimized: quoted sheet name literal") {
    inline val quotedSheet = "'Q1 Sales'" // Literal with quotes
    val result = ref"$quotedSheet!A1"

    result match
      case RefType.QualifiedCell(name, _) =>
        assertEquals(name.value, "Q1 Sales")
      case other => fail(s"Expected optimized QualifiedCell, got $other")
  }

  test("Optimized: escaped quotes in sheet name") {
    inline val escaped = "'It''s Q1'" // Excel escaping: '' → '
    val result = ref"$escaped!A1"

    result match
      case RefType.QualifiedCell(name, _) =>
        assertEquals(name.value, "It's Q1")
      case other => fail(s"Expected optimized with escaped quotes, got $other")
  }

  test("Optimized: multiple interpolations in complex ref") {
    inline val sheet = "Q1"
    inline val col = "B"
    inline val row = 42
    val result = ref"$sheet!$col$row"

    result match
      case RefType.QualifiedCell(name, cellRef) =>
        assertEquals(name.value, "Q1")
        assertEquals(cellRef.toA1, "B42")
      case other => fail(s"Expected optimized QualifiedCell, got $other")
  }

  test("Optimized: range with numeric row literals") {
    inline val startCol = "A"
    inline val startRow = 1
    inline val endCol = "B"
    inline val endRow = 10
    val result = ref"$startCol$startRow:$endCol$endRow"

    // Optimization works
    val range: CellRange = result
    assertEquals(range.toA1, "A1:B10")
  }

  // ===== Runtime Path Preserved (Phase 1 Behavior) =====

  test("Runtime: function call uses runtime path") {
    def getDynamicSheet(): String = "Sales" // Simulates runtime value
    val result = ref"${getDynamicSheet()}!A1"

    // Should return Either (runtime path)
    result match
      case Right(RefType.QualifiedCell(name, _)) =>
        assertEquals(name.value, "Sales")
      case Left(err) => fail(s"Should parse valid ref: $err")
      case other => fail(s"Expected Either (runtime path), got $other")
  }

  test("Runtime: mixed literal + function call uses runtime path") {
    inline val sheet = "Sales" // Literal
    def getDynamicCell(): String = "A1" // Runtime
    val result = ref"$sheet!${getDynamicCell()}"

    // Should use runtime path (any runtime var forces runtime)
    result match
      case Right(_) => () // Expected Either
      case Left(err) => fail(s"Should parse: $err")
  }

  // ===== Property Tests (Identity Law) =====

  test("SUCCESS: Compile-time optimization IS working") {
    // Direct string literals in interpolation are optimized!
    inline val col = "A"
    inline val row = 1
    val result = ref"$col$row"

    // Optimization works - returns unwrapped ARef (not Either)
    val aref: ARef = result // Type checks! Optimization confirmed!
    assertEquals(aref.toA1, "A1")
  }

  test("Round-trip: various literal interpolations preserve values") {
    // Test cell
    inline val col1 = "A"
    inline val row1 = 1
    assertEquals(ref"$col1$row1", ref"A1")

    // Test range
    inline val start = "A1"
    inline val end = "B10"
    assertEquals(ref"$start:$end", ref"A1:B10")

    // Test qualified cell
    inline val sheet1 = "Sales"
    inline val cell1 = "A1"
    assertEquals(ref"$sheet1!$cell1", ref"Sales!A1")

    // Test qualified range
    inline val sheet2 = "Q1"
    inline val range1 = "A1:B10"
    assertEquals(ref"$sheet2!$range1", ref"Q1!A1:B10")
  }

  // ===== Edge Cases =====

  test("Edge: Maximum valid column literal (XFD)") {
    inline val col = "XFD"
    inline val row = 1
    val aref: ARef = ref"$col$row" // Optimized
    assertEquals(aref.col.index0, 16383)
  }

  test("Edge: Maximum valid row literal (1048576)") {
    inline val col = "A"
    inline val row = 1048576
    val aref: ARef = ref"$col$row" // Optimized
    assertEquals(aref.row.index0, 1048575)
  }

  test("Edge: Sheet name with spaces and parens (literal)") {
    inline val sheet = "'Q1 (2025) Sales'"
    val result = ref"$sheet!A1"

    result match
      case RefType.QualifiedCell(name, _) =>
        assertEquals(name.value, "Q1 (2025) Sales")
      case other => fail(s"Expected QualifiedCell, got $other")
  }

  test("Edge: Empty parts with literals") {
    inline val a = "A"
    inline val one = "1"
    val aref: ARef = ref"$a$one" // Optimized
    assertEquals(aref.toA1, "A1")
  }

  test("Edge: Prefix + literal") {
    inline val cell = "A1"
    val result = ref"Sales!$cell"

    result match
      case RefType.QualifiedCell(name, cellRef) =>
        assertEquals(name.value, "Sales")
        assertEquals(cellRef.toA1, "A1")
      case other => fail(s"Expected QualifiedCell, got $other")
  }

  test("Edge: Literal + suffix") {
    inline val sheet = "Sales"
    val result = ref"$sheet!A1"

    result match
      case RefType.QualifiedCell(name, cellRef) =>
        assertEquals(name.value, "Sales")
        assertEquals(cellRef.toA1, "A1")
      case other => fail(s"Expected QualifiedCell, got $other")
  }

  // Note: Compile-time error tests are commented because they would fail compilation.
  // To manually verify, uncomment and check that compilation fails with expected error.

  /*
  test("Compile error: invalid literal in interpolation") {
    val invalid = "NOT_VALID!@#$"
    val result = ref"$invalid"  // Should NOT compile
    // Expected: "Invalid ref literal in interpolation: 'NOT_VALID!@#$'"
  }

  test("Compile error: empty sheet name") {
    val empty = ""
    val result = ref"$empty!A1"  // Should NOT compile
    // Expected: "Empty sheet name"
  }

  test("Compile error: missing ref after bang") {
    val sheet = "Sales"
    val result = ref"$sheet!"  // Should NOT compile
    // Expected: "Missing reference after '!'"
  }
  */
