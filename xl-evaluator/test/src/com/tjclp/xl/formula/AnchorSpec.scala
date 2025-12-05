package com.tjclp.xl.formula

import munit.FunSuite
import com.tjclp.xl.{ARef, Anchor, CellRange}

/**
 * Tests for Excel-style formula anchoring ($) support.
 *
 * Anchor modes:
 *   - Relative (A1): Both column and row adjust during drag
 *   - AbsCol ($A1): Column fixed, row adjusts
 *   - AbsRow (A$1): Column adjusts, row fixed
 *   - Absolute ($A$1): Both fixed
 */
class AnchorSpec extends FunSuite:

  // ===== Anchor.parse tests =====

  test("Anchor.parse: relative reference A1"):
    val (clean, anchor) = Anchor.parse("A1")
    assertEquals(clean, "A1")
    assertEquals(anchor, Anchor.Relative)

  test("Anchor.parse: absolute column $A1"):
    val (clean, anchor) = Anchor.parse("$A1")
    assertEquals(clean, "A1")
    assertEquals(anchor, Anchor.AbsCol)

  test("Anchor.parse: absolute row A$1"):
    val (clean, anchor) = Anchor.parse("A$1")
    assertEquals(clean, "A1")
    assertEquals(anchor, Anchor.AbsRow)

  test("Anchor.parse: absolute both $A$1"):
    val (clean, anchor) = Anchor.parse("$A$1")
    assertEquals(clean, "A1")
    assertEquals(anchor, Anchor.Absolute)

  test("Anchor.parse: multi-letter column $AB$10"):
    val (clean, anchor) = Anchor.parse("$AB$10")
    assertEquals(clean, "AB10")
    assertEquals(anchor, Anchor.Absolute)

  test("Anchor.parse: mixed case handled"):
    val (clean, anchor) = Anchor.parse("$a$1")
    assertEquals(clean, "a1")
    assertEquals(anchor, Anchor.Absolute)

  // ===== FormulaParser anchor tests =====

  test("FormulaParser: parse relative reference"):
    val result = FormulaParser.parse("=A1")
    assert(result.isRight, s"Parse failed: $result")
    result.foreach { expr =>
      expr match
        case TExpr.PolyRef(_, anchor) =>
          assertEquals(anchor, Anchor.Relative)
        case other =>
          fail(s"Expected PolyRef, got $other")
    }

  test("FormulaParser: parse absolute column $A1"):
    val result = FormulaParser.parse("=$A1")
    assert(result.isRight, s"Parse failed: $result")
    result.foreach { expr =>
      expr match
        case TExpr.PolyRef(_, anchor) =>
          assertEquals(anchor, Anchor.AbsCol)
        case other =>
          fail(s"Expected PolyRef, got $other")
    }

  test("FormulaParser: parse absolute row A$1"):
    val result = FormulaParser.parse("=A$1")
    assert(result.isRight, s"Parse failed: $result")
    result.foreach { expr =>
      expr match
        case TExpr.PolyRef(_, anchor) =>
          assertEquals(anchor, Anchor.AbsRow)
        case other =>
          fail(s"Expected PolyRef, got $other")
    }

  test("FormulaParser: parse absolute both $A$1"):
    val result = FormulaParser.parse("=$A$1")
    assert(result.isRight, s"Parse failed: $result")
    result.foreach { expr =>
      expr match
        case TExpr.PolyRef(_, anchor) =>
          assertEquals(anchor, Anchor.Absolute)
        case other =>
          fail(s"Expected PolyRef, got $other")
    }

  test("FormulaParser: parse mixed anchors in formula"):
    val result = FormulaParser.parse("=A1*$B$2")
    assert(result.isRight, s"Parse failed: $result")
    // Should parse without error, contains both relative and absolute refs

  test("FormulaParser: parse anchored range reference"):
    val result = FormulaParser.parse("=SUM($A$1:B10)")
    assert(result.isRight, s"Parse failed: $result")

  // ===== FormulaPrinter anchor tests =====

  test("FormulaPrinter: print relative reference"):
    val ref = ARef.parse("A1").toOption.get
    val expr = TExpr.PolyRef(ref, Anchor.Relative)
    assertEquals(FormulaPrinter.print(expr), "=A1")

  test("FormulaPrinter: print absolute column $A1"):
    val ref = ARef.parse("A1").toOption.get
    val expr = TExpr.PolyRef(ref, Anchor.AbsCol)
    assertEquals(FormulaPrinter.print(expr), "=$A1")

  test("FormulaPrinter: print absolute row A$1"):
    val ref = ARef.parse("A1").toOption.get
    val expr = TExpr.PolyRef(ref, Anchor.AbsRow)
    assertEquals(FormulaPrinter.print(expr), "=A$1")

  test("FormulaPrinter: print absolute both $A$1"):
    val ref = ARef.parse("A1").toOption.get
    val expr = TExpr.PolyRef(ref, Anchor.Absolute)
    assertEquals(FormulaPrinter.print(expr), "=$A$1")

  // ===== Round-trip tests =====

  test("round-trip: relative reference A1"):
    val original = "=A1"
    val parsed = FormulaParser.parse(original)
    assert(parsed.isRight)
    assertEquals(FormulaPrinter.print(parsed.toOption.get), original)

  test("round-trip: absolute column $A1"):
    val original = "=$A1"
    val parsed = FormulaParser.parse(original)
    assert(parsed.isRight)
    assertEquals(FormulaPrinter.print(parsed.toOption.get), original)

  test("round-trip: absolute row A$1"):
    val original = "=A$1"
    val parsed = FormulaParser.parse(original)
    assert(parsed.isRight)
    assertEquals(FormulaPrinter.print(parsed.toOption.get), original)

  test("round-trip: absolute both $A$1"):
    val original = "=$A$1"
    val parsed = FormulaParser.parse(original)
    assert(parsed.isRight)
    assertEquals(FormulaPrinter.print(parsed.toOption.get), original)

  test("round-trip: mixed anchors A1*$B$2"):
    val original = "=A1*$B$2"
    val parsed = FormulaParser.parse(original)
    assert(parsed.isRight)
    assertEquals(FormulaPrinter.print(parsed.toOption.get), original)

  // ===== FormulaShifter tests =====

  test("FormulaShifter: shift relative reference right"):
    val parsed = FormulaParser.parse("=A1").toOption.get
    val shifted = FormulaShifter.shift(parsed, colDelta = 1, rowDelta = 0)
    assertEquals(FormulaPrinter.print(shifted), "=B1")

  test("FormulaShifter: shift relative reference down"):
    val parsed = FormulaParser.parse("=A1").toOption.get
    val shifted = FormulaShifter.shift(parsed, colDelta = 0, rowDelta = 1)
    assertEquals(FormulaPrinter.print(shifted), "=A2")

  test("FormulaShifter: shift relative reference diagonal"):
    val parsed = FormulaParser.parse("=A1").toOption.get
    val shifted = FormulaShifter.shift(parsed, colDelta = 2, rowDelta = 3)
    assertEquals(FormulaPrinter.print(shifted), "=C4")

  test("FormulaShifter: absolute column $A1 - column stays fixed"):
    val parsed = FormulaParser.parse("=$A1").toOption.get
    val shifted = FormulaShifter.shift(parsed, colDelta = 5, rowDelta = 2)
    assertEquals(FormulaPrinter.print(shifted), "=$A3")

  test("FormulaShifter: absolute row A$1 - row stays fixed"):
    val parsed = FormulaParser.parse("=A$1").toOption.get
    val shifted = FormulaShifter.shift(parsed, colDelta = 2, rowDelta = 5)
    assertEquals(FormulaPrinter.print(shifted), "=C$1")

  test("FormulaShifter: absolute both $A$1 - both stay fixed"):
    val parsed = FormulaParser.parse("=$A$1").toOption.get
    val shifted = FormulaShifter.shift(parsed, colDelta = 10, rowDelta = 10)
    assertEquals(FormulaPrinter.print(shifted), "=$A$1")

  test("FormulaShifter: mixed anchors in formula"):
    val parsed = FormulaParser.parse("=A1*$B$2").toOption.get
    val shifted = FormulaShifter.shift(parsed, colDelta = 1, rowDelta = 1)
    assertEquals(FormulaPrinter.print(shifted), "=B2*$B$2")

  test("FormulaShifter: complex formula with multiple refs"):
    val parsed = FormulaParser.parse("=A1+$B1+C$1+$D$1").toOption.get
    val shifted = FormulaShifter.shift(parsed, colDelta = 2, rowDelta = 3)
    // A1 (relative) -> C4
    // $B1 (abs col) -> $B4
    // C$1 (abs row) -> E$1
    // $D$1 (absolute) -> $D$1
    assertEquals(FormulaPrinter.print(shifted), "=C4+$B4+E$1+$D$1")

  test("FormulaShifter: shift with zero delta is identity"):
    val original = "=A1*$B$2+C3"
    val parsed = FormulaParser.parse(original).toOption.get
    val shifted = FormulaShifter.shift(parsed, colDelta = 0, rowDelta = 0)
    assertEquals(FormulaPrinter.print(shifted), original)

  test("FormulaShifter: range reference shifts"):
    val parsed = FormulaParser.parse("=SUM(A1:B2)").toOption.get
    val shifted = FormulaShifter.shift(parsed, colDelta = 1, rowDelta = 1)
    assertEquals(FormulaPrinter.print(shifted), "=SUM(B2:C3)")

  test("FormulaShifter: clamps to valid range (no negative indices)"):
    val parsed = FormulaParser.parse("=A1").toOption.get
    val shifted = FormulaShifter.shift(parsed, colDelta = -5, rowDelta = -5)
    // Should clamp to A1 (0,0) not go negative
    assertEquals(FormulaPrinter.print(shifted), "=A1")

  // ===== Anchor extension method tests =====

  test("Anchor.isColAbsolute: true for AbsCol"):
    assert(Anchor.AbsCol.isColAbsolute)

  test("Anchor.isColAbsolute: true for Absolute"):
    assert(Anchor.Absolute.isColAbsolute)

  test("Anchor.isColAbsolute: false for Relative"):
    assert(!Anchor.Relative.isColAbsolute)

  test("Anchor.isColAbsolute: false for AbsRow"):
    assert(!Anchor.AbsRow.isColAbsolute)

  test("Anchor.isRowAbsolute: true for AbsRow"):
    assert(Anchor.AbsRow.isRowAbsolute)

  test("Anchor.isRowAbsolute: true for Absolute"):
    assert(Anchor.Absolute.isRowAbsolute)

  test("Anchor.isRowAbsolute: false for Relative"):
    assert(!Anchor.Relative.isRowAbsolute)

  test("Anchor.isRowAbsolute: false for AbsCol"):
    assert(!Anchor.AbsCol.isRowAbsolute)

  // ===== CellRange anchor tests =====

  test("CellRange.parse: anchored range $A$1:B10"):
    val result = CellRange.parse("$A$1:B10")
    assert(result.isRight, s"Parse failed: $result")
    result.foreach { range =>
      assertEquals(range.startAnchor, Anchor.Absolute)
      assertEquals(range.endAnchor, Anchor.Relative)
    }

  test("CellRange.parse: mixed anchors $A1:B$10"):
    val result = CellRange.parse("$A1:B$10")
    assert(result.isRight, s"Parse failed: $result")
    result.foreach { range =>
      assertEquals(range.startAnchor, Anchor.AbsCol)
      assertEquals(range.endAnchor, Anchor.AbsRow)
    }

  test("CellRange.parse: both endpoints absolute $A$1:$B$10"):
    val result = CellRange.parse("$A$1:$B$10")
    assert(result.isRight, s"Parse failed: $result")
    result.foreach { range =>
      assertEquals(range.startAnchor, Anchor.Absolute)
      assertEquals(range.endAnchor, Anchor.Absolute)
    }

  test("CellRange.parse: both endpoints relative A1:B10"):
    val result = CellRange.parse("A1:B10")
    assert(result.isRight, s"Parse failed: $result")
    result.foreach { range =>
      assertEquals(range.startAnchor, Anchor.Relative)
      assertEquals(range.endAnchor, Anchor.Relative)
    }

  test("CellRange.parse: single cell with anchor $A$1"):
    val result = CellRange.parse("$A$1")
    assert(result.isRight, s"Parse failed: $result")
    result.foreach { range =>
      assertEquals(range.startAnchor, Anchor.Absolute)
      assertEquals(range.endAnchor, Anchor.Absolute)
    }

  test("CellRange.toA1Anchored: formats with anchors"):
    val result = CellRange.parse("$A$1:B10")
    assert(result.isRight)
    result.foreach { range =>
      assertEquals(range.toA1Anchored, "$A$1:B10")
    }

  test("CellRange.toA1: formats without anchors (backward compatible)"):
    val result = CellRange.parse("$A$1:B10")
    assert(result.isRight)
    result.foreach { range =>
      assertEquals(range.toA1, "A1:B10")
    }

  // ===== FormulaShifter anchored range tests =====

  test("FormulaShifter: shift anchored range $A$1:B10 - start fixed, end shifts"):
    val parsed = FormulaParser.parse("=SUM($A$1:B10)").toOption.get
    val shifted = FormulaShifter.shift(parsed, colDelta = 2, rowDelta = 3)
    // $A$1 stays fixed (Absolute), B10 → D13 (Relative)
    assertEquals(FormulaPrinter.print(shifted), "=SUM($A$1:D13)")

  test("FormulaShifter: shift mixed anchors $A1:B$10"):
    val parsed = FormulaParser.parse("=SUM($A1:B$10)").toOption.get
    val shifted = FormulaShifter.shift(parsed, colDelta = 2, rowDelta = 3)
    // $A1 → $A4 (col fixed, row shifts), B$10 → D$10 (col shifts, row fixed)
    assertEquals(FormulaPrinter.print(shifted), "=SUM($A4:D$10)")

  test("FormulaShifter: shift all-absolute range $A$1:$B$10 - nothing shifts"):
    val parsed = FormulaParser.parse("=SUM($A$1:$B$10)").toOption.get
    val shifted = FormulaShifter.shift(parsed, colDelta = 5, rowDelta = 5)
    // Both endpoints are Absolute, so nothing shifts
    assertEquals(FormulaPrinter.print(shifted), "=SUM($A$1:$B$10)")

  test("FormulaShifter: shift all-relative range A1:B10 - both endpoints shift"):
    val parsed = FormulaParser.parse("=SUM(A1:B10)").toOption.get
    val shifted = FormulaShifter.shift(parsed, colDelta = 2, rowDelta = 3)
    // Both endpoints are Relative
    assertEquals(FormulaPrinter.print(shifted), "=SUM(C4:D13)")

  // ===== Round-trip tests for anchored ranges =====

  test("round-trip: anchored range $A$1:B10"):
    val original = "=SUM($A$1:B10)"
    val parsed = FormulaParser.parse(original)
    assert(parsed.isRight, s"Parse failed: $parsed")
    assertEquals(FormulaPrinter.print(parsed.toOption.get), original)

  test("round-trip: mixed anchors $A1:B$10"):
    val original = "=SUM($A1:B$10)"
    val parsed = FormulaParser.parse(original)
    assert(parsed.isRight, s"Parse failed: $parsed")
    assertEquals(FormulaPrinter.print(parsed.toOption.get), original)

  test("round-trip: all absolute $A$1:$B$10"):
    val original = "=SUM($A$1:$B$10)"
    val parsed = FormulaParser.parse(original)
    assert(parsed.isRight, s"Parse failed: $parsed")
    assertEquals(FormulaPrinter.print(parsed.toOption.get), original)

  test("round-trip: AbsRow start, AbsCol end A$1:$B10"):
    val original = "=SUM(A$1:$B10)"
    val parsed = FormulaParser.parse(original)
    assert(parsed.isRight, s"Parse failed: $parsed")
    assertEquals(FormulaPrinter.print(parsed.toOption.get), original)
