package com.tjclp.xl.addressing

import com.tjclp.xl.Generators.given
import munit.ScalaCheckSuite
import org.scalacheck.Prop.*

/**
 * Runtime column parsing (GH-361): `Column.parse` is the discoverable runtime counterpart of the
 * compile-time `ref` literal for scripts that compute column letters dynamically.
 */
class ColumnSpec extends ScalaCheckSuite:

  // ========== Column.parse: bare letters ==========

  test("parse accepts a bare letter"):
    assertEquals(Column.parse("D"), Right(Column.from0(3)))

  test("parse accepts multi-letter columns"):
    assertEquals(Column.parse("AA"), Right(Column.from0(26)))
    assertEquals(Column.parse("XFD"), Right(Column.from0(Column.MaxIndex0)))

  test("parse is case-insensitive"):
    assertEquals(Column.parse("d"), Right(Column.from0(3)))
    assertEquals(Column.parse("aA"), Right(Column.from0(26)))

  test("parse tolerates a trailing row number (full A1 refs work directly)"):
    assertEquals(Column.parse("D1"), Right(Column.from0(3)))
    assertEquals(Column.parse("xfd1048576"), Right(Column.from0(Column.MaxIndex0)))

  // ========== Column.parse: rejections ==========

  test("parse rejects empty input"):
    assert(Column.parse("").isLeft)

  test("parse rejects digits-only input"):
    assert(Column.parse("12").isLeft)

  test("parse rejects letters after the row digits"):
    assert(Column.parse("D1D").isLeft)

  test("parse rejects absolute anchors and other punctuation"):
    assert(Column.parse("$D").isLeft)
    assert(Column.parse("D$1").isLeft)
    assert(Column.parse("D:D").isLeft)

  test("parse rejects columns past XFD"):
    assert(Column.parse("XFE").isLeft)
    assert(Column.parse("ZZZZ").isLeft)

  test("parse rejects non-ASCII letters and unicode digits"):
    assert(Column.parse("Δ").isLeft)
    assert(Column.parse("D١").isLeft) // Arabic-Indic digit is not a row number

  // ========== fromLetter edge cases (sibling of parse) ==========

  test("fromLetter rejects empty, mixed, and overflow inputs"):
    assert(Column.fromLetter("").isLeft)
    assert(Column.fromLetter("D1").isLeft) // strict: letters only
    assert(Column.fromLetter("XFE").isLeft)
    assertEquals(Column.fromLetter("xfd"), Right(Column.from0(Column.MaxIndex0)))

  // ========== Laws ==========

  property("parse . toLetter = Right (round-trip for every valid column)") {
    forAll { (col: Column) =>
      assertEquals(Column.parse(col.toLetter), Right(col))
      true
    }
  }

  property("parse tolerates any appended 1-based row number") {
    forAll { (col: Column, row: Row) =>
      assertEquals(Column.parse(s"${col.toLetter}${row.index1}"), Right(col))
      true
    }
  }

  property("parse agrees with fromLetter on letters-only input") {
    forAll { (col: Column) =>
      assertEquals(Column.parse(col.toLetter), Column.fromLetter(col.toLetter))
      true
    }
  }
