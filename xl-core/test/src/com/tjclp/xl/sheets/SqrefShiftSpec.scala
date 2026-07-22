package com.tjclp.xl.sheets

import com.tjclp.xl.{*, given}
import com.tjclp.xl.addressing.{ARef, CellRange, Row}
import munit.FunSuite

/**
 * GH-429: byte-surgical sqref rewrite for Preserved payloads. The guards are all-or-nothing: any
 * doubt returns the payload unchanged (never a partial rewrite); only a full collapse drops the
 * entry.
 */
class SqrefShiftSpec extends FunSuite:

  private def rowInsert(at: Int, count: Int)(r: CellRange): Option[CellRange] =
    Sheet
      .shiftSpan(r.start.row.index0, r.end.row.index0, at, count, deleting = false, Row.MaxIndex0)
      .map((ns, ne) =>
        CellRange(ARef.from0(r.start.col.index0, ns), ARef.from0(r.end.col.index0, ne))
      )

  private def rowDelete(at: Int, count: Int)(r: CellRange): Option[CellRange] =
    Sheet
      .shiftSpan(r.start.row.index0, r.end.row.index0, at, count, deleting = true, Row.MaxIndex0)
      .map((ns, ne) =>
        CellRange(ARef.from0(r.start.col.index0, ns), ARef.from0(r.end.col.index0, ne))
      )

  test("shifts the root sqref and nothing else") {
    val xml =
      """<dataValidation type="whole" sqref="H5:H10"><formula1>1</formula1></dataValidation>"""
    assertEquals(
      SqrefShift.shiftPayload(xml, "dataValidation", rowInsert(1, 2)),
      Some(
        """<dataValidation type="whole" sqref="H7:H12"><formula1>1</formula1></dataValidation>"""
      )
    )
  }

  test("identity shift returns the payload byte-identically") {
    val xml = """<dataValidation sqref="H5:H10" type="whole"/>"""
    assertEquals(SqrefShift.shiftPayload(xml, "dataValidation", Some(_)), Some(xml))
  }

  test("refuse-matrix: missing sqref, two matches, xr:sqref only, corrupt token, label mismatch") {
    val shift = rowInsert(0, 1)
    val missing = """<dataValidation type="whole"><formula1>1</formula1></dataValidation>"""
    assertEquals(SqrefShift.shiftPayload(missing, "dataValidation", shift), Some(missing))
    val two = """<dataValidation sqref="A1" sqref="B2"/>"""
    assertEquals(SqrefShift.shiftPayload(two, "dataValidation", shift), Some(two))
    val xrOnly = """<dataValidation xr:sqref="A1"/>"""
    assertEquals(SqrefShift.shiftPayload(xrOnly, "dataValidation", shift), Some(xrOnly))
    val corrupt = """<dataValidation sqref="NOT A REF"/>"""
    assertEquals(SqrefShift.shiftPayload(corrupt, "dataValidation", shift), Some(corrupt))
    val empty = """<dataValidation sqref=""/>"""
    assertEquals(SqrefShift.shiftPayload(empty, "dataValidation", shift), Some(empty))
    val mismatch = """<conditionalFormatting sqref="A1"/>"""
    assertEquals(SqrefShift.shiftPayload(mismatch, "dataValidation", shift), Some(mismatch))
    val labelPrefix = """<dataValidationX sqref="A1"/>"""
    assertEquals(SqrefShift.shiftPayload(labelPrefix, "dataValidation", shift), Some(labelPrefix))
  }

  test("xr:sqref beside a real sqref: only the real one rewrites") {
    val xml = """<dataValidation xr:sqref="A1" sqref="A1"/>"""
    assertEquals(
      SqrefShift.shiftPayload(xml, "dataValidation", rowInsert(0, 1)),
      Some("""<dataValidation xr:sqref="A1" sqref="A2"/>""")
    )
  }

  test("multi-token sqref: partial drop keeps survivors; all-drop removes the entry") {
    val xml = """<dataValidation sqref="A2:A3 C5"/>"""
    assertEquals(
      SqrefShift.shiftPayload(xml, "dataValidation", rowDelete(1, 2)), // rows 2-3 vanish
      Some("""<dataValidation sqref="C3"/>""")
    )
    val gone = """<dataValidation sqref="A2:A3 B2"/>"""
    assertEquals(SqrefShift.shiftPayload(gone, "dataValidation", rowDelete(1, 3)), None)
  }

  test("full-column token is a byte-identical fixed point under row inserts (post GH-428)") {
    val xml = """<conditionalFormatting sqref="A:C"/>"""
    assertEquals(
      SqrefShift.shiftPayload(xml, "conditionalFormatting", rowInsert(1, 2)),
      Some(xml)
    )
  }

  test("changed tokens print in Excel's shape: full-row and 1x1 forms") {
    // full-row token 2:3 under a row insert above stays full-row shaped
    val fullRow = """<conditionalFormatting sqref="2:3"/>"""
    assertEquals(
      SqrefShift.shiftPayload(fullRow, "conditionalFormatting", rowInsert(0, 1)),
      Some("""<conditionalFormatting sqref="3:4"/>""")
    )
    // a single-cell token stays single-cell shaped
    val cell = """<dataValidation sqref="B2"/>"""
    assertEquals(
      SqrefShift.shiftPayload(cell, "dataValidation", rowInsert(0, 3)),
      Some("""<dataValidation sqref="B5"/>""")
    )
  }

  test("adversarial sqref lookalike inside a child rides untouched (canonical escaping)") {
    // canonical payloads escape every '"' as &quot; outside attribute delimiters, so a
    // formula1 containing `sqref="H5"` can never match the start-tag regex
    val xml =
      """<dataValidation sqref="H5:H10"><formula1>"sqref=&quot;H5&quot;"</formula1></dataValidation>"""
    assertEquals(
      SqrefShift.shiftPayload(xml, "dataValidation", rowInsert(0, 1)),
      Some(
        """<dataValidation sqref="H6:H11"><formula1>"sqref=&quot;H5&quot;"</formula1></dataValidation>"""
      )
    )
  }
