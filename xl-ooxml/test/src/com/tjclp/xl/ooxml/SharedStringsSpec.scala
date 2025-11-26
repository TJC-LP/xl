package com.tjclp.xl.ooxml

import munit.FunSuite
import com.tjclp.xl.api.*
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.codec.CellCodec.given
import com.tjclp.xl.macros.ref

/** Tests for SharedStrings count vs uniqueCount */
class SharedStringsSpec extends FunSuite:

  test("fromWorkbook counts total instances correctly") {
    val sheet = Sheet("Test")
      .put(ref"A1", CellValue.Text("Hello"))
      .put(ref"A2", CellValue.Text("World"))
      .put(ref"A3", CellValue.Text("Hello")) // Duplicate
      .put(ref"A4", CellValue.Text("Hello")) // Another duplicate

    val wb = Workbook(Vector(sheet))
    val sst = SharedStrings.fromWorkbook(wb)

    // Should have 4 total instances (including duplicates)
    assertEquals(sst.totalCount, 4, "totalCount should count all instances")

    // Should have 2 unique strings
    assertEquals(sst.strings.size, 2, "uniqueCount should be 2")
  }

  test("toXml emits correct count and uniqueCount") {
    val sheet = Sheet("Test")
      .put(ref"A1", CellValue.Text("Alpha"))
      .put(ref"A2", CellValue.Text("Beta"))
      .put(ref"A3", CellValue.Text("Alpha")) // Duplicate

    val wb = Workbook(Vector(sheet))
    val sst = SharedStrings.fromWorkbook(wb)
    val xml = sst.toXml

    // Check count attribute (total instances)
    val count = (xml \ "@count").text.toInt
    assertEquals(count, 3, "count attribute should be 3 (total instances)")

    // Check uniqueCount attribute (unique strings)
    val uniqueCount = (xml \ "@uniqueCount").text.toInt
    assertEquals(uniqueCount, 2, "uniqueCount attribute should be 2 (unique strings)")
  }

  test("workbook with duplicate strings has count > uniqueCount") {
    val sheet = Sheet("Test")
      .put(ref"A1", CellValue.Text("Duplicate"))
      .put(ref"A2", CellValue.Text("Duplicate"))
      .put(ref"A3", CellValue.Text("Duplicate"))
      .put(ref"A4", CellValue.Text("Unique"))

    val wb = Workbook(Vector(sheet))
    val sst = SharedStrings.fromWorkbook(wb)

    assert(sst.totalCount > sst.strings.size,
      s"count (${sst.totalCount}) should be > uniqueCount (${sst.strings.size})")
    assertEquals(sst.totalCount, 4)
    assertEquals(sst.strings.size, 2)
  }

  test("workbook with no duplicates has count == uniqueCount") {
    val sheet = Sheet("Test")
      .put(ref"A1", CellValue.Text("One"))
      .put(ref"A2", CellValue.Text("Two"))
      .put(ref"A3", CellValue.Text("Three"))

    val wb = Workbook(Vector(sheet))
    val sst = SharedStrings.fromWorkbook(wb)

    assertEquals(sst.totalCount, sst.strings.size,
      "count should equal uniqueCount when no duplicates")
    assertEquals(sst.totalCount, 3)
    assertEquals(sst.strings.size, 3)
  }

  test("empty SharedStrings has count and uniqueCount of 0") {
    val sst = SharedStrings.empty
    val xml = sst.toXml

    val count = (xml \ "@count").text.toInt
    val uniqueCount = (xml \ "@uniqueCount").text.toInt

    assertEquals(count, 0, "empty SST should have count=0")
    assertEquals(uniqueCount, 0, "empty SST should have uniqueCount=0")
  }

  test("fromStrings with explicit totalCount") {
    val strings = Vector("A", "B", "A", "C", "A") // 5 total, 3 unique
    val sst = SharedStrings.fromStrings(strings, Some(5))

    assertEquals(sst.totalCount, 5, "totalCount should match explicit value")
    assertEquals(sst.strings.size, 3, "uniqueCount should be 3")
  }

  test("fromStrings without totalCount defaults to input size") {
    val strings = Vector("X", "Y", "Z")
    val sst = SharedStrings.fromStrings(strings)

    // Without explicit totalCount, defaults to strings.size (3 in this case)
    assertEquals(sst.totalCount, 3)
    assertEquals(sst.strings.size, 3)
  }
