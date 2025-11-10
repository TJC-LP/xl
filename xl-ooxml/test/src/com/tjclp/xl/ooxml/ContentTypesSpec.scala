package com.tjclp.xl.ooxml

import munit.FunSuite

class ContentTypesSpec extends FunSuite:
  test("forSheetIndices emits overrides for given sheet indices") {
    val contentTypes = ContentTypes.forSheetIndices(Seq(3, 1, 3), hasStyles = true, hasSharedStrings = false)

    val overrides = contentTypes.overrides
    assert(overrides.contains("/xl/workbook.xml"), "Workbook override missing")
    assert(overrides.contains("/xl/styles.xml"), "Styles override missing")
    assertEquals(overrides.contains("/xl/worksheets/sheet3.xml"), true)
    assertEquals(overrides.contains("/xl/worksheets/sheet1.xml"), true)
    // Ensure no unwanted sheet2 entry when not requested
    assertEquals(overrides.contains("/xl/worksheets/sheet2.xml"), false)
  }
