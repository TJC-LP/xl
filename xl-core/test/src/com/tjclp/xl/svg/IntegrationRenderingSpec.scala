package com.tjclp.xl.svg

import com.tjclp.xl.{*, given}
import munit.FunSuite

class IntegrationRenderingSpec extends FunSuite:
  test("GH-505: a roomy numeric cell retains its backend's natural glyph advances") {
    val sheet = Sheet("Data")
      .put(ref"A1", CellValue.Number(BigDecimal("1234567.9")))
      .setColumnProperties(Column.from0(0), ColumnProperties(width = Some(20.0)))
    assert(!sheet.toSvg(ref"A1:A1").contains("textLength="))
  }
