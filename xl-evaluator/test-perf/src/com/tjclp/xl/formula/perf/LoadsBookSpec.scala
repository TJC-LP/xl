package com.tjclp.xl.formula.perf

import com.tjclp.xl.cells.CellValue
import munit.FunSuite

/** #767: the synthetic generator is deterministic and builds the shape it describes. */
class LoadsBookSpec extends FunSuite:

  private def formulaCells(built: LoadsBook.Built): Int =
    built.workbook.sheets.iterator
      .map(_.cells.valuesIterator.count(_.value match
        case _: CellValue.Formula => true
        case _ => false))
      .sum

  test("a shape builds the same workbook every time") {
    assertEquals(LoadsBook.build(LoadsBook.Tiny), LoadsBook.build(LoadsBook.Tiny))
  }

  test("a different seed builds a different workbook") {
    assertNotEquals(
      LoadsBook.build(LoadsBook.Tiny).workbook,
      LoadsBook.build(LoadsBook.Tiny.copy(seed = 1L)).workbook
    )
  }

  test("Tiny: the formula census matches the cells") {
    val built = LoadsBook.build(LoadsBook.Tiny)
    assertEquals(built.dataFormulas, 3 * 400)
    assertEquals(built.summaryFormulas, 64)
    assertEquals(built.rhFormulas, 24 * 14)
    assertEquals(built.totalFormulas, 1600)
    assertEquals(formulaCells(built), built.totalFormulas)
    assertEquals(built.sampleRefs.size, 40)
    assertEquals(built.sampleRefs.distinct.size, 40)
  }

  test("raw criteria without helpers: no Loads formulas, the same Summary and Rate History") {
    val built = LoadsBook.build(LoadsBook.Tiny.copy(helpers = false, crit = LoadsBook.Crit.Raw))
    assertEquals(built.dataFormulas, 0)
    assertEquals(formulaCells(built), 64 + 24 * 14)
    assert(!built.workbook.metadata.definedNames.exists(_.name == "ld_spend"))
  }

  test("lanes and carriers scale with rows, so equality buckets stay flat") {
    assertEquals(LoadsBook.Tiny.nLanes, 20)
    assertEquals(LoadsBook.R20.nLanes, 400)
    assertEquals(LoadsBook.R20.nCarriers, 500)
    assertEquals(LoadsBook.Real.nLanes, 664)
    assertEquals(LoadsBook.R20.rows / LoadsBook.R20.nLanes, 50)
  }

  test("presets are found by name, case-insensitively") {
    assertEquals(LoadsBook.preset("h10"), Some(LoadsBook.H10))
    assertEquals(LoadsBook.preset("nope"), None)
  }

  test("whole-column names cover the column; bounded names cover the data rows") {
    def target(shape: LoadsBook.Shape) =
      LoadsBook.build(shape).workbook.metadata.definedNames.find(_.name == "ld_cost").map(_.formula)
    assertEquals(target(LoadsBook.Tiny), Some("Loads!$F$2:$F$401"))
    assertEquals(target(LoadsBook.Tiny.copy(wholeColumns = true)), Some("Loads!$F:$F"))
  }
