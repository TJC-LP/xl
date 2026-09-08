package com.tjclp.xl.cells

import munit.FunSuite

/**
 * GH-630: `CellError` models every error value Excel writes — the classic seven and the modern
 * `#SPILL!`, `#CALC!`, `#FIELD!`, `#CONNECT!`, `#BLOCKED!`, `#UNKNOWN!`, `#GETTING_DATA` — with the
 * spelling Excel uses in a `t="e"` cell and the number `ERROR.TYPE` assigns each.
 */
class CellErrorSpec extends FunSuite:

  test("every error round-trips through its Excel spelling, case-insensitively") {
    CellError.values.foreach { err =>
      assertEquals(CellError.parse(err.toExcel), Right(err))
      assertEquals(CellError.parse(err.toExcel.toLowerCase), Right(err))
    }
    assertEquals(CellError.values.map(_.toExcel).distinct.length, CellError.values.length)
  }

  test("the modern codes spell as Excel writes them") {
    assertEquals(CellError.Spill.toExcel, "#SPILL!")
    assertEquals(CellError.Calc.toExcel, "#CALC!")
    assertEquals(CellError.Field.toExcel, "#FIELD!")
    assertEquals(CellError.Connect.toExcel, "#CONNECT!")
    assertEquals(CellError.Blocked.toExcel, "#BLOCKED!")
    assertEquals(CellError.Unknown.toExcel, "#UNKNOWN!")
    assertEquals(CellError.GettingData.toExcel, "#GETTING_DATA")
  }

  test("an unknown spelling is a Left naming it") {
    assertEquals(CellError.parse("#BOGUS!"), Left("Unknown error: #BOGUS!"))
    assertEquals(CellError.parse(""), Left("Unknown error: "))
  }

  test("ERROR.TYPE numbers follow Microsoft's table and are distinct") {
    val numbers = CellError.values.toVector.map(e => e.errorTypeNumber -> e).sortBy(_._1)
    assertEquals(numbers.map(_._1), (1 to 14).toVector)
    assertEquals(
      numbers.map(_._2),
      Vector(
        CellError.Null,
        CellError.Div0,
        CellError.Value,
        CellError.Ref,
        CellError.Name,
        CellError.Num,
        CellError.NA,
        CellError.GettingData,
        CellError.Spill,
        CellError.Connect,
        CellError.Blocked,
        CellError.Unknown,
        CellError.Field,
        CellError.Calc
      )
    )
  }
