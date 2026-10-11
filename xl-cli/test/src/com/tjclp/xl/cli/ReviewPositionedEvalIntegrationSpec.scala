package com.tjclp.xl.cli

import cats.effect.IO
import munit.CatsEffectSuite

import com.tjclp.xl.{*, given}
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.cli.commands.ReadCommands

class ReviewPositionedEvalIntegrationSpec extends CatsEffectSuite:
  private val sheet = Sheet("S").put(ref"D5", CellValue.Number(BigDecimal(99)))
  private val wb = Workbook(sheet)

  List("ROW(D5)" -> 5, "COLUMN(D5)" -> 4, "ROWS(D5:D6)" -> 2).foreach { (formula, expected) =>
    test(s"positioned $formula can inspect its own address without reading its value") {
      ReadCommands.evalData(wb, Some(sheet), formula, Some("D5"), Nil).map { text =>
        assertEquals(ujson.read(text)("result")("value").num, expected.toDouble)
      }
    }
  }

  test("positioned ordinary scalar reference still refuses a real circular read") {
    ReadCommands.evalData(wb, Some(sheet), "D5+1", Some("D5"), Nil).attempt.map { result =>
      assert(result.isLeft, result.toString)
    }
  }

  test("positioned ROW without a reference composes with array semantics changes") {
    ReadCommands.evalData(wb, Some(sheet), "ROW()+COLUMN()", Some("D5"), Nil).map { text =>
      assertEquals(ujson.read(text)("result")("value").num, 9.0)
    }
  }
