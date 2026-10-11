package com.tjclp.xl.cli

import cats.effect.IO
import munit.CatsEffectSuite

import com.tjclp.xl.{*, given}
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.cli.commands.ReadCommands

class PositionedReferenceInspectionSpec extends CatsEffectSuite:
  private val sheet = Sheet("S").put(ref"D5", CellValue.Number(BigDecimal(99)))
  private val wb = Workbook(sheet)

  List(
    "ROW(D5)" -> 5,
    "COLUMN(D5)" -> 4,
    "ROWS(D5:D6)" -> 2,
    "CELL(\"row\",D5)" -> 5,
    "CELL(\"col\",D5)" -> 4
  ).foreach { (formula, expected) =>
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

  test("inspection nested in arithmetic does not create a circular value dependency") {
    ReadCommands.evalData(wb, Some(sheet), "ROW(D5)+ROWS(D5:D6)", Some("D5"), Nil).map { text =>
      assertEquals(ujson.read(text)("result")("value").num, 7.0)
    }
  }

  test("an INDEX selector that reads the placed cell is still circular") {
    ReadCommands.evalData(wb, Some(sheet), "ROW(INDEX(A1:A10,D5))", Some("D5"), Nil).attempt.map {
      result =>
        assert(result.isLeft, result.toString)
    }
  }

  test("a precedent may inspect the placement address without reading its old value") {
    val source = sheet.put(ref"A1", CellValue.Formula("ROW(D5)", None))
    ReadCommands.evalData(Workbook(source), Some(source), "A1+1", Some("D5"), Nil).map { text =>
      assertEquals(ujson.read(text)("result")("value").num, 6.0)
    }
  }

  test("an existing self-inspection formula does not poison the closure order") {
    val source = sheet.put(ref"A1", CellValue.Formula("ROW(A1)", None))
    ReadCommands.evalData(Workbook(source), Some(source), "A1+1", Some("D5"), Nil).map { text =>
      assertEquals(ujson.read(text)("result")("value").num, 2.0)
    }
  }

  List(
    "LET(chosen,D5,ROWS(chosen))" -> 1,
    "LET(chosen,INDEX(D5:D6,1),ROWS(chosen))" -> 1,
    "LET(outer,D5,LET(inner,outer,COLUMNS(inner)))" -> 1
  ).foreach { (formula, expected) =>
    test(s"positioned $formula follows LET references without a circular value read") {
      ReadCommands.evalData(wb, Some(sheet), formula, Some("D5"), Nil).map { text =>
        assertEquals(ujson.read(text)("result")("value").num, expected.toDouble)
      }
    }
  }

  List(
    "LET(chosen,D5,chosen+1)",
    "LET(chosen,INDEX(A1:A10,D5),ROWS(chosen))",
    "LET(unused,D5+1,ROW(A1))"
  ).foreach { formula =>
    test(s"positioned $formula retains actual LET value reads") {
      ReadCommands.evalData(wb, Some(sheet), formula, Some("D5"), Nil).attempt.map { result =>
        assert(result.left.exists(_.getMessage.contains("Circular")), result.toString)
      }
    }
  }
