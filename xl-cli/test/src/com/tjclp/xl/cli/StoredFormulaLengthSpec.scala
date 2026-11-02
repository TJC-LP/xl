package com.tjclp.xl.cli

import cats.effect.{IO, unsafe}
import com.tjclp.xl.{Sheet, Workbook}
import com.tjclp.xl.cli.commands.WriteCommands
import com.tjclp.xl.cli.helpers.{BatchParser, ValueParser}
import com.tjclp.xl.formula.FormulaParser
import com.tjclp.xl.ooxml.FormulaStorage
import com.tjclp.xl.ooxml.writer.WriterConfig
import java.nio.file.Files
import munit.FunSuite

class StoredFormulaLengthSpec extends FunSuite:
  given unsafe.IORuntime = unsafe.IORuntime.global

  // One SUM argument avoids Excel's argument-count limit; the 699 operators fit XL's cap.
  // A shallow LET with many references: valid input length, excessive storage prefixes.
  private def formula(references: Int): String =
    s"LET(local,1,SUM(${List.fill(references)("local").mkString("+")}))"
  private val oversized = formula(700)

  test("GH-693: premise is parseable input whose actual stored text exceeds 8192") {
    assert(oversized.length < 8192)
    assert(FormulaParser.parse(oversized).isRight)
    assert(FormulaStorage.toStored(oversized).length > 8192)
  }

  test("GH-693: putf refuses oversized stored text before creating output") {
    val dir = Files.createTempDirectory("stored-formula-")
    val output = dir.resolve("out.xlsx")
    try
      val wb = Workbook(Sheet("Test"))
      val result = WriteCommands
        .putFormula(wb, wb.sheets.headOption, "A1", List(oversized), output, WriterConfig.default)
        .attempt
        .unsafeRunSync()
      assert(result.isLeft, "oversized formula was unexpectedly accepted")
      assert(result.swap.toOption.exists(_.getMessage.contains("8192")))
      assert(!Files.exists(output))
    finally
      Files.deleteIfExists(output)
      Files.deleteIfExists(dir)
  }

  test("GH-693: batch rejects every formula shape before write or dry-run") {
    val shapes = List(
      ujson.Obj("op" -> "putf", "ref" -> "A1", "value" -> oversized),
      ujson.Obj("op" -> "putf", "ref" -> "A1:A2", "value" -> oversized, "from" -> "A1"),
      ujson.Obj("op" -> "putf", "ref" -> "A1:A2", "values" -> ujson.Arr("1", oversized)),
      ujson.Obj("op" -> "putf", "ref" -> "A1", "value" -> oversized, "array" -> true)
    )
    shapes.foreach { op =>
      val result = BatchParser.parseBatchJson(ujson.Arr(op).render())
      assert(result.isLeft, "oversized formula was unexpectedly accepted")
    }
  }

  test("GH-693: storage length boundary and already-prefixed input are exact") {
    val base = formula(650)
    val padding = 8192 - FormulaStorage.toStored(base).length
    assert(padding >= 0)
    val boundary = base.replace("SUM(", "SUM(" + " " * padding)
    assertEquals(FormulaStorage.toStored(boundary).length, 8192)
    assertEquals(ValueParser.storedFormulaLengthError(boundary), None)
    assert(ValueParser.storedFormulaLengthError(boundary.replace("SUM(", "SUM( ")).isDefined)
    assertEquals(ValueParser.storedFormulaLengthError("=SUM(1,2)"), None)
    assertEquals(ValueParser.storedFormulaLengthError(FormulaStorage.toStored(base)), None)
  }

  test("GH-693: streaming putf refuses storage expansion without touching input or output") {
    val dir = Files.createTempDirectory("stored-formula-stream-")
    val source = dir.resolve("input.xlsx")
    val output = dir.resolve("out.xlsx")
    try
      val wb = Workbook(Sheet("Test"))
      com.tjclp.xl.io.ExcelIO.instance[IO].write(wb, source).unsafeRunSync()
      val before = Files.readAllBytes(source).toVector
      val result = com.tjclp.xl.cli.commands.StreamingWriteCommands
        .putFormula(source, output, Some("Test"), "A1", List(oversized))
        .attempt
        .unsafeRunSync()
      assert(result.isLeft)
      assert(result.swap.toOption.exists(_.getMessage.contains("8192")))
      assert(!Files.exists(output))
      assertEquals(Files.readAllBytes(source).toVector, before)
    finally
      Files.deleteIfExists(output)
      Files.deleteIfExists(source)
      Files.deleteIfExists(dir)
  }
