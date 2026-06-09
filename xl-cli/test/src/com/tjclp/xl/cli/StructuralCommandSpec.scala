package com.tjclp.xl.cli

import java.nio.file.Files

import cats.effect.IO
import munit.CatsEffectSuite

import com.tjclp.xl.{*, given}
import com.tjclp.xl.cells.{CellError, CellValue}
import com.tjclp.xl.cli.commands.WriteCommands
import com.tjclp.xl.io.ExcelIO
import com.tjclp.xl.macros.ref
import com.tjclp.xl.ooxml.writer.WriterConfig

/**
 * Integration tests for the `insert-rows` / `delete-rows` / `insert-cols` / `delete-cols` CLI
 * commands. These round-trip through a real `.xlsx` so they exercise the full path INCLUDING the
 * SourceContext interaction — i.e. they regress the bug where a structural edit on a freshly-read
 * (clean) workbook was silently dropped by the writer's verbatim-copy fast-path.
 */
@SuppressWarnings(Array("org.wartremover.warts.IterableOps"))
class StructuralCommandSpec extends CatsEffectSuite:

  private val excel = ExcelIO.instance[IO]
  private val config = WriterConfig.default

  private def tmp(tag: String) =
    val p = Files.createTempFile(s"struct-$tag-", ".xlsx")
    p.toFile.deleteOnExit()
    p

  test("insert-rows: read -> shift cells & rewrite formulas -> write (no stale verbatim copy)") {
    val wb = Workbook(
      Vector(
        Sheet("S")
          .put(ref"A1", CellValue.Number(10))
          .put(ref"A3", CellValue.Number(30))
          .put(ref"B1", CellValue.Formula("=A1+A3", None))
      )
    )
    val in = tmp("in")
    val out = tmp("out")
    for
      _ <- excel.write(wb, in)
      read <- excel.read(in) // sourceContext present & clean
      _ <- WriteCommands.insertRows(read, read.sheets.headOption, 2, 1, out, config)
      result <- excel.read(out)
    yield
      val s = result.sheets.head
      assertEquals(s(ref"B1").value, CellValue.Formula("=A1+A4", None)) // A3 -> A4
      assertEquals(s(ref"A4").value, CellValue.Number(30)) // cell shifted down
      assertEquals(s(ref"A1").value, CellValue.Number(10)) // unchanged
  }

  test("delete-rows: a reference into the deleted row becomes #REF!") {
    val wb = Workbook(
      Vector(
        Sheet("S")
          .put(ref"A4", CellValue.Number(40))
          .put(ref"B1", CellValue.Formula("=A4", None))
      )
    )
    val in = tmp("in2")
    val out = tmp("out2")
    for
      _ <- excel.write(wb, in)
      read <- excel.read(in)
      _ <- WriteCommands.deleteRows(read, read.sheets.headOption, 4, 1, out, config)
      result <- excel.read(out)
    yield assertEquals(result.sheets.head(ref"B1").value, CellValue.Error(CellError.Ref))
  }

  test("insert-cols: column references at/after the insertion point shift right") {
    val wb = Workbook(
      Vector(
        Sheet("S")
          .put(ref"A1", CellValue.Number(1))
          .put(ref"A2", CellValue.Formula("=C1", None))
      )
    )
    val in = tmp("in3")
    val out = tmp("out3")
    for
      _ <- excel.write(wb, in)
      read <- excel.read(in)
      _ <- WriteCommands.insertColumns(read, read.sheets.headOption, "B", 1, out, config)
      result <- excel.read(out)
    yield assertEquals(result.sheets.head(ref"A2").value, CellValue.Formula("=D1", None)) // C1 -> D1
  }

  test("delete-cols: range form C:E removes the whole span (GH-129)") {
    val wb = Workbook(
      Vector(
        Sheet("S")
          .put(ref"A1", CellValue.Number(1))
          .put(ref"B1", CellValue.Number(2))
          .put(ref"C1", CellValue.Number(3))
          .put(ref"D1", CellValue.Number(4))
          .put(ref"E1", CellValue.Number(5))
          .put(ref"F1", CellValue.Number(6))
      )
    )
    val in = tmp("in4")
    val out = tmp("out4")
    for
      _ <- excel.write(wb, in)
      read <- excel.read(in)
      // count arg is ignored when a range is given; C:E deletes 3 columns
      _ <- WriteCommands.deleteColumns(read, read.sheets.headOption, "C:E", 1, out, config)
      result <- excel.read(out)
    yield
      val s = result.sheets.head
      assertEquals(s(ref"A1").value, CellValue.Number(1)) // unchanged
      assertEquals(s(ref"B1").value, CellValue.Number(2)) // unchanged
      assertEquals(s(ref"C1").value, CellValue.Number(6)) // F shifted left into C
      assert(!s.contains(ref"D1")) // only 3 columns remain
  }

  test("delete-cols: reversed range E:C normalizes to the same span") {
    val wb = Workbook(
      Vector(
        Sheet("S")
          .put(ref"A1", CellValue.Number(1))
          .put(ref"B1", CellValue.Number(2))
          .put(ref"C1", CellValue.Number(3))
          .put(ref"D1", CellValue.Number(4))
          .put(ref"E1", CellValue.Number(5))
          .put(ref"F1", CellValue.Number(6))
      )
    )
    val in = tmp("in5")
    val out = tmp("out5")
    for
      _ <- excel.write(wb, in)
      read <- excel.read(in)
      _ <- WriteCommands.deleteColumns(read, read.sheets.headOption, "E:C", 1, out, config)
      result <- excel.read(out)
    yield
      val s = result.sheets.head
      assertEquals(s(ref"C1").value, CellValue.Number(6)) // identical to C:E
      assert(!s.contains(ref"D1"))
  }

  test("delete-cols: malformed ranges are rejected (total Left, never crash)") {
    val wb = Workbook(Vector(Sheet("S").put(ref"A1", CellValue.Number(1))))
    val in = tmp("in6")
    val out = tmp("out6")
    for
      _ <- excel.write(wb, in)
      read <- excel.read(in)
      r1 <- WriteCommands.deleteColumns(read, read.sheets.headOption, "C:E:F", 1, out, config).attempt
      r2 <- WriteCommands.deleteColumns(read, read.sheets.headOption, "C:", 1, out, config).attempt
    yield
      assert(r1.isLeft) // three-part range
      assert(r2.isLeft) // trailing colon (empty endpoint)
  }
