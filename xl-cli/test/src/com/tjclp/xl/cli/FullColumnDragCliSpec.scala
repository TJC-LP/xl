package com.tjclp.xl.cli

import java.nio.file.{Files, Path}
import java.util.zip.ZipFile

import cats.effect.{IO, unsafe}
import com.tjclp.xl.{*, given}
import com.tjclp.xl.cli.commands.{StreamingWriteCommands, WriteCommands}
import com.tjclp.xl.cli.helpers.BatchParser
import com.tjclp.xl.cli.helpers.BatchParser.BatchOp
import com.tjclp.xl.io.ExcelIO
import com.tjclp.xl.ooxml.writer.WriterConfig

import munit.FunSuite

/**
 * GH-612: the CLI's three formula-dragging paths (`putf` over a range, batch `putf … from`, and the
 * streaming batch) keep whole-column / whole-row references in their Excel form. Before the fix
 * `putf Z1:Z3 '=COUNTIF($A:$A,B1)'` wrote `$A2:$A1048577` — a row that does not exist.
 */
class FullColumnDragCliSpec extends FunSuite:
  given unsafe.IORuntime = unsafe.IORuntime.global
  private val config = WriterConfig.default

  private def num(value: Int): CellValue = CellValue.Number(BigDecimal(value))

  /** A1:A6 holds 1,2,2,3,3,3; B1:B3 holds 1,2,3 — COUNTIF($A:$A,Bn) is 1, 2, 3. */
  private def fixture: Workbook =
    Workbook(
      Sheet("Data")
        .put(ref"A1", num(1))
        .put(ref"A2", num(2))
        .put(ref"A3", num(2))
        .put(ref"A4", num(3))
        .put(ref"A5", num(3))
        .put(ref"A6", num(3))
        .put(ref"B1", num(1))
        .put(ref"B2", num(2))
        .put(ref"B3", num(3))
    )

  private def withTemp(test: Path => Unit): Unit =
    val path = Files.createTempFile("full-column-drag", ".xlsx")
    try test(path)
    finally Files.deleteIfExists(path)

  private def read(path: Path): Workbook = ExcelIO.instance[IO].read(path).unsafeRunSync()

  private def formulaAt(wb: Workbook, ref: ARef): (String, Option[CellValue]) =
    wb.sheets.headOption.map(_(ref).value) match
      case Some(CellValue.Formula(text, cached, _)) => (text, cached)
      case other => fail(s"Expected a formula at ${ref.toA1}, got $other")

  private def sheetXml(path: Path): String =
    val zip = new ZipFile(path.toFile)
    try
      val entry = zip.getEntry("xl/worksheets/sheet1.xml")
      new String(zip.getInputStream(entry).readAllBytes(), "UTF-8")
    finally zip.close()

  test("putf over a range drags COUNTIF($A:$A,B1) along rows without touching the column") {
    withTemp { out =>
      WriteCommands
        .putFormula(
          fixture,
          fixture.sheets.headOption,
          "Z1:Z3",
          List("=COUNTIF($A:$A,B1)"),
          out,
          config
        )
        .unsafeRunSync()
      val wb = read(out)
      assertEquals(formulaAt(wb, ref"Z1"), ("COUNTIF($A:$A,B1)", Some(num(1))))
      assertEquals(formulaAt(wb, ref"Z2"), ("COUNTIF($A:$A,B2)", Some(num(2))))
      assertEquals(formulaAt(wb, ref"Z3"), ("COUNTIF($A:$A,B3)", Some(num(3))))
      // the <f> text in the package is the Excel form, not a corner range
      val xml = sheetXml(out)
      assert(xml.contains("<f>COUNTIF($A:$A,B2)</f>"), xml)
      assert(!xml.contains("1048577"), xml)
    }
  }

  test("putf over a range drags SUM(1:1) along columns without touching the row") {
    withTemp { out =>
      WriteCommands
        .putFormula(fixture, fixture.sheets.headOption, "Z5:AA5", List("=SUM(1:1)"), out, config)
        .unsafeRunSync()
      val wb = read(out)
      assertEquals(formulaAt(wb, ref"Z5")._1, "SUM(1:1)")
      assertEquals(formulaAt(wb, ref"AA5")._1, "SUM(1:1)")
      assert(!sheetXml(out).contains("XFE"))
    }
  }

  test("batch putf with 'from' keeps the whole-column form") {
    val ops = Vector(BatchOp.PutFormulaDragging("Z1:Z3", "=COUNTIF($A:$A,B1)", "Z1"))
    val wb =
      BatchParser.applyBatchOperations(fixture, fixture.sheets.headOption, ops).unsafeRunSync()
    assertEquals(formulaAt(wb, ref"Z1"), ("COUNTIF($A:$A,B1)", Some(num(1))))
    assertEquals(formulaAt(wb, ref"Z2"), ("COUNTIF($A:$A,B2)", Some(num(2))))
    assertEquals(formulaAt(wb, ref"Z3"), ("COUNTIF($A:$A,B3)", Some(num(3))))
  }

  test("batch putf with 'from' writes #REF! for a reference dragged before row 1") {
    // from Z3, the Z1 copy shifts B3 up two rows to B1 — fine — but A2 up two rows falls off
    val ops = Vector(BatchOp.PutFormulaDragging("Z1:Z3", "=B3+A2", "Z3"))
    val wb =
      BatchParser.applyBatchOperations(fixture, fixture.sheets.headOption, ops).unsafeRunSync()
    assertEquals(formulaAt(wb, ref"Z1"), ("B1+#REF!", Some(CellValue.Error(CellError.Ref))))
    assertEquals(formulaAt(wb, ref"Z2")._1, "B2+A1")
    assertEquals(formulaAt(wb, ref"Z3")._1, "B3+A2")
  }

  test("streaming batch putf with 'from' keeps the whole-column form") {
    withTemp { source =>
      withTemp { out =>
        ExcelIO.instance[IO].write(fixture, source).unsafeRunSync()
        val json = Files.createTempFile("full-column-drag", ".json")
        try
          Files.writeString(
            json,
            """[{"op":"putf","ref":"Z1:Z3","value":"=COUNTIF($A:$A,B1)","from":"Z1"}]"""
          )
          StreamingWriteCommands.batch(source, out, Some("Data"), json.toString).unsafeRunSync()
          val wb = read(out)
          assertEquals(formulaAt(wb, ref"Z1")._1, "COUNTIF($A:$A,B1)")
          assertEquals(formulaAt(wb, ref"Z2")._1, "COUNTIF($A:$A,B2)")
          assertEquals(formulaAt(wb, ref"Z3")._1, "COUNTIF($A:$A,B3)")
        finally Files.deleteIfExists(json)
      }
    }
  }
