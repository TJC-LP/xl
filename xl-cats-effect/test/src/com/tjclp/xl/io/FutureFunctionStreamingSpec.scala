package com.tjclp.xl.io

import cats.effect.IO
import munit.CatsEffectSuite

import java.nio.file.{Files, Path}

import com.tjclp.xl.{*, given}
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.macros.ref
import com.tjclp.xl.ooxml.{XlsxReader, XlsxWriter}

/**
 * GH-556: the streaming paths agree with the DOM path on the `_xlfn.` storage form — the SAX reader
 * hands the model the bare formula-bar spelling, and the streaming writer emits the prefix.
 */
class FutureFunctionStreamingSpec extends CatsEffectSuite:

  val tempDir: FunFixture[Path] = FunFixture[Path](
    setup = _ => Files.createTempDirectory("xl-xlfn-stream-"),
    teardown = dir =>
      Files
        .walk(dir)
        .sorted(java.util.Comparator.reverseOrder())
        .forEach(Files.delete)
  )

  private val maxifs = CellValue.Formula("MAXIFS(B1:B3,A1:A3,C1)", Some(CellValue.Number(1)))

  private def workbook: Workbook =
    Workbook(
      Vector(
        Sheet("Data")
          .put(ref"A1", CellValue.Text("a"))
          .put(ref"B1", CellValue.Number(1))
          .put(ref"C1", CellValue.Text("a"))
          .put(ref"D1", maxifs)
      )
    )

  private def formulaAt(cells: Map[Int, CellValue], col: Int): String =
    cells.get(col) match
      case Some(CellValue.Formula(expr, _, _)) => expr
      case other => fail(s"expected a formula in column $col, got $other")

  tempDir.test("GH-556: readStream strips _xlfn. like the DOM reader") { dir =>
    val path = dir.resolve("xlfn.xlsx")
    IO(XlsxWriter.write(workbook, path)).flatMap { _ =>
      ExcelIO.instance[IO].readStream(path).compile.toList.map { rows =>
        val row = rows.headOption.getOrElse(fail("expected one streamed row"))
        assertEquals(formulaAt(row.cells, 3), "MAXIFS(B1:B3,A1:A3,C1)")
        // DOM parity
        val dom = XlsxReader.read(path).fold(err => fail(s"read failed: $err"), identity)
        dom.sheets(0)(ref"D1").value match
          case CellValue.Formula(expr, _, _) => assertEquals(expr, "MAXIFS(B1:B3,A1:A3,C1)")
          case other => fail(s"expected formula at D1, got $other")
      }
    }
  }

  tempDir.test("GH-556: the streaming writer emits the _xlfn. prefix") { dir =>
    val path = dir.resolve("xlfn-stream-write.xlsx")
    val excel = ExcelIO.instance[IO]
    val rows = fs2.Stream.emit(
      RowData(
        1,
        Map(
          0 -> CellValue.Text("a"),
          1 -> CellValue.Number(1),
          2 -> CellValue.Text("a"),
          3 -> maxifs
        )
      )
    )
    rows.through(excel.writeStream(path, "Data")).compile.drain.flatMap { _ =>
      IO {
        val zip = new java.util.zip.ZipFile(path.toFile)
        try
          val entry = Option(zip.getEntry("xl/worksheets/sheet1.xml"))
            .getOrElse(fail("sheet1.xml missing"))
          val xml = new String(zip.getInputStream(entry).readAllBytes(), "UTF-8")
          assert(xml.contains("_xlfn.MAXIFS(B1:B3,A1:A3,C1)"), xml)
        finally zip.close()
      }
    }
  }
