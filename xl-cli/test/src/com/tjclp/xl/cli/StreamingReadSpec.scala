package com.tjclp.xl.cli

import java.nio.file.{Files, Path}

import cats.effect.IO
import munit.CatsEffectSuite

import com.tjclp.xl.{Workbook, Sheet, given}
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.cli.commands.{ReadCommands, StreamingReadCommands}
import com.tjclp.xl.io.ExcelIO
import com.tjclp.xl.macros.ref

class StreamingReadSpec extends CatsEffectSuite:

  private def withTempWorkbook[A](wb: Workbook)(test: (Workbook, Path) => IO[A]): IO[A] =
    IO.blocking {
      val tempFile = Files.createTempFile("xl-cli-stream-read-", ".xlsx")
      tempFile.toFile.deleteOnExit()
      tempFile
    }.flatMap { tempFile =>
      ExcelIO.instance[IO].write(wb, tempFile) *> test(wb, tempFile)
    }

  test("streaming view csv matches in-memory output without extra trailing newline") {
    val sheet = Sheet("Test")
      .put(ref"A1", CellValue.Text("Hello"))
      .put(ref"A2", CellValue.Number(BigDecimal(42)))
      .put(ref"A3", CellValue.Text("World"))
    val wb = Workbook(Vector(sheet))

    withTempWorkbook(wb) { case (workbook, path) =>
      val sheetOpt = workbook.sheets.headOption
      for
        inMemory <- ReadCommands.view(
          workbook,
          sheetOpt,
          "A1:A3",
          showFormulas = false,
          evalFormulas = false,
          strict = false,
          limit = 100,
          format = ViewFormat.Csv,
          printScale = false,
          showGridlines = false,
          showLabels = false,
          dpi = 96,
          quality = 90,
          rasterOutput = None,
          skipEmpty = false,
          headerRow = None
        )
        streaming <- StreamingReadCommands.view(
          path,
          Some("Test"),
          "A1:A3",
          showFormulas = false,
          limit = 100,
          format = ViewFormat.Csv,
          showLabels = false,
          skipEmpty = false,
          headerRow = None
        )
      yield
        assertEquals(streaming, inMemory)
        assert(!streaming.endsWith("\n"), s"streaming CSV should not end with newline: '$streaming'")
    }
  }
