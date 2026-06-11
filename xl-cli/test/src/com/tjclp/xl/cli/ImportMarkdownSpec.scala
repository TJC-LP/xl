package com.tjclp.xl.cli

import java.nio.file.{Files, Path}

import cats.effect.IO
import munit.CatsEffectSuite

import com.tjclp.xl.{Sheet, Workbook, given}
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.cli.commands.ImportCommands
import com.tjclp.xl.io.ExcelIO
import com.tjclp.xl.macros.ref
import com.tjclp.xl.ooxml.writer.WriterConfig
import com.tjclp.xl.styles.alignment.HAlign
import com.tjclp.xl.styles.numfmt.NumFmt

/**
 * Integration tests for the import-md command (GH-159).
 *
 * Covers: positioned import, --new-sheet, smart type detection (numbers, currency, percent,
 * dates), --no-type-inference opt-out, --skip-header, and GFM alignment markers mapping to cell
 * alignment.
 */
@SuppressWarnings(Array("org.wartremover.warts.OptionPartial"))
class ImportMarkdownSpec extends CatsEffectSuite:

  private val config: WriterConfig = WriterConfig.default

  private def withOutput[A](test: Path => IO[A]): IO[A] =
    IO.blocking {
      val p = Files.createTempFile("xl-import-md", ".xlsx")
      p.toFile.deleteOnExit()
      p
    }.flatMap(test)

  private def readBack(path: Path): IO[Workbook] =
    ExcelIO.instance[IO].read(path)

  private def importMd(
    wb: Workbook,
    sheetOpt: Option[Sheet],
    content: String,
    startRef: Option[String] = None,
    skipHeader: Boolean = false,
    newSheet: Option[String] = None,
    noTypeInference: Boolean = false,
    outputPath: Path
  ): IO[String] =
    ImportCommands.importMarkdownContent(
      wb,
      sheetOpt,
      content,
      "table.md",
      startRef,
      skipHeader,
      newSheet,
      noTypeInference,
      outputPath,
      config,
      stream = false
    )

  private val basicTable = Seq(
    "| Item | Qty |",
    "|------|-----|",
    "| Apple | 3 |",
    "| Pear | 5 |"
  ).mkString("\n")

  test("import-md: imports header and body at A1 by default with typed values") {
    withOutput { out =>
      val sheet = Sheet("Data")
      val wb = Workbook(Vector(sheet))
      for
        msg <- importMd(wb, Some(sheet), basicTable, outputPath = out)
        imported <- readBack(out)
        s = imported.sheets.head
      yield
        assert(msg.contains("table.md"), s"Message should mention source: $msg")
        assert(msg.contains("3 rows"), s"Message should report rows: $msg")
        assertEquals(s.cells.get(ref"A1").map(_.value), Some(CellValue.Text("Item")))
        assertEquals(s.cells.get(ref"B1").map(_.value), Some(CellValue.Text("Qty")))
        assertEquals(s.cells.get(ref"A2").map(_.value), Some(CellValue.Text("Apple")))
        assertEquals(s.cells.get(ref"B2").map(_.value), Some(CellValue.Number(BigDecimal(3))))
        assertEquals(s.cells.get(ref"B3").map(_.value), Some(CellValue.Number(BigDecimal(5))))
    }
  }

  test("import-md: --start offsets the table position") {
    withOutput { out =>
      val sheet = Sheet("Data")
      val wb = Workbook(Vector(sheet))
      for
        _ <- importMd(wb, Some(sheet), basicTable, startRef = Some("C5"), outputPath = out)
        imported <- readBack(out)
        s = imported.sheets.head
      yield
        assertEquals(s.cells.get(ref"C5").map(_.value), Some(CellValue.Text("Item")))
        assertEquals(s.cells.get(ref"D6").map(_.value), Some(CellValue.Number(BigDecimal(3))))
        assertEquals(s.cells.get(ref"A1"), None)
    }
  }

  test("import-md: --skip-header drops the header row") {
    withOutput { out =>
      val sheet = Sheet("Data")
      val wb = Workbook(Vector(sheet))
      for
        _ <- importMd(wb, Some(sheet), basicTable, skipHeader = true, outputPath = out)
        imported <- readBack(out)
        s = imported.sheets.head
      yield
        assertEquals(s.cells.get(ref"A1").map(_.value), Some(CellValue.Text("Apple")))
        assertEquals(s.cells.get(ref"B2").map(_.value), Some(CellValue.Number(BigDecimal(5))))
        assertEquals(s.cells.get(ref"A3"), None)
    }
  }

  test("import-md: --new-sheet creates a new sheet") {
    withOutput { out =>
      val existing = Sheet("Existing")
      val wb = Workbook(Vector(existing))
      for
        msg <- importMd(wb, None, basicTable, newSheet = Some("Imported"), outputPath = out)
        imported <- readBack(out)
      yield
        assert(msg.contains("Imported"), s"Message should mention new sheet: $msg")
        assertEquals(imported.sheets.map(_.name.value).toSet, Set("Existing", "Imported"))
        val s = imported.sheets.find(_.name.value == "Imported").get
        assertEquals(s.cells.get(ref"A1").map(_.value), Some(CellValue.Text("Item")))
    }
  }

  test("import-md: --new-sheet rejects existing sheet name") {
    withOutput { out =>
      val existing = Sheet("Data")
      val wb = Workbook(Vector(existing))
      importMd(wb, None, basicTable, newSheet = Some("Data"), outputPath = out).attempt.map {
        case Left(err) =>
          assert(err.getMessage.contains("already exists"), s"Got: ${err.getMessage}")
        case Right(_) => fail("Duplicate sheet name should fail")
      }
    }
  }

  test("import-md: smart detection applies currency, percent, and date formats") {
    val table = Seq(
      "| Price | Margin | Date |",
      "|-------|--------|------|",
      "| $1,234.56 | 45.5% | 2025-01-15 |"
    ).mkString("\n")
    withOutput { out =>
      val sheet = Sheet("Data")
      val wb = Workbook(Vector(sheet))
      for
        _ <- importMd(wb, Some(sheet), table, outputPath = out)
        imported <- readBack(out)
        s = imported.sheets.head
      yield
        assertEquals(s.cells.get(ref"A2").map(_.value), Some(CellValue.Number(BigDecimal("1234.56"))))
        val a2Fmt = s.cells.get(ref"A2").flatMap(_.styleId).flatMap(s.styleRegistry.get).map(_.numFmt)
        assertEquals(a2Fmt, Some(NumFmt.Currency))
        assertEquals(s.cells.get(ref"B2").map(_.value), Some(CellValue.Number(BigDecimal("0.455"))))
        val b2Fmt = s.cells.get(ref"B2").flatMap(_.styleId).flatMap(s.styleRegistry.get).map(_.numFmt)
        assertEquals(b2Fmt, Some(NumFmt.Percent))
        // Dates serialize as Excel serial numbers; the Date numFmt carries the semantics
        val c2Fmt = s.cells.get(ref"C2").flatMap(_.styleId).flatMap(s.styleRegistry.get).map(_.numFmt)
        assertEquals(c2Fmt, Some(NumFmt.Date))
        assert(
          s.cells.get(ref"C2").map(_.value).exists {
            case CellValue.DateTime(_) | CellValue.Number(_) => true
            case _ => false
          },
          s"C2 should be a date-valued cell, got ${s.cells.get(ref"C2").map(_.value)}"
        )
    }
  }

  test("import-md: --no-type-inference keeps every value as text") {
    val table = Seq(
      "| Price | Qty |",
      "|-------|-----|",
      "| $1,234.56 | 3 |"
    ).mkString("\n")
    withOutput { out =>
      val sheet = Sheet("Data")
      val wb = Workbook(Vector(sheet))
      for
        _ <- importMd(wb, Some(sheet), table, noTypeInference = true, outputPath = out)
        imported <- readBack(out)
        s = imported.sheets.head
      yield
        assertEquals(s.cells.get(ref"A2").map(_.value), Some(CellValue.Text("$1,234.56")))
        assertEquals(s.cells.get(ref"B2").map(_.value), Some(CellValue.Text("3")))
    }
  }

  test("import-md: alignment markers map to cell horizontal alignment") {
    val table = Seq(
      "| L | C | R |",
      "|:---|:---:|---:|",
      "| a | b | c |"
    ).mkString("\n")
    withOutput { out =>
      val sheet = Sheet("Data")
      val wb = Workbook(Vector(sheet))
      for
        _ <- importMd(wb, Some(sheet), table, outputPath = out)
        imported <- readBack(out)
        s = imported.sheets.head
      yield
        def halign(r: com.tjclp.xl.addressing.ARef) =
          s.cells.get(r).flatMap(_.styleId).flatMap(s.styleRegistry.get).map(_.align.horizontal)
        assertEquals(halign(ref"A2"), Some(HAlign.Left))
        assertEquals(halign(ref"B2"), Some(HAlign.Center))
        assertEquals(halign(ref"C2"), Some(HAlign.Right))
    }
  }

  test("import-md: import to existing sheet without --sheet fails with guidance") {
    withOutput { out =>
      val wb = Workbook(Vector(Sheet("Data")))
      importMd(wb, None, basicTable, outputPath = out).attempt.map {
        case Left(err) =>
          assert(err.getMessage.contains("--sheet"), s"Got: ${err.getMessage}")
        case Right(_) => fail("Missing sheet should fail")
      }
    }
  }

  test("import-md: input without a table fails with a clear error") {
    withOutput { out =>
      val sheet = Sheet("Data")
      val wb = Workbook(Vector(sheet))
      importMd(wb, Some(sheet), "no table here", outputPath = out).attempt.map {
        case Left(err) =>
          assert(
            err.getMessage.toLowerCase.contains("table"),
            s"Error should mention table: ${err.getMessage}"
          )
        case Right(_) => fail("Missing table should fail")
      }
    }
  }

  test("import-md: importMarkdown reads from a file path") {
    withOutput { out =>
      for
        mdFile <- IO.blocking {
          val p = Files.createTempFile("xl-table", ".md")
          p.toFile.deleteOnExit()
          Files.writeString(p, basicTable)
          p
        }
        sheet = Sheet("Data")
        wb = Workbook(Vector(sheet))
        _ <- ImportCommands.importMarkdown(
          wb,
          Some(sheet),
          mdFile.toString,
          None,
          false,
          None,
          false,
          out,
          config,
          stream = false
        )
        imported <- readBack(out)
        s = imported.sheets.head
      yield assertEquals(s.cells.get(ref"A1").map(_.value), Some(CellValue.Text("Item")))
    }
  }
