package com.tjclp.xl.cli

import munit.FunSuite

import java.nio.file.{Files, Path}

import cats.effect.{IO, unsafe}
import com.tjclp.xl.{Sheet, Workbook}
import com.tjclp.xl.addressing.{Column, Row}
import com.tjclp.xl.cli.commands.WriteCommands
import com.tjclp.xl.cli.helpers.BatchParser
import com.tjclp.xl.io.ExcelIO
import com.tjclp.xl.ooxml.writer.WriterConfig

/**
 * E2E tests for row/column outline grouping (GH-421): `group-rows` / `group-cols` / `ungroup-rows`
 * / `ungroup-cols` commands and their batch twins, writing through the existing
 * `RowProperties`/`ColumnProperties` outlineLevel + collapsed fields.
 */
@SuppressWarnings(
  Array("org.wartremover.warts.OptionPartial", "org.wartremover.warts.IterableOps")
)
class GroupingCommandSpec extends FunSuite:

  given unsafe.IORuntime = unsafe.IORuntime.global

  val config: WriterConfig = WriterConfig.default

  private def withTempOutput[A](f: Path => A): A =
    val out = Files.createTempFile("grouping-test", ".xlsx")
    try f(out)
    finally Files.deleteIfExists(out)

  private def readBack(path: Path): Workbook =
    ExcelIO.instance[IO].read(path).unsafeRunSync()

  private def sheetXml(path: Path): String =
    val zip = new java.util.zip.ZipFile(path.toFile)
    try
      val entry = zip.getEntry("xl/worksheets/sheet1.xml")
      new String(
        zip.getInputStream(entry).readAllBytes(),
        java.nio.charset.StandardCharsets.UTF_8
      )
    finally zip.close()

  // ========== group-rows / ungroup-rows commands ==========

  test("group-rows 10:12 writes outlineLevel attrs and round-trips") {
    withTempOutput { out =>
      val sheet = Sheet("Data")
      val wb = Workbook(sheet)
      val result = WriteCommands
        .groupRows(wb, Some(sheet), "10:12", level = 1, collapsed = false, out, config)
        .unsafeRunSync()

      assert(result.contains("Grouped rows 10:12"), result)

      val xml = sheetXml(out)
      (10 to 12).foreach { r =>
        assert(
          xml.contains(s"""<row outlineLevel="1" r="$r"/>""") ||
            xml.contains(s"""<row r="$r" outlineLevel="1"/>"""),
          s"row $r should carry outlineLevel in: $xml"
        )
      }

      val props = readBack(out).sheets.head.rowProperties
      (10 to 12).foreach { r =>
        assertEquals(props.get(Row.from1(r)).flatMap(_.outlineLevel), Some(1))
      }
    }
  }

  test("group-rows --collapsed hides members and marks the summary row collapsed") {
    withTempOutput { out =>
      val sheet = Sheet("Data")
      val wb = Workbook(sheet)
      WriteCommands
        .groupRows(wb, Some(sheet), "5:7", level = 2, collapsed = true, out, config)
        .unsafeRunSync()

      val xml = sheetXml(out)
      assert(xml.contains("collapsed=\"1\""), s"summary row should be collapsed in: $xml")

      val props = readBack(out).sheets.head.rowProperties
      (5 to 7).foreach { r =>
        val p = props.get(Row.from1(r))
        assertEquals(p.flatMap(_.outlineLevel), Some(2), s"row $r outline level")
        assert(p.exists(_.hidden), s"row $r should be hidden")
      }
      assert(props.get(Row.from1(8)).exists(_.collapsed), "row 8 carries the collapse marker")
      assertEquals(props.get(Row.from1(8)).flatMap(_.outlineLevel), None)
    }
  }

  test("ungroup-rows clears outlineLevel and collapse markers from the file") {
    withTempOutput { grouped =>
      withTempOutput { out =>
        val sheet = Sheet("Data")
        val wb = Workbook(sheet)
        WriteCommands
          .groupRows(wb, Some(sheet), "5:7", level = 1, collapsed = true, grouped, config)
          .unsafeRunSync()

        val wb1 = readBack(grouped)
        val result = WriteCommands
          .ungroupRows(wb1, Some(wb1.sheets.head), "5:7", out, config)
          .unsafeRunSync()

        assert(result.contains("Ungrouped rows 5:7"), result)
        val xml = sheetXml(out)
        assert(!xml.contains("outlineLevel"), s"no outlineLevel should remain in: $xml")
        assert(!xml.contains("collapsed"), s"no collapse marker should remain in: $xml")

        val props = readBack(out).sheets.head.rowProperties
        (5 to 7).foreach { r =>
          assertEquals(props.get(Row.from1(r)).flatMap(_.outlineLevel), None)
          // Collapse-hidden members stay hidden (Excel semantics) — unhide via row --show.
          assert(props.get(Row.from1(r)).exists(_.hidden), s"row $r stays hidden after ungroup")
        }
        assert(!props.get(Row.from1(8)).exists(_.collapsed), "summary marker cleared")
      }
    }
  }

  test("group-rows accepts a single row spec") {
    withTempOutput { out =>
      val sheet = Sheet("Data")
      val wb = Workbook(sheet)
      WriteCommands
        .groupRows(wb, Some(sheet), "10", level = 1, collapsed = false, out, config)
        .unsafeRunSync()

      val props = readBack(out).sheets.head.rowProperties
      assertEquals(props.get(Row.from1(10)).flatMap(_.outlineLevel), Some(1))
      assertEquals(props.get(Row.from1(11)), None)
    }
  }

  test("group-rows preserves existing row properties (height survives grouping)") {
    withTempOutput { out =>
      val sheet = Sheet("Data").setRowProperties(
        Row.from1(10),
        com.tjclp.xl.sheets.RowProperties(height = Some(30.0))
      )
      val wb = Workbook(sheet)
      WriteCommands
        .groupRows(wb, Some(sheet), "10:11", level = 1, collapsed = false, out, config)
        .unsafeRunSync()

      val props = readBack(out).sheets.head.rowProperties
      assertEquals(props.get(Row.from1(10)).flatMap(_.height), Some(30.0))
      assertEquals(props.get(Row.from1(10)).flatMap(_.outlineLevel), Some(1))
    }
  }

  test("group-rows: level out of range yields clean error") {
    withTempOutput { out =>
      val wb = Workbook(Sheet("Data"))
      val result = WriteCommands
        .groupRows(wb, Some(wb.sheets.head), "1:3", level = 8, collapsed = false, out, config)
        .attempt
        .unsafeRunSync()

      assert(result.isLeft)
      val msg = result.swap.toOption.get.getMessage
      assert(msg.contains("1-7"), msg)
      assert(msg.contains("8"), msg)
    }
  }

  test("group-rows: malformed spec yields clean error") {
    withTempOutput { out =>
      val wb = Workbook(Sheet("Data"))
      val result = WriteCommands
        .groupRows(wb, Some(wb.sheets.head), "abc", level = 1, collapsed = false, out, config)
        .attempt
        .unsafeRunSync()

      assert(result.isLeft)
      assert(
        result.swap.toOption.get.getMessage.contains("abc"),
        result.swap.toOption.get.getMessage
      )
    }
  }

  // ========== group-cols / ungroup-cols commands ==========

  test("group-cols E:H writes a span with outlineLevel and round-trips") {
    withTempOutput { out =>
      val sheet = Sheet("Data")
      val wb = Workbook(sheet)
      val result = WriteCommands
        .groupCols(wb, Some(sheet), "E:H", level = 1, collapsed = false, out, config)
        .unsafeRunSync()

      assert(result.contains("Grouped columns E:H"), result)

      val xml = sheetXml(out)
      // E..H with identical props merge into one span: min=5, max=8 (1-based).
      assert(
        xml.contains("min=\"5\"") && xml.contains("max=\"8\"") &&
          xml.contains("outlineLevel=\"1\""),
        xml
      )

      val props = readBack(out).sheets.head.columnProperties
      ('E' to 'H').foreach { c =>
        val col = Column.fromLetter(c.toString).toOption.get
        assertEquals(props.get(col).flatMap(_.outlineLevel), Some(1), s"col $c")
      }
    }
  }

  test("group-cols --collapsed hides members and marks the summary column") {
    withTempOutput { out =>
      val sheet = Sheet("Data")
      val wb = Workbook(sheet)
      WriteCommands
        .groupCols(wb, Some(sheet), "E:H", level = 1, collapsed = true, out, config)
        .unsafeRunSync()

      val props = readBack(out).sheets.head.columnProperties
      ('E' to 'H').foreach { c =>
        val col = Column.fromLetter(c.toString).toOption.get
        assert(props.get(col).exists(_.hidden), s"col $c should be hidden")
      }
      val summaryCol = Column.fromLetter("I").toOption.get
      assert(props.get(summaryCol).exists(_.collapsed), "col I carries the collapse marker")
    }
  }

  test("ungroup-cols clears grouping from a written file (no resurrection from preserved cols)") {
    withTempOutput { grouped =>
      withTempOutput { out =>
        val sheet = Sheet("Data")
        val wb = Workbook(sheet)
        WriteCommands
          .groupCols(wb, Some(sheet), "E:H", level = 1, collapsed = false, grouped, config)
          .unsafeRunSync()

        val wb1 = readBack(grouped)
        val result = WriteCommands
          .ungroupCols(wb1, Some(wb1.sheets.head), "E:H", out, config)
          .unsafeRunSync()

        assert(result.contains("Ungrouped columns E:H"), result)
        val xml = sheetXml(out)
        assert(!xml.contains("outlineLevel"), s"no outlineLevel should remain in: $xml")
        val props = readBack(out).sheets.head.columnProperties
        props.foreach { case (col, p) =>
          assertEquals(p.outlineLevel, None, s"col ${col.toLetter}")
          assert(!p.collapsed, s"col ${col.toLetter} should not be collapsed")
        }
      }
    }
  }

  // ========== batch ops: parse ==========

  test("batch: group-rows and group-cols parse without warnings") {
    val json =
      """[
        {"op": "group-rows", "rows": "10:20", "level": 2, "collapsed": true},
        {"op": "group-cols", "cols": "E:H"}
      ]"""
    val result = BatchParser.parseBatchJson(json)
    assert(result.isRight, s"Should parse: $result")
    assertEquals(result.toOption.get.ops.size, 2)
    assertEquals(result.toOption.get.warnings, Vector.empty[String])
  }

  test("batch: ungroup-rows and ungroup-cols parse without warnings") {
    val json =
      """[
        {"op": "ungroup-rows", "rows": "10:20"},
        {"op": "ungroup-cols", "cols": "E:H"}
      ]"""
    val result = BatchParser.parseBatchJson(json)
    assert(result.isRight, s"Should parse: $result")
    assertEquals(result.toOption.get.ops.size, 2)
    assertEquals(result.toOption.get.warnings, Vector.empty[String])
  }

  test("batch: group-rows unknown property warns") {
    val json = """[{"op": "group-rows", "rows": "1:3", "depth": 2}]"""
    val result = BatchParser.parseBatchJson(json)
    assert(result.isRight, s"Should parse: $result")
    assert(result.toOption.get.warnings.exists(_.contains("depth")), result.toString)
  }

  // ========== batch ops: apply ==========

  test("batch: group-rows sets outlineLevel on every row in the range") {
    val sheet = Sheet("Data")
    val wb = Workbook(sheet)
    val ops = BatchParser
      .parseBatchJson("""[{"op": "group-rows", "rows": "10:12", "level": 2}]""")
      .toOption
      .get
      .ops
    val updated = BatchParser.applyBatchOperations(wb, Some(sheet), ops).unsafeRunSync()
    val props = updated.sheets.head.rowProperties
    (10 to 12).foreach { r =>
      assertEquals(
        props.get(Row.from1(r)).flatMap(_.outlineLevel),
        Some(2),
        s"row $r should be at outline level 2"
      )
    }
    assertEquals(props.get(Row.from1(9)), None)
    assertEquals(props.get(Row.from1(13)), None)
  }

  test("batch: group-cols defaults to level 1") {
    val sheet = Sheet("Data")
    val wb = Workbook(sheet)
    val ops = BatchParser
      .parseBatchJson("""[{"op": "group-cols", "cols": "E:H"}]""")
      .toOption
      .get
      .ops
    val updated = BatchParser.applyBatchOperations(wb, Some(sheet), ops).unsafeRunSync()
    val props = updated.sheets.head.columnProperties
    ('E' to 'H').foreach { c =>
      assertEquals(
        Column.fromLetter(c.toString).toOption.flatMap(props.get).flatMap(_.outlineLevel),
        Some(1),
        s"col $c should be at outline level 1"
      )
    }
  }

  test("batch: ungroup-rows clears grouping state set by group-rows") {
    val sheet = Sheet("Data")
    val wb = Workbook(sheet)
    val group = BatchParser
      .parseBatchJson("""[{"op": "group-rows", "rows": "5:8", "collapsed": true}]""")
      .toOption
      .get
      .ops
    val ungroup = BatchParser
      .parseBatchJson("""[{"op": "ungroup-rows", "rows": "5:8"}]""")
      .toOption
      .get
      .ops
    val grouped = BatchParser.applyBatchOperations(wb, Some(sheet), group).unsafeRunSync()
    val cleared = BatchParser
      .applyBatchOperations(grouped, Some(grouped.sheets.head), ungroup)
      .unsafeRunSync()
    val props = cleared.sheets.head.rowProperties
    (5 to 8).foreach { r =>
      assertEquals(
        props.get(Row.from1(r)).flatMap(_.outlineLevel),
        None,
        s"row $r should have no outline level"
      )
      assert(!props.get(Row.from1(r)).exists(_.collapsed), s"row $r should not be collapsed")
    }
    assert(!props.get(Row.from1(9)).exists(_.collapsed), "summary row 9 marker should be cleared")
  }

  test("batch: grouping ops require --sheet") {
    val ops = BatchParser
      .parseBatchJson("""[{"op": "group-rows", "rows": "1:2"}]""")
      .toOption
      .get
      .ops
    val wb = Workbook(Sheet("Data"))
    val result = BatchParser.applyBatchOperations(wb, None, ops).attempt.unsafeRunSync()
    assert(result.isLeft)
    assert(
      result.swap.toOption.get.getMessage.contains("--sheet"),
      result.swap.toOption.get.getMessage
    )
  }

  test("batch: group-rows invalid level fails cleanly at apply time") {
    val sheet = Sheet("Data")
    val wb = Workbook(sheet)
    val ops = BatchParser
      .parseBatchJson("""[{"op": "group-rows", "rows": "1:3", "level": 9}]""")
      .toOption
      .get
      .ops
    val result = BatchParser.applyBatchOperations(wb, Some(sheet), ops).attempt.unsafeRunSync()
    assert(result.isLeft)
    assert(
      result.swap.toOption.get.getMessage.contains("1-7"),
      result.swap.toOption.get.getMessage
    )
  }
