package com.tjclp.xl.cli

import munit.FunSuite

import java.nio.file.{Files, Path}

import cats.effect.{IO, unsafe}
import com.tjclp.xl.{Workbook, Sheet}
import com.tjclp.xl.cli.helpers.BatchParser
import com.tjclp.xl.cli.commands.WriteCommands
import com.tjclp.xl.io.ExcelIO
import com.tjclp.xl.ooxml.writer.WriterConfig
import com.tjclp.xl.styles.color.{Color, ThemeSlot}

/**
 * E2E tests for the sheet appearance & print setup commands (GH-358): sheet-view, tab-color,
 * page-setup, header-footer — CLI handlers plus their batch ops, with write → re-read model
 * assertions.
 */
@SuppressWarnings(
  Array("org.wartremover.warts.OptionPartial", "org.wartremover.warts.IterableOps")
)
class AppearanceCommandsSpec extends FunSuite:

  given unsafe.IORuntime = unsafe.IORuntime.global

  val config: WriterConfig = WriterConfig.default

  private def tempXlsx(): Path = Files.createTempFile("appearance-test", ".xlsx")

  private def withTempOutput[A](f: Path => A): A =
    val out = tempXlsx()
    try f(out)
    finally Files.deleteIfExists(out)

  private def readBack(path: Path): Workbook =
    ExcelIO.instance[IO].read(path).unsafeRunSync()

  // ========== ColorParser theme syntax (GH-358) ==========

  test("ColorParser: theme:accent1 parses to Theme color with zero tint") {
    assertEquals(ColorParser.parse("theme:accent1"), Right(Color.Theme(ThemeSlot.Accent1, 0.0)))
  }

  test("ColorParser: theme:accent2:0.25 parses slot and tint") {
    assertEquals(
      ColorParser.parse("theme:accent2:0.25"),
      Right(Color.Theme(ThemeSlot.Accent2, 0.25))
    )
  }

  test("ColorParser: theme:dark1:-0.5 accepts negative tint") {
    assertEquals(
      ColorParser.parse("theme:dark1:-0.5"),
      Right(Color.Theme(ThemeSlot.Dark1, -0.5))
    )
  }

  test("ColorParser: unknown theme slot yields clean error") {
    val result = ColorParser.parse("theme:accent9")
    assert(result.isLeft, s"Expected error, got $result")
    assert(result.swap.toOption.get.contains("accent9"), result.toString)
  }

  test("ColorParser: out-of-range tint yields clean error") {
    val result = ColorParser.parse("theme:accent1:1.5")
    assert(result.isLeft, s"Expected error, got $result")
    assert(result.swap.toOption.get.contains("1.5"), result.toString)
  }

  // ========== sheet-view command ==========

  test("sheet-view: gridlines off + zoom write and round-trip") {
    withTempOutput { out =>
      val wb = Workbook(Sheet("Test"))
      val result = WriteCommands
        .sheetView(
          wb,
          Some(wb.sheets.head),
          gridlines = Some(false),
          zoom = Some(85),
          tabSelected = None,
          out,
          config
        )
        .unsafeRunSync()

      assert(result.contains("Saved"), result)

      val sheet = readBack(out).sheets.head
      val view = sheet.viewSettings
      assert(view.isDefined, "viewSettings should round-trip")
      assertEquals(view.get.showGridLines, false)
      assertEquals(view.get.zoomScale, Some(85))
    }
  }

  test("sheet-view: zoom out of range yields clean error (no stack trace)") {
    withTempOutput { out =>
      val wb = Workbook(Sheet("Test"))
      val result = WriteCommands
        .sheetView(wb, Some(wb.sheets.head), None, Some(5), None, out, config)
        .attempt
        .unsafeRunSync()

      assert(result.isLeft)
      val msg = result.swap.toOption.get.getMessage
      assert(msg.contains("10-400"), msg)
      assert(msg.contains("5"), msg)
    }
  }

  test("sheet-view: no options yields clean error") {
    withTempOutput { out =>
      val wb = Workbook(Sheet("Test"))
      val result = WriteCommands
        .sheetView(wb, Some(wb.sheets.head), None, None, None, out, config)
        .attempt
        .unsafeRunSync()

      assert(result.isLeft)
      assert(
        result.swap.toOption.get.getMessage.contains("at least one"),
        result.swap.toOption.get.getMessage
      )
    }
  }

  test("sheet-view: preserves unspecified fields from existing view settings") {
    withTempOutput { out =>
      val sheet = Sheet("Test").withViewSettings(
        com.tjclp.xl.sheets.SheetView(showGridLines = false, zoomScale = Some(70))
      )
      val wb = Workbook(sheet)
      WriteCommands
        .sheetView(wb, Some(sheet), gridlines = None, zoom = Some(90), None, out, config)
        .unsafeRunSync()

      val view = readBack(out).sheets.head.viewSettings.get
      assertEquals(view.showGridLines, false, "gridlines=off must survive a zoom-only edit")
      assertEquals(view.zoomScale, Some(90))
    }
  }

  // ========== tab-color command ==========

  test("tab-color: hex color writes and round-trips") {
    withTempOutput { out =>
      val wb = Workbook(Sheet("Test"))
      val result = WriteCommands
        .tabColor(wb, Some(wb.sheets.head), Some("#1F4E79"), clear = false, out, config)
        .unsafeRunSync()

      assert(result.contains("Saved"), result)
      val sheet = readBack(out).sheets.head
      assertEquals(sheet.tabColor, Some(Color.Rgb(0xff1f4e79)))
    }
  }

  test("tab-color: theme color round-trips through file") {
    withTempOutput { out =>
      val wb = Workbook(Sheet("Test"))
      WriteCommands
        .tabColor(wb, Some(wb.sheets.head), Some("theme:accent2:0.25"), clear = false, out, config)
        .unsafeRunSync()

      val sheet = readBack(out).sheets.head
      assertEquals(sheet.tabColor, Some(Color.Theme(ThemeSlot.Accent2, 0.25)))
    }
  }

  test("tab-color: --clear clears a modeled color") {
    withTempOutput { out =>
      val sheet = Sheet("Test").withTabColor(Color.Rgb(0xff1f4e79))
      val wb = Workbook(sheet)
      val result = WriteCommands
        .tabColor(wb, Some(sheet), None, clear = true, out, config)
        .unsafeRunSync()

      assert(result.contains("Cleared"), result)
      assertEquals(readBack(out).sheets.head.tabColor, None)
    }
  }

  test("tab-color: color and --clear are mutually exclusive") {
    withTempOutput { out =>
      val wb = Workbook(Sheet("Test"))
      val result = WriteCommands
        .tabColor(wb, Some(wb.sheets.head), Some("red"), clear = true, out, config)
        .attempt
        .unsafeRunSync()

      assert(result.isLeft)
      assert(
        result.swap.toOption.get.getMessage.contains("mutually exclusive"),
        result.swap.toOption.get.getMessage
      )
    }
  }

  test("tab-color: neither color nor --clear yields clean error") {
    withTempOutput { out =>
      val wb = Workbook(Sheet("Test"))
      val result = WriteCommands
        .tabColor(wb, Some(wb.sheets.head), None, clear = false, out, config)
        .attempt
        .unsafeRunSync()

      assert(result.isLeft)
    }
  }

  test("tab-color: invalid color yields clean error") {
    withTempOutput { out =>
      val wb = Workbook(Sheet("Test"))
      val result = WriteCommands
        .tabColor(wb, Some(wb.sheets.head), Some("notacolor"), clear = false, out, config)
        .attempt
        .unsafeRunSync()

      assert(result.isLeft)
      assert(
        result.swap.toOption.get.getMessage.contains("notacolor"),
        result.swap.toOption.get.getMessage
      )
    }
  }

  // ========== page-setup command ==========

  test("page-setup: landscape + fit 1x1 writes and round-trips") {
    withTempOutput { out =>
      val wb = Workbook(Sheet("Test"))
      val result = WriteCommands
        .pageSetup(
          wb,
          Some(wb.sheets.head),
          orientation = Some("landscape"),
          scale = None,
          fitToWidth = Some(1),
          fitToHeight = Some(1),
          fitToPage = None,
          out,
          config
        )
        .unsafeRunSync()

      assert(result.contains("Saved"), result)
      val setup = readBack(out).sheets.head.pageSetup
      assert(setup.isDefined, "pageSetup should round-trip")
      assertEquals(setup.get.orientation, Some("landscape"))
      assertEquals(setup.get.fitToWidth, Some(1))
      assertEquals(setup.get.fitToHeight, Some(1))
    }
  }

  test("page-setup: scale bounds pre-validated with clean error") {
    withTempOutput { out =>
      val wb = Workbook(Sheet("Test"))
      val result = WriteCommands
        .pageSetup(wb, Some(wb.sheets.head), None, Some(500), None, None, None, out, config)
        .attempt
        .unsafeRunSync()

      assert(result.isLeft)
      val msg = result.swap.toOption.get.getMessage
      assert(msg.contains("10-400"), msg)
    }
  }

  test("page-setup: invalid orientation yields clean error") {
    withTempOutput { out =>
      val wb = Workbook(Sheet("Test"))
      val result = WriteCommands
        .pageSetup(wb, Some(wb.sheets.head), Some("sideways"), None, None, None, None, out, config)
        .attempt
        .unsafeRunSync()

      assert(result.isLeft)
      val msg = result.swap.toOption.get.getMessage
      assert(msg.contains("portrait"), msg)
    }
  }

  test("page-setup: no options yields clean error") {
    withTempOutput { out =>
      val wb = Workbook(Sheet("Test"))
      val result = WriteCommands
        .pageSetup(wb, Some(wb.sheets.head), None, None, None, None, None, out, config)
        .attempt
        .unsafeRunSync()

      assert(result.isLeft)
    }
  }

  // ========== header-footer command ==========

  test("header-footer: odd footer with codes writes and round-trips") {
    withTempOutput { out =>
      val wb = Workbook(Sheet("Test"))
      val footer = "&LProprietary & Confidential&RPage &P of &N"
      val result = WriteCommands
        .headerFooter(
          wb,
          Some(wb.sheets.head),
          oddHeader = None,
          oddFooter = Some(footer),
          evenHeader = None,
          evenFooter = None,
          firstHeader = None,
          firstFooter = None,
          differentOddEven = false,
          differentFirst = false,
          out,
          config
        )
        .unsafeRunSync()

      assert(result.contains("Saved"), result)
      val hf = readBack(out).sheets.head.pageSetup.flatMap(_.headerFooter)
      assert(hf.isDefined, "headerFooter should round-trip")
      assertEquals(hf.get.oddFooter, Some(footer))
    }
  }

  test("header-footer: even text auto-sets differentOddEven") {
    withTempOutput { out =>
      val wb = Workbook(Sheet("Test"))
      WriteCommands
        .headerFooter(
          wb,
          Some(wb.sheets.head),
          None,
          Some("&COdd"),
          None,
          Some("&CEven"),
          None,
          None,
          differentOddEven = false,
          differentFirst = false,
          out,
          config
        )
        .unsafeRunSync()

      val hf = readBack(out).sheets.head.pageSetup.flatMap(_.headerFooter).get
      assertEquals(hf.evenFooter, Some("&CEven"))
      assert(hf.differentOddEven, "differentOddEven should be set when even text is provided")
    }
  }

  test("header-footer: preserves existing page setup fields") {
    withTempOutput { out =>
      val sheet = Sheet("Test").withPageSetup(
        com.tjclp.xl.sheets.PageSetup(orientation = Some("landscape"))
      )
      val wb = Workbook(sheet)
      WriteCommands
        .headerFooter(
          wb,
          Some(sheet),
          None,
          Some("&CFooter"),
          None,
          None,
          None,
          None,
          differentOddEven = false,
          differentFirst = false,
          out,
          config
        )
        .unsafeRunSync()

      val setup = readBack(out).sheets.head.pageSetup.get
      assertEquals(setup.orientation, Some("landscape"), "orientation must survive footer edit")
      assertEquals(setup.headerFooter.flatMap(_.oddFooter), Some("&CFooter"))
    }
  }

  test("header-footer: no options yields clean error") {
    withTempOutput { out =>
      val wb = Workbook(Sheet("Test"))
      val result = WriteCommands
        .headerFooter(
          wb,
          Some(wb.sheets.head),
          None,
          None,
          None,
          None,
          None,
          None,
          differentOddEven = false,
          differentFirst = false,
          out,
          config
        )
        .attempt
        .unsafeRunSync()

      assert(result.isLeft)
    }
  }

  // ========== batch ops ==========

  test("batch: sheet-view op parses without unknown-prop warnings") {
    val json = """[{"op": "sheet-view", "gridlines": false, "zoom": 85}]"""
    val result = BatchParser.parseBatchJson(json)
    assert(result.isRight, s"Should parse: $result")
    assertEquals(result.toOption.get.warnings, Vector.empty[String])
  }

  test("batch: tab-color, page-setup, header-footer ops parse") {
    val json =
      """[
        {"op": "tab-color", "color": "#1F4E79"},
        {"op": "page-setup", "orientation": "landscape", "fitToWidth": 1, "fitToHeight": 1},
        {"op": "header-footer", "oddFooter": "&LConfidential&RPage &P"}
      ]"""
    val result = BatchParser.parseBatchJson(json)
    assert(result.isRight, s"Should parse: $result")
    assertEquals(result.toOption.get.ops.size, 3)
    assertEquals(result.toOption.get.warnings, Vector.empty[String])
  }

  test("batch: sheet-view invalid zoom fails cleanly at apply time") {
    val json = """[{"op": "sheet-view", "zoom": 5}]"""
    val parsed = BatchParser.parseBatchJson(json)
    assert(parsed.isRight, s"Should parse: $parsed")
    val sheet = Sheet("Test")
    val wb = Workbook(sheet)
    val result = BatchParser
      .applyBatchOperations(wb, Some(sheet), parsed.toOption.get.ops)
      .attempt
      .unsafeRunSync()
    assert(result.isLeft)
    assert(
      result.swap.toOption.get.getMessage.contains("10-400"),
      result.swap.toOption.get.getMessage
    )
  }

  test("batch: deliverable finish combo applies via one batch file (GH-358 acceptance)") {
    withTempOutput { out =>
      val json =
        """[
          {"op": "sheet-view", "gridlines": false, "zoom": 85},
          {"op": "tab-color", "color": "#1F4E79"},
          {"op": "page-setup", "orientation": "landscape", "fitToWidth": 1, "fitToHeight": 1},
          {"op": "header-footer", "oddFooter": "&LProprietary & Confidential&RPage &P of &N"}
        ]"""
      val jsonPath = Files.createTempFile("finish", ".json")
      Files.writeString(jsonPath, json)
      try
        val wb = Workbook(Sheet("Model"))
        val result = WriteCommands
          .batch(wb, Some(wb.sheets.head), jsonPath.toString, out, config)
          .unsafeRunSync()

        assert(result.contains("Applied 4 operations"), result)

        val sheet = readBack(out).sheets.head
        val view = sheet.viewSettings.get
        assertEquals(view.showGridLines, false)
        assertEquals(view.zoomScale, Some(85))
        assertEquals(sheet.tabColor, Some(Color.Rgb(0xff1f4e79)))
        val setup = sheet.pageSetup.get
        assertEquals(setup.orientation, Some("landscape"))
        assertEquals(setup.fitToWidth, Some(1))
        assertEquals(setup.fitToHeight, Some(1))
        assertEquals(
          setup.headerFooter.flatMap(_.oddFooter),
          Some("&LProprietary & Confidential&RPage &P of &N")
        )
      finally Files.deleteIfExists(jsonPath)
    }
  }

  test("batch: tab-color clear form clears a modeled color") {
    withTempOutput { out =>
      val sheet = Sheet("Test").withTabColor(Color.Rgb(0xffff0000))
      val wb = Workbook(sheet)
      val ops = BatchParser
        .parseBatchJson("""[{"op": "tab-color", "clear": true}]""")
        .toOption
        .get
        .ops
      val updated = BatchParser.applyBatchOperations(wb, Some(sheet), ops).unsafeRunSync()
      assertEquals(updated.sheets.head.tabColor, None)
    }
  }

  test("deliverable finish: raw worksheet XML carries all four elements (GH-358)") {
    withTempOutput { out =>
      val json =
        """[
          {"op": "sheet-view", "gridlines": false, "zoom": 85},
          {"op": "tab-color", "color": "theme:accent2:0.25"},
          {"op": "page-setup", "orientation": "landscape", "fitToWidth": 1, "fitToHeight": 1},
          {"op": "header-footer", "oddFooter": "&LConfidential&RPage &P"}
        ]"""
      val jsonPath = Files.createTempFile("finish-xml", ".json")
      Files.writeString(jsonPath, json)
      try
        val wb = Workbook(Sheet("Model"))
        WriteCommands
          .batch(wb, Some(wb.sheets.head), jsonPath.toString, out, config)
          .unsafeRunSync()

        val zip = new java.util.zip.ZipFile(out.toFile)
        val xml =
          try
            val entry = zip.getEntry("xl/worksheets/sheet1.xml")
            new String(
              zip.getInputStream(entry).readAllBytes(),
              java.nio.charset.StandardCharsets.UTF_8
            )
          finally zip.close()

        assert(xml.contains("showGridLines=\"0\""), xml)
        assert(xml.contains("zoomScale=\"85\""), xml)
        assert(xml.contains("<tabColor") && xml.contains("theme=\"5\""), xml)
        assert(xml.contains("orientation=\"landscape\""), xml)
        assert(xml.contains("fitToWidth=\"1\"") && xml.contains("fitToHeight=\"1\""), xml)
        assert(xml.contains("fitToPage=\"1\""), "fitToPage should derive from fitTo* — " + xml)
        assert(xml.contains("<oddFooter>&amp;LConfidential&amp;RPage &amp;P</oddFooter>"), xml)
      finally Files.deleteIfExists(jsonPath)
    }
  }

  test("batch: appearance ops require --sheet") {
    val ops = BatchParser
      .parseBatchJson("""[{"op": "sheet-view", "zoom": 85}]""")
      .toOption
      .get
      .ops
    val wb = Workbook(Sheet("Test"))
    val result = BatchParser.applyBatchOperations(wb, None, ops).attempt.unsafeRunSync()
    assert(result.isLeft)
    assert(
      result.swap.toOption.get.getMessage.contains("--sheet"),
      result.swap.toOption.get.getMessage
    )
  }
