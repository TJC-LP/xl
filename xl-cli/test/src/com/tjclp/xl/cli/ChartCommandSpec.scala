package com.tjclp.xl.cli

import java.nio.file.{Files, Path}

import cats.effect.IO
import munit.CatsEffectSuite

import com.tjclp.xl.{*, given}
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.charts.{BarGrouping, ChartType, DataRef, SeriesName}
import com.tjclp.xl.cli.commands.ChartCommands
import com.tjclp.xl.io.ExcelIO
import com.tjclp.xl.macros.ref
import com.tjclp.xl.ooxml.writer.WriterConfig
import com.tjclp.xl.styles.units.Emu

/**
 * Integration tests for `chart add` and `add-image` (GH-222). These round-trip through real .xlsx
 * files so they exercise the full path including chart part emission and re-read.
 */
@SuppressWarnings(Array("org.wartremover.warts.IterableOps"))
class ChartCommandSpec extends CatsEffectSuite:

  private val excel = ExcelIO.instance[IO]
  private val config = WriterConfig.default

  // Self-contained byte templates (xl-cli tests do not see xl-core's test fixtures)
  private def hexBytes(s: String): Array[Byte] =
    s.grouped(2).map(Integer.parseInt(_, 16).toByte).toArray

  /** 2x3 RGB PNG (the shared TestImages template, inlined). */
  private val pngBytes = hexBytes(
    "89504e470d0a1a0a0000000d4948445200000002000000030802000000368849d60000000e49444154785e63f8cf8001fe0300150001ff0bfeb2140000000049454e44ae426082"
  )

  /** WMF placeable header: detectable format, no sniffable dimensions. */
  private val wmfBytes = hexBytes("d7cdc69a000000000000000000000000")

  private def tmp(tag: String, suffix: String = ".xlsx"): Path =
    val p = Files.createTempFile(s"chart-cli-$tag-", suffix)
    p.toFile.deleteOnExit()
    p

  private def dataWorkbook(name: String = "Data"): Workbook =
    val sheet = Sheet(SheetName.unsafe(name))
      .put(ref"A1", CellValue.Text("Quarter"))
      .put(ref"A2", CellValue.Text("Q1"))
      .put(ref"A3", CellValue.Text("Q2"))
      .put(ref"A4", CellValue.Text("Q3"))
      .put(ref"B1", CellValue.Text("North"))
      .put(ref"B2", CellValue.Number(BigDecimal(1)))
      .put(ref"B3", CellValue.Number(BigDecimal(2)))
      .put(ref"B4", CellValue.Number(BigDecimal(3)))
      .put(ref"C1", CellValue.Text("South"))
      .put(ref"C2", CellValue.Number(BigDecimal(4)))
      .put(ref"C3", CellValue.Number(BigDecimal(5)))
      .put(ref"C4", CellValue.Number(BigDecimal(6)))
    Workbook(Vector(sheet))

  private def roundTrip(wb: Workbook)(
    run: (Workbook, Option[Sheet], Path) => IO[String]
  ): IO[Workbook] =
    val in = tmp("in")
    val out = tmp("out")
    for
      _ <- excel.write(wb, in)
      read <- excel.read(in)
      _ <- run(read, read.sheets.headOption, out)
      result <- excel.read(out)
    yield result

  test("chart add happy path: column chart with categories splits one series per data column") {
    roundTrip(dataWorkbook()) { (wb, sheet, out) =>
      ChartCommands.chartAdd(
        wb,
        sheet,
        "column",
        None,
        "B2:C4",
        Some("A2:A4"),
        Some("North,South"),
        None,
        Some("Revenue"),
        None,
        "E2:K15",
        out,
        config
      )
    }.map { result =>
      val frames = result.sheets.head.charts
      assertEquals(frames.size, 1)
      val chart = frames.head.chart
      assertEquals(chart.chartType, ChartType.Bar())
      assertEquals(chart.title, Some("Revenue"))
      assertEquals(chart.series.size, 2)
      val data = SheetName.unsafe("Data")
      assertEquals(chart.series(0).values, DataRef(data, ref"B2:B4"))
      assertEquals(chart.series(1).values, DataRef(data, ref"C2:C4"))
      assertEquals(chart.series(0).categories, Some(DataRef(data, ref"A2:A4")))
      assertEquals(chart.series(0).name, Some(SeriesName.Literal("North")))
      assertEquals(chart.series(1).name, Some(SeriesName.Literal("South")))
    }
  }

  test("chart add: stacked grouping, single-cell --at uses the default extent") {
    roundTrip(dataWorkbook()) { (wb, sheet, out) =>
      ChartCommands.chartAdd(
        wb,
        sheet,
        "bar",
        Some("stacked"),
        "B2:C4",
        None,
        None,
        None,
        None,
        Some("none"),
        "E2",
        out,
        config
      )
    }.map { result =>
      val frame = result.sheets.head.charts.head
      assertEquals(
        frame.chart.chartType,
        ChartType.Bar(com.tjclp.xl.charts.BarDirection.Bar, BarGrouping.Stacked)
      )
      assertEquals(frame.chart.legend, None)
      frame.anchor match
        case DrawingAnchor.OneCell(from, extent) =>
          assertEquals(from.cell, ref"E2")
          assertEquals(extent, Extent(Emu(5400000L), Emu(2700000L)))
        case other => fail(s"expected OneCell anchor with default extent, got $other")
    }
  }

  test("chart add: quoted sheet name flows through to emitted formulas") {
    roundTrip(dataWorkbook("Q1 Report")) { (wb, sheet, out) =>
      ChartCommands.chartAdd(
        wb,
        sheet,
        "line",
        None,
        "B2:B4",
        Some("A2:A4"),
        None,
        None,
        None,
        None,
        "E2:K15",
        out,
        config
      )
    }.map { result =>
      val chart = result.sheets.head.charts.head.chart
      val q1 = SheetName.unsafe("Q1 Report")
      assertEquals(chart.series.head.values, DataRef(q1, ref"B2:B4"))
      assertEquals(chart.series.head.values.toFormula, "'Q1 Report'!$B$2:$B$4")
    }
  }

  test("chart add: row categories split one series per data row") {
    roundTrip(dataWorkbook()) { (wb, sheet, out) =>
      ChartCommands.chartAdd(
        wb,
        sheet,
        "column",
        None,
        "B2:C4",
        Some("B1:C1"),
        None,
        None,
        None,
        None,
        "E2:K15",
        out,
        config
      )
    }.map { result =>
      val chart = result.sheets.head.charts.head.chart
      assertEquals(chart.series.size, 3) // one per data row 2..4
      assertEquals(chart.series(0).values, DataRef(SheetName.unsafe("Data"), ref"B2:C2"))
    }
  }

  test(
    "GH-407: --series-colors stamps explicit fills that round-trip; short lists leave the rest"
  ) {
    import com.tjclp.xl.styles.color.Color
    roundTrip(dataWorkbook()) { (wb, sheet, out) =>
      ChartCommands.chartAdd(
        wb,
        sheet,
        "column",
        None,
        "B2:C4",
        Some("A2:A4"),
        Some("North,South"),
        Some("#307FE2"), // one color for two series: second keeps the accent-cycle default
        None,
        None,
        "E2:K15",
        out,
        config
      )
    }.map { result =>
      val chart = result.sheets.head.charts.head.chart
      assertEquals(chart.series.size, 2)
      assertEquals(chart.series(0).fill, Some(Color.Rgb(0xff307fe2)): Option[Color.Rgb])
      // unset second series reads back as the writer's accent2 default
      assertEquals(chart.series(1).fill, Some(Color.Rgb(0xffed7d31)): Option[Color.Rgb])
    }
  }

  test("GH-418: pie --series-colors maps per-slice; explicit dPt colors round-trip") {
    import com.tjclp.xl.styles.color.Color
    roundTrip(dataWorkbook()) { (wb, sheet, out) =>
      ChartCommands.chartAdd(
        wb,
        sheet,
        "pie",
        None,
        "B2:B4",
        Some("A2:A4"),
        None,
        Some("#307FE2,#005670,#FF0000"),
        None,
        None,
        "E2:K15",
        out,
        config
      )
    }.map { result =>
      val chart = result.sheets.head.charts.head.chart
      assertEquals(chart.chartType, ChartType.Pie)
      assertEquals(chart.series.size, 1)
      // colors land per-SLICE (c:dPt), never as the series-level fill
      assertEquals(chart.series(0).fill, None: Option[Color.Rgb])
      assertEquals(
        chart.series(0).pointFills,
        Vector(Color.Rgb(0xff307fe2), Color.Rgb(0xff005670), Color.Rgb(0xffff0000))
      )
    }
  }

  test("GH-418: fewer pie colors than slices leaves the remainder on the accent cycle") {
    import com.tjclp.xl.styles.color.Color
    roundTrip(dataWorkbook()) { (wb, sheet, out) =>
      ChartCommands.chartAdd(
        wb,
        sheet,
        "pie",
        None,
        "B2:B4",
        None,
        None,
        Some("#307FE2"),
        None,
        None,
        "E2",
        out,
        config
      )
    }.map { result =>
      val chart = result.sheets.head.charts.head.chart
      // slice 0 explicit; slices 1..2 ride the accent cycle (re-derived, so not captured back)
      assertEquals(chart.series(0).pointFills, Vector(Color.Rgb(0xff307fe2)))
    }
  }

  test("GH-407: batch chart op mirrors chart add (same construction path)") {
    import com.tjclp.xl.cli.helpers.BatchParser
    import com.tjclp.xl.styles.color.Color
    val json =
      """[{"op":"chart","type":"column","data":"B2:C4","categories":"A2:A4",""" +
        """"seriesNames":"North,South","seriesColors":"#307FE2,#005670",""" +
        """"title":"Revenue","legend":"right","at":"E2:K15"}]"""
    val wb = dataWorkbook()
    for
      parsed <- IO.fromEither(BatchParser.parseBatchJson(json))
      _ = assertEquals(parsed.warnings, Vector.empty[String])
      updated <- BatchParser.applyBatchOperations(wb, wb.sheets.headOption, parsed.ops)
      // write→read so the assertion covers the full emission path, like the CLI tests above
      out = tmp("batch-chart")
      _ <- excel.write(updated, out)
      result <- excel.read(out)
    yield
      val frames = result.sheets.head.charts
      assertEquals(frames.size, 1)
      val chart = frames.head.chart
      assertEquals(chart.title, Some("Revenue"))
      assertEquals(chart.series.size, 2)
      assertEquals(chart.series(0).name, Some(SeriesName.Literal("North")))
      assertEquals(chart.series(0).fill, Some(Color.Rgb(0xff307fe2)): Option[Color.Rgb])
      assertEquals(chart.series(1).fill, Some(Color.Rgb(0xff005670)): Option[Color.Rgb])
  }

  test("GH-407: batch chart op surfaces the shared validation errors") {
    import com.tjclp.xl.cli.helpers.BatchParser
    // four colors for a three-slice pie: the shared per-slice count check rejects (GH-418)
    val json =
      """[{"op":"chart","type":"pie","data":"B2:B4",""" +
        """"seriesColors":"#307FE2,#005670,#FF0000,#00AA00","at":"E2"}]"""
    val wb = dataWorkbook()
    for
      parsed <- IO.fromEither(BatchParser.parseBatchJson(json))
      err <- BatchParser.applyBatchOperations(wb, wb.sheets.headOption, parsed.ops).attempt
    yield assert(
      err.left.exists(_.getMessage.contains("must not exceed the pie slice count")),
      err.toString
    )
  }

  test(
    "chart add errors: bad type, bad grouping placement, dimension mismatch, name count, pie multi-series"
  ) {
    val wb = dataWorkbook()
    val out = tmp("err")
    def attempt(run: IO[String]): IO[Either[Throwable, String]] = run.attempt
    for
      sheet <- IO.pure(wb.sheets.headOption)
      badType <- attempt(
        ChartCommands.chartAdd(
          wb,
          sheet,
          "donut",
          None,
          "B2:C4",
          None,
          None,
          None,
          None,
          None,
          "E2",
          out,
          config
        )
      )
      badGrouping <- attempt(
        ChartCommands.chartAdd(
          wb,
          sheet,
          "line",
          Some("stacked"),
          "B2:C4",
          None,
          None,
          None,
          None,
          None,
          "E2",
          out,
          config
        )
      )
      mismatch <- attempt(
        ChartCommands.chartAdd(
          wb,
          sheet,
          "column",
          None,
          "B2:C4",
          Some("A2:A9"),
          None,
          None,
          None,
          None,
          "E2",
          out,
          config
        )
      )
      nameCount <- attempt(
        ChartCommands.chartAdd(
          wb,
          sheet,
          "column",
          None,
          "B2:C4",
          None,
          Some("Only One"),
          None,
          None,
          None,
          "E2",
          out,
          config
        )
      )
      colorCount <- attempt(
        ChartCommands.chartAdd(
          wb,
          sheet,
          "column",
          None,
          "B2:C4",
          None,
          None,
          Some("#307FE2,#005670,#FF0000"),
          None,
          None,
          "E2",
          out,
          config
        )
      )
      badColor <- attempt(
        ChartCommands.chartAdd(
          wb,
          sheet,
          "column",
          None,
          "B2:C4",
          None,
          None,
          Some("#307FE2,not-a-color"),
          None,
          None,
          "E2",
          out,
          config
        )
      )
      pieColors <- attempt(
        ChartCommands.chartAdd(
          wb,
          sheet,
          "pie",
          None,
          "B2:B4",
          None,
          None,
          Some("#307FE2,#005670,#FF0000,#00AA00"), // 4 colors, 3 slices (GH-418)
          None,
          None,
          "E2",
          out,
          config
        )
      )
      pieMulti <- attempt(
        ChartCommands.chartAdd(
          wb,
          sheet,
          "pie",
          None,
          "B2:C4",
          None,
          None,
          None,
          None,
          None,
          "E2",
          out,
          config
        )
      )
      missingSheet <- attempt(
        ChartCommands
          .chartAdd(
            wb,
            None,
            "column",
            None,
            "B2:C4",
            None,
            None,
            None,
            None,
            None,
            "E2",
            out,
            config
          )
      )
    yield
      assert(badType.left.exists(_.getMessage.contains("Invalid --type")), badType.toString)
      assert(
        badGrouping.left.exists(_.getMessage.contains("--grouping is only valid")),
        badGrouping.toString
      )
      assert(mismatch.left.exists(_.getMessage.contains("Dimension mismatch")), mismatch.toString)
      assert(
        nameCount.left.exists(_.getMessage.contains("--series-names count")),
        nameCount.toString
      )
      assert(
        colorCount.left.exists(_.getMessage.contains("--series-colors count")),
        colorCount.toString
      )
      assert(
        badColor.left.exists(_.getMessage.contains("Invalid --series-colors entry")),
        badColor.toString
      )
      assert(
        pieColors.left.exists(
          _.getMessage.contains("must not exceed the pie slice count (3)")
        ),
        pieColors.toString
      )
      assert(
        pieMulti.left.exists(_.getMessage.contains("pie charts require exactly 1")),
        pieMulti.toString
      )
      assert(
        missingSheet.left.exists(_.getMessage.contains("requires --sheet")),
        missingSheet.toString
      )
  }

  // ===== add-image =====

  private def pngFile(): Path =
    val p = tmp("img", ".png")
    Files.write(p, pngBytes)
    p

  test("add-image: natural size at a single cell") {
    roundTrip(dataWorkbook()) { (wb, sheet, out) =>
      ChartCommands.addImage(wb, sheet, pngFile(), "E2", None, out, config)
    }.map { result =>
      val pics = result.sheets.head.pictures
      assertEquals(pics.size, 1)
      pics.head.anchor match
        case DrawingAnchor.OneCell(from, extent) =>
          assertEquals(from.cell, ref"E2")
          assertEquals(extent, Extent.fromPx(2, 3)) // sniffed natural size
        case other => fail(s"expected OneCell anchor, got $other")
    }
  }

  test("add-image: explicit --size overrides natural size; range --at stretches") {
    roundTrip(dataWorkbook()) { (wb, sheet, out) =>
      ChartCommands.addImage(wb, sheet, pngFile(), "E2", Some("320x240"), out, config) *>
        IO.pure("")
    }.map { result =>
      result.sheets.head.pictures.head.anchor match
        case DrawingAnchor.OneCell(_, extent) => assertEquals(extent, Extent.fromPx(320, 240))
        case other => fail(s"expected OneCell anchor, got $other")
    } *>
      roundTrip(dataWorkbook()) { (wb, sheet, out) =>
        ChartCommands.addImage(wb, sheet, pngFile(), "E2:H8", None, out, config)
      }.map { result =>
        result.sheets.head.pictures.head.anchor match
          case DrawingAnchor.TwoCell(from, to, _) =>
            assertEquals(from.cell, ref"E2")
            assertEquals(to.cell, ref"I9") // one past the range end
          case other => fail(s"expected TwoCell anchor, got $other")
      }
  }

  test("add-image errors: sniff failure surfaces actionable message; bad --size; size+range") {
    val wb = dataWorkbook()
    val out = tmp("imgerr")
    val sheet = wb.sheets.headOption
    val notAnImage = tmp("bad", ".bin")
    Files.write(notAnImage, Array[Byte](1, 2, 3, 4))
    // wmf header: sniffable format detection but no natural size
    val wmf = tmp("wmf", ".wmf")
    Files.write(wmf, wmfBytes)
    for
      undetected <- ChartCommands.addImage(wb, sheet, notAnImage, "E2", None, out, config).attempt
      unsniffable <- ChartCommands.addImage(wb, sheet, wmf, "E2", None, out, config).attempt
      badSize <- ChartCommands
        .addImage(wb, sheet, pngFile(), "E2", Some("huge"), out, config)
        .attempt
      sizePlusRange <- ChartCommands
        .addImage(wb, sheet, pngFile(), "E2:H8", Some("320x240"), out, config)
        .attempt
    yield
      assert(undetected.isLeft, undetected.toString)
      assert(
        unsniffable.left.exists(_.getMessage.contains("cannot sniff natural size")),
        unsniffable.toString
      )
      assert(badSize.left.exists(_.getMessage.contains("Invalid --size")), badSize.toString)
      assert(
        sizePlusRange.left.exists(_.getMessage.contains("--size cannot be combined")),
        sizePlusRange.toString
      )
  }
