package com.tjclp.xl.cli.commands

import java.nio.file.{Files, Path}

import cats.effect.IO
import com.tjclp.xl.{*, given}
import com.tjclp.xl.addressing.{ARef, CellRange}
import com.tjclp.xl.charts.{
  BarDirection,
  BarGrouping,
  Chart,
  ChartType,
  DataRef,
  Legend,
  LegendPosition,
  Series,
  SeriesName
}
import com.tjclp.xl.cli.ColorParser
import com.tjclp.xl.cli.helpers.SheetResolver
import com.tjclp.xl.io.ExcelIO
import com.tjclp.xl.ooxml.writer.WriterConfig
import com.tjclp.xl.styles.color.Color
import com.tjclp.xl.styles.units.Emu

/**
 * Drawing-layer command handlers (GH-222): `chart add` and `add-image`.
 *
 * Series split for `chart add` is deterministic and documented: orientation follows the categories
 * vector — column categories (a 1-wide range) make one series per data COLUMN, row categories one
 * per data ROW, absent categories default to per-column. Dimension mismatches are hard errors,
 * never guesses.
 */
object ChartCommands:

  /** The fixture's default chart extent for a single-cell `--at` (~5.6 x 2.8 cm). */
  private val defaultChartExtent = Extent(Emu(5400000L), Emu(2700000L))

  private def writeWorkbook(
    wb: Workbook,
    outputPath: Path,
    config: WriterConfig,
    stream: Boolean
  ): IO[Unit] =
    val excel = ExcelIO.instance[IO]
    if stream then excel.writeWorkbookStream(wb, outputPath, config)
    else excel.writeWith(wb, outputPath, config)

  private def saveSuffix(outputPath: Path, stream: Boolean): String =
    if stream then s"Saved to $outputPath (streaming)" else s"Saved to $outputPath"

  private def fail[A](message: String): IO[A] = IO.raiseError(new Exception(message))

  /** Resolve a CLI ref argument to (sheet, inclusive range) — single cells become 1x1 ranges. */
  private def resolveRange(
    wb: Workbook,
    sheetOpt: Option[Sheet],
    refStr: String,
    context: String
  ): IO[(Sheet, CellRange)] =
    SheetResolver.resolveRef(wb, sheetOpt, refStr, context).map {
      case (sheet, Left(ref)) => (sheet, CellRange(ref, ref))
      case (sheet, Right(range)) => (sheet, range)
    }

  private def parseChartType(typeStr: String, groupingOpt: Option[String]): IO[ChartType] =
    val grouping: IO[BarGrouping] = groupingOpt match
      case None | Some("clustered") => IO.pure(BarGrouping.Clustered)
      case Some("stacked") => IO.pure(BarGrouping.Stacked)
      case Some("percent-stacked") => IO.pure(BarGrouping.PercentStacked)
      case Some(other) =>
        fail(s"Invalid --grouping '$other'. Valid: clustered, stacked, percent-stacked")
    typeStr match
      case "column" => grouping.map(g => ChartType.Bar(BarDirection.Col, g))
      case "bar" => grouping.map(g => ChartType.Bar(BarDirection.Bar, g))
      case "line" | "pie" =>
        if groupingOpt.isDefined then
          fail(s"--grouping is only valid with --type column or bar (got --type $typeStr)")
        else IO.pure(if typeStr == "line" then ChartType.Line else ChartType.Pie)
      case other =>
        fail(s"Invalid --type '$other'. Valid: column, bar, line, pie")

  private def parseLegend(legendOpt: Option[String]): IO[Option[Legend]] =
    legendOpt match
      case None => IO.pure(Some(Legend()))
      case Some("none") => IO.pure(None)
      case Some("right") => IO.pure(Some(Legend(LegendPosition.Right)))
      case Some("left") => IO.pure(Some(Legend(LegendPosition.Left)))
      case Some("top") => IO.pure(Some(Legend(LegendPosition.Top)))
      case Some("bottom") => IO.pure(Some(Legend(LegendPosition.Bottom)))
      case Some("top-right") => IO.pure(Some(Legend(LegendPosition.TopRight)))
      case Some(other) =>
        fail(s"Invalid --legend '$other'. Valid: right, left, top, bottom, top-right, none")

  /**
   * Split the data range into series. Column categories (width 1) => one series per data column;
   * row categories (height 1, width > 1) => one per data row; absent => per column.
   */
  private def splitSeries(
    dataSheet: SheetName,
    data: CellRange,
    cats: Option[(SheetName, CellRange)]
  ): IO[Vector[Series]] =
    val dataWidth = data.end.col.index0 - data.start.col.index0 + 1
    val dataHeight = data.end.row.index0 - data.start.row.index0 + 1
    def columnSlice(c: Int): CellRange =
      CellRange(
        ARef.from0(c, data.start.row.index0),
        ARef.from0(c, data.end.row.index0)
      )
    def rowSlice(r: Int): CellRange =
      CellRange(
        ARef.from0(data.start.col.index0, r),
        ARef.from0(data.end.col.index0, r)
      )
    val catRef = cats.map((sheet, range) => DataRef(sheet, range))
    val perColumn = (data.start.col.index0 to data.end.col.index0).toVector
      .map(c => Series(DataRef(dataSheet, columnSlice(c)), catRef, None))
    cats match
      case None => IO.pure(perColumn)
      case Some((_, catRange)) =>
        val catWidth = catRange.end.col.index0 - catRange.start.col.index0 + 1
        val catHeight = catRange.end.row.index0 - catRange.start.row.index0 + 1
        if catWidth > 1 && catHeight > 1 then
          fail(s"--categories must be a vector (one row or one column), got ${catRange.toA1}")
        else if catWidth == 1 then
          // column categories: one series per data COLUMN; lengths must match
          if catHeight != dataHeight then
            fail(
              s"Dimension mismatch: ${catHeight} categories vs ${dataHeight} data rows " +
                s"(categories ${catRange.toA1}, data ${data.toA1})"
            )
          else IO.pure(perColumn)
        else
          // row categories: one series per data ROW; lengths must match
          if catWidth != dataWidth then
            fail(
              s"Dimension mismatch: ${catWidth} categories vs ${dataWidth} data columns " +
                s"(categories ${catRange.toA1}, data ${data.toA1})"
            )
          else
            IO.pure(
              (data.start.row.index0 to data.end.row.index0).toVector
                .map(r => Series(DataRef(dataSheet, rowSlice(r)), catRef, None))
            )

  private def applyNames(series: Vector[Series], namesOpt: Option[String]): IO[Vector[Series]] =
    namesOpt match
      case None => IO.pure(series)
      case Some(raw) =>
        val names = raw.split(',').toVector.map(_.trim)
        if names.sizeIs != series.size then
          fail(
            s"--series-names count (${names.size}) must match the series count (${series.size})"
          )
        else IO.pure(series.zip(names).map { (s, n) => s.copy(name = Some(SeriesName.Literal(n))) })

  /** One `--series-colors` entry: any ColorParser dialect that resolves to a concrete RGB. */
  private def parseSeriesColor(token: String): IO[Color.Rgb] =
    ColorParser.parse(token) match
      case Right(rgb: Color.Rgb) => IO.pure(rgb)
      case Right(_: Color.Theme) =>
        fail(
          s"--series-colors entry '$token': theme colors are not supported for series fills; " +
            "use a concrete color like #307FE2"
        )
      case Left(err) => fail(s"Invalid --series-colors entry '$token': $err")

  /**
   * Stamp explicit fills positionally (GH-407). Fewer colors than series is fine (the rest keep the
   * writer's accent cycle); more is an error. Pie rejects: slices are colored per-slice from the
   * accent cycle, so a single series color would repaint the whole pie.
   */
  private def applyColors(
    series: Vector[Series],
    chartType: ChartType,
    colorsOpt: Option[String]
  ): IO[Vector[Series]] =
    colorsOpt match
      case None => IO.pure(series)
      case Some(_) if chartType == ChartType.Pie =>
        fail(
          "--series-colors is not supported for pie charts: slices are colored per-slice from " +
            "the theme accent cycle (a pie has one series, so a series color would paint every " +
            "slice the same)"
        )
      case Some(raw) =>
        val tokens = raw.split(',').toVector.map(_.trim)
        if tokens.sizeIs > series.size then
          fail(
            s"--series-colors count (${tokens.size}) must not exceed the series count (${series.size})"
          )
        else
          tokens
            .foldLeft(IO.pure(Vector.empty[Color.Rgb])) { (acc, token) =>
              acc.flatMap(colors => parseSeriesColor(token).map(colors :+ _))
            }
            .map { colors =>
              series.zipWithIndex.map { (s, i) =>
                colors.lift(i).fold(s)(c => s.copy(fill = Some(c)))
              }
            }

  /**
   * Shared construction path for `chart add` and the batch `chart` op (GH-407): parse, resolve
   * ranges, split series, apply names/colors, validate, anchor. Returns the updated workbook and
   * the human summary line — no write.
   */
  def buildChartAdd(
    wb: Workbook,
    sheetOpt: Option[Sheet],
    typeStr: String,
    groupingOpt: Option[String],
    dataStr: String,
    categoriesOpt: Option[String],
    seriesNamesOpt: Option[String],
    seriesColorsOpt: Option[String],
    titleOpt: Option[String],
    legendOpt: Option[String],
    atStr: String
  ): IO[(Workbook, String)] =
    for
      chartType <- parseChartType(typeStr, groupingOpt)
      legend <- parseLegend(legendOpt)
      dataResolved <- resolveRange(wb, sheetOpt, dataStr, "chart add --data")
      (dataSheet, dataRange) = dataResolved
      catsResolved <- categoriesOpt match
        case None => IO.pure(None)
        case Some(c) =>
          resolveRange(wb, sheetOpt, c, "chart add --categories")
            .map((s, r) => Some((s.name, r)))
      baseSeries <- splitSeries(dataSheet.name, dataRange, catsResolved)
      named <- applyNames(baseSeries, seriesNamesOpt)
      series <- applyColors(named, chartType, seriesColorsOpt)
      chart <- IO.fromEither(
        Chart
          .validated(chartType, series, titleOpt, legend)
          .left
          .map(err => new Exception(err.message))
      )
      atResolved <- SheetResolver.resolveRef(wb, sheetOpt, atStr, "chart add --at")
      (hostSheet, anchorAt) = atResolved
      anchor = anchorAt match
        case Left(ref) => DrawingAnchor.at(ref, defaultChartExtent)
        case Right(range) => DrawingAnchor.over(range)
      updatedWb = wb.put(hostSheet.addChart(chart, anchor))
    yield
      val seriesNote = s"${series.size} series"
      (
        updatedWb,
        s"Added $typeStr chart ($seriesNote) at $atStr on sheet '${hostSheet.name.value}'"
      )

  /** `chart add`: build a typed chart from sheet ranges, anchor it at `--at`, and write. */
  def chartAdd(
    wb: Workbook,
    sheetOpt: Option[Sheet],
    typeStr: String,
    groupingOpt: Option[String],
    dataStr: String,
    categoriesOpt: Option[String],
    seriesNamesOpt: Option[String],
    seriesColorsOpt: Option[String],
    titleOpt: Option[String],
    legendOpt: Option[String],
    atStr: String,
    outputPath: Path,
    config: WriterConfig,
    stream: Boolean = false
  ): IO[String] =
    for
      built <- buildChartAdd(
        wb,
        sheetOpt,
        typeStr,
        groupingOpt,
        dataStr,
        categoriesOpt,
        seriesNamesOpt,
        seriesColorsOpt,
        titleOpt,
        legendOpt,
        atStr
      )
      (updatedWb, summary) = built
      _ <- writeWorkbook(updatedWb, outputPath, config, stream)
    yield s"$summary\n${saveSuffix(outputPath, stream)}"

  /** `add-image`: embed a picture at `--at` (natural size, explicit --size, or over a range). */
  def addImage(
    wb: Workbook,
    sheetOpt: Option[Sheet],
    imagePath: Path,
    atStr: String,
    sizeOpt: Option[String],
    outputPath: Path,
    config: WriterConfig,
    stream: Boolean = false
  ): IO[String] =
    for
      bytes <- IO
        .blocking(Files.readAllBytes(imagePath))
        .adaptError(e => new Exception(s"Cannot read image file $imagePath: ${e.getMessage}"))
      image <- IO.fromEither(
        ImageData
          .detect(scala.collection.immutable.ArraySeq.unsafeWrapArray(bytes))
          .left
          .map(err => new Exception(err.message))
      )
      atResolved <- SheetResolver.resolveRef(wb, sheetOpt, atStr, "add-image --at")
      (hostSheet, anchorAt) = atResolved
      updatedSheet <- (anchorAt, sizeOpt) match
        case (Right(range), Some(_)) =>
          fail(s"--size cannot be combined with a range --at (${range.toA1} stretches the image)")
        case (Right(range), None) =>
          IO.pure(hostSheet.addImage(image, range))
        case (Left(ref), None) =>
          IO.fromEither(
            hostSheet.addImage(image, ref).left.map(err => new Exception(err.message))
          )
        case (Left(ref), Some(size)) =>
          size match
            case s"${w}x${h}" if w.toIntOption.exists(_ > 0) && h.toIntOption.exists(_ > 0) =>
              val extent = Extent.fromPx(w.toIntOption.getOrElse(1), h.toIntOption.getOrElse(1))
              IO.pure(hostSheet.addImage(image, ref, extent))
            case other =>
              fail(s"Invalid --size '$other'. Expected WxH in pixels, e.g. 320x240")
      _ <- writeWorkbook(wb.put(updatedSheet), outputPath, config, stream)
    yield s"Added ${image.format} image at $atStr on sheet '${hostSheet.name.value}'\n" +
      saveSuffix(outputPath, stream)
