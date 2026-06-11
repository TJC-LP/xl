package com.tjclp.xl.charts

import com.tjclp.xl.addressing.{ARef, CellRange, RefParser, RefType, SheetName}
import com.tjclp.xl.error.{XLError, XLResult}

/**
 * A sheet-qualified data range for a chart series (GH-222).
 *
 * Anchors on `range` are SEMANTICALLY IGNORED: [[toFormula]] always prints absolute references
 * (`$A$2:$A$5`, Excel's own chart output) regardless of `startAnchor`/`endAnchor`, and [[parse]]
 * produces Relative anchors. Note that `CellRange` equality includes anchors — compare DataRef
 * ranges anchor-normalized when the provenance mixes authored and parsed values.
 */
final case class DataRef(sheet: SheetName, range: CellRange) derives CanEqual:

  /** True when the range is a 1×N or N×1 vector (chart series require vectors). */
  def isVector: Boolean =
    range.start.col == range.end.col || range.start.row == range.end.row

  /** Number of cells covered (used for chart cache `ptCount`). */
  def cellCount: Int =
    (range.end.col.index0 - range.start.col.index0 + 1) *
      (range.end.row.index0 - range.start.row.index0 + 1)

  /**
   * Chart-dialect formula: quoted-when-needed sheet name (GH-263, via [[SheetName.quoteForFormula]]
   * — the only quoting authority) + absolute A1 refs. Single-cell ranges collapse to one ref
   * (`'S'!$B$2`), Excel's own shape.
   */
  def toFormula: String =
    val sheetPart = SheetName.quoteForFormula(sheet.value)
    val refPart =
      if range.start == range.end then DataRef.absCell(range.start)
      else s"${DataRef.absCell(range.start)}:${DataRef.absCell(range.end)}"
    s"$sheetPart!$refPart"

object DataRef:

  /** Absolute A1 rendering of one cell (`$B$2`). */
  private[xl] def absCell(ref: ARef): String =
    s"$$${ref.col.toLetter}$$${ref.row.index1}"

  /**
   * Parse a chart `c:f` formula into a DataRef. Total; None for anything outside the supported
   * shape: single-area, sheet-qualified A1 references only (no unions, no external-workbook
   * `[Book]` prefixes — those fail sheet-name validation naturally).
   *
   * The `$` anchors are stripped from the REFERENCE part only, after splitting at the unquoted `!`
   * — sheet names may legally contain `$` (TRAP-1: `RefParser` rejects `$` in cell literals, and
   * its grammar must not change — it backs the `ref""` macro).
   */
  def parse(f: String): Option[DataRef] =
    val bang = RefParser.findUnquotedBang(f)
    if bang <= 0 || bang >= f.length - 1 then None
    else
      val sheetPart = f.substring(0, bang)
      val refPart = f.substring(bang + 1).replace("$", "")
      RefType.parse(s"$sheetPart!$refPart").toOption.flatMap {
        case RefType.QualifiedCell(sheet, ref) => Some(DataRef(sheet, CellRange(ref, ref)))
        case RefType.QualifiedRange(sheet, range) => Some(DataRef(sheet, range))
        case _ => None
      }

  /**
   * Parse a chart `c:f` formula that must denote a SINGLE cell (series-name `tx/strRef`). None for
   * multi-cell ranges.
   */
  def parseSingleCell(f: String): Option[(SheetName, ARef)] =
    parse(f).collect {
      case DataRef(sheet, range) if range.start == range.end => (sheet, range.start)
    }

/**
 * How a series is named: a literal string (`c:tx/c:v`) or a single source cell (`c:tx/c:strRef`).
 */
enum SeriesName derives CanEqual:
  case Literal(text: String)
  case FromCell(sheet: SheetName, ref: ARef)

/**
 * One chart series: a values vector, optional categories vector, optional name. Values and
 * categories must be 1×N or N×1 vectors (enforced by [[Chart.validated]]; the writer stays total
 * regardless).
 */
final case class Series(
  values: DataRef,
  categories: Option[DataRef] = None,
  name: Option[SeriesName] = None
) derives CanEqual

/** Bar chart direction (`c:barDir`): Col = vertical columns, Bar = horizontal bars. */
enum BarDirection derives CanEqual:
  case Col, Bar

/** Bar chart grouping (`c:grouping`). */
enum BarGrouping derives CanEqual:
  case Clustered, Stacked, PercentStacked

/**
 * Supported chart types (GH-222 v1). Axes are write-time constants derived from the chart type — no
 * user-facing Axis model yet (an `Option[Axis]` field is non-breaking to add later).
 */
enum ChartType derives CanEqual:
  case Bar(
    direction: BarDirection = BarDirection.Col,
    grouping: BarGrouping = BarGrouping.Clustered
  )
  case Line
  case Pie

/** Legend placement (`c:legendPos`): r | l | t | b | tr. */
enum LegendPosition derives CanEqual:
  case Right, Left, Top, Bottom, TopRight

/** Chart legend; `overlay` lets the plot area extend under the legend. */
final case class Legend(
  position: LegendPosition = LegendPosition.Right,
  overlay: Boolean = false
) derives CanEqual

/**
 * A typed chart (GH-222): pure data, no zip paths, relationship ids, formula strings, or cache
 * values. Placement on a sheet comes from the enclosing
 * [[com.tjclp.xl.drawings.Drawing.ChartFrame]].
 */
final case class Chart(
  chartType: ChartType,
  series: Vector[Series],
  title: Option[String] = None,
  legend: Option[Legend] = Some(Legend())
) derives CanEqual

object Chart:

  /**
   * Validated construction: at least one series; Pie requires exactly one; every values/categories
   * range must be a vector (1×N or N×1).
   */
  def validated(
    chartType: ChartType,
    series: Vector[Series],
    title: Option[String] = None,
    legend: Option[Legend] = Some(Legend())
  ): XLResult[Chart] =
    if series.isEmpty then Left(XLError.Other("Invalid chart: at least one series is required"))
    else if chartType == ChartType.Pie && series.sizeIs != 1 then
      Left(XLError.Other(s"Invalid chart: pie charts require exactly 1 series, got ${series.size}"))
    else
      series.zipWithIndex
        .collectFirst {
          case (s, i) if !s.values.isVector =>
            XLError.Other(
              s"Invalid chart: series $i values ${s.values.toFormula} is not a vector (1×N or N×1)"
            )
          case (s, i) if s.categories.exists(!_.isVector) =>
            XLError.Other(
              s"Invalid chart: series $i categories is not a vector (1×N or N×1)"
            )
        }
        .toLeft(Chart(chartType, series, title, legend))

  /** Bar/column chart over `series`. */
  def bar(
    series: Vector[Series],
    direction: BarDirection = BarDirection.Col,
    grouping: BarGrouping = BarGrouping.Clustered,
    title: Option[String] = None,
    legend: Option[Legend] = Some(Legend())
  ): XLResult[Chart] =
    validated(ChartType.Bar(direction, grouping), series, title, legend)

  /** Line chart over `series`. */
  def line(
    series: Vector[Series],
    title: Option[String] = None,
    legend: Option[Legend] = Some(Legend())
  ): XLResult[Chart] =
    validated(ChartType.Line, series, title, legend)

  /** Pie chart over a single series. */
  def pie(
    series: Series,
    title: Option[String] = None,
    legend: Option[Legend] = Some(Legend())
  ): XLResult[Chart] =
    validated(ChartType.Pie, Vector(series), title, legend)
