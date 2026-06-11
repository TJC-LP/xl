package com.tjclp.xl.ooxml.chart

import scala.xml.*

import com.tjclp.xl.addressing.{ARef, CellRange, SheetName}
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.charts.{BarDirection, BarGrouping, Chart, ChartType, DataRef, SeriesName}
import com.tjclp.xl.ooxml.{SaxSerializable, SaxWriter, XmlUtil, XmlWritable}
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.workbooks.Workbook

/**
 * Part model and emitter for `xl/charts/chartN.xml` (GH-222) — a pure function of
 * `(Chart, ChartCacheData)`.
 *
 * Dialect: the fixture/openpyxl layout — default namespace = chart, `a` → drawingml-main, no `r`
 * binding (nothing in the emitted subset references it). Element order follows the ECMA-376
 * sequences; attributes and namespace declarations print sorted on both backends. The backends
 * agree semantically and on declaration order (whole-part byte equality is impossible for ANY part
 * today — the writers differ globally in XML declaration and empty-element minimization, the house
 * worksheet-parity precedent); each backend is individually write-twice byte-stable.
 *
 * Every emitted shape is inside [[ChartReader]]'s whitelist — the self-coherence law
 * (`ChartReader.parse(emit(chart)) == Some(chart)`) verifies this mechanically. Axis ids are the
 * FIXED 10/100 pair (Excel's random 9-digit ids would break write-twice stability); we emit
 * Excel-correct `axPos` values (the fixture's l/l quirk is read-tolerated, never reproduced).
 */
private[ooxml] final case class OoxmlChart(chart: Chart, caches: ChartCacheData)
    extends XmlWritable,
      SaxSerializable:

  def toXml: Elem =
    Elem(
      null,
      "chartSpace",
      Null,
      OoxmlChart.rootScope,
      minimizeEmpty = true,
      OoxmlChart.chartElem(chart, caches)
    )

  def writeSax(writer: SaxWriter): Unit =
    import com.tjclp.xl.ooxml.SaxSupport.*
    writer.startDocument()
    writer.startElement("chartSpace")
    SaxWriter.withAttributes(writer, writer.combinedAttributes(OoxmlChart.rootScope, Null)*) {
      writer.writeElem(OoxmlChart.chartElem(chart, caches))
    }
    writer.endElement()
    writer.endDocument()
    writer.flush()

private[ooxml] object OoxmlChart:

  /**
   * Canonical chartSpace scope: bindings print in sorted xmlns-attribute order on both backends.
   */
  val rootScope: NamespaceBinding =
    // sorted: "xmlns" < "xmlns:a" — fold right so the first sorted binding is the chain head
    NamespaceBinding(
      null,
      XmlUtil.nsChart,
      NamespaceBinding("a", XmlUtil.nsDrawingMain, TopScope)
    )

  // ===== element construction (attribute print order = sorted, matching the SAX path) =====

  private def attrs(pairs: Seq[(String, String)]): MetaData =
    pairs.sortBy(_._1).foldRight(Null: MetaData) { case ((k, v), acc) =>
      new UnprefixedAttribute(k, v, acc)
    }

  private def c(label: String, attributes: (String, String)*)(children: Node*): Elem =
    Elem(null, label, attrs(attributes), TopScope, minimizeEmpty = true, children*)

  private def a(label: String, attributes: (String, String)*)(children: Node*): Elem =
    Elem("a", label, attrs(attributes), TopScope, minimizeEmpty = true, children*)

  private def valEl(label: String, v: String): Elem = c(label, "val" -> v)()

  // ===== chart =====

  private[chart] def chartElem(chart: Chart, caches: ChartCacheData): Elem =
    val titleEls = chart.title.toList.map(titleElem)
    // autoTitleDeleted=1 suppresses Excel's auto-title for Chart(title = None)
    val atd = valEl("autoTitleDeleted", if chart.title.isDefined then "0" else "1")
    val legendEls = chart.legend.toList.map { legend =>
      val pos = legend.position match
        case com.tjclp.xl.charts.LegendPosition.Right => "r"
        case com.tjclp.xl.charts.LegendPosition.Left => "l"
        case com.tjclp.xl.charts.LegendPosition.Top => "t"
        case com.tjclp.xl.charts.LegendPosition.Bottom => "b"
        case com.tjclp.xl.charts.LegendPosition.TopRight => "tr"
      c("legend")(valEl("legendPos", pos), valEl("overlay", if legend.overlay then "1" else "0"))
    }
    val tail = Vector(valEl("plotVisOnly", "1"), valEl("dispBlanksAs", "gap"))
    c("chart")(
      (titleEls ++ Vector(atd, plotAreaElem(chart, caches)) ++ legendEls ++ tail)*
    )

  private def titleElem(text: String): Elem =
    c("title")(
      c("tx")(
        c("rich")(
          a("bodyPr")(),
          a("lstStyle")(),
          a("p")(
            a("pPr")(a("defRPr")()),
            a("r")(a("t")(Text(text)))
          )
        )
      ),
      valEl("overlay", "0")
    )

  // ===== plotArea =====

  private def plotAreaElem(chart: Chart, caches: ChartCacheData): Elem =
    val sers = chart.series.zipWithIndex.map { case (s, i) =>
      serElem(s, i, caches.series.lift(i), line = chart.chartType == ChartType.Line)
    }
    chart.chartType match
      case ChartType.Bar(direction, grouping) =>
        val dir = direction match
          case BarDirection.Col => "col"
          case BarDirection.Bar => "bar"
        val grp = grouping match
          case BarGrouping.Clustered => "clustered"
          case BarGrouping.Stacked => "stacked"
          case BarGrouping.PercentStacked => "percentStacked"
        // overlap=100 is REQUIRED for stacked groupings (without it the bars misrender
        // side-by-side); clustered must omit it
        val overlap = grouping match
          case BarGrouping.Clustered => Vector.empty[Elem]
          case _ => Vector(valEl("overlap", "100"))
        val group = c("barChart")(
          (Vector(valEl("barDir", dir), valEl("grouping", grp), valEl("varyColors", "0")) ++
            sers ++ Vector(valEl("gapWidth", "150")) ++ overlap ++
            Vector(valEl("axId", "10"), valEl("axId", "100")))*
        )
        c("plotArea")((Vector(group) ++ axesElems(direction))*)
      case ChartType.Line =>
        val group = c("lineChart")(
          (Vector(valEl("grouping", "standard"), valEl("varyColors", "0")) ++ sers ++
            Vector(valEl("marker", "1"), valEl("axId", "10"), valEl("axId", "100")))*
        )
        c("plotArea")((Vector(group) ++ axesElems(BarDirection.Col))*)
      case ChartType.Pie =>
        val group = c("pieChart")(
          (Vector(valEl("varyColors", "1")) ++ sers ++ Vector(valEl("firstSliceAng", "0")))*
        )
        c("plotArea")(group)

  /** catAx + valAx with fixed ids 10/100; horizontal bar charts swap axis positions. */
  private def axesElems(direction: BarDirection): Vector[Elem] =
    val (catPos, valPos) = direction match
      case BarDirection.Col => ("b", "l")
      case BarDirection.Bar => ("l", "b")
    val catAx = c("catAx")(
      valEl("axId", "10"),
      c("scaling")(valEl("orientation", "minMax")),
      valEl("delete", "0"),
      valEl("axPos", catPos),
      valEl("majorTickMark", "out"),
      valEl("minorTickMark", "none"),
      valEl("tickLblPos", "nextTo"),
      valEl("crossAx", "100"),
      valEl("crosses", "autoZero"),
      valEl("auto", "1"),
      valEl("lblAlgn", "ctr"),
      valEl("lblOffset", "100"),
      valEl("noMultiLvlLbl", "0")
    )
    val valAx = c("valAx")(
      valEl("axId", "100"),
      c("scaling")(valEl("orientation", "minMax")),
      valEl("delete", "0"),
      valEl("axPos", valPos),
      c("majorGridlines")(),
      c("numFmt", "formatCode" -> "General", "sourceLinked" -> "1")(),
      valEl("majorTickMark", "out"),
      valEl("minorTickMark", "none"),
      valEl("tickLblPos", "nextTo"),
      valEl("crossAx", "10"),
      valEl("crosses", "autoZero"),
      valEl("crossBetween", "between")
    )
    Vector(catAx, valAx)

  // ===== series =====

  private def serElem(
    series: com.tjclp.xl.charts.Series,
    index: Int,
    cache: Option[SeriesCache],
    line: Boolean
  ): Elem =
    val txEls = series.name.toList.map {
      case SeriesName.Literal(text) => c("tx")(c("v")(Text(text)))
      case SeriesName.FromCell(sheet, ref) =>
        val f = DataRef(sheet, CellRange(ref, ref)).toFormula
        c("tx")(
          c("strRef")(
            (Vector(c("f")(Text(f))) ++ strCacheEls(cache.flatMap(_.tx)))*
          )
        )
    }
    // CT_LineSer order: marker sits BEFORE cat, smooth AFTER val
    val markerEls =
      if line then Vector(c("marker")(valEl("symbol", "none"))) else Vector.empty[Elem]
    val catEls = series.categories.toList.map { cats =>
      val numeric = cache.exists(_.catNumeric)
      val refLabel = if numeric then "numRef" else "strRef"
      val cacheEls =
        if numeric then numCacheEls(cache.flatMap(_.cat)) else strCacheEls(cache.flatMap(_.cat))
      c("cat")(c(refLabel)((Vector(c("f")(Text(cats.toFormula))) ++ cacheEls)*))
    }
    val valElem = c("val")(
      c("numRef")(
        (Vector(c("f")(Text(series.values.toFormula))) ++ numCacheEls(cache.flatMap(_.vals)))*
      )
    )
    val smoothEls = if line then Vector(valEl("smooth", "0")) else Vector.empty[Elem]
    c("ser")(
      (Vector(valEl("idx", index.toString), valEl("order", index.toString)) ++
        txEls ++ markerEls ++ catEls ++ Vector(valElem) ++ smoothEls)*
    )

  private def numCacheEls(payload: Option[CachePayload]): List[Elem] =
    payload.toList.map { p =>
      c("numCache")(
        (Vector(c("formatCode")(Text("General")), valEl("ptCount", p.ptCount.toString)) ++
          ptEls(p))*
      )
    }

  private def strCacheEls(payload: Option[CachePayload]): List[Elem] =
    payload.toList.map { p =>
      c("strCache")((Vector(valEl("ptCount", p.ptCount.toString)) ++ ptEls(p))*)
    }

  private def ptEls(payload: CachePayload): Vector[Elem] =
    payload.pts.map { case (idx, v) => c("pt", "idx" -> idx.toString)(c("v")(Text(v))) }

/** One emitted cache: total cell count of the range + the non-skipped (idx, rendered) points. */
private[ooxml] final case class CachePayload(ptCount: Int, pts: Vector[(Int, String)])

/**
 * Resolved caches for one series. `catNumeric` decides numRef-vs-strRef for the categories (Excel's
 * rule: all resolvable non-blank cells numeric → numRef); a `None` payload means the referenced
 * sheet is absent from the workbook and the cache element is omitted entirely (the openpyxl-proven
 * bare `c:f` fallback — the writer stays total).
 */
private[ooxml] final case class SeriesCache(
  tx: Option[CachePayload],
  catNumeric: Boolean,
  cat: Option[CachePayload],
  vals: Option[CachePayload]
)

/** Write-only cache data per series, positionally aligned with `Chart.series`. */
private[ooxml] final case class ChartCacheData(series: Vector[SeriesCache])

private[ooxml] object ChartCacheData:
  val empty: ChartCacheData = ChartCacheData(Vector.empty)

/**
 * Pure cache resolution from STORED cell values (GH-222) — no evaluator: xl-ooxml never depends on
 * xl-evaluator; stored values suffice (same as Excel-without-recalc; Excel recalcs on open). Caches
 * are never stored in [[Chart]] and never compared (the formula-cache precedent).
 */
private[ooxml] object ChartCaches:

  def resolve(workbook: Workbook, chart: Chart): ChartCacheData =
    ChartCacheData(chart.series.map { series =>
      val txCache = series.name.flatMap {
        case SeriesName.FromCell(sheetName, ref) =>
          sheetOf(workbook, sheetName).map { sheet =>
            CachePayload(1, strText(valueAt(sheet, ref)).map(0 -> _).toList.toVector)
          }
        case _: SeriesName.Literal => None
      }
      val catInfo = series.categories.flatMap { cats =>
        sheetOf(workbook, cats.sheet).map { sheet =>
          val values = walk(cats.range).map(valueAt(sheet, _))
          val nonBlank = values.filterNot(_ == CellValue.Empty)
          val numeric = nonBlank.nonEmpty && nonBlank.forall(isNumericValue)
          val pts = values.zipWithIndex.flatMap { case (v, k) =>
            (if numeric then numText(v) else strText(v)).map(k -> _)
          }
          (numeric, CachePayload(values.size, pts))
        }
      }
      val valCache = sheetOf(workbook, series.values.sheet).map { sheet =>
        val values = walk(series.values.range).map(valueAt(sheet, _))
        CachePayload(
          values.size,
          values.zipWithIndex.flatMap { case (v, k) => numText(v).map(k -> _) }
        )
      }
      SeriesCache(
        tx = txCache,
        catNumeric = catInfo.exists(_._1),
        cat = catInfo.map(_._2),
        vals = valCache
      )
    })

  /** Sheet lookup is case-insensitive (Excel sheet-reference semantics). */
  private def sheetOf(workbook: Workbook, name: SheetName): Option[Sheet] =
    workbook.sheets.find(_.name.value.equalsIgnoreCase(name.value))

  /** Row-major walk = top→bottom for column vectors, left→right for row vectors. */
  private def walk(range: CellRange): Vector[ARef] =
    (for
      row <- range.start.row.index0 to range.end.row.index0
      col <- range.start.col.index0 to range.end.col.index0
    yield ARef.from0(col, row)).toVector

  private def valueAt(sheet: Sheet, ref: ARef): CellValue =
    sheet.cells.get(ref).map(_.value).getOrElse(CellValue.Empty)

  private def isNumericValue(value: CellValue): Boolean = value match
    case _: CellValue.Number => true
    case _: CellValue.DateTime => true
    case CellValue.Formula(_, Some(cached)) => isNumericValue(cached)
    case _ => false

  /** numCache pt rendering; None skips the pt (it still counts in ptCount). TRAP-5: plainNumber. */
  private def numText(value: CellValue): Option[String] = value match
    case CellValue.Number(n) => Some(XmlUtil.plainNumber(n))
    case CellValue.Bool(b) => Some(if b then "1" else "0")
    case CellValue.DateTime(dt) => Some(XmlUtil.plainNumber(CellValue.dateTimeToExcelSerial(dt)))
    case CellValue.Formula(_, Some(cached)) =>
      cached match
        case CellValue.Number(n) => Some(XmlUtil.plainNumber(n))
        case CellValue.Text(t) => Some(t)
        case _ => None
    case _ => None

  /** strCache pt rendering; None skips the pt. */
  private def strText(value: CellValue): Option[String] = value match
    case CellValue.Text(t) => Some(t)
    case CellValue.RichText(rt) => Some(rt.toPlainText)
    case CellValue.Number(n) => Some(XmlUtil.plainNumber(n))
    case CellValue.Bool(b) => Some(if b then "TRUE" else "FALSE")
    case CellValue.DateTime(dt) => Some(XmlUtil.plainNumber(CellValue.dateTimeToExcelSerial(dt)))
    case CellValue.Formula(_, Some(cached)) =>
      cached match
        case CellValue.Text(t) => Some(t)
        case CellValue.Number(n) => Some(XmlUtil.plainNumber(n))
        case CellValue.Bool(b) => Some(if b then "TRUE" else "FALSE")
        case _ => None
    case _ => None
