package com.tjclp.xl.ooxml.chart

import scala.xml.*

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
import com.tjclp.xl.ooxml.{XmlSecurity, XmlUtil}

/**
 * Parser for `xl/charts/chartN.xml` parts (GH-222). TOTAL: anything outside a STRICT whitelist of
 * the fixture/Excel bar|line|pie dialect yields None and the whole hosting anchor stays
 * [[com.tjclp.xl.drawings.Drawing.Preserved]] — never half-typed.
 *
 * Prefix-agnostic: elements match by local label + namespace URI (the openpyxl fixture uses the
 * DEFAULT chart namespace, Excel uses `c:`). Within the whitelist, presentational details that the
 * canonical emitter re-derives are accepted-and-dropped: `idx`/`order` (re-derived 0..n-1), axis
 * ids (re-derived 10/100), any `axPos` (TRAP-4: the fixture writes the l/l quirk Excel ignores),
 * tick marks, label alignment/offset, `numFmt`, empty `majorGridlines`, num/str caches anywhere
 * (write-only hints), the fixture's no-op `ser/spPr` of exactly `a:ln/a:prstDash val="solid"`, and
 * `lang`/`roundedCorners`. The loss only materializes on a dirty regeneration of the chart part
 * (the GH-221 drawing-layer contract).
 *
 * Deliberately rejected (visible state the model cannot represent): every other chart group,
 * `c:style`/spPr/txPr formatting, externalData, dLbls/dPt/trendlines/errBars, hidden or inverted
 * axes, non-default gapWidth/overlap, smooth lines, real markers, multi-level categories,
 * `mc:AlternateContent`, and `autoTitleDeleted val="0"` without a title (Excel's auto-title state
 * is unrepresentable).
 */
object ChartReader:

  /** Parse a chart part from raw XML. Malformed or out-of-fence XML yields None. */
  def parse(chartXml: String): Option[Chart] =
    XmlSecurity.parseSafe(chartXml, "xl/charts").toOption.flatMap(fromElem)

  /** Parse an already-parsed `chartSpace` element. Total. */
  def fromElem(chartSpace: Elem): Option[Chart] =
    for
      _ <- Option.when(chartSpace.label == "chartSpace" && isC(chartSpace))(())
      kids = elems(chartSpace)
      _ <- Option.when(
        kids.forall(k => isC(k) && Set("chart", "lang", "roundedCorners")(k.label))
      )(())
      chartEl <- single(kids.filter(_.label == "chart"))
      chart <- parseChart(chartEl)
    yield chart

  // ===== chart =====

  private def parseChart(chartEl: Elem): Option[Chart] =
    val allowed =
      Set("title", "autoTitleDeleted", "plotArea", "legend", "plotVisOnly", "dispBlanksAs")
    val kids = elems(chartEl)
    for
      _ <- Option.when(kids.forall(k => isC(k) && allowed(k.label)))(())
      _ <- Option.when(allowed.forall(l => kids.count(_.label == l) <= 1))(())
      title <- kids.find(_.label == "title") match
        case None => Some(None)
        case Some(t) => parseTitle(t).map(Some(_))
      _ <- kids.find(_.label == "autoTitleDeleted") match
        case None => Some(())
        case Some(atd) =>
          valAttr(atd) match
            case Some("1") => Some(())
            case Some("0") if title.isDefined => Some(())
            case _ => None // val=0 without a title: Excel auto-title is unrepresentable
      _ <- kids
        .find(_.label == "plotVisOnly")
        .fold(Option(()))(e => Option.when(valAttr(e).contains("1"))(()))
      _ <- kids
        .find(_.label == "dispBlanksAs")
        .fold(Option(()))(e => Option.when(valAttr(e).contains("gap"))(()))
      plotAreaEl <- kids.find(_.label == "plotArea")
      typeAndSeries <- parsePlotArea(plotAreaEl)
      legend <- kids.find(_.label == "legend") match
        case None => Some(None)
        case Some(l) => parseLegend(l).map(Some(_))
    yield Chart(typeAndSeries._1, typeAndSeries._2, title, legend)

  // ===== title =====

  private def parseTitle(title: Elem): Option[String] =
    val kids = elems(title)
    for
      _ <- Option.when(kids.forall(k => isC(k) && Set("tx", "overlay")(k.label)))(())
      _ <- kids.filter(_.label == "overlay") match
        case Seq() => Some(())
        case Seq(o) => Option.when(valAttr(o).contains("0"))(())
        case _ => None
      tx <- single(kids.filter(_.label == "tx")) // c:title without c:tx -> reject
      rich <- single(elems(tx)).filter(e => e.label == "rich" && isC(e))
      text <- parseRich(rich)
    yield text

  private def parseRich(rich: Elem): Option[String] =
    val kids = elems(rich)
    val bodyPrs = kids.filter(k => k.label == "bodyPr" && isA(k))
    val lstStyles = kids.filter(k => k.label == "lstStyle" && isA(k))
    val ps = kids.filter(k => k.label == "p" && isA(k))
    for
      _ <- Option.when(kids.sizeIs == bodyPrs.size + lstStyles.size + ps.size)(())
      _ <- Option.when(bodyPrs.sizeIs == 1 && bodyPrs.forall(e => elems(e).isEmpty))(())
      _ <- Option.when(lstStyles.sizeIs <= 1 && lstStyles.forall(e => elems(e).isEmpty))(())
      p <- single(ps)
      text <- parseParagraph(p)
    yield text

  private def parseParagraph(p: Elem): Option[String] =
    val kids = elems(p)
    val pPrs = kids.filter(k => k.label == "pPr" && isA(k))
    val runs = kids.filter(k => k.label == "r" && isA(k))
    for
      _ <- Option.when(kids.sizeIs == pPrs.size + runs.size && pPrs.sizeIs <= 1)(())
      _ <- Option.when(pPrs.forall(plainPPr))(())
      texts <- traverseOpt(runs)(parseRun)
    yield texts.mkString

  /** `a:pPr` whose only content is an attribute-free, empty `a:defRPr` (no formatting). */
  private def plainPPr(pPr: Elem): Boolean =
    pPr.attributes == Null && (elems(pPr) match
      case Seq(defRPr) if defRPr.label == "defRPr" && isA(defRPr) =>
        defRPr.attributes == Null && elems(defRPr).isEmpty
      case Seq() => true
      case _ => false)

  /** `a:r` = optional formatting-free `a:rPr` + exactly one `a:t`. */
  private def parseRun(r: Elem): Option[String] =
    val kids = elems(r)
    val rPrs = kids.filter(k => k.label == "rPr" && isA(k))
    val ts = kids.filter(k => k.label == "t" && isA(k))
    for
      _ <- Option.when(kids.sizeIs == rPrs.size + ts.size && rPrs.sizeIs <= 1)(())
      _ <- Option.when(rPrs.forall(e => e.attributes == Null && elems(e).isEmpty))(())
      t <- single(ts)
    yield t.text

  // ===== plotArea =====

  private def parsePlotArea(plotArea: Elem): Option[(ChartType, Vector[Series])] =
    val allowed = Set("layout", "barChart", "lineChart", "pieChart", "catAx", "valAx")
    val kids = elems(plotArea)
    for
      _ <- Option.when(kids.forall(k => isC(k) && allowed(k.label)))(())
      _ <- Option.when(
        kids.filter(_.label == "layout").forall(e => e.attributes == Null && elems(e).isEmpty)
      )(())
      group <- single(kids.filter(k => Set("barChart", "lineChart", "pieChart")(k.label)))
      catAxes = kids.filter(_.label == "catAx")
      valAxes = kids.filter(_.label == "valAx")
      result <- group.label match
        case "pieChart" =>
          for
            _ <- Option.when(catAxes.isEmpty && valAxes.isEmpty)(())
            series <- parsePie(group)
          yield (ChartType.Pie, series)
        case "barChart" =>
          for
            parsed <- parseBar(group)
            _ <- checkAxes(catAxes, valAxes, parsed._3)
          yield (parsed._1, parsed._2)
        case _ =>
          for
            parsed <- parseLine(group)
            _ <- checkAxes(catAxes, valAxes, parsed._2)
          yield (ChartType.Line, parsed._1)
    yield result

  // ===== chart groups =====

  private def parseBar(group: Elem): Option[(ChartType, Vector[Series], Vector[String])] =
    val allowed = Set("barDir", "grouping", "varyColors", "ser", "gapWidth", "overlap", "axId")
    val kids = elems(group)
    for
      _ <- Option.when(kids.forall(k => isC(k) && allowed(k.label)))(())
      _ <- Option.when(
        (allowed - "ser" - "axId").forall(l => kids.count(_.label == l) <= 1)
      )(())
      dirEl <- kids.find(_.label == "barDir")
      direction <- valAttr(dirEl).getOrElse("col") match
        case "col" => Some(BarDirection.Col)
        case "bar" => Some(BarDirection.Bar)
        case _ => None
      grouping <- kids.find(_.label == "grouping") match
        case None => Some(BarGrouping.Clustered)
        case Some(g) =>
          valAttr(g).getOrElse("clustered") match
            case "clustered" => Some(BarGrouping.Clustered)
            case "stacked" => Some(BarGrouping.Stacked)
            case "percentStacked" => Some(BarGrouping.PercentStacked)
            case _ => None
      _ <- varyColors(kids, expected = "0")
      _ <- kids
        .find(_.label == "gapWidth")
        .fold(Option(()))(g => Option.when(valAttr(g).contains("150"))(()))
      // overlap is VISIBLE: stacked groupings require exactly 100, clustered requires absent
      _ <- (kids.find(_.label == "overlap").flatMap(valAttr), grouping) match
        case (None, BarGrouping.Clustered) => Some(())
        case (Some("100"), BarGrouping.Stacked | BarGrouping.PercentStacked) => Some(())
        case _ => None
      series <- traverseOpt(kids.filter(_.label == "ser"))(parseSer(_, line = false))
      axIds = kids.filter(_.label == "axId").flatMap(valAttr)
      _ <- Option.when(axIds.sizeIs == 2)(())
    yield (ChartType.Bar(direction, grouping), series, axIds.toVector)

  private def parseLine(group: Elem): Option[(Vector[Series], Vector[String])] =
    val allowed = Set("grouping", "varyColors", "ser", "marker", "axId")
    val kids = elems(group)
    for
      _ <- Option.when(kids.forall(k => isC(k) && allowed(k.label)))(())
      _ <- Option.when(
        (allowed - "ser" - "axId").forall(l => kids.count(_.label == l) <= 1)
      )(())
      _ <- kids.find(_.label == "grouping") match
        case None => Some(())
        case Some(g) => Option.when(valAttr(g).getOrElse("standard") == "standard")(())
      _ <- varyColors(kids, expected = "0")
      _ <- kids
        .find(_.label == "marker")
        .fold(Option(()))(m => Option.when(valAttr(m).contains("1") && elems(m).isEmpty)(()))
      series <- traverseOpt(kids.filter(_.label == "ser"))(parseSer(_, line = true))
      axIds = kids.filter(_.label == "axId").flatMap(valAttr)
      _ <- Option.when(axIds.sizeIs == 2)(())
    yield (series, axIds.toVector)

  private def parsePie(group: Elem): Option[Vector[Series]] =
    val allowed = Set("varyColors", "ser", "firstSliceAng")
    val kids = elems(group)
    for
      _ <- Option.when(kids.forall(k => isC(k) && allowed(k.label)))(())
      _ <- varyColors(kids, expected = "1")
      _ <- kids
        .find(_.label == "firstSliceAng")
        .fold(Option(()))(f => Option.when(valAttr(f).contains("0"))(()))
      ser <- single(kids.filter(_.label == "ser"))
      series <- parseSer(ser, line = false)
    yield Vector(series)

  private def varyColors(kids: Seq[Elem], expected: String): Option[Unit] =
    kids
      .find(_.label == "varyColors")
      .fold(Option(()))(v => Option.when(valAttr(v).contains(expected))(()))

  // ===== series =====

  private def parseSer(ser: Elem, line: Boolean): Option[Series] =
    val base = Set("idx", "order", "tx", "spPr", "cat", "val")
    val allowed = if line then base ++ Set("marker", "smooth") else base
    val kids = elems(ser)
    for
      _ <- Option.when(kids.forall(k => isC(k) && allowed(k.label)))(())
      _ <- Option.when(allowed.forall(l => kids.count(_.label == l) <= 1))(())
      _ <- Option.when(kids.filter(_.label == "spPr").forall(noOpSpPr))(())
      _ <-
        if line then
          for
            _ <- kids.find(_.label == "marker") match
              case None => Some(())
              case Some(m) =>
                elems(m) match
                  case Seq(sym) if sym.label == "symbol" && isC(sym) =>
                    Option.when(valAttr(sym).contains("none"))(())
                  case _ => None
            _ <- kids
              .find(_.label == "smooth")
              .fold(Option(()))(s => Option.when(valAttr(s).contains("0"))(()))
          yield ()
        else Some(())
      name <- kids.find(_.label == "tx") match
        case None => Some(None)
        case Some(tx) => parseTx(tx).map(Some(_))
      cats <- kids.find(_.label == "cat") match
        case None => Some(None)
        case Some(cat) => parseCatRef(cat).map(Some(_))
      valEl <- kids.find(_.label == "val")
      values <- parseValRef(valEl)
    yield Series(values, cats, name)

  /**
   * The fixture's no-op `ser/spPr`: exactly `a:ln/a:prstDash val="solid"`. Anything else rejects.
   */
  private def noOpSpPr(spPr: Elem): Boolean =
    spPr.attributes == Null && (elems(spPr) match
      case Seq(ln) if ln.label == "ln" && isA(ln) && ln.attributes == Null =>
        elems(ln) match
          case Seq(dash) if dash.label == "prstDash" && isA(dash) =>
            valAttr(dash).contains("solid") && elems(dash).isEmpty
          case _ => false
      case _ => false)

  /** `c:tx` = literal `c:v` or single-cell `c:strRef/c:f` (cache optional, dropped — TRAP-7). */
  private def parseTx(tx: Elem): Option[SeriesName] =
    elems(tx) match
      case Seq(v) if v.label == "v" && isC(v) && elems(v).isEmpty =>
        Some(SeriesName.Literal(v.text))
      case Seq(strRef) if strRef.label == "strRef" && isC(strRef) =>
        for
          f <- refFormula(strRef, cacheLabel = "strCache")
          cell <- DataRef.parseSingleCell(f)
        yield SeriesName.FromCell(cell._1, cell._2)
      case _ => None

  /** `c:cat` = strRef or numRef over a vector range (kind dropped, canonicalized on emission). */
  private def parseCatRef(cat: Elem): Option[DataRef] =
    elems(cat) match
      case Seq(r) if r.label == "strRef" && isC(r) =>
        refFormula(r, cacheLabel = "strCache").flatMap(parseVectorRef)
      case Seq(r) if r.label == "numRef" && isC(r) =>
        refFormula(r, cacheLabel = "numCache").flatMap(parseVectorRef)
      case _ => None

  /** `c:val` = numRef over a vector range. */
  private def parseValRef(valEl: Elem): Option[DataRef] =
    elems(valEl) match
      case Seq(r) if r.label == "numRef" && isC(r) =>
        refFormula(r, cacheLabel = "numCache").flatMap(parseVectorRef)
      case _ => None

  private def parseVectorRef(f: String): Option[DataRef] =
    DataRef.parse(f).filter(_.isVector)

  /** Children of a strRef/numRef: exactly one `c:f` plus an optional dropped cache subtree. */
  private def refFormula(ref: Elem, cacheLabel: String): Option[String] =
    val kids = elems(ref)
    for
      _ <- Option.when(kids.forall(k => isC(k) && Set("f", cacheLabel)(k.label)))(())
      _ <- Option.when(kids.count(_.label == cacheLabel) <= 1)(())
      f <- single(kids.filter(_.label == "f"))
    yield f.text

  // ===== axes =====

  /** Exactly one catAx + one valAx, ids consistent with the group's two axId values. */
  private def checkAxes(
    catAxes: Seq[Elem],
    valAxes: Seq[Elem],
    groupAxIds: Vector[String]
  ): Option[Unit] =
    for
      catAx <- single(catAxes)
      valAx <- single(valAxes)
      cat <- parseAxis(catAx)
      value <- parseAxis(valAx)
      (catId, catCross) = cat
      (valId, valCross) = value
      _ <- Option.when(
        catId != valId && groupAxIds.toSet == Set(catId, valId) &&
          catCross == valId && valCross == catId
      )(())
    yield ()

  /** Whitelisted axis; returns (axId, crossAx) for wiring checks (values otherwise dropped). */
  private def parseAxis(ax: Elem): Option[(String, String)] =
    val allowed = Set(
      "axId",
      "scaling",
      "delete",
      "axPos",
      "majorGridlines",
      "majorTickMark",
      "minorTickMark",
      "tickLblPos",
      "crossAx",
      "crosses",
      "crossBetween",
      "auto",
      "lblAlgn",
      "lblOffset",
      "noMultiLvlLbl",
      "numFmt"
    )
    val kids = elems(ax)
    for
      _ <- Option.when(kids.forall(k => isC(k) && allowed(k.label)))(())
      _ <- Option.when(allowed.forall(l => kids.count(_.label == l) <= 1))(())
      axId <- kids.find(_.label == "axId").flatMap(valAttr)
      crossAx <- kids.find(_.label == "crossAx").flatMap(valAttr)
      _ <- kids.find(_.label == "scaling") match
        case None => Some(())
        case Some(sc) =>
          Option.when(
            elems(sc).forall(k =>
              isC(k) && k.label == "orientation" && valAttr(k).contains("minMax") &&
                elems(k).isEmpty
            )
          )(())
      _ <- kids
        .find(_.label == "delete")
        .fold(Option(()))(d => Option.when(valAttr(d).contains("0"))(()))
      _ <- kids
        .find(_.label == "crosses")
        .fold(Option(()))(c => Option.when(valAttr(c).contains("autoZero"))(()))
      _ <- Option.when(kids.filter(_.label == "majorGridlines").forall(g => elems(g).isEmpty))(())
    yield (axId, crossAx)

  // ===== legend =====

  private def parseLegend(legend: Elem): Option[Legend] =
    val allowed = Set("legendPos", "overlay", "layout")
    val kids = elems(legend)
    for
      _ <- Option.when(kids.forall(k => isC(k) && allowed(k.label)))(())
      _ <- Option.when(allowed.forall(l => kids.count(_.label == l) <= 1))(())
      _ <- Option.when(
        kids.filter(_.label == "layout").forall(e => e.attributes == Null && elems(e).isEmpty)
      )(())
      position <- kids.find(_.label == "legendPos") match
        case None => Some(LegendPosition.Right)
        case Some(p) =>
          valAttr(p) match
            case Some("r") => Some(LegendPosition.Right)
            case Some("l") => Some(LegendPosition.Left)
            case Some("t") => Some(LegendPosition.Top)
            case Some("b") => Some(LegendPosition.Bottom)
            case Some("tr") => Some(LegendPosition.TopRight)
            case _ => None
      overlay <- kids.find(_.label == "overlay") match
        case None => Some(false)
        case Some(o) =>
          valAttr(o) match
            case Some("0") => Some(false)
            case Some("1") => Some(true)
            case _ => None
    yield Legend(position, overlay)

  // ===== helpers =====

  private def elems(e: Elem): Seq[Elem] = e.child.collect { case c: Elem => c }

  private def single(es: Seq[Elem]): Option[Elem] = es match
    case Seq(e) => Some(e)
    case _ => None

  private def valAttr(e: Elem): Option[String] = XmlUtil.getAttrOpt(e, "val")

  private def isC(e: Elem): Boolean = hasNs(e, XmlUtil.nsChart)
  private def isA(e: Elem): Boolean = hasNs(e, XmlUtil.nsDrawingMain)

  /** True when the element's own prefix resolves to `ns` in its scope. */
  private def hasNs(e: Elem, ns: String): Boolean =
    Option(e.scope).flatMap(s => Option(s.getURI(e.prefix))).contains(ns)

  private def traverseOpt[A, B](as: Seq[A])(f: A => Option[B]): Option[Vector[B]] =
    as.foldLeft(Option(Vector.empty[B])) { (acc, a) => acc.flatMap(v => f(a).map(v :+ _)) }
