package com.tjclp.xl.ooxml.drawing

import java.nio.file.Paths
import scala.collection.immutable.ArraySeq
import scala.xml.*

import com.tjclp.xl.addressing.{ARef, Column, Row}
import com.tjclp.xl.drawings.{
  AnchorPoint,
  Drawing,
  DrawingAnchor,
  EditAs,
  Extent,
  ImageData,
  ImageFormat
}
import com.tjclp.xl.ooxml.chart.ChartReader
import com.tjclp.xl.ooxml.worksheet.rebindUsedNamespaces
import com.tjclp.xl.ooxml.{Relationships, XmlSecurity, XmlUtil}
import com.tjclp.xl.styles.units.Emu

/**
 * Access to retained `xl/charts` content during drawing parsing (GH-222) — the `media` accessor
 * pattern: `xml` fetches a chart part's raw XML by zip path, `hasRels` reports whether the part has
 * its OWN rels file (Excel's colors1/style1/userShapes would orphan on regeneration, so such charts
 * stay Preserved).
 */
final case class ChartPartAccess(
  xml: String => Option[String],
  hasRels: String => Boolean
)

object ChartPartAccess:
  val empty: ChartPartAccess = ChartPartAccess(_ => None, _ => false)

/**
 * Result of parsing one drawing part (GH-222): the drawings in document order, the chart provenance
 * for typed ChartFrames (anchor index → (relId, chart part path), feeding
 * SourceContext.chartSnapshots), and the paths of structurally-malformed chart parts found behind
 * otherwise-gated graphicFrames (surfaced as read warnings; the anchors stay Preserved).
 */
final case class ParsedDrawingPart(
  drawings: Vector[Drawing],
  chartRefs: Map[Int, (String, String)],
  malformedChartParts: Vector[String]
)

/**
 * Parser for `xl/drawings/drawingN.xml` parts (GH-221/GH-222). TOTAL: never fails a read.
 *
 * A STRICT typed subset of CT_Drawing parses to [[Drawing.Picture]] or [[Drawing.ChartFrame]];
 * every other anchor — shapes, group shapes, connectors, pictures with
 * crops/effects/hyperlinks/links, charts outside [[ChartReader]]'s fence, mc:AlternateContent — is
 * captured whole as [[Drawing.Preserved]] carrying scope-self-contained canonical XML, re-emitted
 * verbatim in document order on write.
 *
 * Children are matched by local label + namespace URI, prefix-agnostic: openpyxl uses a default
 * namespace, Excel uses the `xdr:` prefix; both parse identically.
 *
 * Accepted-and-dropped within the typed subset (the loss only materializes on a dirty regeneration
 * of the part): `a:blip/@cstate`, a plain `a:xfrm` (off/ext only), any `cNvPicPr` /
 * `cNvGraphicFramePr` content (e.g. `a:picLocks`, `a:graphicFrameLocks`), and a graphicFrame's
 * empty `macro` attribute.
 */
object DrawingReader:

  /**
   * Parse a drawing part from raw XML. Malformed XML yields an empty result (the part still rides
   * byte-preservation; the caller surfaces a read warning).
   */
  def parse(
    partXml: String,
    rels: Relationships,
    media: String => Option[ArraySeq[Byte]],
    charts: ChartPartAccess
  ): ParsedDrawingPart =
    XmlSecurity
      .parseSafe(partXml, "xl/drawings")
      .toOption
      .map(fromElem(_, rels, media, charts))
      .getOrElse(ParsedDrawingPart(Vector.empty, Map.empty, Vector.empty))

  /** Parse an already-parsed `wsDr` element. Total. */
  def fromElem(
    wsDr: Elem,
    rels: Relationships,
    media: String => Option[ArraySeq[Byte]],
    charts: ChartPartAccess
  ): ParsedDrawingPart =
    val anchors = wsDr.child.collect { case e: Elem => e }.toVector
    val drawings = Vector.newBuilder[Drawing]
    val chartRefs = Map.newBuilder[Int, (String, String)]
    val malformed = Vector.newBuilder[String]
    anchors.zipWithIndex.foreach { case (child, idx) =>
      val result = parseAnchor(child, rels, media, charts)
      drawings += result.drawing
      result.chartRef.foreach(ref => chartRefs += idx -> ref)
      malformed ++= result.malformedChartPart
    }
    ParsedDrawingPart(drawings.result(), chartRefs.result(), malformed.result())

  /** Resolve a drawing-rels image target ("/xl/media/x.png" or "../media/x.png") to a zip path. */
  private[ooxml] def resolveMediaTarget(target: String): Option[String] =
    val cleaned = if target.startsWith("/") then target.drop(1) else target
    val resolved =
      if cleaned.startsWith("xl/") || cleaned.startsWith("xl\\") then Paths.get(cleaned)
      else Paths.get("xl/drawings").resolve(cleaned)
    val normalized = resolved.normalize().toString.replace('\\', '/')
    Option.when(normalized.startsWith("xl/"))(normalized)

  /**
   * Resolve a drawing-rels chart target to a zip path. The fixture writes an ABSOLUTE
   * `/xl/charts/chart1.xml`, Excel writes `../charts/chartN.xml`; both resolve here.
   */
  private[ooxml] def resolveChartTarget(target: String): Option[String] =
    val cleaned = if target.startsWith("/") then target.drop(1) else target
    val resolved =
      if cleaned.startsWith("xl/") || cleaned.startsWith("xl\\") then Paths.get(cleaned)
      else Paths.get("xl/drawings").resolve(cleaned)
    val normalized = resolved.normalize().toString.replace('\\', '/')
    Option.when(normalized.startsWith("xl/"))(normalized)

  // ===== anchor dispatch =====

  /** One parsed anchor: the drawing plus chart provenance / malformed-chart-part bookkeeping. */
  private final case class AnchorResult(
    drawing: Drawing,
    chartRef: Option[(String, String)] = None,
    malformedChartPart: Option[String] = None
  )

  private def parseAnchor(
    child: Elem,
    rels: Relationships,
    media: String => Option[ArraySeq[Byte]],
    charts: ChartPartAccess
  ): AnchorResult =
    val typed = child.label match
      case "oneCellAnchor" | "twoCellAnchor" | "absoluteAnchor"
          if hasNs(child, XmlUtil.nsSpreadsheetDrawing) =>
        parseTypedAnchor(child, rels, media, charts)
      case _ => Left(None)
    typed match
      case Right(result) => result
      case Left(malformedChart) =>
        AnchorResult(
          Drawing.Preserved(XmlUtil.compact(rebindUsedNamespaces(child, includeDefault = true))),
          malformedChartPart = malformedChart
        )

  /**
   * Left(malformedChartPathOpt) = stay Preserved (the path is set only when the rejection cause was
   * structurally-malformed chart XML behind an otherwise-gated graphicFrame).
   */
  private def parseTypedAnchor(
    anchor: Elem,
    rels: Relationships,
    media: String => Option[ArraySeq[Byte]],
    charts: ChartPartAccess
  ): Either[Option[String], AnchorResult] =
    val children = anchor.child.collect { case e: Elem => e }
    def sd(label: String): Option[Elem] =
      children.find(e => e.label == label && hasNs(e, XmlUtil.nsSpreadsheetDrawing))
    val structural = Set("from", "to", "ext", "pos", "clientData")
    val objects = children.filterNot(e =>
      structural.contains(e.label) && hasNs(e, XmlUtil.nsSpreadsheetDrawing)
    )
    val drawingAnchorOpt: Option[DrawingAnchor] = anchor.label match
      case "oneCellAnchor" =>
        for
          from <- sd("from").flatMap(parseMarker)
          extent <- sd("ext").flatMap(parseExtent)
          if sd("to").isEmpty && sd("pos").isEmpty
        yield DrawingAnchor.OneCell(from, extent)
      case "twoCellAnchor" =>
        for
          from <- sd("from").flatMap(parseMarker)
          to <- sd("to").flatMap(parseMarker)
          editAs <- parseEditAs(anchor)
          if sd("ext").isEmpty && sd("pos").isEmpty
        yield DrawingAnchor.TwoCell(from, to, editAs)
      case _ => // absoluteAnchor
        for
          pos <- sd("pos").flatMap(parsePos)
          extent <- sd("ext").flatMap(parseExtent)
          if sd("from").isEmpty && sd("to").isEmpty
        yield DrawingAnchor.Absolute(pos._1, pos._2, extent)
    (drawingAnchorOpt, objects) match
      case (Some(drawingAnchor), Seq(pic))
          if pic.label == "pic" && hasNs(pic, XmlUtil.nsSpreadsheetDrawing) =>
        parsePic(pic, drawingAnchor, rels, media).map(AnchorResult(_)).toRight(None)
      case (Some(drawingAnchor), Seq(frame))
          if frame.label == "graphicFrame" && hasNs(frame, XmlUtil.nsSpreadsheetDrawing) =>
        parseGraphicFrame(frame, drawingAnchor, rels, charts)
      case _ => Left(None)

  private def parseEditAs(anchor: Elem): Option[EditAs] =
    XmlUtil.getAttrOpt(anchor, "editAs") match
      case None | Some("twoCell") => Some(EditAs.TwoCell)
      case Some("oneCell") => Some(EditAs.OneCell)
      case Some("absolute") => Some(EditAs.Absolute)
      case Some(_) => None

  // ===== structural pieces =====

  private def parseMarker(marker: Elem): Option[AnchorPoint] =
    def part(label: String): Option[String] =
      marker.child.collectFirst {
        case e: Elem if e.label == label && hasNs(e, XmlUtil.nsSpreadsheetDrawing) =>
          e.text.trim
      }
    for
      col <- part("col").flatMap(_.toIntOption)
      colOff <- part("colOff").flatMap(_.toLongOption)
      row <- part("row").flatMap(_.toIntOption)
      rowOff <- part("rowOff").flatMap(_.toLongOption)
      // CT_Marker is 0-based, matching ARef.from0 — no ±1
      if col >= 0 && col <= Column.MaxIndex0 && row >= 0 && row <= Row.MaxIndex0
    yield AnchorPoint(ARef.from0(col, row), Emu(colOff), Emu(rowOff))

  private def parseExtent(ext: Elem): Option[Extent] =
    for
      cx <- XmlUtil.getAttrOpt(ext, "cx").flatMap(_.toLongOption)
      cy <- XmlUtil.getAttrOpt(ext, "cy").flatMap(_.toLongOption)
    yield Extent(Emu(cx), Emu(cy))

  private def parsePos(pos: Elem): Option[(Emu, Emu)] =
    for
      x <- XmlUtil.getAttrOpt(pos, "x").flatMap(_.toLongOption)
      y <- XmlUtil.getAttrOpt(pos, "y").flatMap(_.toLongOption)
    yield (Emu(x), Emu(y))

  // ===== graphicFrame -> ChartFrame (strict subset, GH-222) =====

  /**
   * Types a `graphicFrame` anchor to [[Drawing.ChartFrame]] IFF every gate below holds; ANY failure
   * leaves the whole anchor Preserved (never half-typed):
   *   - frame attrs ⊆ {macro=""} (dropped); children = nvGraphicFramePr + optional plain xfrm +
   *     a:graphic
   *   - nvGraphicFramePr = cNvPr with NO element children + cNvGraphicFramePr (content dropped)
   *   - xfrm absent, empty, or attribute-free plain off/ext (dropped)
   *   - a:graphic/a:graphicData[@uri == nsChart] containing exactly `c:chart @r:id`
   *   - the rel resolves with type relTypeChart (absolute or ../ target)
   *   - the chart part has NO own rels file (colors/style/userShapes would orphan)
   *   - [[ChartReader]] accepts the chart XML (malformed XML additionally surfaces a warning path)
   */
  private def parseGraphicFrame(
    frame: Elem,
    anchor: DrawingAnchor,
    rels: Relationships,
    charts: ChartPartAccess
  ): Either[Option[String], AnchorResult] =
    val gated: Option[(String, String)] = // (name, relId)
      val children = frame.child.collect { case e: Elem => e }
      def sd(label: String): Option[Elem] =
        children.find(e => e.label == label && hasNs(e, XmlUtil.nsSpreadsheetDrawing))
      for
        _ <- Option.when(
          frame.attributes.asAttrMap.forall((k, v) => k == "macro" && v.isEmpty)
        )(())
        nv <- sd("nvGraphicFramePr")
        graphic <- children.find(e => e.label == "graphic" && hasNs(e, XmlUtil.nsDrawingMain))
        xfrms = children.filter(e => e.label == "xfrm" && hasNs(e, XmlUtil.nsSpreadsheetDrawing))
        _ <- Option.when(children.sizeIs == 2 + xfrms.size && xfrms.sizeIs <= 1)(())
        _ <- Option.when(xfrms.forall(plainFrameXfrm))(())
        name <- parseNvGraphicFramePr(nv)
        relId <- parseGraphicChartRef(graphic)
      yield (name, relId)
    gated match
      case None => Left(None)
      case Some((name, relId)) =>
        val resolved = for
          rel <- rels.findById(relId)
          _ <- Option.when(rel.`type` == XmlUtil.relTypeChart)(())
          path <- resolveChartTarget(rel.target)
          _ <- Option.when(!charts.hasRels(path))(())
          xml <- charts.xml(path)
        yield (path, xml)
        resolved match
          case None => Left(None)
          case Some((path, xml)) =>
            XmlSecurity.parseSafe(xml, path) match
              case Left(_) => Left(Some(path)) // malformed chart XML: warn, stay Preserved
              case Right(elem) =>
                ChartReader.fromElem(elem) match
                  case None => Left(None) // subset rejection: silent Preserved
                  case Some(chart) =>
                    Right(
                      AnchorResult(
                        Drawing.ChartFrame(anchor, chart, name),
                        chartRef = Some((relId, path))
                      )
                    )

  /** `nvGraphicFramePr` = `cNvPr` (no element children) + `cNvGraphicFramePr` (content dropped). */
  private def parseNvGraphicFramePr(nv: Elem): Option[String] =
    val children = nv.child.collect { case e: Elem => e }
    for
      _ <- Option.when(children.sizeIs == 2)(())
      cNvPr <- children.find(e => e.label == "cNvPr" && hasNs(e, XmlUtil.nsSpreadsheetDrawing))
      _ <- Option.when(
        children.exists(e =>
          e.label == "cNvGraphicFramePr" && hasNs(e, XmlUtil.nsSpreadsheetDrawing)
        )
      )(())
      _ <- Option.when(cNvPr.child.collect { case e: Elem => e }.isEmpty)(())
    yield XmlUtil.getAttrOpt(cNvPr, "name").getOrElse("")

  /** Frame-level `xfrm`: no attributes (rot/flip would be visible), children ⊆ plain {off, ext}. */
  private def plainFrameXfrm(xfrm: Elem): Boolean =
    xfrm.attributes == Null && xfrm.child
      .collect { case e: Elem => e }
      .forall(c => (c.label == "off" || c.label == "ext") && hasNs(c, XmlUtil.nsDrawingMain))

  /** `a:graphic/a:graphicData[@uri == nsChart]` containing exactly `c:chart @r:id`. */
  private def parseGraphicChartRef(graphic: Elem): Option[String] =
    val graphicKids = graphic.child.collect { case e: Elem => e }
    for
      _ <- Option.when(graphic.attributes == Null)(())
      data <- graphicKids match
        case Seq(d) if d.label == "graphicData" && hasNs(d, XmlUtil.nsDrawingMain) => Some(d)
        case _ => None
      _ <- Option.when(XmlUtil.getAttrOpt(data, "uri").contains(XmlUtil.nsChart))(())
      chartEl <- data.child.collect { case e: Elem => e } match
        case Seq(ce) if ce.label == "chart" && hasNs(ce, XmlUtil.nsChart) => Some(ce)
        case _ => None
      relId <- chartEl.attribute(XmlUtil.nsRelationships, "id").map(_.text)
    yield relId

  // ===== pic (strict subset) =====

  private def parsePic(
    pic: Elem,
    anchor: DrawingAnchor,
    rels: Relationships,
    media: String => Option[ArraySeq[Byte]]
  ): Option[Drawing.Picture] =
    val children = pic.child.collect { case e: Elem => e }
    def sd(label: String): Option[Elem] =
      children.find(e => e.label == label && hasNs(e, XmlUtil.nsSpreadsheetDrawing))
    for
      _ <- Option.when(children.sizeIs == 3)(())
      nvPicPr <- sd("nvPicPr")
      blipFill <- sd("blipFill")
      spPr <- sd("spPr")
      (name, descr) <- parseNvPicPr(nvPicPr)
      embedId <- parseBlipFill(blipFill)
      _ <- Option.when(plainSpPr(spPr))(())
      image <- resolveImage(embedId, rels, media)
    yield Drawing.Picture(anchor, image, name, descr)

  /**
   * `nvPicPr` = `cNvPr` + `cNvPicPr`. The cNvPr must carry NO element children — hlinkClick (a
   * hyperlinked picture) or extLst content would be silently lost on a dirty regeneration, so such
   * pictures stay Preserved. cNvPicPr content (a:picLocks) is accepted and dropped.
   */
  private def parseNvPicPr(nvPicPr: Elem): Option[(String, String)] =
    val children = nvPicPr.child.collect { case e: Elem => e }
    for
      _ <- Option.when(children.sizeIs == 2)(())
      cNvPr <- children.find(e => e.label == "cNvPr" && hasNs(e, XmlUtil.nsSpreadsheetDrawing))
      _ <- Option.when(
        children.exists(e => e.label == "cNvPicPr" && hasNs(e, XmlUtil.nsSpreadsheetDrawing))
      )(())
      _ <- Option.when(cNvPr.child.collect { case e: Elem => e }.isEmpty)(())
    yield (
      XmlUtil.getAttrOpt(cNvPr, "name").getOrElse(""),
      XmlUtil.getAttrOpt(cNvPr, "descr").getOrElse("")
    )

  /**
   * `blipFill` = `a:blip r:embed` (no `r:link`, optional dropped `cstate`, no children) +
   * `a:stretch`/`a:fillRect`. Crops (srcRect), tiles, and alpha effects reject to Preserved.
   * Returns the embed relationship id.
   */
  private def parseBlipFill(blipFill: Elem): Option[String] =
    val children = blipFill.child.collect { case e: Elem => e }
    def a(label: String): Option[Elem] =
      children.find(e => e.label == label && hasNs(e, XmlUtil.nsDrawingMain))
    for
      _ <- Option.when(children.sizeIs == 2)(())
      blip <- a("blip")
      stretch <- a("stretch")
      _ <- Option.when(blip.child.collect { case e: Elem => e }.isEmpty)(())
      _ <- Option.when(blip.attribute(XmlUtil.nsRelationships, "link").isEmpty)(())
      stretchKids = stretch.child.collect { case e: Elem => e }
      _ <- Option.when(
        stretchKids.forall(e =>
          e.label == "fillRect" && hasNs(e, XmlUtil.nsDrawingMain) && e.child.collect {
            case c: Elem => c
          }.isEmpty
        )
      )(())
      embed <- blip.attribute(XmlUtil.nsRelationships, "embed").map(_.text)
    yield embed

  /**
   * `spPr` with no attributes, children ⊆ {plain `a:xfrm` (off/ext only, no rot/flip — dropped),
   * `a:prstGeom prst="rect"` (empty or with an empty `a:avLst`)}. prstGeom is required;
   * effects/line/fill children reject to Preserved.
   */
  private def plainSpPr(spPr: Elem): Boolean =
    val children = spPr.child.collect { case e: Elem => e }
    val noAttrs = spPr.attributes == Null
    val prstGeom = children.filter(e => e.label == "prstGeom" && hasNs(e, XmlUtil.nsDrawingMain))
    val xfrms = children.filter(e => e.label == "xfrm" && hasNs(e, XmlUtil.nsDrawingMain))
    def plainXfrm(e: Elem): Boolean =
      e.attributes == Null && e.child
        .collect { case c: Elem => c }
        .forall(c => (c.label == "off" || c.label == "ext") && hasNs(c, XmlUtil.nsDrawingMain))
    def rectGeom(e: Elem): Boolean =
      XmlUtil.getAttrOpt(e, "prst").contains("rect") && e.child
        .collect { case c: Elem => c }
        .forall(c =>
          c.label == "avLst" && hasNs(c, XmlUtil.nsDrawingMain) && c.child.collect { case g: Elem =>
            g
          }.isEmpty
        )
    noAttrs &&
    children.sizeIs == (prstGeom.size + xfrms.size) &&
    prstGeom.sizeIs == 1 && prstGeom.forall(rectGeom) &&
    xfrms.sizeIs <= 1 && xfrms.forall(plainXfrm)

  private def resolveImage(
    embedId: String,
    rels: Relationships,
    media: String => Option[ArraySeq[Byte]]
  ): Option[ImageData] =
    for
      rel <- rels.findById(embedId)
      _ <- Option.when(rel.`type` == XmlUtil.relTypeImage)(())
      path <- resolveMediaTarget(rel.target)
      bytes <- media(path)
      format <- ImageFormat
        .detect(bytes)
        .orElse(ImageFormat.fromExtension(extensionOf(path)))
    yield ImageData(bytes, format)

  private def extensionOf(path: String): String =
    val name = path.split('/').lastOption.getOrElse(path)
    val dot = name.lastIndexOf('.')
    if dot >= 0 && dot < name.length - 1 then name.substring(dot + 1) else ""

  /**
   * True when the element's own prefix resolves to `ns` in its scope (default ns for unprefixed).
   */
  private def hasNs(e: Elem, ns: String): Boolean =
    Option(e.scope).flatMap(s => Option(s.getURI(e.prefix))).contains(ns)
