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
import com.tjclp.xl.ooxml.worksheet.rebindUsedNamespaces
import com.tjclp.xl.ooxml.{Relationships, XmlSecurity, XmlUtil}
import com.tjclp.xl.styles.units.Emu

/**
 * Parser for `xl/drawings/drawingN.xml` parts (GH-221). TOTAL: never fails a read.
 *
 * A STRICT typed subset of CT_Drawing parses to [[Drawing.Picture]]; every other anchor — shapes,
 * charts (graphicFrame), group shapes, connectors, pictures with crops/effects/hyperlinks/links,
 * mc:AlternateContent — is captured whole as [[Drawing.Preserved]] carrying scope-self-contained
 * canonical XML, re-emitted verbatim in document order on write.
 *
 * Children are matched by local label + namespace URI, prefix-agnostic: openpyxl uses a default
 * namespace, Excel uses the `xdr:` prefix; both parse identically.
 *
 * Accepted-and-dropped within the typed subset (the loss only materializes on a dirty regeneration
 * of the part): `a:blip/@cstate`, a plain `a:xfrm` (off/ext only), and any `cNvPicPr` content (e.g.
 * `a:picLocks`).
 */
object DrawingReader:

  /**
   * Parse a drawing part from raw XML. Malformed XML yields Vector.empty (the part still rides
   * byte-preservation; the caller surfaces a read warning).
   */
  def parse(
    partXml: String,
    rels: Relationships,
    media: String => Option[ArraySeq[Byte]]
  ): Vector[Drawing] =
    XmlSecurity
      .parseSafe(partXml, "xl/drawings")
      .toOption
      .map(fromElem(_, rels, media))
      .getOrElse(Vector.empty)

  /** Parse an already-parsed `wsDr` element. Total. */
  def fromElem(
    wsDr: Elem,
    rels: Relationships,
    media: String => Option[ArraySeq[Byte]]
  ): Vector[Drawing] =
    wsDr.child.collect { case e: Elem => e }.map(parseAnchor(_, rels, media)).toVector

  /** Resolve a drawing-rels image target ("/xl/media/x.png" or "../media/x.png") to a zip path. */
  private[ooxml] def resolveMediaTarget(target: String): Option[String] =
    val cleaned = if target.startsWith("/") then target.drop(1) else target
    val resolved =
      if cleaned.startsWith("xl/") || cleaned.startsWith("xl\\") then Paths.get(cleaned)
      else Paths.get("xl/drawings").resolve(cleaned)
    val normalized = resolved.normalize().toString.replace('\\', '/')
    Option.when(normalized.startsWith("xl/"))(normalized)

  // ===== anchor dispatch =====

  private def parseAnchor(
    child: Elem,
    rels: Relationships,
    media: String => Option[ArraySeq[Byte]]
  ): Drawing =
    val typed = child.label match
      case "oneCellAnchor" | "twoCellAnchor" | "absoluteAnchor"
          if hasNs(child, XmlUtil.nsSpreadsheetDrawing) =>
        parseTypedAnchor(child, rels, media)
      case _ => None
    typed.getOrElse(
      Drawing.Preserved(XmlUtil.compact(rebindUsedNamespaces(child, includeDefault = true)))
    )

  private def parseTypedAnchor(
    anchor: Elem,
    rels: Relationships,
    media: String => Option[ArraySeq[Byte]]
  ): Option[Drawing.Picture] =
    val children = anchor.child.collect { case e: Elem => e }
    def sd(label: String): Option[Elem] =
      children.find(e => e.label == label && hasNs(e, XmlUtil.nsSpreadsheetDrawing))
    val structural = Set("from", "to", "ext", "pos", "clientData")
    val objects = children.filterNot(e =>
      structural.contains(e.label) && hasNs(e, XmlUtil.nsSpreadsheetDrawing)
    )
    objects match
      case Seq(pic) if pic.label == "pic" && hasNs(pic, XmlUtil.nsSpreadsheetDrawing) =>
        for
          drawingAnchor <- anchor.label match
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
          picture <- parsePic(pic, drawingAnchor, rels, media)
        yield picture
      case _ => None

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
