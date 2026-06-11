package com.tjclp.xl.ooxml.drawing

import scala.xml.*

import com.tjclp.xl.drawings.{Drawing, DrawingAnchor, EditAs, Extent, AnchorPoint}
import com.tjclp.xl.ooxml.worksheet.cleanNamespaces
import com.tjclp.xl.ooxml.worksheet.usedPrefixBindings
import com.tjclp.xl.ooxml.worksheet.usedDefaultNamespace
import com.tjclp.xl.ooxml.{SaxSerializable, SaxWriter, XmlSecurity, XmlUtil, XmlWritable}

/**
 * Part model and emitter for `xl/drawings/drawingN.xml` (GH-221).
 *
 * `children` are fully-resolved anchor elements (typed pictures rendered, preserved fragments
 * re-parsed and namespace-cleaned) in document order = z-order. `rootScope` is a CANONICAL chain:
 * bindings print in sorted xmlns-attribute order on both writer backends, so ScalaXml and SaxStax
 * produce byte-identical parts and write-twice is trivially stable.
 *
 * Determinism contract: part content is a pure function of (drawings vector, embed rel ids, source
 * root scope) — cNvPr ids are assigned max-over-part + 1 sequentially, names default to "Image
 * {ordinal}", attributes print in sorted order.
 */
private[ooxml] final case class OoxmlDrawing(
  children: Vector[Elem],
  rootScope: NamespaceBinding
) extends XmlWritable,
      SaxSerializable:

  def toXml: Elem =
    Elem(null, "wsDr", Null, rootScope, minimizeEmpty = true, children*)

  def writeSax(writer: SaxWriter): Unit =
    import com.tjclp.xl.ooxml.SaxSupport.*
    writer.startDocument()
    writer.startElement("wsDr")
    SaxWriter.withAttributes(writer, writer.combinedAttributes(rootScope, Null)*) {
      children.foreach(writer.writeElem)
    }
    writer.endElement()
    writer.endDocument()
    writer.flush()

private[ooxml] object OoxmlDrawing:

  /** Fresh-part dialect: the fixture/openpyxl layout (default ns = spreadsheetDrawing). */
  private val requiredPictureBindings: Seq[(Option[String], String)] = Seq(
    None -> XmlUtil.nsSpreadsheetDrawing,
    Some("a") -> XmlUtil.nsDrawingMain,
    Some("r") -> XmlUtil.nsRelationships
  )

  /**
   * TRAP-6 (GH-222): `xmlns:c` must join the wsDr root scope whenever a ChartFrame emits
   * (fixture-proven binding) — the emitted `c:chart` element would otherwise carry an unresolvable
   * prefix. Pictures' default/a/r bindings are also required (graphicFrame uses all three).
   */
  private val requiredChartBindings: Seq[(Option[String], String)] =
    Seq(Some("c") -> XmlUtil.nsChart)

  /**
   * Pure string scan for relationship references inside a preserved fragment — fragments whose rIds
   * have no rels to resolve against (fresh parts without the source zip) must be dropped.
   */
  def hasRelationshipRefs(preservedXml: String): Boolean =
    preservedXml.contains("r:embed") || preservedXml.contains("r:id") ||
      preservedXml.contains("r:link")

  /**
   * Build the part from a sheet's drawings.
   *
   * @param drawings
   *   document-order drawings to emit (the caller pre-filters rel-referencing Preserved fragments
   *   for fresh parts)
   * @param embedIds
   *   anchor index -> r:embed relationship id for typed Pictures (allocated by the writer's media
   *   planner); a Picture without an entry is skipped (defensive — the planner always allocates)
   * @param chartRelIds
   *   anchor index -> r:id relationship id for typed ChartFrames (allocated by the writer's chart
   *   planner, GH-222); a ChartFrame without an entry is skipped (same defensive precedent)
   * @param sourceScope
   *   root scope of the source part for same-path regeneration; None = fresh dialect
   */
  def build(
    drawings: Vector[Drawing],
    embedIds: Map[Int, String],
    chartRelIds: Map[Int, String],
    sourceScope: Option[NamespaceBinding]
  ): OoxmlDrawing =
    // Re-parse preserved fragments (canonicalized at read; a user-constructed garbage payload is
    // dropped — documented contract on Drawing.Preserved)
    val resolved: Vector[Either[Elem, (Drawing, String)]] =
      drawings.zipWithIndex.flatMap {
        case (pic: Drawing.Picture, idx) =>
          embedIds.get(idx).map(relId => Right(pic -> relId))
        case (frame: Drawing.ChartFrame, idx) =>
          chartRelIds.get(idx).map(relId => Right(frame -> relId))
        case (Drawing.Preserved(xml), _) =>
          XmlSecurity.parseSafe(xml, "preserved drawing fragment").toOption.map(Left(_))
      }

    val fragments = resolved.collect { case Left(e) => e }
    val hasPictures = resolved.exists {
      case Right((_: Drawing.Picture, _)) => true
      case _ => false
    }
    val hasCharts = resolved.exists {
      case Right((_: Drawing.ChartFrame, _)) => true
      case _ => false
    }

    // cNvPr ids: max over the part (preserved ids scanned) + 1, sequential, nonzero.
    // Picture and chart ordinals are counted SEPARATELY ("Image 1", "Chart 1", ...).
    val usedIds = fragments.flatMap(collectCNvPrIds)
    val firstFreeId = (usedIds.maxOption.getOrElse(0L)).max(0L) + 1L

    val children = resolved
      .foldLeft((Vector.empty[Elem], firstFreeId, 1, 1)) {
        case ((acc, nextId, picOrdinal, chartOrdinal), entry) =>
          entry match
            case Left(fragment) =>
              (acc :+ cleanNamespaces(fragment), nextId, picOrdinal, chartOrdinal)
            case Right((pic: Drawing.Picture, relId)) =>
              (
                acc :+ anchorElem(pic, relId, nextId, picOrdinal),
                nextId + 1,
                picOrdinal + 1,
                chartOrdinal
              )
            case Right((frame: Drawing.ChartFrame, relId)) =>
              (
                acc :+ chartAnchorElem(frame, relId, nextId, chartOrdinal),
                nextId + 1,
                picOrdinal,
                chartOrdinal + 1
              )
            case Right((_: Drawing.Preserved, _)) => (acc, nextId, picOrdinal, chartOrdinal)
      }
      ._1

    // Root scope, first-binding-wins: source bindings (same-path regen) or the fresh dialect;
    // then picture/chart-required bindings (fills gaps when an xdr-dialect source lacks them);
    // then bindings hoisted from preserved fragments (GH-291 machinery, default ns included).
    val baseBindings =
      sourceScope.map(flattenScope).getOrElse(requiredPictureBindings)
    val pictureBindings = if hasPictures then requiredPictureBindings else Seq.empty
    val chartBindings =
      if hasCharts then requiredPictureBindings ++ requiredChartBindings else Seq.empty
    val fragmentBindings: Seq[(Option[String], String)] =
      fragments.flatMap(usedPrefixBindings).map((p, u) => (Some(p), u)) ++
        fragments.flatMap(usedDefaultNamespace).map(u => (None, u))
    val all =
      (baseBindings ++ pictureBindings ++ chartBindings ++ fragmentBindings).distinctBy(_._1)
    OoxmlDrawing(children, canonicalScope(all))

  // ===== scope helpers =====

  private def flattenScope(scope: NamespaceBinding): Seq[(Option[String], String)] =
    @annotation.tailrec
    def loop(
      s: NamespaceBinding,
      acc: List[(Option[String], String)]
    ): List[(Option[String], String)] =
      if s == null || s == TopScope then acc.reverse
      else loop(s.parent, (Option(s.prefix).filter(_.nonEmpty), s.uri) :: acc)
    loop(scope, Nil).distinctBy(_._1)

  /**
   * Build a scope chain whose PRINT order is the sorted xmlns-attribute order on both backends:
   * scala.xml prints the chain head first, and the SAX path sorts namespace attributes by name — so
   * the first sorted binding becomes the chain HEAD (foldRight).
   */
  private def canonicalScope(bindings: Seq[(Option[String], String)]): NamespaceBinding =
    val sorted = bindings.distinctBy(_._1).sortBy {
      case (None, _) => "xmlns"
      case (Some(p), _) => s"xmlns:$p"
    }
    sorted.foldRight(TopScope: NamespaceBinding) { case ((prefix, uri), parent) =>
      NamespaceBinding(prefix.orNull, uri, parent)
    }

  private def collectCNvPrIds(root: Elem): Seq[Long] =
    val own =
      if root.label == "cNvPr" then XmlUtil.getAttrOpt(root, "id").flatMap(_.toLongOption).toList
      else Seq.empty
    own ++ root.child.collect { case e: Elem => e }.flatMap(collectCNvPrIds)

  // ===== element construction (attribute print order = sorted, matching the SAX path) =====

  private def attrs(pairs: Seq[(String, String)]): MetaData =
    pairs.sortBy(_._1).foldRight(Null: MetaData) { case ((k, v), acc) =>
      new UnprefixedAttribute(k, v, acc)
    }

  private def sd(label: String, attributes: (String, String)*)(children: Node*): Elem =
    Elem(null, label, attrs(attributes), TopScope, minimizeEmpty = true, children*)

  private def a(label: String, attributes: (String, String)*)(children: Node*): Elem =
    Elem("a", label, attrs(attributes), TopScope, minimizeEmpty = true, children*)

  private def markerElem(label: String, point: AnchorPoint): Elem =
    sd(label)(
      sd("col")(Text(point.cell.col.index0.toString)),
      sd("colOff")(Text(point.dx.value.toString)),
      sd("row")(Text(point.cell.row.index0.toString)),
      sd("rowOff")(Text(point.dy.value.toString))
    )

  private def extentElem(extent: Extent): Elem =
    sd("ext", "cx" -> extent.cx.value.toString, "cy" -> extent.cy.value.toString)()

  private def anchorElem(pic: Drawing.Picture, relId: String, cnvId: Long, ordinal: Int): Elem =
    wrapAnchor(pic.anchor, buildPic(pic, relId, cnvId, ordinal))

  /** ChartFrame anchor (GH-222): the fixture byte-shape graphicFrame hosting `c:chart r:id`. */
  private def chartAnchorElem(
    frame: Drawing.ChartFrame,
    relId: String,
    cnvId: Long,
    ordinal: Int
  ): Elem =
    wrapAnchor(frame.anchor, buildGraphicFrame(frame, relId, cnvId, ordinal))

  private def wrapAnchor(anchor: DrawingAnchor, objectElem: Elem): Elem =
    val clientData = sd("clientData")()
    anchor match
      case DrawingAnchor.OneCell(from, extent) =>
        sd("oneCellAnchor")(markerElem("from", from), extentElem(extent), objectElem, clientData)
      case DrawingAnchor.TwoCell(from, to, editAs) =>
        val editAttrs = editAs match
          case EditAs.TwoCell => Seq.empty // schema default omitted
          case EditAs.OneCell => Seq("editAs" -> "oneCell")
          case EditAs.Absolute => Seq("editAs" -> "absolute")
        sd("twoCellAnchor", editAttrs*)(
          markerElem("from", from),
          markerElem("to", to),
          objectElem,
          clientData
        )
      case DrawingAnchor.Absolute(x, y, extent) =>
        sd("absoluteAnchor")(
          sd("pos", "x" -> x.value.toString, "y" -> y.value.toString)(),
          extentElem(extent),
          objectElem,
          clientData
        )

  private def buildGraphicFrame(
    frame: Drawing.ChartFrame,
    relId: String,
    cnvId: Long,
    ordinal: Int
  ): Elem =
    val name = if frame.name.nonEmpty then frame.name else s"Chart $ordinal"
    val chartRef = Elem(
      "c",
      "chart",
      new PrefixedAttribute("r", "id", relId, Null),
      TopScope,
      minimizeEmpty = true
    )
    sd("graphicFrame")(
      sd("nvGraphicFramePr")(
        sd("cNvPr", "id" -> cnvId.toString, "name" -> name)(),
        sd("cNvGraphicFramePr")()
      ),
      sd("xfrm")(),
      a("graphic")(a("graphicData", "uri" -> XmlUtil.nsChart)(chartRef))
    )

  private def buildPic(pic: Drawing.Picture, relId: String, cnvId: Long, ordinal: Int): Elem =
    val name = if pic.name.nonEmpty then pic.name else s"Image $ordinal"
    val cNvPrAttrs = Seq("id" -> cnvId.toString, "name" -> name) ++
      (if pic.description.nonEmpty then Seq("descr" -> pic.description) else Seq.empty)
    val blip = Elem(
      "a",
      "blip",
      new PrefixedAttribute("r", "embed", relId, Null),
      TopScope,
      minimizeEmpty = true
    )
    sd("pic")(
      sd("nvPicPr")(sd("cNvPr", cNvPrAttrs*)(), sd("cNvPicPr")()),
      sd("blipFill")(blip, a("stretch")(a("fillRect")())),
      sd("spPr")(a("prstGeom", "prst" -> "rect")())
    )
