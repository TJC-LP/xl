package com.tjclp.xl.ooxml

import scala.xml.*
import XmlUtil.*

/**
 * Content Types for [Content_Types].xml
 *
 * Defines MIME types for parts in the OOXML package. This is a required part at the root of every
 * XLSX file.
 */
case class ContentTypes(
  defaults: Map[String, String], // extension → contentType
  overrides: Map[String, String] // partName → contentType
) extends XmlWritable,
      SaxSerializable:

  def withCommentOverrides(sheetsWithComments: Set[Int]): ContentTypes =
    if sheetsWithComments.isEmpty then this
    else
      val overridesToAdd =
        ContentTypes.commentOverrides(sheetsWithComments) ++ ContentTypes.vmlOverrides(
          sheetsWithComments
        )
      copy(overrides = overrides ++ overridesToAdd)

  /**
   * Register the comment + VML parts a write actually EMITS, by exact part path (GH-315): fresh
   * comment parts allocate numbers above everything the source claims, so index-derived overrides
   * can name the wrong part. Conservative — an Override is added only when the part is not already
   * covered (an existing Override; for VML, also a matching extension Default) — so preserved
   * content types ride through byte-identical when the source already declared the part.
   */
  def withEmittedCommentParts(commentPaths: Set[String], vmlPaths: Set[String]): ContentTypes =
    val commentAdds = commentPaths
      .filterNot(p => overrides.contains(s"/$p"))
      .map(p => s"/$p" -> ctComments)
    val vmlAdds = vmlPaths
      .filterNot { p =>
        overrides.contains(s"/$p") ||
        ContentTypes
          .extensionOf(p)
          .exists(ext => defaults.keysIterator.exists(_.equalsIgnoreCase(ext)))
      }
      .map(p => s"/$p" -> ctVmlDrawing)
    if commentAdds.isEmpty && vmlAdds.isEmpty then this
    else copy(overrides = overrides ++ commentAdds ++ vmlAdds)

  def withTableOverrides(tableCount: Int): ContentTypes =
    if tableCount == 0 then this
    else
      val overridesToAdd = ContentTypes.tableOverrides(tableCount)
      copy(overrides = overrides ++ overridesToAdd)

  /**
   * Register drawing-part overrides (GH-221). `partPaths` are zip paths without the leading slash
   * ("xl/drawings/drawing1.xml"). Idempotent.
   */
  def withDrawingOverrides(partPaths: Set[String]): ContentTypes =
    if partPaths.isEmpty then this
    else copy(overrides = overrides ++ partPaths.map(p => s"/$p" -> ctDrawing))

  /**
   * Register chart-part overrides (GH-222). `partPaths` are zip paths without the leading slash
   * ("xl/charts/chart1.xml"). Idempotent.
   */
  def withChartOverrides(partPaths: Set[String]): ContentTypes =
    if partPaths.isEmpty then this
    else copy(overrides = overrides ++ partPaths.map(p => s"/$p" -> ctChart))

  /**
   * Register media extension Defaults (GH-221), e.g. "png" -> "image/png", derived from the media
   * filenames actually written or reused by the drawing layer. Idempotent.
   */
  def withImageDefaults(extToCt: Map[String, String]): ContentTypes =
    if extToCt.isEmpty then this
    else copy(defaults = defaults ++ extToCt)

  /**
   * Reconcile docProps overrides with what the writer actually emits (GH-242).
   *
   * Idempotent: adds the override when the part is emitted, removes any stale override when it is
   * not (e.g. a preserved [Content_Types].xml whose source had docProps that the model no longer
   * carries).
   */
  def withDocPropsOverrides(hasCore: Boolean, hasApp: Boolean): ContentTypes =
    def adjust(ov: Map[String, String], present: Boolean, part: String, ct: String) =
      if present then ov + (part -> ct) else ov - part
    copy(overrides =
      adjust(
        adjust(overrides, hasCore, "/docProps/core.xml", ctCoreProperties),
        hasApp,
        "/docProps/app.xml",
        ctExtendedProperties
      )
    )

  def toXml: Elem =
    val defaultElems = defaults.toSeq.sortBy(_._1).map { (ext, ct) =>
      elem("Default", "Extension" -> ext, "ContentType" -> ct)()
    }

    val overrideElems = overrides.toSeq.sortBy(_._1).map { (partName, ct) =>
      elem("Override", "PartName" -> partName, "ContentType" -> ct)()
    }

    elem("Types", "xmlns" -> nsContentTypes)(defaultElems ++ overrideElems*)

  def writeSax(writer: SaxWriter): Unit =
    writer.startDocument()
    writer.startElement("Types")
    SaxWriter.withAttributes(writer, "xmlns" -> nsContentTypes) {
      defaults.toSeq.sortBy(_._1).foreach { case (ext, ct) =>
        writer.startElement("Default")
        writer.writeAttribute("ContentType", ct)
        writer.writeAttribute("Extension", ext)
        writer.endElement()
      }

      overrides.toSeq.sortBy(_._1).foreach { case (partName, ct) =>
        writer.startElement("Override")
        writer.writeAttribute("ContentType", ct)
        writer.writeAttribute("PartName", partName)
        writer.endElement()
      }
    }
    writer.endElement()
    writer.endDocument()
    writer.flush()

object ContentTypes extends XmlReadable[ContentTypes]:
  /** Create minimal content types for a workbook */
  def minimal(
    hasStyles: Boolean = false,
    hasSharedStrings: Boolean = false,
    sheetCount: Int = 1,
    sheetsWithComments: Set[Int] = Set.empty
  ): ContentTypes =
    forSheetIndices((1 to sheetCount).toSeq, hasStyles, hasSharedStrings, sheetsWithComments)

  /** Create content types using explicit sheet indices (supports non-sequential sheets). */
  def forSheetIndices(
    sheetIndices: Seq[Int],
    hasStyles: Boolean = false,
    hasSharedStrings: Boolean = false,
    sheetsWithComments: Set[Int] = Set.empty
  ): ContentTypes =
    val baseDefaults = Map(
      "rels" -> ctRelationships,
      "xml" -> "application/xml",
      "vml" -> ctVmlDrawing
    )

    val sheetOverrides = sheetIndices.distinct.sorted.map { idx =>
      s"/xl/worksheets/sheet$idx.xml" -> ctWorksheet
    }

    // Add comment overrides for sheets with comments
    val commentOverrides = ContentTypes.commentOverrides(sheetsWithComments)
    val vmlOverrides = ContentTypes.vmlOverrides(sheetsWithComments)

    val baseOverrides = Map(
      "/xl/workbook.xml" -> ctWorkbook
    ) ++ sheetOverrides ++ commentOverrides ++ vmlOverrides

    val stylesOverride = if hasStyles then Map("/xl/styles.xml" -> ctStyles) else Map.empty
    val sstOverride =
      if hasSharedStrings then Map("/xl/sharedStrings.xml" -> ctSharedStrings) else Map.empty

    ContentTypes(
      defaults = baseDefaults,
      overrides = baseOverrides ++ stylesOverride ++ sstOverride
    )

  /**
   * Reconcile a PRESERVED [Content_Types].xml with the MODEL-required content types (GH-314).
   *
   * Used when a write regenerates the structural parts (metadata-modified surgical writes) but
   * exotic preserved parts (pivots, custom XML, macro payloads) still ride the verbatim copy loop:
   * their registrations must survive, or the package is corrupted.
   *
   * Union of Defaults and Overrides from both sides; the model wins on conflict — it knows exactly
   * which worksheets/styles/sharedStrings/comments/tables ship — EXCEPT `/xl/workbook.xml`, whose
   * content type encodes the package dialect (macro-enabled, template) that the domain model does
   * not track.
   */
  def reconcile(preserved: ContentTypes, model: ContentTypes): ContentTypes =
    val workbookDialect =
      preserved.overrides.get(workbookPartName).map(workbookPartName -> _)
    ContentTypes(
      defaults = preserved.defaults ++ model.defaults,
      overrides = preserved.overrides ++ model.overrides ++ workbookDialect
    )

  private val workbookPartName = "/xl/workbook.xml"

  def fromXml(elem: Elem): Either[String, ContentTypes] =
    val defaults = getChildren(elem, "Default").map { e =>
      for
        ext <- getAttr(e, "Extension")
        ct <- getAttr(e, "ContentType")
      yield ext -> ct
    }

    val overrides = getChildren(elem, "Override").map { e =>
      for
        partName <- getAttr(e, "PartName")
        ct <- getAttr(e, "ContentType")
      yield partName -> ct
    }

    // Collect errors
    val defaultErrors = defaults.collect { case Left(err) => err }
    val overrideErrors = overrides.collect { case Left(err) => err }
    val errors = defaultErrors ++ overrideErrors

    if errors.nonEmpty then Left(s"ContentTypes parse errors: ${errors.mkString(", ")}")
    else
      Right(
        ContentTypes(
          defaults = defaults.collect { case Right(pair) => pair }.toMap,
          overrides = overrides.collect { case Right(pair) => pair }.toMap
        )
      )

  private def commentOverrides(sheetsWithComments: Set[Int]): Seq[(String, String)] =
    sheetsWithComments.toSeq.sorted.map { idx =>
      s"/xl/comments$idx.xml" -> ctComments
    }

  private def vmlOverrides(sheetsWithComments: Set[Int]): Seq[(String, String)] =
    sheetsWithComments.toSeq.sorted.map { idx =>
      s"/xl/drawings/vmlDrawing$idx.vml" -> ctVmlDrawing
    }

  private def tableOverrides(tableCount: Int): Seq[(String, String)] =
    (1 to tableCount).map { idx =>
      s"/xl/tables/table$idx.xml" -> ctTable
    }

  /** File extension of a zip part path ("xl/drawings/vmlDrawing2.vml" -> "vml"). */
  private def extensionOf(path: String): Option[String] =
    val slash = path.lastIndexOf('/')
    val dot = path.lastIndexOf('.')
    Option.when(dot > slash && dot < path.length - 1)(path.substring(dot + 1))
