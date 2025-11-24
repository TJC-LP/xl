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

  def withTableOverrides(tableCount: Int): ContentTypes =
    if tableCount == 0 then this
    else
      val overridesToAdd = ContentTypes.tableOverrides(tableCount)
      copy(overrides = overrides ++ overridesToAdd)

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
