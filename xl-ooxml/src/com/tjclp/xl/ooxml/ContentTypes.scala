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
) extends XmlWritable:

  def toXml: Elem =
    val defaultElems = defaults.toSeq.sortBy(_._1).map { (ext, ct) =>
      elem("Default", "Extension" -> ext, "ContentType" -> ct)()
    }

    val overrideElems = overrides.toSeq.sortBy(_._1).map { (partName, ct) =>
      elem("Override", "PartName" -> partName, "ContentType" -> ct)()
    }

    elem("Types", "xmlns" -> nsContentTypes)(defaultElems ++ overrideElems*)

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
    val commentOverrides = sheetsWithComments.toSeq.sorted.map { idx =>
      s"/xl/comments$idx.xml" -> ctComments
    }

    val baseOverrides = Map(
      "/xl/workbook.xml" -> ctWorkbook
    ) ++ sheetOverrides ++ commentOverrides

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
