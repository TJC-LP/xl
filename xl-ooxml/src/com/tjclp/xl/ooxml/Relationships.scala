package com.tjclp.xl.ooxml

import scala.xml.*
import XmlUtil.*

/** A single relationship in an OOXML package */
final case class Relationship(
  id: String, // Relationship ID (e.g., "rId1")
  `type`: String, // Type URI
  target: String, // Target path
  targetMode: Option[String] = None // Optional target mode (External, etc.)
)

/**
 * Relationships for .rels files
 *
 * Maps relationship IDs to targets. Every significant part can have a corresponding .rels file in a
 * _rels/ subdirectory.
 */
final case class Relationships(
  relationships: Seq[Relationship]
) extends XmlWritable,
      SaxSerializable:

  def toXml: Elem =
    val relElems = relationships.sortBy(_.id).map { rel =>
      val baseAttrs = Seq(
        "Id" -> rel.id,
        "Type" -> rel.`type`,
        "Target" -> rel.target
      )
      val attrs = rel.targetMode match
        case Some(mode) => baseAttrs :+ ("TargetMode" -> mode)
        case None => baseAttrs

      elem("Relationship", attrs*)()
    }

    elem("Relationships", "xmlns" -> nsPackageRels)(relElems*)

  def writeSax(writer: SaxWriter): Unit =
    writer.startDocument()
    writer.startElement("Relationships")
    SaxWriter.withAttributes(writer, "xmlns" -> nsPackageRels) {
      relationships.sortBy(_.id).foreach { rel =>
        writer.startElement("Relationship")
        writer.writeAttribute("Id", rel.id)
        writer.writeAttribute("Target", rel.target)
        writer.writeAttribute("Type", rel.`type`)
        rel.targetMode.foreach(writer.writeAttribute("TargetMode", _))
        writer.endElement()
      }
    }
    writer.endElement()
    writer.endDocument()
    writer.flush()

  /** Find relationship by ID */
  def findById(id: String): Option[Relationship] =
    relationships.find(_.id == id)

  /** Find relationship by type */
  def findByType(`type`: String): Option[Relationship] =
    relationships.find(_.`type` == `type`)

  /** Get all relationships of a given type */
  def findAllByType(`type`: String): Seq[Relationship] =
    relationships.filter(_.`type` == `type`)

object Relationships extends XmlReadable[Relationships]:
  val empty: Relationships = Relationships(Seq.empty)

  /** Create root .rels file pointing to workbook */
  def root(workbookPath: String = "xl/workbook.xml"): Relationships =
    Relationships(
      Seq(
        Relationship("rId1", relTypeOfficeDocument, workbookPath)
      )
    )

  /** Create workbook .rels with sheets (for sequential sheet indices) */
  def workbook(
    sheetCount: Int,
    hasStyles: Boolean = false,
    hasSharedStrings: Boolean = false
  ): Relationships =
    workbookWithIndices((1 to sheetCount).toSeq, hasStyles, hasSharedStrings)

  /** Create workbook .rels with specific sheet indices (supports non-sequential indices) */
  def workbookWithIndices(
    sheetIndices: Seq[Int],
    hasStyles: Boolean = false,
    hasSharedStrings: Boolean = false
  ): Relationships =
    val sheets = sheetIndices.zipWithIndex.map { case (sheetIdx, i) =>
      Relationship(s"rId${i + 1}", relTypeWorksheet, s"worksheets/sheet$sheetIdx.xml")
    }

    val nextId = sheetIndices.size + 1
    val styles =
      if hasStyles then Seq(Relationship(s"rId$nextId", relTypeStyles, "styles.xml")) else Seq.empty
    val sst =
      if hasSharedStrings then
        Seq(Relationship(s"rId${nextId + 1}", relTypeSharedStrings, "sharedStrings.xml"))
      else Seq.empty

    Relationships(sheets ++ styles ++ sst)

  def fromXml(elem: Elem): Either[String, Relationships] =
    val rels = getChildren(elem, "Relationship").map { e =>
      for
        id <- getAttr(e, "Id")
        relType <- getAttr(e, "Type")
        target <- getAttr(e, "Target")
        targetMode = getAttrOpt(e, "TargetMode")
      yield Relationship(id, relType, target, targetMode)
    }

    val errors = rels.collect { case Left(err) => err }

    if errors.nonEmpty then Left(s"Relationships parse errors: ${errors.mkString(", ")}")
    else Right(Relationships(rels.collect { case Right(r) => r }))
