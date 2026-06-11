package com.tjclp.xl.ooxml

import java.nio.file.Paths

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

  /**
   * Reconcile package-level docProps relationships with what the writer emits (GH-242).
   *
   * Existing relationships of the right type are kept verbatim (stable rIds for preserved
   * `_rels/.rels`); missing ones are appended with the next free numeric rId; stale ones (part no
   * longer emitted) are removed. Deterministic for a given input.
   */
  def withDocProps(hasCore: Boolean, hasApp: Boolean): Relationships =
    def nextId(rels: Seq[Relationship]): String =
      val maxNum = rels.flatMap(r => r.id.stripPrefix("rId").toIntOption).maxOption.getOrElse(0)
      s"rId${maxNum + 1}"
    def ensure(
      rels: Seq[Relationship],
      present: Boolean,
      relType: String,
      target: String
    ): Seq[Relationship] =
      if present then
        if rels.exists(_.`type` == relType) then rels
        else rels :+ Relationship(nextId(rels), relType, target)
      else rels.filterNot(_.`type` == relType)

    val withCore = ensure(relationships, hasCore, relTypeCoreProperties, "docProps/core.xml")
    val withApp = ensure(withCore, hasApp, relTypeExtendedProperties, "docProps/app.xml")
    Relationships(withApp)

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

  /**
   * Resolve a workbook-rels target to a package path against the xl/ base. A leading slash is
   * package-absolute (openpyxl writes `/xl/worksheets/sheet1.xml`); anything else resolves relative
   * to xl/. Single source of truth shared by the reader's sheet resolution and the writer's rels
   * reconciliation (GH-320) — they must never disagree on what a target names.
   */
  def resolveWorkbookTarget(target: String): String =
    val cleaned = if target.startsWith("/") then target.drop(1) else target
    val resolvedPath =
      if cleaned.startsWith("xl/") || cleaned.startsWith("xl\\") then Paths.get(cleaned)
      else Paths.get("xl").resolve(cleaned)
    resolvedPath.normalize().toString.replace('\\', '/')

  /** A package path as a workbook-rels target (relative to xl/). */
  private def workbookTargetOf(packagePath: String): String =
    packagePath.stripPrefix("xl/")

  /**
   * Regenerate workbook.xml.rels in the same pass as workbook.xml (GH-320).
   *
   * Sheet rIds are REUSED from the preserved rels — each live sheet matched by its
   * identity-resolved source part path — and never renumbered: LibreOffice numbers its rels
   * theme-first (rId1 = theme, sheets from rId3), so renumbering sheets rId1..N against a verbatim
   * rels copy made the first sheet resolve to the theme part. Non-sheet rels (theme, styles,
   * sharedStrings, calcChain, externalLinks, pivotCache, ...) ride through untouched — their ids
   * stay stable, so `r:id` references inside preserved workbook elements (pivotCaches,
   * externalReferences) keep resolving. Worksheet rels not claimed by any live sheet are dropped
   * (deleted sheets); new sheets allocate ids ABOVE every preserved numeric id, in sheet order
   * (deterministic). The styles/sharedStrings rels are ensured by type with the same allocation
   * (presence follows what the writer actually emits — the withDocProps precedent).
   *
   * @param preserved
   *   the source workbook.xml.rels
   * @param sheetSourcePaths
   *   per live sheet, its identity-resolved SOURCE part path (None = new sheet)
   * @param sheetOutputPaths
   *   per live sheet, the package path the writer emits the worksheet at
   * @param ensureSharedStrings
   *   true when the output ships xl/sharedStrings.xml
   * @return
   *   (rId per sheet index — these MUST feed the workbook.xml `<sheet r:id>` emission, so every
   *   reference resolves by construction; the regenerated relationships)
   */
  def reconcileWorkbook(
    preserved: Relationships,
    sheetSourcePaths: Vector[Option[String]],
    sheetOutputPaths: Vector[String],
    ensureSharedStrings: Boolean
  ): (Vector[String], Relationships) =
    val worksheetRelByPath: Map[String, Relationship] =
      preserved.relationships.iterator
        .filter(_.`type` == relTypeWorksheet)
        .map(rel => resolveWorkbookTarget(rel.target) -> rel)
        .toMap
    val maxNumericId: Int = preserved.relationships
      .flatMap(_.id.stripPrefix("rId").toIntOption)
      .maxOption
      .getOrElse(0)

    // Per sheet: keep the matched rel verbatim when its target already names the output part
    // (byte-stable for untouched dialects), re-target it when the sheet moved, or allocate fresh.
    val (sheetRels, afterSheets) = sheetSourcePaths
      .zip(sheetOutputPaths)
      .foldLeft((Vector.empty[Relationship], maxNumericId)) {
        case ((acc, maxId), (sourcePath, outputPath)) =>
          sourcePath.flatMap(worksheetRelByPath.get) match
            case Some(rel) if resolveWorkbookTarget(rel.target) == outputPath =>
              (acc :+ rel, maxId)
            case Some(rel) =>
              (acc :+ rel.copy(target = workbookTargetOf(outputPath)), maxId)
            case None =>
              val fresh =
                Relationship(s"rId${maxId + 1}", relTypeWorksheet, workbookTargetOf(outputPath))
              (acc :+ fresh, maxId + 1)
      }

    val nonSheet = preserved.relationships.filterNot(_.`type` == relTypeWorksheet)

    def ensure(
      state: (Seq[Relationship], Int),
      present: Boolean,
      relType: String,
      target: String
    ): (Seq[Relationship], Int) =
      val (rels, maxId) = state
      if present then
        if rels.exists(_.`type` == relType) then state
        else (rels :+ Relationship(s"rId${maxId + 1}", relType, target), maxId + 1)
      else (rels.filterNot(_.`type` == relType), maxId)

    val base = (nonSheet ++ sheetRels, afterSheets)
    val withStyles = ensure(base, present = true, relTypeStyles, "styles.xml")
    val (finalRels, _) =
      ensure(withStyles, ensureSharedStrings, relTypeSharedStrings, "sharedStrings.xml")

    (sheetRels.map(_.id), Relationships(finalRels))

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
