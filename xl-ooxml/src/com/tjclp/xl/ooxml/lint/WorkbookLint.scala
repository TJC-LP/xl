package com.tjclp.xl.ooxml.lint

import java.io.ByteArrayInputStream
import java.nio.charset.StandardCharsets
import java.nio.file.{Path, Paths}
import java.util.zip.{ZipFile, ZipInputStream}

import scala.xml.Elem

import com.tjclp.xl.error.{XLError, XLResult}
import com.tjclp.xl.ooxml.{OoxmlWorkbook, Relationships, XmlSecurity, XmlUtil}
import com.tjclp.xl.ooxml.worksheet.worksheetCanonicalOrder

/**
 * Category of a structural lint finding (GH-397).
 *
 * These are the Excel-repair classes every lenient reader (xl included) accepts silently: Excel
 * shows the repair dialog and strips the offending structure, so a deliverable carrying one of
 * these reads as "corrupt file" to the recipient even though library round-trips look fine.
 */
enum LintCategory derives CanEqual:
  /** A part's child-element sequence violates its CT schema order (ECMA-376 Part 1). */
  case ChildOrder

  /** An `r:id` reference has no matching `<Relationship>` in the part's paired `.rels`. */
  case UnresolvedRelId

  /** An `r:id` reference resolves to a relationship (or part content) of the wrong type. */
  case WrongRelType

  /** A resolved internal relationship targets a part that is absent from the package. */
  case MissingPart

  /** Stable kebab-case identifier used in CLI text and JSON output. */
  def slug: String = this match
    case LintCategory.ChildOrder => "child-order"
    case LintCategory.UnresolvedRelId => "unresolved-rel-id"
    case LintCategory.WrongRelType => "wrong-rel-type"
    case LintCategory.MissingPart => "missing-part"

/**
 * A single structural lint finding.
 *
 * @param part
 *   zip entry path of the offending part (e.g. "xl/workbook.xml")
 * @param category
 *   the finding class (see [[LintCategory]])
 * @param locator
 *   the offending element with its identifying attributes (e.g. `<sheet name="Q1" r:id="rId99">`)
 *   so the finding can be located and acted on without re-deriving it
 * @param message
 *   human-readable explanation naming expected vs actual
 */
final case class Finding(
  part: String,
  category: LintCategory,
  locator: String,
  message: String
) derives CanEqual

/**
 * Structural validation of an XLSX package against the corruption classes Excel repairs loudly
 * (GH-397): CT_Workbook / CT_Worksheet child-order violations and r:id references that do not
 * resolve in the paired `.rels` (or resolve to the wrong relationship/part type).
 *
 * Lint runs on the RAW ZIP PARTS, never on the parsed domain model — a full read would
 * repair/normalize the very structure lint inspects (the reader silently falls back on unresolved
 * sheet r:ids and drops dangling hyperlink r:ids). Total: package-level failures (unreadable zip,
 * missing/malformed core part) surface as `Left`, never as a throw; structural problems in a
 * parseable package surface as findings.
 */
object WorkbookLint:

  private val workbookPart = "xl/workbook.xml"
  private val workbookRelsPart = "xl/_rels/workbook.xml.rels"

  /** Access to package parts by zip entry name. */
  private trait PartSource:
    def has(name: String): Boolean
    def read(name: String): XLResult[Option[String]]

  private final class ZipFilePartSource(zip: ZipFile) extends PartSource:
    def has(name: String): Boolean = Option(zip.getEntry(name)).isDefined
    def read(name: String): XLResult[Option[String]] =
      Option(zip.getEntry(name)) match
        case None => Right(None)
        case Some(entry) =>
          try
            val is = zip.getInputStream(entry)
            try Right(Some(new String(is.readAllBytes(), StandardCharsets.UTF_8)))
            finally is.close()
          catch
            case e: Exception =>
              Left(XLError.IOError(s"Failed to read zip entry $name: ${e.getMessage}"))

  private final class MapPartSource(names: Set[String], parts: Map[String, String])
      extends PartSource:
    def has(name: String): Boolean = names.contains(name)
    def read(name: String): XLResult[Option[String]] = Right(parts.get(name))

  /** Lint an XLSX file on disk. Only the structural parts are read (workbook, worksheets, rels). */
  def lint(path: Path): XLResult[Vector[Finding]] =
    openZipFile(path).flatMap { zip =>
      try lintSource(ZipFilePartSource(zip))
      finally zip.close()
    }

  /** Lint an XLSX package from raw bytes. */
  def lintBytes(bytes: Array[Byte]): XLResult[Vector[Finding]] =
    readEntries(bytes).flatMap(lintSource)

  /** Safely open a ZIP file, converting exceptions to XLResult errors (metadata-reader pattern). */
  private def openZipFile(path: Path): XLResult[ZipFile] =
    try Right(new ZipFile(path.toFile))
    catch
      case e: java.util.zip.ZipException =>
        Left(XLError.ParseError(path.toString, s"Invalid ZIP file: ${e.getMessage}"))
      case e: java.io.IOException =>
        Left(XLError.IOError(s"Failed to open file: ${e.getMessage}"))

  /**
   * Index a zip from bytes: all entry names, plus content for the XML/rels parts lint may inspect
   * (media and other binary payloads are never loaded).
   */
  @SuppressWarnings(Array("org.wartremover.warts.Var", "org.wartremover.warts.While"))
  private def readEntries(bytes: Array[Byte]): XLResult[MapPartSource] =
    try
      val zip = new ZipInputStream(new ByteArrayInputStream(bytes))
      val names = Set.newBuilder[String]
      val parts = Map.newBuilder[String, String]
      try
        var entry = zip.getNextEntry
        var sawEntry = false
        while entry != null do
          if !entry.isDirectory then
            sawEntry = true
            val name = entry.getName
            names += name
            if name.endsWith(".xml") || name.endsWith(".rels") then
              parts += name -> new String(zip.readAllBytes(), StandardCharsets.UTF_8)
          entry = zip.getNextEntry
        // ZipInputStream yields zero entries (not an error) for non-zip bytes — reject explicitly
        if !sawEntry then Left(XLError.ParseError("package", "Not a ZIP archive (no entries)"))
        else Right(MapPartSource(names.result(), parts.result()))
      finally zip.close()
    catch case e: Exception => Left(XLError.IOError(s"Failed to read ZIP bytes: ${e.getMessage}"))

  // ===== Pipeline =====

  private def lintSource(parts: PartSource): XLResult[Vector[Finding]] =
    for
      wbXmlOpt <- parts.read(workbookPart)
      wbXml <- wbXmlOpt.toRight(
        XLError.ParseError(workbookPart, s"Missing required part: $workbookPart")
      )
      wbElem <- XmlSecurity.parseSafe(wbXml, workbookPart)
      _ <- Either.cond(
        wbElem.label == "workbook",
        (),
        XLError.ParseError(workbookPart, s"Root element is <${wbElem.label}>, expected <workbook>")
      )
      wbRels <- readRelationships(parts, workbookRelsPart)
      sheetFindings <- lintWorksheets(wbElem, wbRels, parts)
    yield checkChildOrder(workbookPart, wbElem, OoxmlWorkbook.canonicalChildOrder, "CT_Workbook") ++
      checkRelRefs(workbookPart, workbookRelRefs(wbElem), wbRels, workbookRelsPart, parts, "xl") ++
      sheetFindings

  /** Parse a `.rels` part; absent part means no relationships (every r:id is then unresolved). */
  private def readRelationships(parts: PartSource, relsPath: String): XLResult[Relationships] =
    parts.read(relsPath).flatMap {
      case None => Right(Relationships.empty)
      case Some(xml) =>
        XmlSecurity
          .parseSafe(xml, relsPath)
          .flatMap { elem =>
            Relationships.fromXml(elem).left.map(err => XLError.ParseError(relsPath, err))
          }
    }

  /**
   * Lint every worksheet part reachable from a worksheet-typed sheet rel: CT_Worksheet child order
   * plus the sheet-level r:id references against the sheet's own `.rels`. Sheets whose rel is
   * unresolved / wrong-typed / missing are already reported at workbook level and skipped here.
   */
  private def lintWorksheets(
    wbElem: Elem,
    wbRels: Relationships,
    parts: PartSource
  ): XLResult[Vector[Finding]] =
    val sheetPaths = nestedElems(wbElem, "sheets", "sheet")
      .flatMap { e =>
        relIdOf(e)
          .flatMap(wbRels.findById)
          .filter(rel => rel.`type` == XmlUtil.relTypeWorksheet)
          .map(rel => Relationships.resolveWorkbookTarget(rel.target))
      }
      .distinct
      .filter(parts.has)

    sheetPaths.foldLeft[XLResult[Vector[Finding]]](Right(Vector.empty)) { (acc, path) =>
      for
        found <- acc
        xmlOpt <- parts.read(path)
        xml <- xmlOpt.toRight(XLError.ParseError(path, s"Missing part: $path"))
        elem <- XmlSecurity.parseSafe(xml, path)
        relsPath = siblingRelsPath(path)
        rels <- readRelationships(parts, relsPath)
      yield
        // A worksheet-typed rel pointing at non-worksheet CONTENT (root is <styleSheet>, ...) is
        // the zip-patch renumber class: the rel Type lies about what the target actually is.
        if elem.label != "worksheet" then
          found :+ Finding(
            path,
            LintCategory.WrongRelType,
            s"<${elem.label}>",
            s"Part is the target of a worksheet-typed relationship but its root element is <${elem.label}>, expected <worksheet>"
          )
        else
          found ++
            checkChildOrder(path, elem, worksheetCanonicalOrder, "CT_Worksheet") ++
            checkRelRefs(path, worksheetRelRefs(elem), rels, relsPath, parts, parentDir(path))
    }

  // ===== Child-order check =====

  /**
   * The observed sequence of RECOGNIZED child labels must be non-decreasing in canonical-order
   * index (repeats allowed: cols, conditionalFormatting, fileRecoveryPr are maxOccurs>1). Unknown
   * labels — mc:AlternateContent, xr:revisionPtr, vendor extensions — are position-transparent.
   * Each adjacent inversion yields one finding naming the misordered pair.
   */
  private def checkChildOrder(
    part: String,
    root: Elem,
    canonical: Vector[String],
    schemaName: String
  ): Vector[Finding] =
    val index: Map[String, Int] = canonical.zipWithIndex.toMap
    val recognized = root.child.toVector
      .collect { case e: Elem => e }
      .zipWithIndex
      .collect {
        case (e, i) if inMainNamespace(e) && index.contains(e.label) => (e.label, i + 1)
      }
    recognized.zip(recognized.drop(1)).collect {
      case ((prevLabel, _), (label, pos)) if index(label) < index(prevLabel) =>
        Finding(
          part,
          LintCategory.ChildOrder,
          s"<$label> (element #$pos)",
          s"$schemaName schema order requires <$label> before <$prevLabel>, but it appears after it (Excel repairs this)"
        )
    }

  /** True when the element is in the main SpreadsheetML namespace (or none — lenient). */
  private def inMainNamespace(e: Elem): Boolean =
    Option(e.namespace).forall(ns => ns.isEmpty || ns == XmlUtil.nsSpreadsheetML)

  // ===== r:id resolution checks =====

  /** One r:id-bearing element to verify against its part's paired `.rels`. */
  private final case class RelRef(
    elem: Elem,
    label: String,
    relIdRequired: Boolean,
    expectedTypes: Set[String],
    expectedDesc: String
  )

  /** A `<sheet>` may target a worksheet, chartsheet, or dialogsheet part. */
  private val sheetRelTypes: Set[String] =
    Set(XmlUtil.relTypeWorksheet, XmlUtil.relTypeChartsheet, XmlUtil.relTypeDialogsheet)

  private def workbookRelRefs(wbElem: Elem): Vector[RelRef] =
    nestedElems(wbElem, "sheets", "sheet").map(
      RelRef(_, "sheet", relIdRequired = true, sheetRelTypes, "worksheet")
    ) ++
      nestedElems(wbElem, "externalReferences", "externalReference").map(
        RelRef(
          _,
          "externalReference",
          relIdRequired = true,
          Set(XmlUtil.relTypeExternalLink),
          "externalLink"
        )
      ) ++
      nestedElems(wbElem, "pivotCaches", "pivotCache").map(
        RelRef(
          _,
          "pivotCache",
          relIdRequired = true,
          Set(XmlUtil.relTypePivotCacheDefinition),
          "pivotCacheDefinition"
        )
      )

  private def worksheetRelRefs(wsElem: Elem): Vector[RelRef] =
    def direct(label: String, types: Set[String], desc: String): Vector[RelRef] =
      childElems(wsElem, label).map(RelRef(_, label, relIdRequired = true, types, desc))
    direct("drawing", Set(XmlUtil.relTypeDrawing), "drawing") ++
      direct("legacyDrawing", Set(XmlUtil.relTypeVmlDrawing), "vmlDrawing") ++
      direct("legacyDrawingHF", Set(XmlUtil.relTypeVmlDrawing), "vmlDrawing") ++
      direct("picture", Set(XmlUtil.relTypeImage), "image") ++
      // hyperlink r:id is OPTIONAL — internal links carry location only (GH-235)
      nestedElems(wsElem, "hyperlinks", "hyperlink").map(
        RelRef(_, "hyperlink", relIdRequired = false, Set(XmlUtil.relTypeHyperlink), "hyperlink")
      ) ++
      nestedElems(wsElem, "tableParts", "tablePart").map(
        RelRef(_, "tablePart", relIdRequired = true, Set(XmlUtil.relTypeTable), "table")
      )

  /**
   * Resolve each reference in the paired `.rels`: the id must exist (UnresolvedRelId), the
   * relationship Type must match the referencing element (WrongRelType), and an internal-mode
   * target must exist in the package (MissingPart). External-mode targets are never checked for
   * existence.
   */
  private def checkRelRefs(
    part: String,
    refs: Vector[RelRef],
    rels: Relationships,
    relsPath: String,
    parts: PartSource,
    baseDir: String
  ): Vector[Finding] =
    refs.flatMap { ref =>
      val relId = relIdOf(ref.elem)
      val nameAttr = XmlUtil.getAttrOpt(ref.elem, "name").fold("")(n => s""" name="$n"""")
      val idAttr = relId.fold("")(id => s""" r:id="$id"""")
      val locator = s"<${ref.label}$nameAttr$idAttr>"
      relId match
        case None if ref.relIdRequired =>
          Vector(
            Finding(
              part,
              LintCategory.UnresolvedRelId,
              locator,
              s"<${ref.label}> has no r:id attribute, so it cannot resolve in $relsPath"
            )
          )
        case None => Vector.empty
        case Some(id) =>
          rels.findById(id) match
            case None =>
              Vector(
                Finding(
                  part,
                  LintCategory.UnresolvedRelId,
                  locator,
                  s"""r:id "$id" does not resolve: no Relationship with that Id in $relsPath"""
                )
              )
            case Some(rel) if !ref.expectedTypes.contains(rel.`type`) =>
              Vector(
                Finding(
                  part,
                  LintCategory.WrongRelType,
                  locator,
                  s"""r:id "$id" resolves to a ${shortType(
                      rel.`type`
                    )}-typed relationship (target "${rel.target}"), expected ${ref.expectedDesc}"""
                )
              )
            case Some(rel) if rel.targetMode.contains("External") => Vector.empty
            case Some(rel) =>
              val target = resolveTarget(baseDir, rel.target)
              if parts.has(target) then Vector.empty
              else
                Vector(
                  Finding(
                    part,
                    LintCategory.MissingPart,
                    locator,
                    s"""r:id "$id" resolves to target "${rel.target}" but part "$target" is not in the package"""
                  )
                )
    }

  // ===== Small helpers =====

  private def relIdOf(e: Elem): Option[String] =
    e.attribute(XmlUtil.nsRelationships, "id").map(_.text)

  /** Direct child elements with the given label (namespace-lenient, label match). */
  private def childElems(parent: Elem, label: String): Vector[Elem] =
    parent.child.toVector.collect { case e: Elem if e.label == label => e }

  /** Grandchild elements: children of `containerLabel` children with `childLabel`. */
  private def nestedElems(parent: Elem, containerLabel: String, childLabel: String): Vector[Elem] =
    childElems(parent, containerLabel).flatMap(childElems(_, childLabel))

  /** Last path segment of a relationship type URI, for readable messages ("styles", "table"). */
  private def shortType(relType: String): String =
    relType.split('/').lastOption.getOrElse(relType)

  /**
   * `_rels` sibling path for a part: xl/worksheets/sheet1.xml ->
   * xl/worksheets/_rels/sheet1.xml.rels
   */
  private def siblingRelsPath(partPath: String): String =
    val idx = partPath.lastIndexOf('/')
    if idx < 0 then s"_rels/$partPath.rels"
    else s"${partPath.substring(0, idx)}/_rels/${partPath.substring(idx + 1)}.rels"

  private def parentDir(partPath: String): String =
    val idx = partPath.lastIndexOf('/')
    if idx < 0 then "" else partPath.substring(0, idx)

  /**
   * Resolve a rels target to a package path against the referencing part's directory. A leading
   * slash is package-absolute; `..` segments are normalized (sheet rels use `../drawings/...`).
   */
  private def resolveTarget(baseDir: String, target: String): String =
    val resolved =
      if target.startsWith("/") then Paths.get(target.drop(1))
      else Paths.get(baseDir).resolve(target)
    resolved.normalize().toString.replace('\\', '/')
