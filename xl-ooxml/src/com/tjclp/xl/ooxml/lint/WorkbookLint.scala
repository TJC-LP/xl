package com.tjclp.xl.ooxml.lint

import java.io.{ByteArrayInputStream, InputStream}
import java.nio.charset.StandardCharsets
import java.nio.file.{Path, Paths}
import java.util.zip.{ZipFile, ZipInputStream}

import scala.xml.Elem

import com.tjclp.xl.addressing.{ARef, Column, Row}
import com.tjclp.xl.cells.FormulaKind
import com.tjclp.xl.error.{XLError, XLResult}
import com.tjclp.xl.ooxml.{
  ContentTypes,
  FormulaKindCodec,
  OoxmlWorkbook,
  Relationships,
  XmlSecurity,
  XmlUtil
}
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

  /** A present-and-referenced part has no Override and no extension Default in [Content_Types]. */
  case MissingContentType

  /** A ref/sqref/dimension token lies past row 1048576 or column XFD (GH-428 corruption class). */
  case RefOutOfBounds

  /** A `<f t="dataTable">` record whose grid no longer holds together (GH-442). */
  case DataTableTorn

  /** An uncached data-table interior in a `calcMode="autoNoTable"` book — opens BLANK (GH-442). */
  case DataTableUnseeded

  /**
   * `<f>` text stored with the display form's leading '=' — non-spec, misreads in strict tools
   * (GH-456).
   */
  case FormulaLeadingEquals

  /** Stable kebab-case identifier used in CLI text and JSON output. */
  def slug: String = this match
    case LintCategory.ChildOrder => "child-order"
    case LintCategory.UnresolvedRelId => "unresolved-rel-id"
    case LintCategory.WrongRelType => "wrong-rel-type"
    case LintCategory.MissingPart => "missing-part"
    case LintCategory.MissingContentType => "missing-content-type"
    case LintCategory.RefOutOfBounds => "ref-out-of-bounds"
    case LintCategory.DataTableTorn => "data-table-torn"
    case LintCategory.DataTableUnseeded => "data-table-unseeded"
    case LintCategory.FormulaLeadingEquals => "formula-leading-equals"

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
 * (GH-397, extended by GH-413):
 *
 *   - CT_Workbook / CT_Worksheet / CT_Chartsheet / CT_Dialogsheet child-order violations
 *   - `r:id` references that do not resolve in the paired `.rels` (or resolve to the wrong
 *     relationship/part type), including the externalLink part's own `<externalBook r:id>` hop
 *   - present-and-referenced parts missing their [Content_Types].xml registration (a part covered
 *     only by a generic extension Default is treated as registered — content-type correctness is
 *     out of scope)
 *   - ref/sqref/dimension tokens past row 1048576 / column XFD (the GH-428 overflow class Excel
 *     refuses as corrupt)
 *   - `<f t="dataTable">` records whose grid is torn or whose interior is uncached in a book that
 *     never recomputes tables (GH-442) — the only checks here that are silently WRONG output rather
 *     than an Excel repair: Excel refuses these edits in the UI, so a file carrying one was written
 *     past its own guards
 *   - `<f>` text stored with the display form's leading '=' (GH-456) — non-spec; Excel tolerates it
 *     but strict readers (openpyxl) misread it, and re-writing the file with xl heals it
 *
 * Lint runs on the RAW ZIP PARTS, never on the parsed domain model — a full read would
 * repair/normalize the very structure lint inspects (the reader silently falls back on unresolved
 * sheet r:ids and drops dangling hyperlink r:ids). Total: package-level failures (unreadable zip,
 * missing/malformed core part) surface as `Left`, never as a throw; structural problems in a
 * parseable package surface as findings.
 *
 * Two scanning modes produce IDENTICAL findings (pinned by the parity suite): the default DOM mode
 * parses each sheet part with [[XmlSecurity.parseSafe]]; the streaming mode ([[lintStream]])
 * SAX-scans sheet-class and table parts so memory stays O(1) in the row count (GH-413 item 4).
 * Workbook, rels, [Content_Types].xml and externalLink parts are always DOM-parsed — they are
 * bounded-size regardless of data volume.
 */
object WorkbookLint:

  private val workbookPart = "xl/workbook.xml"
  private val workbookRelsPart = "xl/_rels/workbook.xml.rels"
  private val rootRelsPart = "_rels/.rels"
  private val contentTypesPart = "[Content_Types].xml"

  /** Access to package parts by zip entry name. */
  private trait PartSource:
    def has(name: String): Boolean
    def read(name: String): XLResult[Option[String]]
    def openStream(name: String): XLResult[Option[InputStream]]

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
    def openStream(name: String): XLResult[Option[InputStream]] =
      Option(zip.getEntry(name)) match
        case None => Right(None)
        case Some(entry) =>
          try Right(Some(zip.getInputStream(entry)))
          catch
            case e: Exception =>
              Left(XLError.IOError(s"Failed to open zip entry $name: ${e.getMessage}"))

  private final class MapPartSource(names: Set[String], parts: Map[String, String])
      extends PartSource:
    def has(name: String): Boolean = names.contains(name)
    def read(name: String): XLResult[Option[String]] = Right(parts.get(name))
    def openStream(name: String): XLResult[Option[InputStream]] =
      Right(parts.get(name).map(s => ByteArrayInputStream(s.getBytes(StandardCharsets.UTF_8))))

  /** Lint an XLSX file on disk. Only the structural parts are read (workbook, worksheets, rels). */
  def lint(path: Path): XLResult[Vector[Finding]] =
    openZipFile(path).flatMap { zip =>
      try lintSource(ZipFilePartSource(zip), streaming = false)
      finally zip.close()
    }

  /** Lint an XLSX package from raw bytes. */
  def lintBytes(bytes: Array[Byte]): XLResult[Vector[Finding]] =
    readEntries(bytes).flatMap(lintSource(_, streaming = false))

  /**
   * Streaming lint (GH-413 item 4): sheet-class and table parts are SAX-scanned instead of
   * DOM-parsed, so memory stays O(1) in the row count — use for 100k+-row files. Findings are
   * identical to [[lint]] (pinned by the parity suite).
   */
  def lintStream(path: Path): XLResult[Vector[Finding]] =
    openZipFile(path).flatMap { zip =>
      try lintSource(ZipFilePartSource(zip), streaming = true)
      finally zip.close()
    }

  /**
   * Streaming lint over raw bytes — the parity entry point for [[lintStream]]. The bytes are
   * already in memory, so this mode saves only the DOM allocation, not the input buffer.
   */
  def lintStreamBytes(bytes: Array[Byte]): XLResult[Vector[Finding]] =
    readEntries(bytes).flatMap(lintSource(_, streaming = true))

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

  private def lintSource(parts: PartSource, streaming: Boolean): XLResult[Vector[Finding]] =
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
      rootRels <- readRelationships(parts, rootRelsPart)
      sheetResult <- lintSheets(wbElem, wbRels, parts, streaming, autoNoTableOf(wbElem))
      externalResult <- lintExternalLinks(wbElem, wbRels, parts)
      referenced =
        presentInternalTargets(rootRels, "", rootRelsPart, parts) ++
          presentInternalTargets(wbRels, "xl", workbookRelsPart, parts) ++
          sheetResult._2 ++ externalResult._2
      ctFindings <- checkContentTypes(parts, referenced)
    yield checkChildOrder(
      workbookPart,
      mainChildLabelsOf(wbElem),
      OoxmlWorkbook.canonicalChildOrder,
      "CT_Workbook"
    ) ++
      checkRelRefs(workbookPart, workbookRelRefs(wbElem), wbRels, workbookRelsPart, parts, "xl") ++
      sheetResult._1 ++ externalResult._1 ++ ctFindings

  /**
   * True when the book declares `calcMode="autoNoTable"` — the house dialect in which Excel never
   * recomputes data tables, so an uncached interior opens BLANK (the GH-419 calc-mode doctrine).
   */
  private def autoNoTableOf(wbElem: Elem): Boolean =
    childElems(wbElem, "calcPr")
      .exists(e => XmlUtil.getAttrOpt(e, "calcMode").exists(_.trim == "autoNoTable"))

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

  // ===== Sheet-class parts (worksheet / chartsheet / dialogsheet) =====

  /** One sheet-part flavor: its rel type, expected root, order table, and checked r:id specs. */
  private final case class SheetKind(
    relType: String,
    expectedRoot: String,
    schemaName: String,
    canonical: Vector[String],
    refSpecs: Vector[RefSpec]
  )

  /** An r:id-bearing element class to verify against the part's paired `.rels`. */
  private final case class RefSpec(
    label: String,
    relIdRequired: Boolean,
    expectedTypes: Set[String],
    expectedDesc: String
  )

  private val drawingSpec = RefSpec("drawing", true, Set(XmlUtil.relTypeDrawing), "drawing")
  private val legacyDrawingSpec =
    RefSpec("legacyDrawing", true, Set(XmlUtil.relTypeVmlDrawing), "vmlDrawing")
  private val legacyDrawingHFSpec =
    RefSpec("legacyDrawingHF", true, Set(XmlUtil.relTypeVmlDrawing), "vmlDrawing")
  private val pictureSpec = RefSpec("picture", true, Set(XmlUtil.relTypeImage), "image")
  // hyperlink r:id is OPTIONAL — internal links carry location only (GH-235)
  private val hyperlinkSpec =
    RefSpec("hyperlink", false, Set(XmlUtil.relTypeHyperlink), "hyperlink")
  private val tablePartSpec = RefSpec("tablePart", true, Set(XmlUtil.relTypeTable), "table")

  private val worksheetKind = SheetKind(
    XmlUtil.relTypeWorksheet,
    "worksheet",
    "CT_Worksheet",
    worksheetCanonicalOrder,
    Vector(
      drawingSpec,
      legacyDrawingSpec,
      legacyDrawingHFSpec,
      pictureSpec,
      hyperlinkSpec,
      tablePartSpec
    )
  )

  /** CT_Chartsheet child sequence (ECMA-376 Part 1 §18.3.1.12, transitional sml.xsd). */
  private val chartsheetCanonicalOrder: Vector[String] = Vector(
    "sheetPr",
    "sheetViews",
    "sheetProtection",
    "customSheetViews",
    "pageMargins",
    "pageSetup",
    "headerFooter",
    "drawing",
    "legacyDrawing",
    "legacyDrawingHF",
    "drawingHF",
    "picture",
    "webPublishItems",
    "extLst"
  )

  private val chartsheetKind = SheetKind(
    XmlUtil.relTypeChartsheet,
    "chartsheet",
    "CT_Chartsheet",
    chartsheetCanonicalOrder,
    Vector(drawingSpec, legacyDrawingSpec, legacyDrawingHFSpec, pictureSpec)
  )

  /** CT_Dialogsheet child sequence (ECMA-376 Part 1 §18.3.1.34, transitional sml.xsd). */
  private val dialogsheetCanonicalOrder: Vector[String] = Vector(
    "sheetPr",
    "sheetViews",
    "sheetFormatPr",
    "sheetProtection",
    "customSheetViews",
    "printOptions",
    "pageMargins",
    "pageSetup",
    "headerFooter",
    "drawing",
    "legacyDrawing",
    "legacyDrawingHF",
    "drawingHF",
    "oleObjects",
    "controls",
    "extLst"
  )

  private val dialogsheetKind = SheetKind(
    XmlUtil.relTypeDialogsheet,
    "dialogsheet",
    "CT_Dialogsheet",
    dialogsheetCanonicalOrder,
    Vector(drawingSpec, legacyDrawingSpec, legacyDrawingHFSpec)
  )

  private val sheetKinds: Vector[SheetKind] = Vector(worksheetKind, chartsheetKind, dialogsheetKind)

  /** A `<sheet>` may target a worksheet, chartsheet, or dialogsheet part. */
  private val sheetRelTypes: Set[String] = sheetKinds.map(_.relType).toSet

  /**
   * Lint every sheet-class part reachable from a sheet rel: CT child order, the part-level r:id
   * references against the part's own `.rels`, ref/sqref bounds, data-table record integrity, and
   * the bounds of any referenced table parts. Sheets whose rel is unresolved / wrong-typed /
   * missing are already reported at workbook level and skipped here. Also returns the internal rel
   * targets seen (for the [Content_Types].xml registration check).
   */
  private def lintSheets(
    wbElem: Elem,
    wbRels: Relationships,
    parts: PartSource,
    streaming: Boolean,
    autoNoTable: Boolean
  ): XLResult[(Vector[Finding], Vector[(String, String)])] =
    val targets: Vector[(String, SheetKind)] = nestedElems(wbElem, "sheets", "sheet")
      .flatMap { e =>
        for
          id <- relIdOf(e)
          rel <- wbRels.findById(id)
          kind <- sheetKinds.find(_.relType == rel.`type`)
        yield (Relationships.resolveWorkbookTarget(rel.target), kind)
      }
      .distinctBy(_._1)
      .filter((path, _) => parts.has(path))

    targets
      .foldLeft[XLResult[(Vector[Finding], Vector[(String, String)], Set[String])]](
        Right((Vector.empty, Vector.empty, Set.empty))
      ) { case (acc, (path, kind)) =>
        for
          found <- acc
          scan <- scanPart(parts, path, streaming)
          relsPath = siblingRelsPath(path)
          rels <- readRelationships(parts, relsPath)
          rootMatches = scan.rootLabel == kind.expectedRoot
          tableResult <-
            if rootMatches then scanTableParts(parts, rels, parentDir(path), streaming, found._3)
            else Right((Vector.empty[Finding], found._3))
        yield
          // A sheet-typed rel pointing at other CONTENT (root is <styleSheet>, ...) is the zip-patch
          // renumber class: the rel Type lies about what the target actually is.
          val findings =
            if !rootMatches then
              Vector(
                Finding(
                  path,
                  LintCategory.WrongRelType,
                  s"<${scan.rootLabel}>",
                  s"Part is the target of a ${kind.expectedRoot}-typed relationship but its root element is <${scan.rootLabel}>, expected <${kind.expectedRoot}>"
                )
              )
            else
              checkChildOrder(path, scan.mainChildLabels, kind.canonical, kind.schemaName) ++
                checkRelRefs(
                  path,
                  relRefsFor(kind, scan.captures),
                  rels,
                  relsPath,
                  parts,
                  parentDir(path)
                ) ++
                scan.boundsFindings ++ scan.formulaFindings ++
                dataTableFindings(path, scan.dataTables, autoNoTable) ++ tableResult._1
          (
            found._1 ++ findings,
            found._2 ++ presentInternalTargets(rels, parentDir(path), relsPath, parts),
            tableResult._2
          )
      }
      .map(acc => (acc._1, acc._2))

  /**
   * Scan referenced table parts for out-of-bounds refs (`<table ref>`, nested autoFilter). Parts in
   * `alreadyScanned` are skipped — a table shared by two sheets' rels is reported once — and the
   * returned set carries every target scanned so far.
   */
  private def scanTableParts(
    parts: PartSource,
    rels: Relationships,
    baseDir: String,
    streaming: Boolean,
    alreadyScanned: Set[String]
  ): XLResult[(Vector[Finding], Set[String])] =
    val targets = rels.relationships.toVector
      .filter(r => r.`type` == XmlUtil.relTypeTable && !r.targetMode.contains("External"))
      .map(r => resolveTarget(baseDir, r.target))
      .distinct
      .filterNot(alreadyScanned.contains)
      .filter(parts.has)
    targets.foldLeft[XLResult[(Vector[Finding], Set[String])]](
      Right((Vector.empty, alreadyScanned))
    ) { (acc, target) =>
      for
        found <- acc
        scan <- scanPart(parts, target, streaming)
      yield (found._1 ++ scan.boundsFindings, found._2 + target)
    }

  // ===== externalLink parts (GH-413 item 2) =====

  /**
   * Valid `<externalBook r:id>` target types: the ECMA externalLinkPath plus the Microsoft
   * xlExternalLinkPath variants Excel itself writes for broken/special external-book paths (GH-458)
   * — a real in-the-wild state, not a corruption class. The expectedTypes membership test fires
   * before the TargetMode=External bail-out, so the set must carry all of them.
   */
  private val externalLinkPathTypes: Set[String] = Set(
    XmlUtil.relTypeExternalLinkPath,
    XmlUtil.relTypeXlPathMissing,
    XmlUtil.relTypeXlLibraryPath,
    XmlUtil.relTypeXlStartupPath,
    XmlUtil.relTypeXlAltStartupPath
  )

  /**
   * Lint each externalLink part reachable from a resolved workbook `externalReference`: the part's
   * own `<externalBook r:id>` must resolve in its sibling `.rels` to an externalLinkPath-typed
   * relationship — the one-more-hop of the workbook-level check. Parts are tiny (no data rows), so
   * both scanning modes DOM-parse them.
   */
  private def lintExternalLinks(
    wbElem: Elem,
    wbRels: Relationships,
    parts: PartSource
  ): XLResult[(Vector[Finding], Vector[(String, String)])] =
    val targets = nestedElems(wbElem, "externalReferences", "externalReference")
      .flatMap(e => relIdOf(e).flatMap(wbRels.findById))
      .filter(rel =>
        rel.`type` == XmlUtil.relTypeExternalLink && !rel.targetMode.contains("External")
      )
      .map(rel => Relationships.resolveWorkbookTarget(rel.target))
      .distinct
      .filter(parts.has)

    targets.foldLeft[XLResult[(Vector[Finding], Vector[(String, String)])]](
      Right((Vector.empty, Vector.empty))
    ) { (acc, path) =>
      for
        found <- acc
        xmlOpt <- parts.read(path)
        xml <- xmlOpt.toRight(XLError.ParseError(path, s"Missing part: $path"))
        elem <- XmlSecurity.parseSafe(xml, path)
        relsPath = siblingRelsPath(path)
        rels <- readRelationships(parts, relsPath)
      yield
        val findings =
          if elem.label != "externalLink" then
            Vector(
              Finding(
                path,
                LintCategory.WrongRelType,
                s"<${elem.label}>",
                s"Part is the target of an externalLink-typed relationship but its root element is <${elem.label}>, expected <externalLink>"
              )
            )
          else
            val refs = childElems(elem, "externalBook").map { book =>
              RelRef(
                "externalBook",
                relIdOf(book),
                None,
                relIdRequired = true,
                externalLinkPathTypes,
                "externalLinkPath (or a Microsoft xlExternalLinkPath variant)"
              )
            }
            checkRelRefs(path, refs, rels, relsPath, parts, parentDir(path))
        (
          found._1 ++ findings,
          found._2 ++ presentInternalTargets(rels, parentDir(path), relsPath, parts)
        )
    }

  // ===== [Content_Types].xml registration (GH-413 item 3) =====

  /** Internal-mode rel targets that exist in the package, tagged with the rels part naming them. */
  private def presentInternalTargets(
    rels: Relationships,
    baseDir: String,
    relsPath: String,
    parts: PartSource
  ): Vector[(String, String)] =
    rels.relationships.toVector
      .filterNot(_.targetMode.contains("External"))
      .map(rel => resolveTarget(baseDir, rel.target))
      .filter(parts.has)
      .map(target => (target, relsPath))

  /**
   * Every present-and-referenced part must be registered in [Content_Types].xml — an Override for
   * its exact part name or a Default for its extension (case-insensitive). A part covered only by a
   * generic Default (e.g. `xml` -> application/xml) counts as registered: content-type CORRECTNESS
   * is out of scope, absence is the Excel-repair class (GH-413 item 3). An absent
   * [Content_Types].xml yields a single finding — the rest of the package is still lintable.
   */
  private def checkContentTypes(
    parts: PartSource,
    referenced: Vector[(String, String)]
  ): XLResult[Vector[Finding]] =
    parts.read(contentTypesPart).flatMap {
      case None =>
        Right(
          Vector(
            Finding(
              contentTypesPart,
              LintCategory.MissingContentType,
              "<Types>",
              s"Package has no $contentTypesPart part — Excel cannot open the file"
            )
          )
        )
      case Some(xml) =>
        for
          elem <- XmlSecurity.parseSafe(xml, contentTypesPart)
          _ <- Either.cond(
            elem.label == "Types",
            (),
            XLError.ParseError(
              contentTypesPart,
              s"Root element is <${elem.label}>, expected <Types>"
            )
          )
          ct <- ContentTypes
            .fromXml(elem)
            .left
            .map(err => XLError.ParseError(contentTypesPart, err))
        yield referenced.distinctBy(_._1).flatMap { (path, via) =>
          if registeredIn(ct, path) then Vector.empty
          else
            Vector(
              Finding(
                contentTypesPart,
                LintCategory.MissingContentType,
                s"""<Override PartName="/$path">""",
                s"""Part "$path" is present and referenced ($via) but $contentTypesPart has no Override for "/$path" and no Default for its extension (Excel repairs this)"""
              )
            )
        }
    }

  private def registeredIn(ct: ContentTypes, path: String): Boolean =
    ct.overrides.contains(s"/$path") ||
      ContentTypes
        .extensionOf(path)
        .exists(ext => ct.defaults.keysIterator.exists(_.equalsIgnoreCase(ext)))

  // ===== Child-order check =====

  /**
   * The observed sequence of RECOGNIZED child labels must be non-decreasing in canonical-order
   * index (repeats allowed: cols, conditionalFormatting, fileRecoveryPr are maxOccurs>1). Unknown
   * labels — mc:AlternateContent, xr:revisionPtr, vendor extensions — are position-transparent.
   * Each adjacent inversion yields one finding naming the misordered pair.
   *
   * @param childLabels
   *   the root's main-namespace child labels with their 1-based position among ALL child elements
   */
  private def checkChildOrder(
    part: String,
    childLabels: Vector[(String, Int)],
    canonical: Vector[String],
    schemaName: String
  ): Vector[Finding] =
    val index: Map[String, Int] = canonical.zipWithIndex.toMap
    val recognized = childLabels.filter((label, _) => index.contains(label))
    recognized.zip(recognized.drop(1)).collect {
      case ((prevLabel, _), (label, pos)) if index(label) < index(prevLabel) =>
        Finding(
          part,
          LintCategory.ChildOrder,
          s"<$label> (element #$pos)",
          s"$schemaName schema order requires <$label> before <$prevLabel>, but it appears after it (Excel repairs this)"
        )
    }

  /** Main-namespace child labels with 1-based positions among all child elements. */
  private def mainChildLabelsOf(root: Elem): Vector[(String, Int)] =
    root.child.toVector
      .collect { case e: Elem => e }
      .zipWithIndex
      .collect { case (e, i) if inMainNamespace(e) => (e.label, i + 1) }

  /** True when the element is in the main SpreadsheetML namespace (or none — lenient). */
  private def inMainNamespace(e: Elem): Boolean =
    Option(e.namespace).forall(ns => ns.isEmpty || ns == XmlUtil.nsSpreadsheetML)

  // ===== Part scanning (shared DOM / SAX substrate) =====

  /** Attributes that carry A1-style range tokens (dimension/autoFilter/mergeCell/CF/DV/...). */
  private val refAttrNames = Vector("ref", "sqref")

  /** Direct-child element labels whose r:id the sheet-level check verifies. */
  private val topLevelRefLabels: Set[String] =
    Set("drawing", "legacyDrawing", "legacyDrawingHF", "picture")

  /** An r:id-bearing element observed while scanning a sheet part. */
  private final case class CapturedRef(label: String, relId: Option[String])

  /**
   * Everything the per-part checks need, produced identically by the DOM and SAX scanners: root
   * label, ordered main-namespace top-level labels (with positions among all top-level elements),
   * the captured r:id-bearing elements, the ref/sqref bounds findings, the leading-'=' formula
   * findings (GH-456), and the data-table record state accumulated over sheetData (GH-442).
   */
  private final case class SheetScan(
    rootLabel: String,
    mainChildLabels: Vector[(String, Int)],
    captures: Vector[CapturedRef],
    boundsFindings: Vector[Finding],
    formulaFindings: Vector[Finding],
    dataTables: Vector[RecordFacts]
  )

  private def scanPart(parts: PartSource, path: String, streaming: Boolean): XLResult[SheetScan] =
    if streaming then
      parts.openStream(path).flatMap {
        case None => Left(XLError.ParseError(path, s"Missing part: $path"))
        case Some(stream) =>
          try SheetStreamScanner.scan(path, stream)
          finally stream.close()
      }
    else
      for
        xmlOpt <- parts.read(path)
        xml <- xmlOpt.toRight(XLError.ParseError(path, s"Missing part: $path"))
        elem <- XmlSecurity.parseSafe(xml, path)
      yield scanElem(path, elem)

  /** DOM scanner: same observation rules as [[SheetStreamScanner]] (parity-pinned). */
  private def scanElem(part: String, root: Elem): SheetScan =
    val children = root.child.toVector.collect { case e: Elem => e }
    val captures = children.flatMap { e =>
      val own =
        if topLevelRefLabels.contains(e.label) then Vector(CapturedRef(e.label, relIdOf(e)))
        else Vector.empty
      val nested = e.label match
        case "hyperlinks" =>
          childElems(e, "hyperlink").map(h => CapturedRef("hyperlink", relIdOf(h)))
        case "tableParts" =>
          childElems(e, "tablePart").map(t => CapturedRef("tablePart", relIdOf(t)))
        case _ => Vector.empty
      own ++ nested
    }
    val bounds = root.descendant_or_self.toVector
      .collect { case e: Elem => e }
      .flatMap { e =>
        refAttrNames.flatMap { attr =>
          XmlUtil.getAttrOpt(e, attr).toList.flatMap(refBoundsFindings(part, e.label, attr, _))
        }
      }
    val cellObs = nestedElems(root, "sheetData", "row")
      .flatMap(childElems(_, "c"))
      .map(domCellObs)
    val dataTables = cellObs.foldLeft(Vector.empty[RecordFacts])(observeCell)
    val formulaEq = cellObs.flatMap(formulaEqualsFindings(part, _))
    SheetScan(root.label, mainChildLabelsOf(root), captures, bounds, formulaEq, dataTables)

  /** One `<c>` element's data-table facts (DOM side of the parity pair). */
  private def domCellObs(cell: Elem): CellObs =
    val formula = childElems(cell, "f").headOption
    CellObs(
      ref = XmlUtil.getAttrOpt(cell, "r").flatMap(ARef.parse(_).toOption),
      record =
        formula.flatMap(f => dataTableKindOf(XmlUtil.getAttrOpt(f, "t"), XmlUtil.getAttrOpt(f, _))),
      hasFormula = formula.isDefined,
      hasValue = childElems(cell, "v").nonEmpty || childElems(cell, "is").nonEmpty,
      leadingEquals = formula.exists(_.text.startsWith("="))
    )

  /**
   * SAX scanner: O(1) memory in the row count — state is the top-level label list, the captured
   * r:id elements, any bounds findings, and one [[RecordFacts]] per data-table record (counters
   * only, never a materialized interior); sheetData rows/cells are otherwise folded and dropped. No
   * early abort: the full parse also validates well-formedness, matching the DOM mode's Left on
   * malformed parts.
   */
  private object SheetStreamScanner:
    import org.xml.sax.{Attributes, InputSource, SAXException}
    import org.xml.sax.helpers.DefaultHandler

    def scan(part: String, stream: InputStream): XLResult[SheetScan] =
      try
        // GH-350: shared XXE hardening + benign-doctype strip, matching the parseSafe path
        val parser = XmlSecurity.secureSaxParserFactory().newSAXParser()
        val handler = new ScanHandler(part)
        parser.parse(InputSource(XmlSecurity.stripLeadingDoctypeStream(stream)), handler)
        handler.result.toRight(XLError.ParseError(part, "Empty document (no root element)"))
      catch
        case e: SAXException =>
          Left(XLError.ParseError(part, s"Malformed XML: ${e.getMessage}"))
        case e: java.io.IOException =>
          Left(XLError.IOError(s"Failed to read $part: ${e.getMessage}"))

    @SuppressWarnings(Array("org.wartremover.warts.Var"))
    private final class ScanHandler(part: String) extends DefaultHandler:
      private var rootLabel: Option[String] = None
      private var depth = 0
      private var parents: List[String] = Nil
      private var topPos = 0
      private val labels = Vector.newBuilder[(String, Int)]
      private val captures = Vector.newBuilder[CapturedRef]
      private val bounds = Vector.newBuilder[Finding]
      private val formulaEq = Vector.newBuilder[Finding]
      private var dataTables: Vector[RecordFacts] = Vector.empty
      private var cell: Option[CellObs] = None
      // True while inside the current cell's FIRST <f> and no text chunk has arrived yet, so
      // characters() inspects exactly the first character of the formula text (GH-456).
      private var awaitingFormulaText = false

      def result: Option[SheetScan] =
        rootLabel.map(
          SheetScan(
            _,
            labels.result(),
            captures.result(),
            bounds.result(),
            formulaEq.result(),
            dataTables
          )
        )

      override def startElement(
        uri: String,
        localName: String,
        qName: String,
        atts: Attributes
      ): Unit =
        val label = if localName.nonEmpty then localName else qName
        if depth == 0 then rootLabel = Some(label)
        else if depth == 1 then
          topPos += 1
          if uri.isEmpty || uri == XmlUtil.nsSpreadsheetML then labels += ((label, topPos))
          if topLevelRefLabels.contains(label) then captures += CapturedRef(label, relIdIn(atts))
        else if depth == 2 then
          val parent = parents.headOption.getOrElse("")
          if (parent == "hyperlinks" && label == "hyperlink") ||
            (parent == "tableParts" && label == "tablePart")
          then captures += CapturedRef(label, relIdIn(atts))
        // GH-442: sheetData cells — <c> sits at depth 3 under <row>, its <f>/<v> at depth 4.
        // The grandparent check keeps the observed set identical to the DOM walk's sheetData/row/c.
        if depth == 3 && label == "c" && parents.take(2) == List("row", "sheetData") then
          cell = Some(
            CellObs(
              ref = Option(atts.getValue("", "r")).flatMap(ARef.parse(_).toOption),
              record = None,
              hasFormula = false,
              hasValue = false,
              leadingEquals = false
            )
          )
        else if depth == 4 && parents.headOption.contains("c") then
          // GH-456: only the cell's first <f> is observed, matching domCellObs's headOption
          if label == "f" && cell.exists(!_.hasFormula) then awaitingFormulaText = true
          cell = cell.map { open =>
            if label == "f" && !open.hasFormula then
              open.copy(
                record = dataTableKindOf(
                  Option(atts.getValue("", "t")),
                  name => Option(atts.getValue("", name))
                ),
                hasFormula = true
              )
            else if label == "v" || label == "is" then open.copy(hasValue = true)
            else open
          }
        refAttrNames.foreach { attr =>
          Option(atts.getValue("", attr)).foreach { value =>
            bounds ++= refBoundsFindings(part, label, attr, value)
          }
        }
        parents = label :: parents
        depth += 1

      override def characters(ch: Array[Char], start: Int, length: Int): Unit =
        if awaitingFormulaText && length > 0 then
          awaitingFormulaText = false
          if ch(start) == '=' then cell = cell.map(_.copy(leadingEquals = true))

      override def endElement(uri: String, localName: String, qName: String): Unit =
        val label = if localName.nonEmpty then localName else qName
        parents = parents.drop(1)
        depth -= 1
        if label == "f" then awaitingFormulaText = false
        if depth == 3 && label == "c" then
          cell.foreach { obs =>
            dataTables = observeCell(dataTables, obs)
            formulaEq ++= formulaEqualsFindings(part, obs)
          }
          cell = None

      private def relIdIn(atts: Attributes): Option[String] =
        Option(atts.getValue(XmlUtil.nsRelationships, "id"))

  // ===== Data-table record integrity (GH-442) =====

  /**
   * One `<c>` element's data-table-relevant facts, produced identically by both scanners and folded
   * in DOCUMENT ORDER so each mode sees the same record state at every cell.
   */
  private final case class CellObs(
    ref: Option[ARef],
    record: Option[FormulaKind.DataTable],
    hasFormula: Boolean,
    hasValue: Boolean,
    leadingEquals: Boolean
  )

  /**
   * GH-456: `<f>` text beginning with '=' — the display form leaked into the store. Excel/LO
   * tolerate it, but strict readers (openpyxl) read the formula back doubled ("==B4*2") and
   * formula-text diff tooling sees phantom differences. Shared by both scanners for parity.
   */
  private def formulaEqualsFindings(part: String, obs: CellObs): Vector[Finding] =
    if !obs.leadingEquals then Vector.empty
    else
      Vector(
        Finding(
          part,
          LintCategory.FormulaLeadingEquals,
          obs.ref.fold("<f>")(r => s"""<c r="${r.toA1}"><f>"""),
          "Formula text is stored with a leading '=' — OOXML carries the expression without it, " +
            "and strict readers (openpyxl) see a doubled '=' formula; re-writing the file with " +
            "xl heals it"
        )
      )

  /**
   * Accumulated facts for one data-table record, keyed by the record's own `ref`. Every field is a
   * counter: the interior is never materialized, so a crafted `ref="B2:XFD1048576"` (~1.7e10 cells)
   * costs the same as a 3x2 sensitivity grid.
   */
  private final case class RecordFacts(
    kind: FormulaKind.DataTable,
    cornerRecord: Boolean,
    cachedCells: Long,
    tornFirst: Option[ARef],
    tornCells: Long
  )

  /** The `<f>` attributes as a data-table record, via the shared CT_CellFormula codec (GH-430). */
  private def dataTableKindOf(
    t: Option[String],
    get: String => Option[String]
  ): Option[FormulaKind.DataTable] =
    FormulaKindCodec.fromAttrs(t, get).collect { case dt: FormulaKind.DataTable => dt }

  /**
   * Fold one cell into a sheet part's record state.
   *
   * A record is registered when its `<f>` is first seen, so interior accounting covers the cells at
   * or after it. Both the corner-only dialect Excel writes and the every-interior dialect found in
   * the field put a record on the ref's top-left, which is the first interior cell in document
   * order; a ref whose top-left carries no record is reported torn, so no tear can hide in the
   * cells that streamed before registration.
   */
  private def observeCell(facts: Vector[RecordFacts], obs: CellObs): Vector[RecordFacts] =
    val registered = obs.record match
      case Some(dt) if !facts.exists(_.kind.ref == dt.ref) =>
        facts :+ RecordFacts(dt, obs.ref.contains(dt.ref.start), 0L, None, 0L)
      case Some(dt) =>
        facts.map(f =>
          if f.kind.ref == dt.ref && obs.ref.contains(dt.ref.start) then f.copy(cornerRecord = true)
          else f
        )
      case None => facts
    // A `<c>` without `r` is positionally addressed; lint is not a syntax validator and skips it.
    obs.ref match
      case None => registered
      case Some(cellRef) =>
        registered.map { f =>
          if !f.kind.ref.contains(cellRef) then f
          else if obs.record.isEmpty && obs.hasFormula then
            f.copy(tornFirst = f.tornFirst.orElse(Some(cellRef)), tornCells = f.tornCells + 1)
          else if obs.hasValue then f.copy(cachedCells = f.cachedCells + 1)
          else f
        }

  /**
   * Findings for the records observed in one sheet part, in row-major record order.
   *
   * `data-table-torn` reports each way the grid stopped holding together. `data-table-unseeded`
   * reports an interior that opens blank, and is suppressed for records seeding cannot fix — the
   * skip conditions of `DataTableSeeder` (del flags, no room for the corner/axes, missing inputs) —
   * so the finding never advises seeding a record that cannot be seeded.
   */
  private def dataTableFindings(
    part: String,
    facts: Vector[RecordFacts],
    autoNoTable: Boolean
  ): Vector[Finding] =
    facts
      .sortBy(f => (f.kind.ref.start.row.index0, f.kind.ref.start.col.index0))
      .flatMap { f =>
        val locator = recordLocator(f.kind)
        tornMessages(f).map(Finding(part, LintCategory.DataTableTorn, locator, _)) ++
          unseededMessage(f, autoNoTable)
            .map(Finding(part, LintCategory.DataTableUnseeded, locator, _))
      }

  /**
   * The record's `<f>` element as the file carries it (bare ref on a 1x1 interior, per the codec).
   */
  private def recordLocator(kind: FormulaKind.DataTable): String =
    val refA1 = if kind.ref.start == kind.ref.end then kind.ref.start.toA1 else kind.ref.toA1
    s"""<f t="dataTable" ref="$refA1">"""

  /** Room for the corner formula one-up-one-left plus both axes (the GH-419 V1 geometry). */
  private def hasRoom(kind: FormulaKind.DataTable): Boolean =
    kind.ref.start.row.index0 >= 1 && kind.ref.start.col.index0 >= 1

  /** Mirrors the skip conditions of `DataTableSeeder` — a record it skips cannot be seeded. */
  private def seedable(kind: FormulaKind.DataTable): Boolean =
    !kind.del1 && !kind.del2 && hasRoom(kind) &&
      kind.r1.isDefined && (!kind.dt2D || kind.r2.isDefined)

  private def tornMessages(f: RecordFacts): Vector[String] =
    val ref = f.kind.ref.toA1
    val deleted = Vector(
      Option.when(f.kind.del1)("del1=\"1\""),
      Option.when(f.kind.del2)("del2=\"1\"")
    ).flatten
    // A del flag legitimately omits its input cell, so only an un-flagged absence is a loss.
    val missing = Vector(
      Option.when(f.kind.r1.isEmpty && !f.kind.del1)("r1"),
      Option.when(f.kind.dt2D && f.kind.r2.isEmpty && !f.kind.del2)("r2")
    ).flatten
    Vector(
      Option.when(!hasRoom(f.kind))(
        s"Data table record $ref leaves no room for its corner formula and axes (an interior " +
          "cannot start in row 1 or column A) — its axis row/column was structurally removed and " +
          "the grid can no longer recalculate"
      ),
      Option.when(deleted.nonEmpty)(
        s"Data table record $ref lost an input cell to a structural delete " +
          s"(${deleted.mkString(", ")}) — the grid can no longer recalculate"
      ),
      Option.when(missing.nonEmpty)(
        s"Data table record $ref is missing its required input reference " +
          s"(${missing.mkString(", ")}) — the grid can no longer recalculate"
      ),
      Option.when(!f.cornerRecord)(
        s"Data table record $ref has no record cell at its top-left ${f.kind.ref.start.toA1} — " +
          "the corner record was overwritten and the grid is torn"
      ),
      Option.when(f.tornCells > 0)(
        s"Data table record $ref has ${f.tornCells} formula cell(s) inside its interior (first " +
          s"${f.tornFirst.fold("unknown")(_.toA1)}) — a formula overwrite tore the record grid; " +
          "Excel refuses this edit (cannot change part of a data table)"
      )
    ).flatten

  private def unseededMessage(f: RecordFacts, autoNoTable: Boolean): Option[String] =
    val area = f.kind.ref.width.toLong * f.kind.ref.height.toLong
    val uncached = math.max(0L, area - f.cachedCells)
    Option.when(autoNoTable && seedable(f.kind) && uncached > 0)(
      s"Data table record ${f.kind.ref.toA1} has $uncached of $area interior cell(s) with no " +
        "cached value and the book is calcMode=\"autoNoTable\", which never recomputes data " +
        "tables — the table opens BLANK (seed it with: xl recalc --tables)"
    )

  // ===== ref/sqref bounds check (GH-428 corruption class) =====

  private val maxRow1Based: Long = Row.MaxIndex0.toLong + 1
  private val maxCol0Based: Long = Column.MaxIndex0.toLong

  /**
   * Flag every whitespace-separated `ref`/`sqref` token whose row exceeds 1048576 or whose column
   * lies beyond XFD — the GH-428 overflow class (a range ending at the sheet edge shifted past it
   * by a structural insert). Excel refuses such files as corrupt; LibreOffice opens them silently.
   * Tokens that are not plain A1-style cells/ranges (row/column spans, anchored refs) are ignored —
   * lint is not a syntax validator.
   */
  private def refBoundsFindings(
    part: String,
    elemLabel: String,
    attrName: String,
    value: String
  ): Vector[Finding] =
    value.split("\\s+").toVector.filter(_.nonEmpty).flatMap { token =>
      val cells = token.split(":", -1).toVector
      val parsed = if cells.sizeIs <= 2 then cells.map(cellTokenBounds) else Vector(None)
      if parsed.exists(_.isEmpty) then Vector.empty
      else
        val bounds = parsed.flatten
        val maxRow = bounds.map(_._2).maxOption.getOrElse(0L)
        val maxCol = bounds.map(_._1).maxOption.getOrElse(0L)
        val problems = Vector(
          Option.when(maxRow > maxRow1Based)(s"row $maxRow > $maxRow1Based"),
          Option.when(maxCol > maxCol0Based)(s"column ${columnName(maxCol)} > XFD")
        ).flatten
        if problems.isEmpty then Vector.empty
        else
          Vector(
            Finding(
              part,
              LintCategory.RefOutOfBounds,
              s"""<$elemLabel $attrName="$token">""",
              s"""$attrName "$token" lies outside the sheet's addressable range (${problems
                  .mkString(", ")}) — Excel refuses the file as corrupt"""
            )
          )
    }

  /** Lenient A1-token parse to (col0, row1) WITHOUT bounds validation — that is the whole point. */
  private def cellTokenBounds(token: String): Option[(Long, Long)] =
    val letters = token.takeWhile(c => (c >= 'A' && c <= 'Z') || (c >= 'a' && c <= 'z'))
    val digits = token.drop(letters.length)
    if letters.isEmpty || letters.length > 4 || digits.isEmpty || digits.length > 9 ||
      !digits.forall(c => c >= '0' && c <= '9')
    then None
    else
      val col0 = letters.foldLeft(0L)((acc, c) => acc * 26 + (c.toUpper - 'A' + 1)) - 1
      digits.toLongOption.map(row1 => (col0, row1))

  /** 0-based column index to letters (16383 -> "XFD", 16384 -> "XFE"). */
  private def columnName(col0: Long): String =
    @annotation.tailrec
    def loop(n: Long, acc: List[Char]): List[Char] =
      if n < 0 then acc
      else loop(n / 26 - 1, ('A' + (n % 26).toInt).toChar :: acc)
    loop(col0, Nil).mkString

  // ===== r:id resolution checks =====

  /** One r:id-bearing element occurrence to verify against its part's paired `.rels`. */
  private final case class RelRef(
    label: String,
    relId: Option[String],
    name: Option[String],
    relIdRequired: Boolean,
    expectedTypes: Set[String],
    expectedDesc: String
  )

  private def workbookRelRefs(wbElem: Elem): Vector[RelRef] =
    def relRef(e: Elem, label: String, types: Set[String], desc: String): RelRef =
      RelRef(label, relIdOf(e), XmlUtil.getAttrOpt(e, "name"), relIdRequired = true, types, desc)
    nestedElems(wbElem, "sheets", "sheet").map(relRef(_, "sheet", sheetRelTypes, "worksheet")) ++
      nestedElems(wbElem, "externalReferences", "externalReference").map(
        relRef(_, "externalReference", Set(XmlUtil.relTypeExternalLink), "externalLink")
      ) ++
      nestedElems(wbElem, "pivotCaches", "pivotCache").map(
        relRef(_, "pivotCache", Set(XmlUtil.relTypePivotCacheDefinition), "pivotCacheDefinition")
      )

  /**
   * Captured sheet-part elements resolved through the kind's spec table (grouped by spec order).
   */
  private def relRefsFor(kind: SheetKind, captures: Vector[CapturedRef]): Vector[RelRef] =
    kind.refSpecs.flatMap { spec =>
      captures.collect {
        case c if c.label == spec.label =>
          RelRef(
            spec.label,
            c.relId,
            None,
            spec.relIdRequired,
            spec.expectedTypes,
            spec.expectedDesc
          )
      }
    }

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
      val nameAttr = ref.name.fold("")(n => s""" name="$n"""")
      val idAttr = ref.relId.fold("")(id => s""" r:id="$id"""")
      val locator = s"<${ref.label}$nameAttr$idAttr>"
      ref.relId match
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
