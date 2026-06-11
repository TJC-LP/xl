package com.tjclp.xl.ooxml

import com.tjclp.xl.addressing.{ARef, Column, Row, SheetName}
import com.tjclp.xl.cells.{Cell, CellError, CellValue}
import com.tjclp.xl.api.{Sheet, Workbook}
import com.tjclp.xl.sheets.{
  ColumnProperties,
  HeaderFooter,
  PageMargins,
  PageSetup,
  RowProperties,
  SheetView
}
import com.tjclp.xl.error.{XLError, XLResult}
import com.tjclp.xl.context.{ModificationTracker, SourceContext, SourceFingerprint}
import com.tjclp.xl.tables.TableSpec

import scala.xml.*
import java.io.{ByteArrayInputStream, FileInputStream, InputStream}
import java.nio.file.{Files, Path, Paths}
import java.security.{DigestInputStream, MessageDigest}
import java.util.zip.ZipInputStream
import scala.collection.mutable
import scala.collection.immutable.ArraySeq
import com.tjclp.xl.styles.StyleRegistry
import com.tjclp.xl.styles.units.StyleId
import com.tjclp.xl.styles.color.ThemePalette
import com.tjclp.xl.workbooks.{DefinedName, WorkbookMetadata}

/**
 * Reader for XLSX files (ZIP parsing)
 *
 * Parses XLSX ZIP structure and converts to domain Workbook.
 *
 * **WARNING**: This is the in-memory reader that loads the entire file into memory. For large files
 * (>10k rows), use `ExcelIO.readStream()` instead for constant-memory streaming.
 *
 * Use this reader for:
 *   - Small files (<10k rows)
 *   - When you need random cell access
 *   - When you need styling information
 *
 * Use `ExcelIO.readStream()` for:
 *   - Large files (100k+ rows)
 *   - Sequential row processing
 *   - Constant-memory requirements (O(1) memory regardless of file size)
 */
object XlsxReader:

  /** Warning emitted while reading XLSX input (non-fatal). */
  enum Warning:
    case MissingStylesXml

    /** A drawing part was structurally unparseable; it rides byte-preservation only (GH-221). */
    case MalformedDrawingPart(path: String)

    /**
     * A chart part behind an otherwise-typeable graphicFrame was structurally unparseable; the
     * anchor stays Preserved and the part rides byte-preservation (GH-222). Charts merely outside
     * the typed subset are silently Preserved (the pic-fence precedent) — no warning.
     */
    case MalformedChartPart(path: String)

  /** Successful read result with accumulated warnings. */
  case class ReadResult(workbook: Workbook, warnings: Vector[Warning])

  /**
   * Configuration for XLSX reader with security limits.
   *
   * These limits protect against ZIP bombs and other denial-of-service attacks. All limits are
   * checked during ZIP entry processing - if any limit is exceeded, reading stops immediately with
   * a SecurityError.
   *
   * @param maxCompressionRatio
   *   Maximum allowed compression ratio (uncompressed/compressed). Default 100:1. Set to 0 to
   *   disable.
   * @param maxUncompressedSize
   *   Maximum total uncompressed size in bytes. Default 100 MB. Set to 0 to disable.
   * @param maxEntryCount
   *   Maximum number of ZIP entries. Default 10,000. Set to 0 to disable.
   * @param maxCellCount
   *   Maximum number of cells per sheet. Default 10 million. Set to 0 to disable.
   * @param maxStringLength
   *   Maximum string length for cell values. Default 32 KB (Excel's limit). Set to 0 to disable.
   */
  case class ReaderConfig(
    maxCompressionRatio: Int = 100,
    maxUncompressedSize: Long = 100_000_000L,
    maxEntryCount: Int = 10_000,
    maxCellCount: Long = 10_000_000L,
    maxStringLength: Int = 32_768
  )

  object ReaderConfig:
    /** Default configuration with sensible security limits. */
    val default: ReaderConfig = ReaderConfig()

    /** Permissive configuration with no limits (use only for trusted files). */
    val permissive: ReaderConfig = ReaderConfig(
      maxCompressionRatio = 0,
      maxUncompressedSize = 0L,
      maxEntryCount = 0,
      maxCellCount = 0L,
      maxStringLength = 0
    )

  /**
   * Handle for accessing source file during read (enables surgical modification).
   *
   * Provides the path to the original file, which allows XlsxReader to create a SourceContext with
   * PreservedPartStore for unknown parts.
   */
  case class SourceHandle(path: Path, size: Long, digest: MessageDigest):
    def finalizeFingerprint(): SourceFingerprint =
      SourceFingerprint(size, ArraySeq.unsafeWrapArray(digest.digest()))

  /** Set of ZIP entry paths that XL knows how to parse. All other entries are preserved. */
  private val knownParts: Set[String] = Set(
    "[Content_Types].xml",
    "_rels/.rels",
    "xl/workbook.xml",
    "xl/_rels/workbook.xml.rels",
    "xl/styles.xml",
    "xl/sharedStrings.xml",
    "xl/theme/theme1.xml",
    // GH-242: docProps are parsed into WorkbookMetadata and re-emitted from the model on write
    // (NOT preserved verbatim — the model wins). docProps/custom.xml stays unparsed/preserved.
    DocProps.corePath,
    DocProps.appPath
  )

  /**
   * Checks if the given ZIP entry path is a known part that XL can parse.
   *
   * Known parts include:
   *   - Core OOXML parts (workbook, styles, sharedStrings)
   *   - Worksheets (pattern: xl/worksheets/sheet*.xml)
   *   - Relationships
   *
   * All other parts (charts, drawings, images, etc.) are preserved byte-for-byte.
   */
  private def isKnownPart(path: String): Boolean =
    knownParts.contains(path) ||
      path.matches("xl/worksheets/sheet\\d+\\.xml") ||
      path.matches("xl/comments\\d+\\.xml") ||
      path.matches("xl/tables/table\\d+\\.xml") ||
      path.matches("xl/worksheets/_rels/sheet\\d+\\.xml\\.rels")

  /**
   * Drawing parts and media retained as raw content during the zip scan (GH-221). These entries
   * stay `recordUnparsed` in the manifest — their unparsed status feeds the surgical writer's
   * preservation machinery (determinePreservableParts), which must keep copying them verbatim. Only
   * the media bytes referenced by parsed Pictures survive into the model; this scan-time container
   * is discarded after parsing.
   */
  private case class RetainedDrawingParts(
    xml: Map[String, String], // drawingN.xml and drawingN.xml.rels, as UTF-8
    charts: Map[String, String], // chartN.xml and chartN.xml.rels (GH-222), as UTF-8
    media: Map[String, ArraySeq[Byte]] // xl/media/*
  )

  private def isDrawingXmlPart(path: String): Boolean =
    path.matches("xl/drawings/drawing\\d+\\.xml") ||
      path.matches("xl/drawings/_rels/drawing\\d+\\.xml\\.rels")

  /**
   * Chart parts retained for the typed chart layer (GH-222); stay `recordUnparsed` like drawings.
   */
  private def isChartXmlPart(path: String): Boolean =
    path.matches("xl/charts/chart\\d+\\.xml") ||
      path.matches("xl/charts/_rels/chart\\d+\\.xml\\.rels")

  /**
   * Read workbook from XLSX file (in-memory).
   *
   * Loads entire file into memory. For large files, use `ExcelIO.readStream()` instead.
   *
   * **Surgical Modification**: When reading from a file (not a stream), XL creates a SourceContext
   * that enables surgical writes. Unknown parts (charts, images, etc.) are indexed but not loaded
   * into memory, allowing them to be preserved byte-for-byte on write.
   *
   * @param inputPath
   *   Path to XLSX file
   * @param config
   *   Reader configuration with security limits. Default applies sensible limits.
   */
  def read(inputPath: Path, config: ReaderConfig = ReaderConfig.default): XLResult[Workbook] =
    readWithWarnings(inputPath, config).map(_.workbook)

  /** Read workbook and surface non-fatal warnings. */
  def readWithWarnings(
    inputPath: Path,
    config: ReaderConfig = ReaderConfig.default
  ): XLResult[ReadResult] =
    try
      val size = Files.size(inputPath)
      val digest = MessageDigest.getInstance("SHA-256")
      val fileStream = new FileInputStream(inputPath.toFile)
      val digestStream = new DigestInputStream(fileStream, digest)
      try
        readFromStreamWithWarnings(
          digestStream,
          Some(SourceHandle(inputPath, size, digest)),
          config
        )
      finally
        digestStream.close()
    catch case e: Exception => Left(XLError.IOError(s"Failed to read XLSX: ${e.getMessage}"))

  /** Read workbook from byte array (for testing) */
  def readFromBytes(
    bytes: Array[Byte],
    config: ReaderConfig = ReaderConfig.default
  ): XLResult[Workbook] =
    readFromBytesWithWarnings(bytes, config).map(_.workbook)

  def readFromBytesWithWarnings(
    bytes: Array[Byte],
    config: ReaderConfig = ReaderConfig.default
  ): XLResult[ReadResult] =
    try readFromStreamWithWarnings(new ByteArrayInputStream(bytes), None, config)
    catch case e: Exception => Left(XLError.IOError(s"Failed to read bytes: ${e.getMessage}"))

  /** Read workbook from input stream (no surgical modification support) */
  @SuppressWarnings(Array("org.wartremover.warts.Var", "org.wartremover.warts.While"))
  private def readFromStream(
    is: InputStream,
    config: ReaderConfig = ReaderConfig.default
  ): XLResult[Workbook] =
    readFromStreamWithWarnings(is, None, config).map(_.workbook)

  /**
   * Read workbook from input stream with optional source handle for surgical modification.
   *
   * @param is
   *   Input stream containing XLSX ZIP data
   * @param source
   *   Optional source handle providing path to original file. If provided, creates SourceContext
   *   with indexed unknown parts for surgical writes.
   * @param config
   *   Reader configuration with security limits
   * @return
   *   ReadResult with workbook and warnings, or SecurityError if limits exceeded
   */
  @SuppressWarnings(Array("org.wartremover.warts.Var", "org.wartremover.warts.While"))
  private def readFromStreamWithWarnings(
    is: InputStream,
    source: Option[SourceHandle],
    config: ReaderConfig
  ): XLResult[ReadResult] =
    val zip = new ZipInputStream(is)
    val parts = mutable.Map[String, String]()
    val drawingXml = mutable.Map[String, String]()
    val chartXml = mutable.Map[String, String]()
    val mediaBytes = mutable.Map[String, ArraySeq[Byte]]()
    val manifestBuilder = PartManifestBuilder.empty

    // Security tracking
    var entryCount = 0
    var totalUncompressedSize = 0L
    var securityError: Option[XLError] = None

    try
      // Read all ZIP entries - index ALL entries, parse only KNOWN ones
      var builder = manifestBuilder
      var entry = zip.getNextEntry
      while entry != null && securityError.isEmpty do
        if !entry.isDirectory then
          val entryName = entry.getName
          entryCount += 1

          // Security check: Entry count limit
          if config.maxEntryCount > 0 && entryCount > config.maxEntryCount then
            securityError = Some(
              XLError.SecurityError(
                s"ZIP entry count ($entryCount) exceeds limit (${config.maxEntryCount})"
              )
            )
          else
            // Record entry metadata in manifest (size, CRC, etc.)
            builder = builder.+=(entry)

            // Read content with size tracking
            val content = zip.readAllBytes()
            val uncompressedSize = content.length.toLong
            totalUncompressedSize += uncompressedSize

            // Security check: Total uncompressed size
            if config.maxUncompressedSize > 0 && totalUncompressedSize > config.maxUncompressedSize
            then
              securityError = Some(
                XLError.SecurityError(
                  s"Total uncompressed size ($totalUncompressedSize bytes) exceeds limit (${config.maxUncompressedSize} bytes)"
                )
              )

            // Security check: Compression ratio (ZIP bomb detection)
            val compressedSize = entry.getCompressedSize
            if securityError.isEmpty && config.maxCompressionRatio > 0 && compressedSize > 0 then
              val ratio = uncompressedSize.toDouble / compressedSize.toDouble
              if ratio > config.maxCompressionRatio then
                securityError = Some(
                  XLError.SecurityError(
                    f"Compression ratio ($ratio%.1f:1) for '$entryName' exceeds limit (${config.maxCompressionRatio}:1) - possible ZIP bomb"
                  )
                )

            // Only parse known parts to save memory
            if securityError.isEmpty then
              if isKnownPart(entryName) then
                parts(entryName) = new String(content, "UTF-8")
                builder = builder.recordParsed(entryName)
              else
                // GH-221/GH-222: retain drawing/chart XML and media bytes for the drawing layer.
                // The entries stay UNPARSED in the manifest (byte-preservation contract); zero
                // extra IO — content was already read for security accounting. Memory stays
                // bounded by ReaderConfig.maxUncompressedSize.
                if isDrawingXmlPart(entryName) then
                  drawingXml(entryName) = new String(content, "UTF-8")
                else if isChartXmlPart(entryName) then
                  chartXml(entryName) = new String(content, "UTF-8")
                else if entryName.startsWith("xl/media/") then
                  mediaBytes(entryName) = ArraySeq.unsafeWrapArray(content)
                // Unknown part - index but don't store content
                builder = builder.recordUnparsed(entryName)

        zip.closeEntry()
        entry = zip.getNextEntry

      // Return early if security error occurred
      securityError match
        case Some(err) => Left(err)
        case None =>
          // Build final manifest
          val manifest = builder.build()

          // Drain the remaining container bytes through the digest before fingerprinting:
          // ZipInputStream stops consuming at the central directory, but copyVerbatim hashes
          // the whole file — without this drain the fingerprints could never match and every
          // clean file→file write failed loudly (GH-261).
          if source.isDefined then
            val drainBuf = new Array[Byte](8192)
            while is.read(drainBuf) != -1 do ()

          // Compute fingerprint if reading from a file
          val fingerprint = source.map(_.finalizeFingerprint())

          // Parse workbook structure from known parts
          parseWorkbook(
            parts.toMap,
            RetainedDrawingParts(drawingXml.toMap, chartXml.toMap, mediaBytes.toMap),
            source,
            manifest,
            fingerprint,
            config
          )

    finally zip.close()

  /**
   * Parse workbook from collected parts and optionally create SourceContext.
   *
   * @param parts
   *   Map of ZIP entry paths to their content (only known parts)
   * @param source
   *   Optional source handle for enabling surgical modification
   * @param manifest
   *   Part manifest with metadata for all ZIP entries (known + unknown)
   * @param config
   *   Reader configuration with security limits
   * @return
   *   ReadResult with workbook (with optional SourceContext) and warnings
   */
  private def parseWorkbook(
    parts: Map[String, String],
    retainedDrawings: RetainedDrawingParts,
    source: Option[SourceHandle],
    manifest: PartManifest,
    fingerprint: Option[SourceFingerprint],
    config: ReaderConfig
  ): XLResult[ReadResult] =
    for
      // Parse workbook.xml
      workbookXml <- parts
        .get("xl/workbook.xml")
        .toRight(XLError.ParseError("xl/workbook.xml", "Missing workbook.xml"))
      workbookElem <- parseXml(workbookXml, "xl/workbook.xml")
      ooxmlWb <- OoxmlWorkbook
        .fromXml(workbookElem)
        .left
        .map(err => XLError.ParseError("xl/workbook.xml", err): XLError)

      // Parse optional shared strings
      sst <- parseOptionalSST(parts)

      // Parse styles (if present)
      (styles, styleWarnings) <- parseStyles(parts)

      // Parse workbook relationships
      workbookRels <- parseWorkbookRelationships(parts)

      // Parse theme (optional, falls back to Office theme)
      theme = parseTheme(parts)

      // Parse sheets and collect comment/drawing mappings (GH-221)
      parsedSheets <- parseSheets(
        parts,
        retainedDrawings,
        ooxmlWb.sheets,
        sst,
        styles,
        workbookRels
      )
      (sheets, commentPathMapping) = (parsedSheets.sheets, parsedSheets.commentPathMapping)

      // Parse defined names from workbook.xml (shared with the surgical writer)
      definedNames = OoxmlWorkbook.parseDefinedNames(ooxmlWb.definedNames)

      // Extract sheet visibility states from OOXML
      sheetStates = ooxmlWb.sheets
        .filter(_.state.nonEmpty)
        .map(ref => ref.name -> ref.state)
        .toMap

      // GH-242: parse document properties (lenient — absent/malformed parts yield empty fields)
      docProps = parseDocProps(parts)

      // GH-294: bookViews/workbookView activeTab → activeSheetIndex, clamped leniently into
      // [0, sheetCount-1] (malformed or out-of-range values must never fail a read)
      activeSheetIndex = OoxmlWorkbook.clampActiveTab(
        OoxmlWorkbook.parseActiveTab(ooxmlWb.bookViews).getOrElse(0),
        sheets.size
      )

      // Assemble workbook with optional SourceContext
      workbook <- assembleWorkbook(
        sheets,
        source,
        manifest,
        fingerprint,
        theme,
        definedNames,
        sheetStates,
        commentPathMapping,
        date1904 = ooxmlWb.date1904,
        activeSheetIndex = activeSheetIndex,
        docProps,
        drawingPathMapping = parsedSheets.drawingPathMapping,
        drawingSnapshots = parsedSheets.drawingSnapshots,
        chartSnapshots = parsedSheets.chartSnapshots
      )
    yield ReadResult(workbook, styleWarnings ++ parsedSheets.warnings)

  /**
   * Parse docProps/core.xml + app.xml into the modeled metadata fields (GH-242).
   *
   * Lenient by design: docProps in the wild are messy and must never fail a read. Absent or
   * malformed parts contribute no fields; unmodeled fields (title, keywords, ...) are ignored.
   */
  private def parseDocProps(parts: Map[String, String]): DocProps.Data =
    def parsePart(path: String, parse: Elem => DocProps.Data): DocProps.Data =
      parts
        .get(path)
        .flatMap(xml => XmlSecurity.parseSafe(xml, path).toOption)
        .map(parse)
        .getOrElse(DocProps.Data.empty)
    DocProps.merge(
      parsePart(DocProps.corePath, DocProps.parseCoreXml),
      parsePart(DocProps.appPath, DocProps.parseAppXml)
    )

  /** Parse optional shared strings table */
  private def parseOptionalSST(parts: Map[String, String]): XLResult[Option[SharedStrings]] =
    parts.get("xl/sharedStrings.xml") match
      case None => Right(None)
      case Some(xml) =>
        for
          elem <- parseXml(xml, "xl/sharedStrings.xml")
          sst <- SharedStrings
            .fromXml(elem)
            .left
            .map(err => XLError.ParseError("xl/sharedStrings.xml", err): XLError)
        yield Some(sst)

  /**
   * Parse worksheet relationships to find comment, table, hyperlink, and drawing references.
   *
   * Returns (commentPath, tableTargets, hyperlinkRels, drawingPath) where:
   *   - commentPath: Optional path to comments file (e.g., "../comments1.xml")
   *   - tableTargets: Sequence of table file paths (e.g., ["../tables/table1.xml",
   *     "../tables/table2.xml"])
   *   - drawingPath: Optional resolved drawing part path (GH-221) — lenient: an unresolvable target
   *     yields None (the part still rides byte-preservation) rather than failing the read.
   */
  private def parseWorksheetRelationships(
    parts: Map[String, String],
    sheetIndex: Int
  ): XLResult[(Option[String], Seq[String], Map[String, String], Option[String])] =
    val relsPath = s"xl/worksheets/_rels/sheet$sheetIndex.xml.rels"
    parts.get(relsPath) match
      case None => Right((None, Seq.empty, Map.empty, None)) // No relationships for this sheet
      case Some(xml) =>
        for
          elem <- parseXml(xml, relsPath)
          rels <- Relationships
            .fromXml(elem)
            .left
            .map(err => XLError.ParseError(relsPath, err): XLError)

          // Extract comment relationship (existing logic)
          commentTarget <- rels.relationships
            .find(_.`type` == XmlUtil.relTypeComments) match
            case None => Right(None)
            case Some(rel) => resolveCommentPath(rel.target, relsPath).map(Some(_))

          // Extract table relationships (NEW)
          tableRels = rels.relationships.filter(_.`type` == XmlUtil.relTypeTable)
          tableTargets <- resolveTablePaths(tableRels, relsPath)

          // GH-235: hyperlink relationships (rId -> external target URL)
          hyperlinkRels = rels.relationships
            .filter(_.`type` == XmlUtil.relTypeHyperlink)
            .map(r => r.id -> r.target)
            .toMap

          // GH-221: drawing relationship — same normalization as comments (the fixture target is
          // ABSOLUTE "/xl/drawings/drawing1.xml"; Excel writes "../drawings/drawing1.xml")
          drawingTarget = rels.relationships
            .find(_.`type` == XmlUtil.relTypeDrawing)
            .flatMap(rel => resolveCommentPath(rel.target, relsPath).toOption)
        yield (commentTarget, tableTargets, hyperlinkRels, drawingTarget)

  private def resolveCommentPath(target: String, relsPath: String): XLResult[String] =
    val cleanedTarget = if target.startsWith("/") then target.drop(1) else target
    val targetPath =
      if cleanedTarget.startsWith("xl/") || cleanedTarget.startsWith("xl\\") then
        Paths.get(cleanedTarget)
      else Paths.get("xl/worksheets").resolve(cleanedTarget)

    val normalized = targetPath.normalize().toString.replace('\\', '/')

    if normalized.startsWith("xl/") then Right(normalized)
    else
      Left(
        XLError.ParseError(
          relsPath,
          s"Invalid comment relationship target outside xl/: $target"
        )
      )

  /**
   * Resolve table relationship target path.
   *
   * Converts relative paths (e.g., "../tables/table1.xml") to absolute paths within xl/. Validates
   * that paths remain within xl/ directory (security check).
   */
  private def resolveTablePath(target: String, relsPath: String): XLResult[String] =
    val cleanedTarget = if target.startsWith("/") then target.drop(1) else target
    val targetPath =
      if cleanedTarget.startsWith("xl/") || cleanedTarget.startsWith("xl\\") then
        Paths.get(cleanedTarget)
      else Paths.get("xl/worksheets").resolve(cleanedTarget)

    val normalized = targetPath.normalize().toString.replace('\\', '/')

    if normalized.startsWith("xl/") then Right(normalized)
    else
      Left(
        XLError.ParseError(
          relsPath,
          s"Invalid table relationship target outside xl/: $target"
        )
      )

  /**
   * Resolve multiple table relationship targets.
   *
   * Processes all table relationships for a worksheet and resolves their paths. Returns sequence of
   * normalized paths or error if any path is invalid.
   */
  private def resolveTablePaths(
    tableRels: Seq[Relationship],
    relsPath: String
  ): XLResult[Seq[String]] =
    tableRels.foldLeft[XLResult[Seq[String]]](Right(Seq.empty)) { (accEither, rel) =>
      for
        acc <- accEither
        path <- resolveTablePath(rel.target, relsPath)
      yield acc :+ path
    }

  /**
   * Parse comments for a single sheet.
   *
   * Resolves relative path from worksheet relationship and parses the comment XML file.
   */
  private def parseCommentsForSheet(
    parts: Map[String, String],
    commentPath: String
  ): XLResult[Map[ARef, com.tjclp.xl.cells.Comment]] =
    parts.get(commentPath) match
      case None => Right(Map.empty) // Missing comment file (graceful degradation)
      case Some(xml) =>
        for
          elem <- parseXml(xml, commentPath)
          ooxmlComments <- OoxmlComments
            .fromXml(elem)
            .left
            .map(err => XLError.ParseError(commentPath, err): XLError)
          domainComments <- convertToDomainComments(ooxmlComments, commentPath)
        yield domainComments

  /**
   * Convert OOXML comments to domain Comment map.
   *
   * Maps author IDs to author names and creates domain Comment objects.
   */
  /**
   * Strip author prefix from comment text if present AND it matches XL's exact format.
   *
   * Only strips if:
   *   1. First run text is exactly "AuthorName:" (colon, no newline)
   *   2. First run is bold
   *   3. Second run starts with newline
   *
   * This ensures we only strip prefixes WE added, not author text from real Excel files.
   */
  private[ooxml] def stripAuthorPrefix(
    text: com.tjclp.xl.richtext.RichText,
    authorName: String
  ): com.tjclp.xl.richtext.RichText =
    text.runs match
      case firstRun +: secondRun +: tail
          if authorPrefixMatches(firstRun, authorName) &&
            firstRun.font.exists(_.bold) &&
            startsWithNewline(secondRun.text) =>
        // Exact match for XL-generated format - strip it
        val cleanedSecondRun = secondRun.copy(text = dropLeadingNewline(secondRun.text))
        com.tjclp.xl.richtext.RichText(Vector(cleanedSecondRun) ++ tail)
      case _ =>
        // Different format or no prefix - preserve as-is (might be real Excel file)
        text

  private def authorPrefixMatches(run: com.tjclp.xl.richtext.TextRun, authorName: String): Boolean =
    val normalized = run.text.replace("\u00A0", " ").trim // Excel can emit nbsp and spaces
    normalized == s"$authorName:" || normalized == s"$authorName: "

  private def startsWithNewline(text: String): Boolean =
    text.startsWith("\n") || text.startsWith("\r\n")

  private def dropLeadingNewline(text: String): String =
    if text.startsWith("\r\n") then text.drop(2)
    else if text.startsWith("\n") then text.drop(1)
    else text

  private[ooxml] def convertToDomainComments(
    ooxmlComments: OoxmlComments,
    commentPath: String = "comments.xml"
  ): XLResult[Map[ARef, com.tjclp.xl.cells.Comment]] =
    val authorMap = ooxmlComments.authors.zipWithIndex.map { case (author, idx) =>
      idx -> author
    }.toMap

    ooxmlComments.comments
      .foldLeft[XLResult[Map[ARef, com.tjclp.xl.cells.Comment]]](Right(Map.empty)) {
        case (accEither, ooxmlComment) =>
          for
            acc <- accEither
            _ <- validateAuthorId(ooxmlComment, ooxmlComments.authors.size, commentPath)
            author =
              authorMap
                .get(ooxmlComment.authorId)
                .filter(_.nonEmpty) // Ignore empty string (unauthored)

            cleanedText = author match
              case Some(authorName) =>
                stripAuthorPrefix(ooxmlComment.text, authorName)
              case None => ooxmlComment.text

            domainComment = com.tjclp.xl.cells.Comment(
              text = cleanedText,
              author = author
            )
          yield acc.updated(ooxmlComment.ref, domainComment)
      }

  /**
   * Parse tables for a single sheet.
   *
   * Resolves relative paths from worksheet relationships and parses all table XML files. Returns a
   * Map[String, TableSpec] keyed by table name.
   */
  private def parseTablesForSheet(
    parts: Map[String, String],
    tablePaths: Seq[String]
  ): XLResult[Map[String, TableSpec]] =
    tablePaths.foldLeft[XLResult[Map[String, TableSpec]]](Right(Map.empty)) {
      case (accEither, tablePath) =>
        for
          acc <- accEither
          xml <- parts
            .get(tablePath)
            .toRight(XLError.ParseError(tablePath, s"Missing table file: $tablePath"))
          elem <- parseXml(xml, tablePath)
          ooxmlTable <- OoxmlTable
            .fromXml(elem)
            .left
            .map(err => XLError.ParseError(tablePath, err): XLError)
          domainTable = TableConversions.fromOoxml(ooxmlTable)
        yield acc.updated(domainTable.name, domainTable)
    }

  private def validateAuthorId(
    ooxmlComment: OoxmlComment,
    authorCount: Int,
    commentPath: String
  ): XLResult[Unit] =
    if ooxmlComment.authorId >= 0 && ooxmlComment.authorId < authorCount then Right(())
    else
      Left(
        XLError.ParseError(
          commentPath,
          s"Invalid authorId ${ooxmlComment.authorId} for comment ${ooxmlComment.ref.toA1} (authors: $authorCount)"
        )
      )

  /** Parse styles table (falls back to default styles if missing) */
  private def parseStyles(parts: Map[String, String]): XLResult[(WorkbookStyles, Vector[Warning])] =
    parts.get("xl/styles.xml") match
      case None => Right((WorkbookStyles.default, Vector(Warning.MissingStylesXml)))
      case Some(xml) =>
        for
          elem <- parseXml(xml, "xl/styles.xml")
          styles <- WorkbookStyles
            .fromXml(elem)
            .left
            .map(err => XLError.ParseError("xl/styles.xml", err): XLError)
        yield (styles, Vector.empty)

  /** Parse workbook relationships (used to resolve worksheet locations) */
  private def parseWorkbookRelationships(parts: Map[String, String]): XLResult[Relationships] =
    parts.get("xl/_rels/workbook.xml.rels") match
      case None => Right(Relationships.empty)
      case Some(xml) =>
        for
          elem <- parseXml(xml, "xl/_rels/workbook.xml.rels")
          rels <- Relationships
            .fromXml(elem)
            .left
            .map(err => XLError.ParseError("xl/_rels/workbook.xml.rels", err): XLError)
        yield rels

  /**
   * Parse theme from xl/theme/theme1.xml.
   *
   * Falls back to Office theme if theme is missing or cannot be parsed. Theme parsing errors are
   * non-fatal since the default Office theme is a reasonable fallback.
   */
  private def parseTheme(parts: Map[String, String]): ThemePalette =
    parts.get("xl/theme/theme1.xml") match
      case None => ThemePalette.office
      case Some(xml) =>
        ThemeParser.parse(xml) match
          case Right(palette) => palette
          case Left(_) => ThemePalette.office

  /** Result of [[parseSheets]]: sheets plus the per-sheet part mappings and read warnings. */
  private case class ParsedSheets(
    sheets: Vector[Sheet],
    commentPathMapping: Map[Int, String],
    drawingPathMapping: Map[Int, String],
    drawingSnapshots: Map[Int, Vector[com.tjclp.xl.drawings.Drawing]],
    chartSnapshots: Map[Int, Vector[com.tjclp.xl.context.ChartSnapshot]],
    warnings: Vector[Warning]
  )

  /**
   * Parse all worksheets and collect comment/drawing path mappings.
   *
   * Excel numbers comment files sequentially (comments1.xml, comments2.xml...) across only sheets
   * that have comments, NOT by sheet index. The mappings preserve the original paths.
   *
   * GH-221: drawing parts are parsed into `Sheet.drawings`; the as-parsed vectors are ALSO recorded
   * (same references) as snapshots for the writer's dirty test. A structurally unparseable drawing
   * part contributes a warning and an empty vector — the part itself still rides byte-preservation.
   *
   * GH-222: typed ChartFrames additionally record (anchorIdx, relId, partPath, chart) snapshots
   * feeding the writer's equality-match part/rel-id reuse.
   */
  private def parseSheets(
    parts: Map[String, String],
    retainedDrawings: RetainedDrawingParts,
    sheetRefs: Seq[SheetRef],
    sst: Option[SharedStrings],
    styles: WorkbookStyles,
    relationships: Relationships
  ): XLResult[ParsedSheets] =
    val relMap = relationships.relationships.map(rel => rel.id -> rel).toMap
    val commentPathBuilder = Map.newBuilder[Int, String]
    val drawingPathBuilder = Map.newBuilder[Int, String]
    val drawingSnapshotBuilder = Map.newBuilder[Int, Vector[com.tjclp.xl.drawings.Drawing]]
    val chartSnapshotBuilder = Map.newBuilder[Int, Vector[com.tjclp.xl.context.ChartSnapshot]]
    val warningBuilder = Vector.newBuilder[Warning]

    sheetRefs.toVector.zipWithIndex
      .traverse { case (ref, idx) =>
        val sheetPath = relMap
          .get(ref.relationshipId)
          .map(rel => resolveSheetPath(rel.target))
          .getOrElse(defaultSheetPath(ref.sheetId))

        for
          xml <- parts
            .get(sheetPath)
            .toRight(XLError.ParseError(sheetPath, s"Missing worksheet: $sheetPath"))
          elem <- parseXml(xml, sheetPath)
          ooxmlSheet <- OoxmlWorksheet
            .fromXmlWithSST(elem, sst)
            .left
            .map(err => XLError.ParseError(sheetPath, err): XLError)
          // Parse worksheet relationships to find comment/table references (1-based sheet index)
          (commentTarget, tableTargets, hyperlinkRels, drawingTarget) <-
            parseWorksheetRelationships(parts, idx + 1)

          // Parse comments if relationship exists
          comments <- commentTarget match
            case Some(target) =>
              // Track the mapping from sheet index to comment file path
              commentPathBuilder += (idx -> target)
              parseCommentsForSheet(parts, target)
            case None => Right(Map.empty)

          // Parse tables if relationships exist
          tables <- parseTablesForSheet(parts, tableTargets)

          // GH-221: parse the drawing part (total — never fails the read)
          drawings = drawingTarget.filter(retainedDrawings.xml.contains) match
            case Some(drawingPath) =>
              drawingPathBuilder += (idx -> drawingPath)
              val parsed = parseDrawingsForSheet(retainedDrawings, drawingPath) match
                case Right(part) =>
                  // GH-222: associate typed ChartFrames with their rel/part provenance
                  part.malformedChartParts
                    .foreach(p => warningBuilder += Warning.MalformedChartPart(p))
                  val snapshots = part.chartRefs.toVector.sortBy(_._1).flatMap {
                    case (anchorIdx, (relId, partPath)) =>
                      part.drawings.lift(anchorIdx).collect {
                        case frame: com.tjclp.xl.drawings.Drawing.ChartFrame =>
                          com.tjclp.xl.context.ChartSnapshot(
                            anchorIdx,
                            relId,
                            partPath,
                            frame.chart
                          )
                      }
                  }
                  if snapshots.nonEmpty then chartSnapshotBuilder += (idx -> snapshots)
                  part.drawings
                case Left(warning) =>
                  warningBuilder += warning
                  Vector.empty
              drawingSnapshotBuilder += (idx -> parsed)
              parsed
            case None => Vector.empty[com.tjclp.xl.drawings.Drawing]

          domainSheet <- convertToDomainSheet(
            ref.name,
            ooxmlSheet,
            sst,
            styles,
            comments,
            tables,
            hyperlinkRels,
            drawings
          )
        yield domainSheet
      }
      .map(sheets =>
        ParsedSheets(
          sheets,
          commentPathBuilder.result(),
          drawingPathBuilder.result(),
          drawingSnapshotBuilder.result(),
          chartSnapshotBuilder.result(),
          warningBuilder.result()
        )
      )

  /**
   * Parse one drawing part with its rels into the typed/preserved drawing vector (GH-221).
   * Left(warning) when the part XML itself is structurally malformed.
   */
  private def parseDrawingsForSheet(
    retainedDrawings: RetainedDrawingParts,
    drawingPath: String
  ): Either[Warning, com.tjclp.xl.ooxml.drawing.ParsedDrawingPart] =
    import com.tjclp.xl.ooxml.drawing.{ChartPartAccess, DrawingReader, ParsedDrawingPart}
    val relsPath = drawingRelsPath(drawingPath)
    val rels = retainedDrawings.xml
      .get(relsPath)
      .flatMap(xml => XmlSecurity.parseSafe(xml, relsPath).toOption)
      .flatMap(elem => Relationships.fromXml(elem).toOption)
      .getOrElse(Relationships.empty)
    // GH-222: chart-part access mirrors the media accessor; hasRels gates Excel charts whose
    // colors1/style1/userShapes rels would orphan on regeneration
    val chartAccess = ChartPartAccess(
      xml = retainedDrawings.charts.get,
      hasRels = path => retainedDrawings.charts.contains(chartRelsPath(path))
    )
    retainedDrawings.xml.get(drawingPath) match
      case None => Right(ParsedDrawingPart(Vector.empty, Map.empty, Vector.empty))
      case Some(xml) =>
        XmlSecurity.parseSafe(xml, drawingPath) match
          case Left(_) => Left(Warning.MalformedDrawingPart(drawingPath))
          case Right(elem) =>
            Right(DrawingReader.fromElem(elem, rels, retainedDrawings.media.get, chartAccess))

  /**
   * Rels path for a drawing part: xl/drawings/drawing1.xml -> xl/drawings/_rels/drawing1.xml.rels
   */
  private def drawingRelsPath(drawingPath: String): String =
    val slash = drawingPath.lastIndexOf('/')
    val (dir, name) = drawingPath.splitAt(slash + 1)
    s"${dir}_rels/$name.rels"

  /** Rels path for a chart part: xl/charts/chart1.xml -> xl/charts/_rels/chart1.xml.rels */
  private def chartRelsPath(chartPath: String): String =
    val slash = chartPath.lastIndexOf('/')
    val (dir, name) = chartPath.splitAt(slash + 1)
    s"${dir}_rels/$name.rels"

  private def defaultSheetPath(sheetId: Int): String = s"xl/worksheets/sheet$sheetId.xml"

  private def resolveSheetPath(target: String): String =
    val cleaned = if target.startsWith("/") then target.drop(1) else target
    val resolvedPath =
      if cleaned.startsWith("xl/") || cleaned.startsWith("xl\\") then Paths.get(cleaned)
      else Paths.get("xl").resolve(cleaned)
    resolvedPath.normalize().toString.replace('\\', '/')

  /**
   * Convert OoxmlWorksheet to domain Sheet.
   *
   * SST resolution now happens in OoxmlWorksheet.fromXml, so ooxmlCell.value is already the final
   * CellValue (Text or RichText).
   */
  private def convertToDomainSheet(
    name: SheetName,
    ooxmlSheet: OoxmlWorksheet,
    sst: Option[SharedStrings],
    styles: WorkbookStyles,
    comments: Map[ARef, com.tjclp.xl.cells.Comment],
    tables: Map[String, TableSpec],
    hyperlinkRels: Map[String, String],
    drawings: Vector[com.tjclp.xl.drawings.Drawing]
  ): XLResult[Sheet] =
    val (preRegisteredRegistry, styleMapping) = buildStyleRegistry(styles)

    // Optimization: Use builder pattern instead of nested foldLeft with Map.updated()
    // Old approach: O(n log n) with 50k cells = ~580k operations
    // New approach: O(n) with 50k cells = 50k operations (10x faster)
    val builder = Map.newBuilder[ARef, Cell]
    for
      row <- ooxmlSheet.rows
      ooxmlCell <- row.cells
    do
      // SST resolution already done in OoxmlWorksheet.fromXml
      val value = ooxmlCell.value
      val styleIdOpt = ooxmlCell.styleIndex.flatMap(styleMapping.get)
      val cell = Cell(ooxmlCell.ref, value, styleIdOpt)
      builder += (cell.ref -> cell)

    val cellsMap = builder.result()

    // GH-235: populate Cell.hyperlink from the worksheet <hyperlinks> element. External targets
    // resolve through the sheet rels (rId -> URL); internal targets use the `location` attribute.
    val cellsWithHyperlinks = ooxmlSheet.preservedKnown.get("hyperlinks") match
      case None => cellsMap
      case Some(hls) =>
        (hls \ "hyperlink").foldLeft(cellsMap) {
          case (m, e: Elem) =>
            val refStr = e \@ "ref"
            val rid = e.attributes.collectFirst {
              case a: PrefixedAttribute if a.key == "id" => a.value.text
            }
            val loc = Option(e \@ "location").filter(_.nonEmpty)
            val target = rid.flatMap(hyperlinkRels.get).orElse(loc)
            (ARef.parse(refStr).toOption, target) match
              case (Some(ref), Some(t)) =>
                val base = m.getOrElse(ref, Cell(ref, com.tjclp.xl.cells.CellValue.Empty))
                m.updated(ref, base.withHyperlink(t))
              case _ => m
          case (m, _) => m
        }

    // Parse column properties from <cols> element
    val columnProperties = ooxmlSheet.cols match
      case Some(colsElem) => parseColumnProperties(colsElem)
      case None => Map.empty[Column, ColumnProperties]

    // Extract row properties from OoxmlRow data
    val rowProperties = extractRowProperties(ooxmlSheet.rows)

    // Parse print setup (scale/orientation/fit, margins, header/footer) — GH-259
    val pageSetup =
      parsePageSetup(ooxmlSheet.pageSetup, ooxmlSheet.pageMargins, ooxmlSheet.headerFooter)

    // Parse view settings (gridlines, zoom) from <sheetViews> — GH-258
    val viewSettings = parseSheetView(ooxmlSheet.sheetViews)

    Right(
      Sheet(
        name = name,
        cells = cellsWithHyperlinks,
        mergedRanges = ooxmlSheet.mergedRanges,
        styleRegistry = preRegisteredRegistry,
        comments = comments,
        tables = tables,
        columnProperties = columnProperties,
        rowProperties = rowProperties,
        pageSetup = pageSetup,
        viewSettings = viewSettings,
        drawings = drawings
      )
    )

  private def buildStyleRegistry(
    styles: WorkbookStyles
  ): (StyleRegistry, Map[Int, StyleId]) =
    // Start with EMPTY registry (not default) to preserve exact source styles
    // This prevents adding an extra "default" style when source style 0 differs from CellStyle.default
    val emptyRegistry = StyleRegistry(Vector.empty, Map.empty)

    // Use foldLeft to accumulate registry and mapping without var
    styles.cellStyles.zipWithIndex.foldLeft((emptyRegistry, Map.empty[Int, StyleId])) {
      case ((reg, map), (style, idx)) =>
        val (nextRegistry, styleId) = reg.register(style)
        (nextRegistry, map + (idx -> styleId))
    }

  /**
   * Parse column properties from <cols> XML element.
   *
   * Each <col> element has min/max (1-indexed) and optional width, hidden, outlineLevel, collapsed.
   * A single <col> can cover multiple columns (min="1" max="3" applies to columns A-C).
   */
  private def parseColumnProperties(colsElem: Elem): Map[Column, ColumnProperties] =
    val builder = Map.newBuilder[Column, ColumnProperties]
    for colElem <- colsElem.child.collect { case e: Elem if e.label == "col" => e } do
      val attrs = colElem.attributes.asAttrMap
      for
        minStr <- attrs.get("min")
        maxStr <- attrs.get("max")
        min <- minStr.toIntOption
        max <- maxStr.toIntOption
      do
        val props = ColumnProperties(
          width = attrs.get("width").flatMap(_.toDoubleOption),
          hidden = attrs.get("hidden").contains("1"),
          outlineLevel = attrs.get("outlineLevel").flatMap(_.toIntOption),
          collapsed = attrs.get("collapsed").contains("1")
        )
        // Expand range: min and max are 1-indexed, Column is 0-indexed
        for colIdx <- min to max do builder += (Column.from1(colIdx) -> props)
    builder.result()

  /**
   * Extract row properties from parsed OoxmlRow data.
   *
   * Only includes rows that have non-default properties (height, hidden, outlineLevel, collapsed).
   */
  private def extractRowProperties(rows: Seq[OoxmlRow]): Map[Row, RowProperties] =
    val builder = Map.newBuilder[Row, RowProperties]
    for row <- rows do
      val props = RowProperties(
        height = row.height,
        hidden = row.hidden,
        outlineLevel = row.outlineLevel,
        collapsed = row.collapsed
      )
      // Only include if not all defaults
      if props.height.isDefined || props.hidden || props.outlineLevel.isDefined || props.collapsed
      then builder += (Row.from1(row.rowIndex) -> props)
    builder.result()

  /**
   * Parse print setup from the <pageSetup>, <pageMargins>, and <headerFooter> XML elements
   * (GH-259). Any one of the three is enough to produce a PageSetup; printArea/repeatRows are
   * lifted from workbook defined names later (PrintNames.extract).
   *
   * The parse is total: out-of-range or unmodeled attribute values (scale outside 10-400,
   * orientation="default", negative margins) fall back to defaults instead of violating the model
   * invariants — a malformed file must never throw.
   */
  private def parsePageSetup(
    pageSetupElem: Option[Elem],
    pageMarginsElem: Option[Elem],
    headerFooterElem: Option[Elem]
  ): Option[PageSetup] =
    val margins = pageMarginsElem.flatMap(parsePageMargins)
    val headerFooter = headerFooterElem.flatMap(parseHeaderFooter)
    val base = pageSetupElem.map { elem =>
      val attrs = elem.attributes.asAttrMap
      PageSetup(
        scale = attrs
          .get("scale")
          .flatMap(_.toIntOption)
          .filter(s => s >= 10 && s <= 400)
          .getOrElse(100),
        orientation = attrs.get("orientation").filter(o => o == "portrait" || o == "landscape"),
        fitToWidth = attrs.get("fitToWidth").flatMap(_.toIntOption).filter(_ >= 0),
        fitToHeight = attrs.get("fitToHeight").flatMap(_.toIntOption).filter(_ >= 0)
      )
    }
    if base.isEmpty && margins.isEmpty && headerFooter.isEmpty then None
    else Some(base.getOrElse(PageSetup()).copy(margins = margins, headerFooter = headerFooter))

  /**
   * Parse <pageMargins>; all six attributes are required by the schema (total: None if invalid).
   */
  private def parsePageMargins(elem: Elem): Option[PageMargins] =
    val attrs = elem.attributes.asAttrMap
    def attr(name: String): Option[Double] =
      attrs.get(name).flatMap(_.toDoubleOption).filter(_ >= 0)
    for
      left <- attr("left")
      right <- attr("right")
      top <- attr("top")
      bottom <- attr("bottom")
      header <- attr("header")
      footer <- attr("footer")
    yield PageMargins(left, right, top, bottom, header, footer)

  /**
   * Parse the modeled header/footer from <headerFooter> (GH-266): all six page parts
   * (odd/even/first × header/footer) plus the differentOddEven/differentFirst flags. Returns Some
   * only when a modeled part or flag is present, so a headerFooter element carrying only unmodeled
   * attributes (scaleWithDoc, ...) rides through as preserved XML.
   */
  private def parseHeaderFooter(elem: Elem): Option[HeaderFooter] =
    def part(label: String): Option[String] =
      (elem \ label).headOption.map(_.text).filter(_.nonEmpty)
    def flag(name: String): Boolean =
      val value = elem \@ name
      value == "1" || value == "true"
    val parsed = HeaderFooter(
      oddHeader = part("oddHeader"),
      oddFooter = part("oddFooter"),
      evenHeader = part("evenHeader"),
      evenFooter = part("evenFooter"),
      firstHeader = part("firstHeader"),
      firstFooter = part("firstFooter"),
      differentOddEven = flag("differentOddEven"),
      differentFirst = flag("differentFirst")
    )
    if parsed == HeaderFooter() then None else Some(parsed)

  /**
   * Parse sheet view settings (GH-258) from the first <sheetView>.
   *
   * Returns Some only when a modeled attribute (showGridLines, zoomScale) is present, so typical
   * files keep viewSettings = None and the raw sheetViews XML rides through unchanged on rewrite
   * (passive preserve, mirroring freezePane semantics). Out-of-range zoom values are dropped rather
   * than violating the SheetView invariant.
   */
  private def parseSheetView(sheetViewsElem: Option[Elem]): Option[SheetView] =
    for
      views <- sheetViewsElem
      view <- (views \ "sheetView").collectFirst { case e: Elem => e }
      attrs = view.attributes.asAttrMap
      showGridLines = attrs.get("showGridLines").map(v => v != "0" && v != "false")
      zoomScale = attrs.get("zoomScale").flatMap(_.toIntOption).filter(z => z >= 10 && z <= 400)
      if showGridLines.isDefined || zoomScale.isDefined
    yield SheetView(showGridLines = showGridLines.getOrElse(true), zoomScale = zoomScale)

  /**
   * Assemble final workbook with optional SourceContext for surgical modification.
   *
   * If a source handle is provided, creates a SourceContext that enables surgical writes by
   * preserving unknown ZIP entries (charts, images, etc.) byte-for-byte.
   *
   * @param sheets
   *   Parsed sheets
   * @param source
   *   Optional source handle (path to original file)
   * @param manifest
   *   Manifest of all ZIP entries (known + unknown)
   * @param theme
   *   Theme palette parsed from xl/theme/theme1.xml
   * @param definedNames
   *   Defined names from workbook.xml
   * @param sheetStates
   *   Sheet visibility states from workbook.xml
   * @param commentPathMapping
   *   Mapping from 0-based sheet index to comment file path (e.g., "xl/comments1.xml")
   * @param date1904
   *   True when workbookPr declares the 1904 date system (GH-243)
   * @param activeSheetIndex
   *   Active tab parsed from bookViews, already clamped to the sheet count (GH-294)
   * @param docProps
   *   Document properties parsed from docProps/core.xml + app.xml (GH-242)
   * @return
   *   Workbook with optional SourceContext
   */
  private def assembleWorkbook(
    sheets: Vector[Sheet],
    source: Option[SourceHandle],
    manifest: PartManifest,
    fingerprint: Option[SourceFingerprint],
    theme: ThemePalette,
    definedNames: Vector[DefinedName],
    sheetStates: Map[SheetName, Option[String]],
    commentPathMapping: Map[Int, String],
    date1904: Boolean,
    activeSheetIndex: Int,
    docProps: DocProps.Data,
    drawingPathMapping: Map[Int, String],
    drawingSnapshots: Map[Int, Vector[com.tjclp.xl.drawings.Drawing]],
    chartSnapshots: Map[Int, Vector[com.tjclp.xl.context.ChartSnapshot]]
  ): XLResult[Workbook] =
    if sheets.isEmpty then Left(XLError.InvalidWorkbook("Workbook must have at least one sheet"))
    else
      val sourceContextEither: XLResult[Option[SourceContext]] =
        (source, fingerprint) match
          case (Some(handle), Some(fp)) =>
            Right(
              Some(
                SourceContext.fromFile(
                  handle.path,
                  manifest,
                  fp,
                  commentPathMapping,
                  drawingPathMapping,
                  drawingSnapshots,
                  chartSnapshots
                )
              )
            )
          case (None, None) => Right(None)
          case (Some(_), None) =>
            Left(XLError.IOError("Missing source fingerprint for workbook"))
          case (None, Some(_)) =>
            Left(XLError.IOError("Unexpected source fingerprint without source handle"))

      // GH-259: lift modelable sheet-scoped print names (_xlnm.Print_Area/_xlnm.Print_Titles)
      // into Sheet.pageSetup; unmodelable forms stay in metadata.definedNames verbatim.
      val (sheetsWithPrint, remainingNames) = PrintNames.extract(sheets, definedNames)

      // GH-242: docProps fields reflect the FILE exactly (None when absent) — never the
      // WorkbookMetadata defaults, so write(read(f)) round-trips document properties.
      val metadata =
        WorkbookMetadata(
          creator = docProps.creator,
          created = docProps.created,
          modified = docProps.modified,
          lastModifiedBy = docProps.lastModifiedBy,
          application = docProps.application,
          appVersion = docProps.appVersion,
          theme = theme,
          definedNames = remainingNames,
          sheetStates = sheetStates,
          date1904 = date1904
        )
      sourceContextEither.map(ctx =>
        Workbook(
          sheets = sheetsWithPrint,
          metadata = metadata,
          activeSheetIndex = activeSheetIndex,
          sourceContext = ctx
        )
      )

  /**
   * Parse XML string to Elem with XXE protection
   *
   * REQUIRES: xmlString is a valid XML string ENSURES:
   *   - Returns Right(elem) if parsing succeeds
   *   - Returns Left(XLError) if parsing fails
   *   - Rejects XML with DOCTYPE declarations (XXE attack prevention)
   *   - Rejects external entity references
   * DETERMINISTIC: Yes (for same input) ERROR CASES: Malformed XML, XXE attempts, parser
   * configuration failures
   */
  private def parseXml(xmlString: String, location: String): XLResult[Elem] =
    XmlSecurity.parseSafe(xmlString, location)

  // parseDefinedNames now lives in OoxmlWorkbook (single source of truth shared with the writer).

  /** Extension for traverse on Vector */
  extension [A](vec: Vector[A])
    private def traverse[B](f: A => XLResult[B]): XLResult[Vector[B]] =
      // Optimization: Use VectorBuilder instead of Vector :+ for O(1) amortized append (was O(n) per append = O(n²) total)
      val builder = Vector.newBuilder[B]
      vec.foldLeft[XLResult[Unit]](Right(())) { (acc, a) =>
        for
          _ <- acc
          b <- f(a)
        yield builder += b
      } match
        case Right(_) => Right(builder.result())
        case Left(err) => Left(err)
