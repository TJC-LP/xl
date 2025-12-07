package com.tjclp.xl.ooxml

import com.tjclp.xl.addressing.{ARef, Column, Row, SheetName}
import com.tjclp.xl.cells.{Cell, CellError, CellValue}
import com.tjclp.xl.api.{Sheet, Workbook}
import com.tjclp.xl.sheets.{ColumnProperties, PageSetup, RowProperties}
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
    "xl/theme/theme1.xml"
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

          // Compute fingerprint if reading from a file
          val fingerprint = source.map(_.finalizeFingerprint())

          // Parse workbook structure from known parts
          parseWorkbook(parts.toMap, source, manifest, fingerprint, config)

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

      // Parse sheets
      sheets <- parseSheets(parts, ooxmlWb.sheets, sst, styles, workbookRels)

      // Parse defined names from workbook.xml
      definedNames = parseDefinedNames(ooxmlWb.definedNames)

      // Assemble workbook with optional SourceContext
      workbook <- assembleWorkbook(sheets, source, manifest, fingerprint, theme, definedNames)
    yield ReadResult(workbook, styleWarnings)

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
   * Parse worksheet relationships to find comment and table references.
   *
   * Returns (commentPath, tableTargets) where:
   *   - commentPath: Optional path to comments file (e.g., "../comments1.xml")
   *   - tableTargets: Sequence of table file paths (e.g., ["../tables/table1.xml",
   *     "../tables/table2.xml"])
   */
  private def parseWorksheetRelationships(
    parts: Map[String, String],
    sheetIndex: Int
  ): XLResult[(Option[String], Seq[String])] =
    val relsPath = s"xl/worksheets/_rels/sheet$sheetIndex.xml.rels"
    parts.get(relsPath) match
      case None => Right((None, Seq.empty)) // No relationships for this sheet
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
        yield (commentTarget, tableTargets)

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

  /** Parse all worksheets */
  private def parseSheets(
    parts: Map[String, String],
    sheetRefs: Seq[SheetRef],
    sst: Option[SharedStrings],
    styles: WorkbookStyles,
    relationships: Relationships
  ): XLResult[Vector[Sheet]] =
    val relMap = relationships.relationships.map(rel => rel.id -> rel).toMap
    sheetRefs.toVector.zipWithIndex.traverse { case (ref, idx) =>
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
        (commentTarget, tableTargets) <- parseWorksheetRelationships(parts, idx + 1)

        // Parse comments if relationship exists
        comments <- commentTarget match
          case Some(target) => parseCommentsForSheet(parts, target)
          case None => Right(Map.empty)

        // Parse tables if relationships exist
        tables <- parseTablesForSheet(parts, tableTargets)

        domainSheet <- convertToDomainSheet(ref.name, ooxmlSheet, sst, styles, comments, tables)
      yield domainSheet
    }

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
    tables: Map[String, TableSpec]
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

    // Parse column properties from <cols> element
    val columnProperties = ooxmlSheet.cols match
      case Some(colsElem) => parseColumnProperties(colsElem)
      case None => Map.empty[Column, ColumnProperties]

    // Extract row properties from OoxmlRow data
    val rowProperties = extractRowProperties(ooxmlSheet.rows)

    // Parse page setup (print scale, orientation, etc.)
    val pageSetup = parsePageSetup(ooxmlSheet.pageSetup)

    Right(
      Sheet(
        name = name,
        cells = cellsMap,
        mergedRanges = ooxmlSheet.mergedRanges,
        styleRegistry = preRegisteredRegistry,
        comments = comments,
        tables = tables,
        columnProperties = columnProperties,
        rowProperties = rowProperties,
        pageSetup = pageSetup
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
   * Parse page setup from <pageSetup> XML element.
   *
   * Extracts print settings like scale, orientation, fitToWidth, fitToHeight.
   */
  private def parsePageSetup(pageSetupElem: Option[Elem]): Option[PageSetup] =
    pageSetupElem.map { elem =>
      val attrs = elem.attributes.asAttrMap
      PageSetup(
        scale = attrs.get("scale").flatMap(_.toIntOption).getOrElse(100),
        orientation = attrs.get("orientation"),
        fitToWidth = attrs.get("fitToWidth").flatMap(_.toIntOption),
        fitToHeight = attrs.get("fitToHeight").flatMap(_.toIntOption)
      )
    }

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
   * @return
   *   Workbook with optional SourceContext
   */
  private def assembleWorkbook(
    sheets: Vector[Sheet],
    source: Option[SourceHandle],
    manifest: PartManifest,
    fingerprint: Option[SourceFingerprint],
    theme: ThemePalette,
    definedNames: Vector[DefinedName]
  ): XLResult[Workbook] =
    if sheets.isEmpty then Left(XLError.InvalidWorkbook("Workbook must have at least one sheet"))
    else
      val sourceContextEither: XLResult[Option[SourceContext]] =
        (source, fingerprint) match
          case (Some(handle), Some(fp)) =>
            Right(Some(SourceContext.fromFile(handle.path, manifest, fp)))
          case (None, None) => Right(None)
          case (Some(_), None) =>
            Left(XLError.IOError("Missing source fingerprint for workbook"))
          case (None, Some(_)) =>
            Left(XLError.IOError("Unexpected source fingerprint without source handle"))

      val metadata = WorkbookMetadata(theme = theme, definedNames = definedNames)
      sourceContextEither.map(ctx =>
        Workbook(sheets = sheets, metadata = metadata, sourceContext = ctx)
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

  /**
   * Parse defined names from the raw XML element.
   *
   * OOXML structure:
   * {{{
   * <definedNames>
   *   <definedName name="SalesTotal" localSheetId="0">Sheet1!$A$1:$A$10</definedName>
   *   <definedName name="TaxRate" hidden="1">0.08</definedName>
   * </definedNames>
   * }}}
   */
  private def parseDefinedNames(rawElem: Option[Elem]): Vector[DefinedName] =
    rawElem match
      case None => Vector.empty
      case Some(elem) =>
        (elem \ "definedName").collect { case e: Elem =>
          val name = (e \@ "name")
          val formula = e.text.trim
          val localSheetId = Option(e \@ "localSheetId").filter(_.nonEmpty).flatMap(_.toIntOption)
          val hidden = (e \@ "hidden") == "1"
          val comment = Option(e \@ "comment").filter(_.nonEmpty)
          DefinedName(
            name = name,
            formula = formula,
            localSheetId = localSheetId,
            hidden = hidden,
            comment = comment
          )
        }.toVector

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
