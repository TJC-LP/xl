package com.tjclp.xl.ooxml

import scala.xml.*
import java.io.{File, FileInputStream, FileOutputStream, ByteArrayOutputStream, OutputStream}
import java.util.zip.{ZipInputStream, ZipOutputStream, ZipEntry, ZipFile}
import java.nio.file.{Path, Files, StandardCopyOption}
import java.nio.charset.StandardCharsets
import scala.collection.mutable
import com.tjclp.xl.api.Workbook
import com.tjclp.xl.error.{XLError, XLResult}
import com.tjclp.xl.SourceContext

/** Shared Strings Table usage policy */
enum SstPolicy derives CanEqual:
  /** Auto-detect based on heuristics (default) */
  case Auto

  /** Always use SST regardless of content */
  case Always

  /** Never use SST (inline strings only) */
  case Never

/**
 * Compression method for ZIP entries.
 *
 * DEFLATED (default) produces 5-10x smaller files with minimal CPU overhead. STORED is useful for
 * debugging (human-readable ZIP contents).
 */
enum Compression derives CanEqual:
  /** No compression (STORED) - faster writes, larger files, requires CRC32 precomputation */
  case Stored

  /** DEFLATE compression (DEFLATED) - smaller files, standard production use */
  case Deflated

  /** ZIP constant for this compression method */
  def zipMethod: Int = this match
    case Stored => ZipEntry.STORED
    case Deflated => ZipEntry.DEFLATED

/** Writer configuration options */
case class WriterConfig(
  sstPolicy: SstPolicy = SstPolicy.Auto,
  compression: Compression = Compression.Deflated,
  prettyPrint: Boolean = false // Compact XML for production
)

object WriterConfig:
  /** Default production configuration: DEFLATED compression + compact XML */
  val default: WriterConfig = WriterConfig()

  /** Debug configuration: STORED compression + pretty XML for manual inspection */
  val debug: WriterConfig = WriterConfig(
    compression = Compression.Stored,
    prettyPrint = true
  )

/**
 * Target for XLSX output (file path or output stream).
 *
 * Enables unified handling of different output destinations in surgical modification. File-based
 * targets support verbatim copy optimization for clean workbooks.
 */
sealed trait OutputTarget:
  /** Get path if this is a file target (None for streams) */
  def asPathOption: Option[Path] = None

/** Output target that writes to a file path */
case class OutputPath(path: Path) extends OutputTarget:
  override def asPathOption: Option[Path] = Some(path)

/** Output target that writes to an output stream */
case class OutputStreamTarget(stream: OutputStream) extends OutputTarget

/**
 * Writer for XLSX files (ZIP assembly)
 *
 * Takes a domain Workbook and produces a valid XLSX file with all required parts.
 *
 * **Surgical Modification** (Phase 4): When a workbook has a SourceContext (from file-based reads),
 * the writer uses a hybrid strategy:
 *   - Clean workbooks → verbatim file copy (11x faster)
 *   - Partially modified → regenerate changed sheets, copy preserved parts (2-5x faster)
 *   - Fully modified → full regeneration (same as before)
 */
object XlsxWriter:

  /** Write workbook to XLSX file with default configuration */
  def write(workbook: Workbook, outputPath: Path): XLResult[Unit] =
    writeWith(workbook, outputPath, WriterConfig())

  /**
   * Write workbook to XLSX file with custom configuration.
   *
   * **Surgical Modification**: Automatically uses hybrid strategy when workbook has SourceContext:
   *   - Clean workbook → verbatim copy (11x faster)
   *   - Partially modified → regenerate changed parts, preserve rest (2-5x faster)
   *   - No SourceContext → full regeneration (current behavior)
   */
  def writeWith(
    workbook: Workbook,
    outputPath: Path,
    config: WriterConfig = WriterConfig()
  ): XLResult[Unit] =
    writeToTarget(workbook, OutputPath(outputPath), config)

  /**
   * Internal dispatch: Choose write strategy based on SourceContext.
   *
   * Strategy selection:
   *   1. No SourceContext → full regeneration (programmatically created workbook)
   *   1. SourceContext + clean + file target → verbatim copy (fast path)
   *   1. SourceContext + dirty → hybrid write (surgical modification)
   */
  private def writeToTarget(
    workbook: Workbook,
    target: OutputTarget,
    config: WriterConfig
  ): XLResult[Unit] =
    try
      workbook.sourceContext match
        case None =>
          // No source context → full regeneration (current behavior)
          regenerateAll(workbook, target, config)

        case Some(ctx) if ctx.isClean =>
          // Clean workbook → check if file target for fast copy
          target match
            case OutputPath(path) =>
              copyVerbatim(ctx.sourcePath, path)
            case OutputStreamTarget(_) =>
              // Can't copy verbatim to stream, fall back to regenerate
              regenerateAll(workbook, target, config)

        case Some(ctx) =>
          // Dirty workbook → hybrid write (surgical modification)
          hybridWrite(workbook, ctx, target, config)

      Right(())

    catch case e: Exception => Left(XLError.IOError(s"Failed to write XLSX: ${e.getMessage}"))

  /**
   * Copy source file verbatim to destination (for clean workbooks).
   *
   * Fast path optimization: When a workbook has no modifications, just copy the source file
   * byte-for-byte instead of regenerating all XML. This is 10-11x faster than full regeneration.
   *
   * Handles edge case where source == dest (no-op).
   */
  private def copyVerbatim(source: Path, dest: Path): Unit =
    if source == dest then
      // Source and destination are the same file - no-op
      ()
    else
      // Simple file copy using NIO
      Files.copy(source, dest, StandardCopyOption.REPLACE_EXISTING)
      ()

  /**
   * Full regeneration of all XLSX parts (current behavior).
   *
   * Used when:
   *   - Workbook has no SourceContext (created programmatically)
   *   - Output target is a stream (can't use file copy optimization)
   *
   * Regenerates all parts from domain model.
   */
  private def regenerateAll(
    workbook: Workbook,
    target: OutputTarget,
    config: WriterConfig
  ): Unit =
    // Build shared strings table based on policy
    val sst = config.sstPolicy match
      case SstPolicy.Always => Some(SharedStrings.fromWorkbook(workbook))
      case SstPolicy.Never => None
      case SstPolicy.Auto =>
        if SharedStrings.shouldUseSST(workbook) then Some(SharedStrings.fromWorkbook(workbook))
        else None

    // Build unified style index with per-sheet remappings
    val (styleIndex, sheetRemappings) = StyleIndex.fromWorkbook(workbook)
    val styles = OoxmlStyles(styleIndex)

    // Convert domain workbook to OOXML
    val ooxmlWb = OoxmlWorkbook.fromDomain(workbook)

    // Convert sheets to OOXML worksheets with style remapping
    val ooxmlSheets = workbook.sheets.zipWithIndex.map { case (sheet, sheetIdx) =>
      val remapping = sheetRemappings.getOrElse(sheetIdx, Map.empty)
      OoxmlWorksheet.fromDomainWithSST(sheet, sst, remapping)
    }

    // Create content types
    val contentTypes = ContentTypes.minimal(
      hasStyles = true, // Always include styles
      hasSharedStrings = sst.isDefined,
      sheetCount = workbook.sheets.size
    )

    // Create relationships
    val rootRels = Relationships.root()
    val workbookRels = Relationships.workbook(
      sheetCount = workbook.sheets.size,
      hasStyles = true,
      hasSharedStrings = sst.isDefined
    )

    // Dispatch based on target type
    target match
      case OutputPath(path) =>
        writeZip(
          path,
          contentTypes,
          rootRels,
          workbookRels,
          ooxmlWb,
          ooxmlSheets,
          styles,
          sst,
          config
        )
      case OutputStreamTarget(stream) =>
        writeZipToStream(
          stream,
          contentTypes,
          rootRels,
          workbookRels,
          ooxmlWb,
          ooxmlSheets,
          styles,
          sst,
          config
        )

  /** Write all parts to ZIP file */
  private def writeZip(
    path: Path,
    contentTypes: ContentTypes,
    rootRels: Relationships,
    workbookRels: Relationships,
    workbook: OoxmlWorkbook,
    sheets: Vector[OoxmlWorksheet],
    styles: OoxmlStyles,
    sst: Option[SharedStrings],
    config: WriterConfig
  ): Unit =
    val zip = new ZipOutputStream(new FileOutputStream(path.toFile))
    try
      // Write parts in canonical order
      writePart(zip, "[Content_Types].xml", contentTypes.toXml, config)
      writePart(zip, "_rels/.rels", rootRels.toXml, config)
      writePart(zip, "xl/workbook.xml", workbook.toXml, config)
      writePart(zip, "xl/_rels/workbook.xml.rels", workbookRels.toXml, config)

      // Write styles
      writePart(zip, "xl/styles.xml", styles.toXml, config)

      // Write shared strings if present
      sst.foreach { sharedStrings =>
        writePart(zip, "xl/sharedStrings.xml", sharedStrings.toXml, config)
      }

      // Write worksheets
      sheets.zipWithIndex.foreach { case (sheet, idx) =>
        writePart(zip, s"xl/worksheets/sheet${idx + 1}.xml", sheet.toXml, config)
      }

    finally zip.close()

  /** Write all parts to ZIP stream (for OutputStreamTarget) */
  private def writeZipToStream(
    stream: OutputStream,
    contentTypes: ContentTypes,
    rootRels: Relationships,
    workbookRels: Relationships,
    workbook: OoxmlWorkbook,
    sheets: Vector[OoxmlWorksheet],
    styles: OoxmlStyles,
    sst: Option[SharedStrings],
    config: WriterConfig
  ): Unit =
    val zip = new ZipOutputStream(stream)
    try
      // Write parts in canonical order
      writePart(zip, "[Content_Types].xml", contentTypes.toXml, config)
      writePart(zip, "_rels/.rels", rootRels.toXml, config)
      writePart(zip, "xl/workbook.xml", workbook.toXml, config)
      writePart(zip, "xl/_rels/workbook.xml.rels", workbookRels.toXml, config)

      // Write styles
      writePart(zip, "xl/styles.xml", styles.toXml, config)

      // Write shared strings if present
      sst.foreach { sharedStrings =>
        writePart(zip, "xl/sharedStrings.xml", sharedStrings.toXml, config)
      }

      // Write worksheets
      sheets.zipWithIndex.foreach { case (sheet, idx) =>
        writePart(zip, s"xl/worksheets/sheet${idx + 1}.xml", sheet.toXml, config)
      }

    finally zip.close()

  /** Write a single XML part to ZIP */
  private def writePart(
    zip: ZipOutputStream,
    entryName: String,
    xml: Elem,
    config: WriterConfig
  ): Unit =
    val entry = new ZipEntry(entryName)
    entry.setTime(0L) // deterministic ZIP metadata for reproducible output
    entry.setMethod(config.compression.zipMethod)

    // Convert XML to bytes with conditional formatting
    val xmlString = if config.prettyPrint then XmlUtil.prettyPrint(xml) else XmlUtil.compact(xml)
    val bytes = xmlString.getBytes(StandardCharsets.UTF_8)

    // For STORED method, must set size and CRC before writing
    // For DEFLATED, these are computed automatically by ZipOutputStream
    config.compression match
      case Compression.Stored =>
        entry.setSize(bytes.length)
        entry.setCompressedSize(bytes.length)
        entry.setCrc(calculateCrc(bytes))
      case Compression.Deflated =>
        // ZipOutputStream computes these automatically for DEFLATED
        ()

    zip.putNextEntry(entry)
    zip.write(bytes)
    zip.closeEntry()

  /** Calculate CRC32 checksum for ZIP entry */
  private def calculateCrc(bytes: Array[Byte]): Long =
    val crc = new java.util.zip.CRC32()
    crc.update(bytes)
    crc.getValue

  // ========== Surgical Modification Methods ==========

  /**
   * Determine which parts can be preserved byte-for-byte during surgical write.
   *
   * A part can be preserved if:
   *   1. It's in the unparsed set (XL didn't parse it)
   *   1. It doesn't reference any modified or deleted sheets
   *
   * @return
   *   Set of ZIP entry paths that can be copied verbatim from source
   */
  private def determinePreservableParts(
    wb: Workbook,
    ctx: SourceContext,
    graph: RelationshipGraph
  ): Set[String] =
    val tracker = ctx.modificationTracker
    val unparsed = ctx.partManifest.unparsedParts

    unparsed.filter { path =>
      val dependentSheets = graph.dependenciesFor(path)
      val touchesDeleted = dependentSheets.exists(tracker.deletedSheets.contains)
      val touchesModified = dependentSheets.exists(tracker.modifiedSheets.contains)

      // Preserve only if doesn't touch any modified/deleted sheets
      !touchesDeleted && !touchesModified
    }

  /**
   * Determine which parts must be regenerated during surgical write.
   *
   * Parts that must be regenerated:
   *   - Structural parts (always): workbook.xml, relationships, content types
   *   - Styles/SST if any sheet modified (indices may change)
   *   - Modified sheets
   *   - Relationships for modified/deleted sheets
   *
   * @return
   *   Set of ZIP entry paths that must be regenerated from domain model
   */
  private def determineRegenerateParts(
    wb: Workbook,
    ctx: SourceContext
  ): Set[String] =
    val tracker = ctx.modificationTracker
    val regenerate = mutable.Set[String]()

    // Always regenerate structural parts
    regenerate ++= Set(
      "[Content_Types].xml",
      "_rels/.rels",
      "xl/workbook.xml",
      "xl/_rels/workbook.xml.rels"
    )

    // Regenerate styles if any sheet modified (style indices may change)
    if tracker.modifiedSheets.nonEmpty then regenerate += "xl/styles.xml"

    // Regenerate SST if any sheet modified (string indices may change)
    if tracker.modifiedSheets.nonEmpty then regenerate += "xl/sharedStrings.xml"

    // Regenerate modified sheets
    tracker.modifiedSheets.foreach { idx =>
      regenerate += s"xl/worksheets/sheet${idx + 1}.xml"
    }

    // Regenerate relationships for modified/deleted sheets
    (tracker.modifiedSheets ++ tracker.deletedSheets).foreach { idx =>
      regenerate += s"xl/worksheets/_rels/sheet${idx + 1}.xml.rels"
    }

    regenerate.toSet

  /**
   * Copy a single preserved part from source ZIP to output ZIP.
   *
   * Streams bytes in 8KB chunks (constant memory). Preserves compression method and metadata.
   */
  @SuppressWarnings(Array("org.wartremover.warts.Var", "org.wartremover.warts.While"))
  private def copyPreservedPart(
    sourcePath: Path,
    partPath: String,
    outputZip: ZipOutputStream
  ): Unit =
    val sourceZip = new ZipFile(sourcePath.toFile)
    try
      val entry = sourceZip.getEntry(partPath)
      if entry == null then throw new IllegalStateException(s"Entry missing from source: $partPath")

      // Create new entry with deterministic metadata
      val newEntry = new ZipEntry(partPath)
      newEntry.setTime(0L) // Deterministic timestamp
      newEntry.setMethod(entry.getMethod)

      if entry.getMethod == ZipEntry.STORED then
        newEntry.setSize(entry.getSize)
        newEntry.setCompressedSize(entry.getCompressedSize)
        newEntry.setCrc(entry.getCrc)

      // Stream bytes directly (preserves compression)
      outputZip.putNextEntry(newEntry)
      val is = sourceZip.getInputStream(entry)
      try
        val buffer = new Array[Byte](8192)
        var read = is.read(buffer)
        while read != -1 do
          outputZip.write(buffer, 0, read)
          read = is.read(buffer)
      finally is.close()
      outputZip.closeEntry()
    finally sourceZip.close()

  /**
   * Re-parse structural files from source ZIP for preservation.
   *
   * Returns (ContentTypes, rootRels, workbookRels, workbook) parsed from the original file. If any
   * file is missing or fails to parse, returns None for that component.
   */
  @SuppressWarnings(Array("org.wartremover.warts.Var"))
  private def parsePreservedStructure(
    sourcePath: Path
  ): (Option[ContentTypes], Option[Relationships], Option[Relationships], Option[OoxmlWorkbook]) =
    var contentTypes: Option[ContentTypes] = None
    var rootRels: Option[Relationships] = None
    var workbookRels: Option[Relationships] = None
    var workbook: Option[OoxmlWorkbook] = None

    val sourceZip = new ZipInputStream(new FileInputStream(sourcePath.toFile))
    try
      var entry = sourceZip.getNextEntry
      while entry != null do
        val entryName = entry.getName

        if entryName == "[Content_Types].xml" then
          val content = new String(sourceZip.readAllBytes(), "UTF-8")
          parseXml(content, "[Content_Types].xml") match
            case Right(elem) =>
              ContentTypes.fromXml(elem).foreach(ct => contentTypes = Some(ct))
            case Left(_) => () // Silently ignore - will use minimal fallback
        else if entryName == "_rels/.rels" then
          val content = new String(sourceZip.readAllBytes(), "UTF-8")
          parseXml(content, "_rels/.rels") match
            case Right(elem) =>
              Relationships.fromXml(elem).foreach(rels => rootRels = Some(rels))
            case Left(_) => () // Silently ignore - will use minimal fallback
        else if entryName == "xl/_rels/workbook.xml.rels" then
          val content = new String(sourceZip.readAllBytes(), "UTF-8")
          parseXml(content, "xl/_rels/workbook.xml.rels") match
            case Right(elem) =>
              Relationships.fromXml(elem).foreach(rels => workbookRels = Some(rels))
            case Left(_) => () // Silently ignore - will use minimal fallback
        else if entryName == "xl/workbook.xml" then
          val content = new String(sourceZip.readAllBytes(), "UTF-8")
          parseXml(content, "xl/workbook.xml") match
            case Right(elem) =>
              OoxmlWorkbook.fromXml(elem).foreach(wb => workbook = Some(wb))
            case Left(_) => () // Silently ignore - will use minimal fallback
        else sourceZip.readAllBytes() // Consume entry

        sourceZip.closeEntry()
        entry = sourceZip.getNextEntry
    finally sourceZip.close()

    (contentTypes, rootRels, workbookRels, workbook)

  /** Parse XML with XXE protection (same as XlsxReader) */
  private def parseXml(xmlString: String, location: String): XLResult[Elem] =
    try
      val factory = javax.xml.parsers.SAXParserFactory.newInstance()
      factory.setFeature("http://apache.org/xml/features/disallow-doctype-decl", true)
      factory.setFeature("http://xml.org/sax/features/external-general-entities", false)
      factory.setFeature("http://xml.org/sax/features/external-parameter-entities", false)
      factory.setFeature("http://apache.org/xml/features/nonvalidating/load-external-dtd", false)
      factory.setXIncludeAware(false)
      factory.setNamespaceAware(true)

      val loader = XML.withSAXParser(factory.newSAXParser())
      Right(loader.loadString(xmlString))
    catch
      case e: Exception => Left(XLError.ParseError(location, s"XML parse error: ${e.getMessage}"))

  /**
   * Parse worksheet XML from source ZIP to extract preserved metadata.
   *
   * Returns the parsed OoxmlWorksheet with all metadata (cols, views, etc.) for merging during
   * regeneration.
   */
  @SuppressWarnings(Array("org.wartremover.warts.Var"))
  private def parsePreservedWorksheet(
    sourcePath: Path,
    sheetPath: String
  ): Option[OoxmlWorksheet] =
    val sourceZip = new ZipInputStream(new FileInputStream(sourcePath.toFile))
    try
      var entry = sourceZip.getNextEntry
      var result: Option[OoxmlWorksheet] = None

      while entry != null && result.isEmpty do
        if entry.getName == sheetPath then
          val content = new String(sourceZip.readAllBytes(), "UTF-8")
          result = parseXml(content, sheetPath).toOption
            .flatMap(OoxmlWorksheet.fromXml(_).toOption)

        sourceZip.closeEntry()
        entry = sourceZip.getNextEntry

      result
    finally sourceZip.close()

  /**
   * Hybrid write: regenerate modified parts, copy preserved parts.
   *
   * Strategy:
   *   1. Build RelationshipGraph to understand dependencies
   *   1. Determine which parts can be preserved vs regenerated
   *   1. Regenerate structural parts + modified sheets
   *   1. Copy unmodified sheets from source
   *   1. Copy unknown parts that don't reference modified/deleted sheets
   */
  @SuppressWarnings(Array("org.wartremover.warts.Var"))
  private def hybridWrite(
    workbook: Workbook,
    ctx: SourceContext,
    target: OutputTarget,
    config: WriterConfig
  ): Unit =
    val tracker = ctx.modificationTracker
    val graph = RelationshipGraph.fromManifest(ctx.partManifest)

    val preservableParts = determinePreservableParts(workbook, ctx, graph)
    val regenerateParts = determineRegenerateParts(workbook, ctx)

    val sharedStringsPath = "xl/sharedStrings.xml"
    val sourceHasSharedStrings = ctx.partManifest.contains(sharedStringsPath)

    // If the source workbook already has shared strings, preserve them byte-for-byte so copied
    // sheets keep their existing string indices. Only regenerate when we don't have a source SST
    // (e.g., programmatic workbook) and policy permits.
    val sst =
      if sourceHasSharedStrings then None
      else
        config.sstPolicy match
          case SstPolicy.Always => Some(SharedStrings.fromWorkbook(workbook))
          case SstPolicy.Never => None
          case SstPolicy.Auto =>
            if SharedStrings.shouldUseSST(workbook) then Some(SharedStrings.fromWorkbook(workbook))
            else None
    val regenerateSharedStrings = sst.isDefined
    val sharedStringsInOutput = sourceHasSharedStrings || regenerateSharedStrings

    val (styleIndex, sheetRemappings) = StyleIndex.fromWorkbook(workbook)
    val styles = OoxmlStyles(styleIndex)

    // Preserve structural parts from source (or fallback to minimal)
    val (preservedContentTypes, preservedRootRels, preservedWorkbookRels, preservedWorkbook) =
      parsePreservedStructure(ctx.sourcePath)

    // Use preserved workbook structure if available, otherwise create minimal
    val ooxmlWb = preservedWorkbook match
      case Some(preserved) =>
        // Update sheets in preserved structure (names/order may have changed)
        preserved.updateSheets(workbook.sheets)
      case None =>
        // Fallback to minimal for programmatically created workbooks
        OoxmlWorkbook.fromDomain(workbook)

    val contentTypes = preservedContentTypes.getOrElse(
      ContentTypes.minimal(
        hasStyles = true,
        hasSharedStrings = sharedStringsInOutput,
        sheetCount = workbook.sheets.size
      )
    )

    val rootRels = preservedRootRels.getOrElse(Relationships.root())

    val workbookRels = preservedWorkbookRels.getOrElse(
      Relationships.workbook(
        sheetCount = workbook.sheets.size,
        hasStyles = true,
        hasSharedStrings = sharedStringsInOutput
      )
    )

    // Open output ZIP
    val zip = target match
      case OutputPath(path) => new ZipOutputStream(new FileOutputStream(path.toFile))
      case OutputStreamTarget(stream) => new ZipOutputStream(stream)

    try
      // Write structural parts (always regenerated)
      writePart(zip, "[Content_Types].xml", contentTypes.toXml, config)
      writePart(zip, "_rels/.rels", rootRels.toXml, config)
      writePart(zip, "xl/workbook.xml", ooxmlWb.toXml, config)
      writePart(zip, "xl/_rels/workbook.xml.rels", workbookRels.toXml, config)
      writePart(zip, "xl/styles.xml", styles.toXml, config)

      if regenerateSharedStrings then
        sst.foreach { sharedStrings =>
          writePart(zip, sharedStringsPath, sharedStrings.toXml, config)
        }
      else if sourceHasSharedStrings then copyPreservedPart(ctx.sourcePath, sharedStringsPath, zip)

      // Write sheets: regenerate modified, copy unmodified
      workbook.sheets.zipWithIndex.foreach { case (sheet, idx) =>
        if tracker.modifiedSheets.contains(idx) then
          // Regenerate modified sheet with preserved metadata
          val sheetPath = graph.pathForSheet(idx)
          val preservedMetadata = parsePreservedWorksheet(ctx.sourcePath, sheetPath)
          val remapping = sheetRemappings.getOrElse(idx, Map.empty)
          val ooxmlSheet =
            OoxmlWorksheet.fromDomainWithMetadata(sheet, sst, remapping, preservedMetadata)
          writePart(zip, s"xl/worksheets/sheet${idx + 1}.xml", ooxmlSheet.toXml, config)
        else
          // Copy unmodified sheet from source
          val sheetPath = graph.pathForSheet(idx)
          copyPreservedPart(ctx.sourcePath, sheetPath, zip)
      }

      // Copy preserved parts (charts, drawings, images, etc.)
      preservableParts.foreach { path =>
        copyPreservedPart(ctx.sourcePath, path, zip)
      }

    finally zip.close()

  /** Write workbook to bytes (for testing) */
  def writeToBytes(workbook: Workbook): XLResult[Array[Byte]] =
    try
      val baos = new ByteArrayOutputStream()
      val tempPath = Files.createTempFile("xl-", ".xlsx")
      try
        write(workbook, tempPath).map { _ =>
          Files.readAllBytes(tempPath)
        }
      finally
        Files.deleteIfExists(tempPath)
    catch case e: Exception => Left(XLError.IOError(s"Failed to write bytes: ${e.getMessage}"))
