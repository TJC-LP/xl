package com.tjclp.xl.ooxml

import scala.xml.*
import java.io.{ByteArrayOutputStream, FileOutputStream, OutputStream}
import java.security.MessageDigest
import java.util.zip.{ZipEntry, ZipFile, ZipOutputStream}
import java.nio.file.{Files, Path}
import java.nio.charset.StandardCharsets
import scala.collection.mutable
import scala.util.{Failure, Success, Try, Using}
import com.tjclp.xl.api.{Workbook, CellValue}
import com.tjclp.xl.errors.{XLError, XLResult}
import com.tjclp.xl.{ModificationTracker, SourceContext}
import com.tjclp.xl.richtext.RichText

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
 * Unified XLSX writer with intelligent surgical modification.
 *
 * Single code path that automatically optimizes based on context:
 *   - **With source** (read → modify → write): Surgical mode
 *     - Clean → verbatim copy (11x faster)
 *     - Partially modified → regenerate changed, copy unchanged (2-5x faster)
 *     - Fully modified → regenerate all with styles preservation
 *   - **Without source** (create → write): Full regeneration (standard path)
 *
 * The writer transparently chooses the optimal strategy - users don't need to decide.
 *
 * **Key Features**:
 *   - Preserves unknown parts (charts, comments, drawings)
 *   - Preserves differential formats (dxfs) for conditional formatting
 *   - Byte-identical copies for unmodified sheets
 *   - RichText in SharedStrings (not inlined)
 *   - Row-level styles preservation
 *   - Excel compression level matching (defS = level 1)
 */
object XlsxWriter:

  /** Write workbooks to XLSX file with default configuration */
  def write(workbook: Workbook, outputPath: Path): XLResult[Unit] =
    writeWith(workbook, outputPath, WriterConfig())

  /**
   * Write workbooks to XLSX file with custom configuration.
   *
   * **Surgical Modification**: Automatically uses hybrid strategy when workbooks has SourceContext:
   *   - Clean workbooks → verbatim copy (11x faster)
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
   *   1. SourceContext + clean + file target → verbatim copy (fastest)
   *   1. All other cases → unified write (surgical if source available, else full regeneration)
   *
   * Unified write automatically optimizes:
   *   - With source: regenerate modified, copy unmodified (surgical)
   *   - Without source: regenerate all (graceful degradation)
   */
  private def writeToTarget(
    workbook: Workbook,
    target: OutputTarget,
    config: WriterConfig
  ): XLResult[Unit] =
    try
      workbook.sourceContext match
        case Some(ctx) if ctx.isClean =>
          // Clean workbooks + file target → verbatim copy (ultra-fast)
          target match
            case OutputPath(path) =>
              copyVerbatim(ctx, path)
            case OutputStreamTarget(_) =>
              // Can't copy to stream, use unified write (will copy all parts)
              unifiedWrite(workbook, workbook.sourceContext, target, config)

        case _ =>
          // All other cases: unified write (surgical if source, else full regeneration)
          unifiedWrite(workbook, workbook.sourceContext, target, config)

      Right(())

    catch case e: Exception => Left(XLError.IOError(s"Failed to write XLSX: ${e.getMessage}"))

  /**
   * Copy source file verbatim to destination (for clean workbooks).
   *
   * Fast path optimization: When a workbooks has no modifications, just copy the source file
   * byte-for-byte instead of regenerating all XML. This is 10-11x faster than full regeneration.
   *
   * Handles edge case where source == dest (no-op).
   */
  private def copyVerbatim(ctx: SourceContext, dest: Path): Unit =
    val source = ctx.sourcePath
    if source != dest then
      val fingerprint = ctx.fingerprint
      val currentSize = Files.size(source)
      if currentSize != fingerprint.size then
        throw new IllegalStateException(
          s"Source file changed size since read (expected ${fingerprint.size} bytes, found $currentSize)"
        )

      val digest = MessageDigest.getInstance("SHA-256")

      val bytesCopied = usingOrThrow(Using.Manager { use =>
        val in = use(Files.newInputStream(source))
        val out = use(Files.newOutputStream(dest))
        val buffer = new Array[Byte](8192)

        def loop(total: Long): Long =
          val read = in.read(buffer)
          if read == -1 then total
          else
            digest.update(buffer, 0, read)
            out.write(buffer, 0, read)
            loop(total + read)

        loop(0L)
      })

      val computedDigest = digest.digest()
      if !fingerprint.matches(bytesCopied, computedDigest) then
        Files.deleteIfExists(dest)
        throw new IllegalStateException("Source file changed since read; refusing to copy verbatim")

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

    // Build unified styles index with per-sheets remappings
    val (styleIndex, sheetRemappings) = StyleIndex.fromWorkbook(workbook)
    val styles = OoxmlStyles(styleIndex)

    // Convert domain workbooks to OOXML
    val ooxmlWb = OoxmlWorkbook.fromDomain(workbook)

    // Convert sheets to OOXML worksheets with styles remapping
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
    zip.setLevel(1) // Match Excel's compression level (super fast)
    try
      // Write parts in canonical order
      writePart(zip, "[Content_Types].xml", contentTypes.toXml, config)
      writePart(zip, "_rels/.rels", rootRels.toXml, config)
      writePart(zip, "xl/workbooks.xml", workbook.toXml, config)
      writePart(zip, "xl/_rels/workbooks.xml.rels", workbookRels.toXml, config)

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
    zip.setLevel(1) // Match Excel's compression level (super fast)
    try
      // Write parts in canonical order
      writePart(zip, "[Content_Types].xml", contentTypes.toXml, config)
      writePart(zip, "_rels/.rels", rootRels.toXml, config)
      writePart(zip, "xl/workbooks.xml", workbook.toXml, config)
      writePart(zip, "xl/_rels/workbooks.xml.rels", workbookRels.toXml, config)

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
   *   - Structural parts (always): workbooks.xml, relationships, content types
   *   - Styles/SST if any sheets modified (indices may change)
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
      "xl/workbooks.xml",
      "xl/_rels/workbooks.xml.rels"
    )

    // Regenerate styles if any sheets modified (styles indices may change)
    if tracker.modifiedSheets.nonEmpty then regenerate += "xl/styles.xml"

    // Regenerate SST if any sheets modified (string indices may change)
    if tracker.modifiedSheets.nonEmpty then regenerate += "xl/sharedStrings.xml"

    // Regenerate modified sheets
    tracker.modifiedSheets.foreach { idx =>
      regenerate += s"xl/worksheets/sheet${idx + 1}.xml"
    }

    // Regenerate relationships for modified/deleted sheets
    (tracker.modifiedSheets ++ tracker.deletedSheets).foreach { idx =>
      regenerate += s"xl/worksheets/_rels/sheets${idx + 1}.xml.rels"
    }

    regenerate.toSet

  private def usingOrThrow[A](result: Try[A]): A =
    result match
      case Success(value) => value
      case Failure(exception) => throw exception

  /** Helper to open a ZipFile with automatic resource management. */
  private def withZipFile[A](path: Path)(f: ZipFile => A): A =
    usingOrThrow(Using.Manager { use =>
      val zip = use(new ZipFile(path.toFile))
      f(zip)
    })

  /** Read the contents of a ZIP entry as UTF-8 string if it exists. */
  private def readZipEntry(zip: ZipFile, entryName: String): Option[String] =
    Option(zip.getEntry(entryName)).map { entry =>
      usingOrThrow(Using.Manager { use =>
        val in = use(zip.getInputStream(entry))
        new String(in.readAllBytes(), StandardCharsets.UTF_8)
      })
    }

  private def parseOptionalEntry[T](
    zip: ZipFile,
    entryName: String
  )(parse: Elem => Either[String, T]): Option[T] =
    readZipEntry(zip, entryName).flatMap { xmlString =>
      XmlSecurity.parseSafe(xmlString, entryName).toOption.flatMap(elem => parse(elem).toOption)
    }

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
    withZipFile(sourcePath) { sourceZip =>
      val entry = Option(sourceZip.getEntry(partPath)).getOrElse {
        throw new IllegalStateException(s"Entry missing from source: $partPath")
      }

      val newEntry = new ZipEntry(partPath)
      newEntry.setTime(0L)
      newEntry.setMethod(entry.getMethod)

      if entry.getMethod == ZipEntry.STORED then
        newEntry.setSize(entry.getSize)
        newEntry.setCompressedSize(entry.getCompressedSize)
        newEntry.setCrc(entry.getCrc)

      var entryOpen = false
      usingOrThrow(Using.Manager { use =>
        val in = use(sourceZip.getInputStream(entry))
        try
          outputZip.putNextEntry(newEntry)
          entryOpen = true
          val buffer = new Array[Byte](8192)
          var read = in.read(buffer)
          while read != -1 do
            outputZip.write(buffer, 0, read)
            read = in.read(buffer)
        finally if entryOpen then outputZip.closeEntry()
      })
    }

  /**
   * Re-parse structural files from source ZIP for preservation.
   *
   * Returns (ContentTypes, rootRels, workbookRels, workbooks) parsed from the original file. If any
   * file is missing or fails to parse, returns None for that component.
   */
  private def parsePreservedStructure(
    sourcePath: Path
  ): (Option[ContentTypes], Option[Relationships], Option[Relationships], Option[OoxmlWorkbook]) =
    withZipFile(sourcePath) { zip =>
      val contentTypes = parseOptionalEntry(zip, "[Content_Types].xml")(ContentTypes.fromXml)
      val rootRels = parseOptionalEntry(zip, "_rels/.rels")(Relationships.fromXml)
      val workbookRels =
        parseOptionalEntry(zip, "xl/_rels/workbooks.xml.rels")(Relationships.fromXml)
      val workbook = parseOptionalEntry(zip, "xl/workbooks.xml")(OoxmlWorkbook.fromXml)
      (contentTypes, rootRels, workbookRels, workbook)
    }

  /**
   * Parse worksheet XML from source ZIP to extract preserved metadata.
   *
   * Returns the parsed OoxmlWorksheet with all metadata (cols, views, etc.) for merging during
   * regeneration.
   */
  private def parsePreservedWorksheet(
    sourcePath: Path,
    sheetPath: String
  ): XLResult[Option[OoxmlWorksheet]] =
    Try(withZipFile(sourcePath)(zip => readZipEntry(zip, sheetPath))) match
      case Failure(e) =>
        Left(XLError.IOError(s"Failed to read preserved worksheet $sheetPath: ${e.getMessage}"))
      case Success(None) => Right(None)
      case Success(Some(xmlString)) =>
        for
          elem <- XmlSecurity.parseSafe(xmlString, sheetPath)
          worksheet <- OoxmlWorksheet
            .fromXml(elem)
            .left
            .map(err => XLError.ParseError(sheetPath, err): XLError)
        yield Some(worksheet)

  /**
   * Parse SharedStrings from source ZIP for modified sheets cell encoding.
   *
   * CRITICAL: Modified sheets need the SST to encode cells as t="s" (SST references) instead of
   * t="inlineStr" (inline strings). We copy the SST verbatim to output, but parse it here so
   * modified sheets can reference the same indices as unmodified sheets.
   */
  @SuppressWarnings(Array("org.wartremover.warts.Var", "org.wartremover.warts.While"))
  private def parsePreservedSST(sourcePath: Path): Option[SharedStrings] =
    Try(withZipFile(sourcePath)(zip => readZipEntry(zip, "xl/sharedStrings.xml"))) match
      case Failure(_) => None
      case Success(None) => None
      case Success(Some(xmlString)) =>
        XmlSecurity
          .parseSafe(xmlString, "xl/sharedStrings.xml")
          .toOption
          .flatMap(SharedStrings.fromXml(_).toOption)

  /**
   * Parse preserved styles.xml to extract namespace metadata and differential formats.
   *
   * Critical for preserving Excel extension namespaces (x14ac, x16r2, xr, mc:Ignorable) and
   * differential formats used by conditional formatting.
   */
  private def parsePreservedStylesMetadata(
    sourcePath: Path
  ): (Option[MetaData], NamespaceBinding, Option[Elem]) =
    Try(withZipFile(sourcePath)(zip => readZipEntry(zip, "xl/styles.xml"))) match
      case Failure(_) => (None, TopScope, None)
      case Success(None) => (None, TopScope, None)
      case Success(Some(xmlString)) =>
        XmlSecurity.parseSafe(xmlString, "xl/styles.xml").toOption match
          case Some(elem) =>
            val dxfs = (elem \ "dxfs").headOption.collect { case e: Elem => e }
            (Some(elem.attributes), elem.scope, dxfs)
          case None => (None, TopScope, None)

  /**
   * Unified write: intelligently regenerates only what changed, preserves the rest.
   *
   * Strategy:
   *   - With source (surgical mode): Regenerate modified sheets, copy unmodified, preserve unknown
   *     parts
   *   - Without source (new workbooks): Regenerate everything (graceful degradation)
   *
   * This unified path handles both surgical modification and normal writes.
   */
  @SuppressWarnings(Array("org.wartremover.warts.Var", "org.wartremover.warts.While"))
  private def unifiedWrite(
    workbook: Workbook,
    sourceContext: Option[SourceContext],
    target: OutputTarget,
    config: WriterConfig
  ): Unit =
    // Determine modification tracking (all sheets modified if no source)
    val tracker = sourceContext.map(_.modificationTracker).getOrElse {
      ModificationTracker(
        modifiedSheets = workbook.sheets.indices.toSet,
        deletedSheets = Set.empty,
        reorderedSheets = false
      )
    }

    val (graph, preservableParts, regenerateParts) = sourceContext match
      case Some(ctx) =>
        val g = RelationshipGraph.fromManifest(ctx.partManifest)
        val pres = determinePreservableParts(workbook, ctx, g)
        val regen = determineRegenerateParts(workbook, ctx)
        (g, pres, regen)
      case None =>
        (RelationshipGraph.empty, Set.empty[String], Set.empty[String])

    val sharedStringsPath = "xl/sharedStrings.xml"
    val sourceHasSharedStrings = sourceContext.exists(_.partManifest.contains(sharedStringsPath))

    // SST Strategy:
    // - If modified sheets have new strings not in preserved SST → regenerate (to avoid inline string corruption)
    // - If modified sheets only use existing SST strings → copy verbatim (byte-perfect preservation)
    // - If no source SST → generate according to policy
    val (sst, regenerateSharedStrings) =
      if sourceHasSharedStrings then
        // Parse preserved SST (sourceContext guaranteed to exist if sourceHasSharedStrings is true)
        val parsedSST = sourceContext.map(ctx => parsePreservedSST(ctx.sourcePath)).getOrElse(None)

        // Check if modified sheets contain NEW strings not in preserved SST
        // Extract entries from modified sheets (normalize for comparison consistency)
        val modifiedSheetEntries: Set[Either[String, RichText]] =
          tracker.modifiedSheets.flatMap { idx =>
            workbook.sheets.lift(idx).toList.flatMap { sheet =>
              sheet.cells.values.flatMap { cell =>
                cell.value match
                  case CellValue.Text(str) => Some(Left(SharedStrings.normalize(str)))
                  case CellValue.RichText(rt) => Some(Right(rt))
                  case _ => None
              }
            }
          }.toSet

        // Get preserved SST entries (normalize for comparison)
        val preservedEntries: Set[Either[String, RichText]] =
          parsedSST
            .map(
              _.strings
                .map {
                  case Left(s) => Left(SharedStrings.normalize(s))
                  case Right(rt) => Right(rt)
                }
                .toSet
            )
            .getOrElse(Set.empty)

        // Determine if modified sheets introduced new strings (normalized comparison)
        val newEntries = modifiedSheetEntries.diff(preservedEntries)

        if newEntries.nonEmpty then
          // New strings detected → build SST from preserved + new entries only
          // Do NOT use SharedStrings.fromWorkbook (includes ALL sheets, even binary system sheets!)
          val combinedEntries =
            parsedSST.map(_.strings).getOrElse(Vector.empty) ++ newEntries.toVector

          // Calculate total reference count: preserved count + new string references
          val originalTotalCount = parsedSST.map(_.totalCount).getOrElse(0)
          val newStringRefCount = newEntries.size // Each new entry used at least once
          val combinedTotalCount = originalTotalCount + newStringRefCount

          // Create new SST with combined entries
          val combinedSST = SharedStrings(
            strings = combinedEntries,
            indexMap = combinedEntries.zipWithIndex.map {
              case (Left(s), idx) => SharedStrings.normalize(s) -> idx
              case (Right(rt), idx) => SharedStrings.normalize(rt.toPlainText) -> idx
            }.toMap,
            totalCount = combinedTotalCount
          )
          (Some(combinedSST), true)
        else
          // No new strings → copy preserved SST verbatim for byte-perfect preservation
          (parsedSST, false)
      else
        // No source SST - generate if policy allows
        val generated = config.sstPolicy match
          case SstPolicy.Always => Some(SharedStrings.fromWorkbook(workbook))
          case SstPolicy.Never => None
          case SstPolicy.Auto =>
            if SharedStrings.shouldUseSST(workbook) then Some(SharedStrings.fromWorkbook(workbook))
            else None
        (generated, generated.isDefined)

    val sharedStringsInOutput = sourceHasSharedStrings || regenerateSharedStrings
    val sstForSheets = if regenerateSharedStrings then sst else None

    // Build styles index (automatic optimization based on sourceContext)
    val (styleIndex, sheetRemappings) = StyleIndex.fromWorkbook(workbook)

    // Parse preserved styles metadata (namespaces and dxfs) if source available
    val (preservedStylesAttrs, preservedStylesScope, preservedDxfs) = sourceContext match
      case Some(ctx) => parsePreservedStylesMetadata(ctx.sourcePath)
      case None => (None, TopScope, None)

    val styles = OoxmlStyles(styleIndex, preservedStylesAttrs, preservedStylesScope, preservedDxfs)

    // Preserve structural parts from source (or fallback to minimal)
    val (preservedContentTypes, preservedRootRels, preservedWorkbookRels, preservedWorkbook) =
      sourceContext match
        case Some(ctx) => parsePreservedStructure(ctx.sourcePath)
        case None => (None, None, None, None)

    // Use preserved workbooks structure if available, otherwise create minimal
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

    // Set compression level to match Excel (level 1 = super fast, matches original defS)
    zip.setLevel(1)

    try
      // Write structural parts (always regenerated)
      writePart(zip, "[Content_Types].xml", contentTypes.toXml, config)
      writePart(zip, "_rels/.rels", rootRels.toXml, config)
      writePart(zip, "xl/workbooks.xml", ooxmlWb.toXml, config)
      writePart(zip, "xl/_rels/workbooks.xml.rels", workbookRels.toXml, config)
      writePart(zip, "xl/styles.xml", styles.toXml, config)

      if regenerateSharedStrings then
        sst.foreach { sharedStrings =>
          writePart(zip, sharedStringsPath, sharedStrings.toXml, config)
        }
      else if sourceHasSharedStrings then
        sourceContext.foreach(ctx => copyPreservedPart(ctx.sourcePath, sharedStringsPath, zip))

      // Write sheets: regenerate modified, copy unmodified (if source available)
      workbook.sheets.zipWithIndex.foreach { case (sheet, idx) =>
        if tracker.modifiedSheets.contains(idx) then
          // Regenerate modified sheets with preserved metadata (if available)
          val sheetPath = graph.pathForSheet(idx)
          val preservedMetadata = sourceContext
            .map(ctx => parsePreservedWorksheet(ctx.sourcePath, sheetPath))
            .getOrElse(Right(None)) match
            case Right(value) => value
            case Left(err) =>
              throw new IllegalStateException(
                s"Failed to parse preserved worksheet $sheetPath: ${err.message}"
              )
          val remapping = sheetRemappings.getOrElse(idx, Map.empty)
          val ooxmlSheet =
            OoxmlWorksheet.fromDomainWithMetadata(sheet, sstForSheets, remapping, preservedMetadata)
          writePart(zip, s"xl/worksheets/sheet${idx + 1}.xml", ooxmlSheet.toXml, config)
        else
          // Copy unmodified sheets from source (only if source available)
          sourceContext.foreach { ctx =>
            val sheetPath = graph.pathForSheet(idx)
            copyPreservedPart(ctx.sourcePath, sheetPath, zip)
          }
      }

      // Copy preserved parts (charts, drawings, images, etc.) if source available
      sourceContext.foreach { ctx =>
        preservableParts.foreach { path =>
          copyPreservedPart(ctx.sourcePath, path, zip)
        }
      }

    finally zip.close()

  /** Write workbooks to bytes (for testing) */
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
