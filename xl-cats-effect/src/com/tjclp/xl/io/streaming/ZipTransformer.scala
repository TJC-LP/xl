package com.tjclp.xl.io.streaming

import cats.effect.Sync
import cats.syntax.all.*
import java.io.{ByteArrayInputStream, ByteArrayOutputStream, FileInputStream, FileOutputStream}
import java.nio.charset.StandardCharsets
import java.nio.file.{Files as JFiles, Path}
import java.util.zip.{ZipEntry, ZipFile, ZipInputStream, ZipOutputStream}
import com.tjclp.xl.addressing.{ARef, CellRange}
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.styles.CellStyle
import scala.collection.mutable
import scala.util.Using

/**
 * ZIP-level transformer for streaming worksheet modifications.
 *
 * Orchestrates the streaming transform of XLSX files:
 *   1. Copy unchanged entries verbatim (O(1) memory)
 *   2. Transform styles.xml with new styles (if needed)
 *   3. Stream-transform target worksheets with cell patches
 *
 * Supports:
 *   - Style changes (transformStyle)
 *   - Value changes (transformValues)
 *   - Combined style + value changes (transform)
 *
 * Memory: O(styles.xml size + patches size) - typically <2MB
 */
object ZipTransformer:

  /**
   * Result of streaming transform operation.
   *
   * @param cellCount
   *   Number of cells modified
   * @param stylesAdded
   *   Number of new styles added to styles.xml
   */
  final case class TransformResult(
    cellCount: Long,
    stylesAdded: Int
  )

  /**
   * Transform XLSX file, applying style to cells in range with O(1) memory.
   *
   * This is the main entry point for streaming style operations.
   *
   * @param source
   *   Input XLSX file
   * @param output
   *   Output XLSX file
   * @param worksheetPath
   *   Worksheet entry path inside the XLSX (e.g., xl/worksheets/sheet1.xml)
   * @param range
   *   Cell range to style
   * @param style
   *   CellStyle to apply
   * @param replace
   *   If true, replace style; if false, merge with existing
   * @return
   *   TransformResult with operation statistics
   */
  def transformStyle[F[_]: Sync](
    source: Path,
    output: Path,
    worksheetPath: String,
    range: CellRange,
    style: CellStyle,
    replace: Boolean
  ): F[TransformResult] =
    Sync[F].delay {
      transformStyleSync(source, output, worksheetPath, range, style, replace)
    }

  /**
   * Transform XLSX file, applying value changes to cells with O(1) memory.
   *
   * This is the main entry point for streaming put operations.
   *
   * @param source
   *   Input XLSX file
   * @param output
   *   Output XLSX file
   * @param worksheetPath
   *   Worksheet entry path inside the XLSX (e.g., xl/worksheets/sheet1.xml)
   * @param values
   *   Map of cell references to new values
   * @return
   *   TransformResult with operation statistics
   */
  def transformValues[F[_]: Sync](
    source: Path,
    output: Path,
    worksheetPath: String,
    values: Map[ARef, CellValue]
  ): F[TransformResult] =
    Sync[F].delay {
      transformValuesSync(source, output, worksheetPath, values)
    }

  /**
   * Transform XLSX file, applying arbitrary cell patches with O(1) memory.
   *
   * This is the most flexible entry point - supports any combination of style and value changes.
   *
   * @param source
   *   Input XLSX file
   * @param output
   *   Output XLSX file
   * @param worksheetPath
   *   Worksheet entry path inside the XLSX (e.g., xl/worksheets/sheet1.xml)
   * @param patches
   *   Map of cell references to patches
   * @param updatedStylesXml
   *   Optional updated styles.xml content (if styles were modified)
   * @return
   *   TransformResult with operation statistics
   */
  def transform[F[_]: Sync](
    source: Path,
    output: Path,
    worksheetPath: String,
    patches: Map[ARef, StreamingTransform.CellPatch],
    updatedStylesXml: Option[String] = None
  ): F[TransformResult] =
    Sync[F].delay {
      transformSync(source, output, worksheetPath, patches, updatedStylesXml)
    }

  /**
   * Synchronous implementation of streaming value transform.
   */
  @SuppressWarnings(Array("org.wartremover.warts.Var", "org.wartremover.warts.While"))
  private def transformValuesSync(
    source: Path,
    output: Path,
    worksheetPath: String,
    values: Map[ARef, CellValue]
  ): TransformResult =
    // Convert values to patches
    val patches = values.map { case (ref, value) =>
      ref -> StreamingTransform.CellPatch.SetValue(value, preserveStyle = true)
    }

    // No style changes needed - pass through original styles.xml
    transformZipNoStyleChanges(source, output, worksheetPath, patches)

    TransformResult(patches.size.toLong, 0)

  /**
   * Synchronous implementation of generic transform.
   */
  @SuppressWarnings(Array("org.wartremover.warts.Var", "org.wartremover.warts.While"))
  private def transformSync(
    source: Path,
    output: Path,
    worksheetPath: String,
    patches: Map[ARef, StreamingTransform.CellPatch],
    updatedStylesXml: Option[String]
  ): TransformResult =
    updatedStylesXml match
      case Some(stylesXml) =>
        transformZip(source, output, worksheetPath, stylesXml, patches)
      case None =>
        transformZipNoStyleChanges(source, output, worksheetPath, patches)

    TransformResult(patches.size.toLong, 0)

  /**
   * Synchronous implementation of streaming style transform.
   */
  @SuppressWarnings(Array("org.wartremover.warts.Var", "org.wartremover.warts.While"))
  private def transformStyleSync(
    source: Path,
    output: Path,
    worksheetPath: String,
    range: CellRange,
    style: CellStyle,
    replace: Boolean
  ): TransformResult =
    // Phase 1: Read styles.xml and build patches
    val (stylesXml, patches, stylesAdded) =
      buildStylePatches(source, worksheetPath, range, style, replace)

    // Phase 2: Transform ZIP with patched styles and worksheet
    transformZip(source, output, worksheetPath, stylesXml, patches)

    TransformResult(patches.size.toLong, stylesAdded)

  /**
   * Build style patches for the target range.
   *
   * For replace mode: All cells get the same new style ID. For merge mode: Each cell's existing
   * style is merged with the new style.
   *
   * @return
   *   (updated styles.xml, Map[ARef -> CellPatch], number of styles added)
   */
  @SuppressWarnings(Array("org.wartremover.warts.Var", "org.wartremover.warts.While"))
  private def buildStylePatches(
    source: Path,
    worksheetPath: String,
    range: CellRange,
    style: CellStyle,
    replace: Boolean
  ): (String, Map[ARef, StreamingTransform.CellPatch], Int) =
    val zipFile = new ZipFile(source.toFile)
    try
      // Read styles.xml
      val stylesEntry = zipFile.getEntry("xl/styles.xml")
      val stylesXml =
        if stylesEntry != null then
          new String(zipFile.getInputStream(stylesEntry).readAllBytes(), StandardCharsets.UTF_8)
        else
          // Minimal styles.xml if none exists
          minimalStylesXml

      val targetCells = range.cells.toSet

      if replace then
        // Replace mode: All cells get same new style
        val (updatedStyles, newStyleId) = StylePatcher.addStyle(stylesXml, style)
        val patches = targetCells.map { ref =>
          ref -> StreamingTransform.CellPatch.SetStyle(newStyleId)
        }.toMap
        (updatedStyles, patches, 1)
      else
        // Merge mode: Two-pass approach
        // Pass 1: Scan existing style IDs for cells in range
        val normalizedWorksheetPath = normalizeWorksheetPath(worksheetPath)
        val worksheetEntry = zipFile.getEntry(normalizedWorksheetPath)
        if worksheetEntry == null then
          throw new Exception(s"Worksheet not found: $normalizedWorksheetPath")

        val existingStyles = StreamingTransform.scanExistingStyles(
          zipFile.getInputStream(worksheetEntry),
          targetCells
        )

        // Group cells by existing style ID for efficient style creation
        val cellsByStyleId = existingStyles.groupBy(_._2)

        // For each unique existing style, merge and add to styles.xml
        var currentStylesXml = stylesXml
        val styleIdMapping = mutable.Map[Int, Int]() // oldStyleId -> newStyleId
        var stylesAdded = 0

        cellsByStyleId.keys.foreach { oldStyleId =>
          val existingStyle = StylePatcher
            .getStyle(currentStylesXml, oldStyleId)
            .getOrElse(CellStyle.default)
          val mergedStyle = StylePatcher.mergeStyles(existingStyle, style)

          val (updated, newId) = StylePatcher.addStyle(currentStylesXml, mergedStyle)
          currentStylesXml = updated
          styleIdMapping(oldStyleId) = newId
          stylesAdded += 1
        }

        // Build patches using the mapping
        val patches = targetCells.map { ref =>
          val oldStyleId = existingStyles.getOrElse(ref, 0)
          val newStyleId = styleIdMapping.getOrElse(
            oldStyleId, {
              // Cell not in scan results - use style for default (0)
              if !styleIdMapping.contains(0) then
                val mergedStyle = StylePatcher.mergeStyles(CellStyle.default, style)
                val (updated, newId) = StylePatcher.addStyle(currentStylesXml, mergedStyle)
                currentStylesXml = updated
                styleIdMapping(0) = newId
                stylesAdded += 1
              styleIdMapping(0)
            }
          )
          ref -> StreamingTransform.CellPatch.SetStyle(newStyleId)
        }.toMap

        (currentStylesXml, patches, stylesAdded)
    finally zipFile.close()

  /**
   * Transform ZIP file, copying unchanged entries and transforming styles + worksheet.
   */
  @SuppressWarnings(Array("org.wartremover.warts.Var", "org.wartremover.warts.While"))
  private def transformZip(
    source: Path,
    output: Path,
    worksheetPath: String,
    stylesXml: String,
    patches: Map[ARef, StreamingTransform.CellPatch]
  ): Unit =
    val normalizedWorksheetPath = normalizeWorksheetPath(worksheetPath)

    // Use temp file for atomic write (prevents corruption if source == output)
    val parent = Option(output.getParent).getOrElse(Path.of("."))
    val tempPath = JFiles.createTempFile(parent, ".xl-stream-", ".tmp")

    try
      val zipIn = new ZipInputStream(new FileInputStream(source.toFile))
      val zipOut = new ZipOutputStream(new FileOutputStream(tempPath.toFile))
      zipOut.setLevel(1) // Match Excel's compression level

      try
        var entry = zipIn.getNextEntry
        while entry != null do
          val entryName = entry.getName

          if entryName == "xl/styles.xml" then
            // Write patched styles.xml
            writeEntry(zipOut, entryName, stylesXml.getBytes(StandardCharsets.UTF_8))
            zipIn.closeEntry()
          else if entryName == normalizedWorksheetPath then
            // Transform worksheet with streaming
            val transformed = transformWorksheetEntry(zipIn, patches)
            writeEntry(zipOut, entryName, transformed)
            zipIn.closeEntry()
          else
            // Copy unchanged entry
            copyEntry(zipIn, zipOut, entry)

          entry = zipIn.getNextEntry
      finally
        zipIn.close()
        zipOut.close()

      // Atomic rename
      try
        JFiles.move(
          tempPath,
          output,
          java.nio.file.StandardCopyOption.REPLACE_EXISTING,
          java.nio.file.StandardCopyOption.ATOMIC_MOVE
        )
      catch
        case _: java.nio.file.AtomicMoveNotSupportedException =>
          JFiles.move(tempPath, output, java.nio.file.StandardCopyOption.REPLACE_EXISTING)
    catch
      case e: Exception =>
        JFiles.deleteIfExists(tempPath)
        throw e

  /**
   * Transform ZIP file without modifying styles.xml.
   *
   * Used for value-only transforms where no style changes are needed.
   */
  @SuppressWarnings(Array("org.wartremover.warts.Var", "org.wartremover.warts.While"))
  private def transformZipNoStyleChanges(
    source: Path,
    output: Path,
    worksheetPath: String,
    patches: Map[ARef, StreamingTransform.CellPatch]
  ): Unit =
    val normalizedWorksheetPath = normalizeWorksheetPath(worksheetPath)

    // Use temp file for atomic write (prevents corruption if source == output)
    val parent = Option(output.getParent).getOrElse(Path.of("."))
    val tempPath = JFiles.createTempFile(parent, ".xl-stream-", ".tmp")

    try
      val zipIn = new ZipInputStream(new FileInputStream(source.toFile))
      val zipOut = new ZipOutputStream(new FileOutputStream(tempPath.toFile))
      zipOut.setLevel(1) // Match Excel's compression level

      try
        var entry = zipIn.getNextEntry
        while entry != null do
          val entryName = entry.getName

          if entryName == normalizedWorksheetPath then
            // Transform worksheet with streaming
            val transformed = transformWorksheetEntry(zipIn, patches)
            writeEntry(zipOut, entryName, transformed)
            zipIn.closeEntry()
          else
            // Copy unchanged entry (including styles.xml)
            copyEntry(zipIn, zipOut, entry)

          entry = zipIn.getNextEntry
      finally
        zipIn.close()
        zipOut.close()

      // Atomic rename
      try
        JFiles.move(
          tempPath,
          output,
          java.nio.file.StandardCopyOption.REPLACE_EXISTING,
          java.nio.file.StandardCopyOption.ATOMIC_MOVE
        )
      catch
        case _: java.nio.file.AtomicMoveNotSupportedException =>
          JFiles.move(tempPath, output, java.nio.file.StandardCopyOption.REPLACE_EXISTING)
    catch
      case e: Exception =>
        JFiles.deleteIfExists(tempPath)
        throw e

  /**
   * Transform worksheet entry using streaming transform with early-abort optimization.
   *
   * For patches targeting only early rows, this aborts SAX parsing after processing target rows and
   * splices the remaining bytes directly (mergeCells, pageMargins, etc.).
   *
   * Note: We read the entry into a buffer first because SAX parsers don't work well with
   * ZipInputStream directly - they may try to close/read past the entry.
   */
  @SuppressWarnings(Array("org.wartremover.warts.Var", "org.wartremover.warts.While"))
  private def transformWorksheetEntry(
    zipIn: ZipInputStream,
    patches: Map[ARef, StreamingTransform.CellPatch]
  ): Array[Byte] =
    // Read entry into buffer first (SAX doesn't work well with ZipInputStream)
    val entryBuffer = new ByteArrayOutputStream()
    val buffer = new Array[Byte](8192)
    var read = zipIn.read(buffer)
    while read != -1 do
      entryBuffer.write(buffer, 0, read)
      read = zipIn.read(buffer)

    val entryBytes = entryBuffer.toByteArray

    // Use early-abort optimization for large files with patches in early rows
    val analysis = StreamingTransform.analyzePatches(patches)
    val useEarlyAbort = analysis.exists { a =>
      // Only use early abort if:
      // 1. File is large enough (>1MB decompressed)
      // 2. Max target row is early in the file (rough heuristic: max row < 10000)
      entryBytes.length > 1_000_000 && a.maxRow < 10_000
    }

    if useEarlyAbort then transformWorksheetEntryWithEarlyAbort(entryBytes, patches)
    else
      // Standard full transform
      val entryIn = new ByteArrayInputStream(entryBytes)
      val baos = new ByteArrayOutputStream()
      StreamingTransform.transformWorksheet(entryIn, baos, patches)
      baos.toByteArray

  /**
   * Transform worksheet entry with early-abort optimization.
   *
   * Uses byte splicing to avoid parsing rows after the target range:
   *
   * ```
   * Input:  <sheetData>...<row r="1">TRANSFORM</row><row r="2">...</row>...</sheetData><mergeCells>...
   * Output: <sheetData>...<row r="1">MODIFIED</row><row r="2">...</row>...</sheetData><mergeCells>...
   *         └────── SAX transform ──────┘└─────────── byte splice from original ──────────────────┘
   * ```
   *
   * The SAX transform handles rows up to the target range, then aborts. Remaining rows are copied
   * directly from the original bytes, preserving them without parsing.
   */
  @SuppressWarnings(
    Array(
      "org.wartremover.warts.Var",
      "org.wartremover.warts.While",
      "org.wartremover.warts.Return"
    )
  )
  private def transformWorksheetEntryWithEarlyAbort(
    entryBytes: Array[Byte],
    patches: Map[ARef, StreamingTransform.CellPatch]
  ): Array[Byte] =
    // Use early-abort transform
    val result = StreamingTransform.transformWorksheetWithEarlyAbort(entryBytes, patches)

    if result.aborted then
      // Find the start of the row where we aborted in original bytes
      // Try common patterns: <row r="N" or <row  r="N" (with space/newline)
      val rowPatterns = Seq(
        s"""<row r="${result.abortedAtRow}"""",
        s"""<row  r="${result.abortedAtRow}"""" // Sometimes has extra whitespace
      ).map(_.getBytes(StandardCharsets.UTF_8))

      val rowStartPosition = rowPatterns.view.flatMap(p => indexOfBytes(entryBytes, p)).headOption

      rowStartPosition match
        case Some(spliceFrom) =>
          // Splice: transformed prefix + original from aborted row to end
          val output = new ByteArrayOutputStream(entryBytes.length)
          output.write(result.outputBytes)
          output.write(entryBytes, spliceFrom, entryBytes.length - spliceFrom)
          output.toByteArray
        case None =>
          // Can't find row - fall back to full transform
          val entryIn = new ByteArrayInputStream(entryBytes)
          val baos = new ByteArrayOutputStream()
          StreamingTransform.transformWorksheet(entryIn, baos, patches)
          baos.toByteArray
    else
      // No early abort - use full transform result
      result.outputBytes

  /**
   * Find index of byte sequence in array.
   *
   * Uses naive search - sufficient for our use case where pattern is short and appears early in
   * large files.
   */
  @SuppressWarnings(
    Array(
      "org.wartremover.warts.Var",
      "org.wartremover.warts.While",
      "org.wartremover.warts.Return"
    )
  )
  private def indexOfBytes(haystack: Array[Byte], needle: Array[Byte]): Option[Int] =
    if needle.isEmpty || haystack.length < needle.length then return None

    val end = haystack.length - needle.length
    var i = 0
    while i <= end do
      var j = 0
      var matches = true
      while j < needle.length && matches do
        if haystack(i + j) != needle(j) then matches = false
        j += 1
      if matches then return Some(i)
      i += 1
    None

  private def normalizeWorksheetPath(path: String): String =
    val trimmed = if path.startsWith("/") then path.drop(1) else path
    if trimmed.startsWith("xl/") then trimmed else s"xl/$trimmed"

  /**
   * Write entry to ZIP with proper metadata.
   */
  private def writeEntry(
    zipOut: ZipOutputStream,
    name: String,
    data: Array[Byte]
  ): Unit =
    val entry = new ZipEntry(name)
    entry.setTime(0L) // Deterministic timestamp
    entry.setMethod(ZipEntry.DEFLATED)
    zipOut.putNextEntry(entry)
    zipOut.write(data)
    zipOut.closeEntry()

  /**
   * Copy entry from input to output ZIP.
   */
  @SuppressWarnings(Array("org.wartremover.warts.Var", "org.wartremover.warts.While"))
  private def copyEntry(
    zipIn: ZipInputStream,
    zipOut: ZipOutputStream,
    sourceEntry: ZipEntry
  ): Unit =
    val newEntry = new ZipEntry(sourceEntry.getName)
    newEntry.setTime(0L) // Deterministic timestamp
    newEntry.setMethod(ZipEntry.DEFLATED)
    zipOut.putNextEntry(newEntry)

    val buffer = new Array[Byte](8192)
    var read = zipIn.read(buffer)
    while read != -1 do
      zipOut.write(buffer, 0, read)
      read = zipIn.read(buffer)

    zipOut.closeEntry()

  /**
   * Minimal styles.xml for files that don't have one.
   */
  private val minimalStylesXml: String =
    """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
      |<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
      |<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>
      |<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>
      |<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
      |<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
      |<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>
      |</styleSheet>""".stripMargin.replaceAll("\n", "")
