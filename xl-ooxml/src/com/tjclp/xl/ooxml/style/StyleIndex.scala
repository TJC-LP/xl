package com.tjclp.xl.ooxml.style

import com.tjclp.xl.api.Workbook
import com.tjclp.xl.context.SourceContext
import com.tjclp.xl.ooxml.XmlSecurity
import com.tjclp.xl.styles.CellStyle
import com.tjclp.xl.styles.border.Border
import com.tjclp.xl.styles.fill.Fill
import com.tjclp.xl.styles.font.Font
import com.tjclp.xl.styles.numfmt.NumFmt
import com.tjclp.xl.styles.units.StyleId

/**
 * Style components and indexing for xl/styles.xml
 *
 * Styles are deduplicated by canonical keys to avoid Excel's 64k style limit. The StyleIndex builds
 * collections of unique fonts, fills, borders, and cellXfs.
 */

/** Index mapping for style components */
final case class StyleIndex(
  fonts: Vector[Font],
  fills: Vector[Fill],
  borders: Vector[Border],
  numFmts: Vector[(Int, NumFmt)], // Custom formats with IDs
  cellStyles: Vector[CellStyle],
  styleToIndex: Map[String, StyleId] // Canonical key -> cellXf index
):
  /** Get style index for a CellStyle (returns 0 if not found - default style) */
  def indexOf(style: CellStyle): StyleId =
    styleToIndex.getOrElse(style.canonicalKey, StyleId(0))

object StyleIndex:
  /** Empty StyleIndex with only default style (useful for testing) */
  val empty: StyleIndex = StyleIndex(
    fonts = Vector.empty,
    fills = Vector.empty,
    borders = Vector.empty,
    numFmts = Vector.empty,
    cellStyles = Vector(CellStyle.default),
    styleToIndex = Map(CellStyle.default.canonicalKey -> StyleId(0))
  )

  /**
   * Build unified style index from workbook with automatic optimization.
   *
   * Strategy (automatic based on workbook.sourceContext):
   *   - **With source**: Preserve original styles for byte-perfect surgical modification
   *   - **Without source**: Full deduplication for optimal compression
   *
   * Users don't choose the strategy - the method transparently optimizes based on available
   * context. This enables read-modify-write workflows to preserve structure automatically while
   * allowing programmatic creation to produce optimal output.
   *
   * @param wb
   *   The workbook to index
   * @return
   *   (StyleIndex for writing, Map[sheetIndex -> Map[localStyleId -> globalStyleId]])
   */
  def fromWorkbook(wb: Workbook): (StyleIndex, Map[Int, Map[Int, Int]]) =
    wb.sourceContext match
      case Some(ctx) =>
        // Has source: surgical mode (preserve original structure)
        fromWorkbookWithSource(wb, ctx)
      case None =>
        // No source: full deduplication (optimal compression)
        fromWorkbookWithoutSource(wb)

  /**
   * Build unified style index from a workbook with full deduplication (no source).
   *
   * Extracts styles from each sheet's StyleRegistry, builds a unified index with deduplication, and
   * creates remapping tables to convert sheet-local styleIds to workbook-level indices.
   *
   * @param wb
   *   The workbook to index
   * @return
   *   (StyleIndex, Map[sheetIndex -> Map[localStyleId -> globalStyleId]])
   */
  @SuppressWarnings(Array("org.wartremover.warts.Var"))
  private def fromWorkbookWithoutSource(wb: Workbook): (StyleIndex, Map[Int, Map[Int, Int]]) =
    import scala.collection.mutable

    // Build unified style index by merging all sheet registries
    // Optimization: Use VectorBuilder instead of Vector :+ for O(1) amortized append (was O(n) per append = O(n^2) total)
    val stylesBuilder = Vector.newBuilder[CellStyle]
    stylesBuilder += CellStyle.default
    var unifiedIndex = Map(CellStyle.default.canonicalKey -> StyleId(0))
    var nextIdx = 1

    // Build per-sheet remapping tables
    val remappings = wb.sheets.zipWithIndex.map { case (sheet, sheetIdx) =>
      val registry = sheet.styleRegistry
      val remapping = mutable.Map[Int, Int]()

      // Map each local styleId to global index
      registry.styles.zipWithIndex.foreach { case (style, localIdx) =>
        val key = style.canonicalKey

        unifiedIndex.get(key) match
          case Some(globalIdx) =>
            // Style already in unified index (deduplication)
            remapping(localIdx) = globalIdx.value
          case None =>
            // New style - add to unified index
            stylesBuilder += style
            unifiedIndex = unifiedIndex + (key -> StyleId(nextIdx))
            remapping(localIdx) = nextIdx
            nextIdx += 1
      }

      sheetIdx -> remapping.toMap
    }.toMap

    val unifiedStyles = stylesBuilder.result()

    // Deduplicate components using LinkedHashSet for O(1) deduplication (60-80% faster than .distinct)
    // Optimization: Single-pass collection instead of three separate passes (was O(3n), now O(n))
    import scala.collection.mutable
    val (fontSet, fillSet, borderSet) = {
      val fonts = mutable.LinkedHashSet.empty[Font]
      val fills = mutable.LinkedHashSet.empty[Fill]
      val borders = mutable.LinkedHashSet.empty[Border]
      unifiedStyles.foreach { style =>
        fonts += style.font
        fills += style.fill
        borders += style.border
      }
      (fonts, fills, borders)
    }
    val uniqueFonts = fontSet.toVector
    val uniqueFills = fillSet.toVector
    val uniqueBorders = borderSet.toVector

    // Collect custom number formats (built-ins don't need entries)
    val customNumFmts = {
      val seen = mutable.LinkedHashSet.empty[String]
      unifiedStyles.foreach { style =>
        style.numFmt match
          case NumFmt.Custom(code) => seen += code
          case _ => ()
      }
      seen.toVector.zipWithIndex.map { case (code, idx) =>
        (164 + idx, NumFmt.Custom(code))
      }
    }

    val styleIndex = StyleIndex(
      uniqueFonts,
      uniqueFills,
      uniqueBorders,
      customNumFmts,
      unifiedStyles,
      unifiedIndex
    )

    (styleIndex, remappings)

  /**
   * Build style index for workbook with source, preserving original styles.
   *
   * This variant is used during surgical modification to avoid corruption:
   *   - Deduplicates styles ONLY from modified sheets (optimal compression)
   *   - Preserves original styles from source for unmodified sheets (no remapping needed)
   *   - Ensures unmodified sheets' style references remain valid after write
   *
   * Strategy:
   *   1. Parse original styles.xml to get complete WorkbookStyles
   *   2. Deduplicate styles from modified sheets only
   *   3. Ensure all original styles are present in output (fill gaps if needed)
   *   4. Return remappings ONLY for modified sheets (unmodified sheets use original IDs)
   *
   * @param wb
   *   The workbook with modified sheets
   * @param ctx
   *   Source context providing modification tracker and original file path
   * @return
   *   (StyleIndex with all original + deduplicated styles, Map[modifiedSheetIdx -> remapping])
   */
  @SuppressWarnings(
    Array(
      "org.wartremover.warts.Var",
      "org.wartremover.warts.While",
      "org.wartremover.warts.IterableOps" // .head is safe - groupBy guarantees non-empty lists
    )
  )
  private def fromWorkbookWithSource(
    wb: Workbook,
    ctx: SourceContext
  ): (StyleIndex, Map[Int, Map[Int, Int]]) =
    // Extract values from context
    val tracker = ctx.modificationTracker
    val modifiedSheetIndices = tracker.modifiedSheets
    val sourcePath = ctx.sourcePath
    import scala.collection.mutable
    import java.util.zip.ZipInputStream
    import java.io.FileInputStream

    // Step 1: Parse original styles.xml to get ALL components (byte-perfect preservation)
    val originalWorkbookStyles: WorkbookStyles = {
      val sourceZip = new ZipInputStream(new FileInputStream(sourcePath.toFile))
      try
        var entry = sourceZip.getNextEntry
        var result: WorkbookStyles = WorkbookStyles.default

        while entry != null && result == WorkbookStyles.default do
          if entry.getName == "xl/styles.xml" then
            val content = new String(sourceZip.readAllBytes(), "UTF-8")
            // Parse styles.xml using WorkbookStyles parser (with XXE protection)
            XmlSecurity.parseSafe(content, "xl/styles.xml").toOption.foreach { parsed =>
              WorkbookStyles.fromXml(parsed).foreach { wbStyles =>
                result = wbStyles
              }
            }

          sourceZip.closeEntry()
          entry = sourceZip.getNextEntry

        result
      finally sourceZip.close()
    }

    val originalStyles = originalWorkbookStyles.cellStyles

    // Step 2: Build index preserving ALL original styles (including duplicates)
    // Use groupBy to map each canonicalKey to ALL indices that have that key
    // This prevents the critical bug where duplicate styles in source get lost
    var unifiedStyles = originalStyles
    val unifiedIndex: Map[String, List[Int]] = originalStyles.zipWithIndex
      .groupBy { case (style, _) => style.canonicalKey }
      .view
      .mapValues(_.map(_._2).toList)
      .toMap
    var nextIdx = originalStyles.size
    var additionalStyles = mutable.Map[String, Int]() // Track styles added after original

    // Step 3: Build mutable component collections starting from original
    // These will grow if new styles introduce new fonts/fills/borders/numFmts
    val fontsBuilder = mutable.ArrayBuffer.from(originalWorkbookStyles.fonts)
    val fillsBuilder = mutable.ArrayBuffer.from(originalWorkbookStyles.fills)
    val bordersBuilder = mutable.ArrayBuffer.from(originalWorkbookStyles.borders)
    val numFmtsBuilder = mutable.ArrayBuffer.from(originalWorkbookStyles.customNumFmts)

    // Build lookup sets for O(1) deduplication of new components
    val fontSet = mutable.LinkedHashSet.from(originalWorkbookStyles.fonts)
    val fillSet = mutable.LinkedHashSet.from(originalWorkbookStyles.fills)
    val borderSet = mutable.LinkedHashSet.from(originalWorkbookStyles.borders)
    val numFmtCodeSet = mutable.Set.from(originalWorkbookStyles.customNumFmts.map(_._2 match {
      case NumFmt.Custom(code) => code
      case _ => ""
    }))
    var nextNumFmtId = if numFmtsBuilder.isEmpty then 164 else numFmtsBuilder.map(_._1).max + 1

    // Step 4: Process ONLY modified sheets for style remapping
    val remappings = wb.sheets.zipWithIndex.flatMap { case (sheet, sheetIdx) =>
      if modifiedSheetIndices.contains(sheetIdx) then
        val registry = sheet.styleRegistry
        val remapping = mutable.Map[Int, Int]()

        // Map each local styleId to global index
        registry.styles.zipWithIndex.foreach { case (style, localIdx) =>
          val key = style.canonicalKey

          // First, check if this key exists in original styles
          unifiedIndex.get(key) match
            case Some(indices) =>
              // Style exists in original - use FIRST matching index
              // This preserves original layout and avoids adding duplicates
              remapping(localIdx) = indices.head
            case None =>
              // Not in original - check if we've already added it
              additionalStyles.get(key) match
                case Some(addedIdx) =>
                  // Already added by earlier sheet processing
                  remapping(localIdx) = addedIdx
                case None =>
                  // Truly new style - add it now
                  unifiedStyles = unifiedStyles :+ style
                  additionalStyles(key) = nextIdx
                  remapping(localIdx) = nextIdx
                  nextIdx += 1

                  // CRITICAL: Also add new font/fill/border/numFmt if not already present
                  // Without this, new styles reference non-existent component indices
                  if !fontSet.contains(style.font) then
                    fontSet += style.font
                    fontsBuilder += style.font

                  if !fillSet.contains(style.fill) then
                    fillSet += style.fill
                    fillsBuilder += style.fill

                  if !borderSet.contains(style.border) then
                    borderSet += style.border
                    bordersBuilder += style.border

                  style.numFmt match
                    case NumFmt.Custom(code) if !numFmtCodeSet.contains(code) =>
                      numFmtCodeSet += code
                      numFmtsBuilder += ((nextNumFmtId, NumFmt.Custom(code)))
                      nextNumFmtId += 1
                    case _ => ()
        }

        Some(sheetIdx -> remapping.toMap)
      else
        // Unmodified sheet - no remapping needed (uses original style IDs)
        None
    }.toMap

    // Step 5: Finalize component vectors (original + any new components)
    val uniqueFonts = fontsBuilder.toVector
    val uniqueFills = fillsBuilder.toVector
    val uniqueBorders = bordersBuilder.toVector
    val customNumFmts = numFmtsBuilder.toVector

    // Convert unifiedIndex back to Map[String, StyleId] for StyleIndex
    // Use first index from each canonicalKey's list (preserves original layout)
    val styleToIndexMap = unifiedIndex.view.mapValues(indices => StyleId(indices.head)).toMap

    val styleIndex = StyleIndex(
      uniqueFonts,
      uniqueFills,
      uniqueBorders,
      customNumFmts,
      unifiedStyles,
      styleToIndexMap
    )

    (styleIndex, remappings)
