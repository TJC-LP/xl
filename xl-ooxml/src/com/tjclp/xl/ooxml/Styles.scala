package com.tjclp.xl.ooxml

import scala.xml.*
import XmlUtil.*
import SaxSupport.*
import com.tjclp.xl.api.*
import com.tjclp.xl.context.SourceContext
import com.tjclp.xl.styles.{CellStyle, StyleRegistry}
import com.tjclp.xl.styles.alignment.{Align, HAlign, VAlign}
import com.tjclp.xl.styles.border.{Border, BorderSide, BorderStyle}
import com.tjclp.xl.styles.color.{Color, ThemeSlot}
import com.tjclp.xl.styles.fill.{Fill, PatternType}
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
  styleToIndex: Map[String, StyleId] // Canonical key → cellXf index
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
    // Optimization: Use VectorBuilder instead of Vector :+ for O(1) amortized append (was O(n) per append = O(n²) total)
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

private val defaultStylesScope =
  NamespaceBinding(null, nsSpreadsheetML, TopScope)

/** Serializer for xl/styles.xml */
final case class OoxmlStyles(
  index: StyleIndex,
  rootAttributes: Option[MetaData] = None,
  rootScope: NamespaceBinding = defaultStylesScope,
  preservedDxfs: Option[Elem] = None // Differential formats for conditional formatting
) extends XmlWritable,
      SaxSerializable:

  def toXml: Elem =
    // Number formats (only custom ones; built-ins are implicit)
    val numFmtsElem =
      if index.numFmts.nonEmpty then
        Some(
          elem("numFmts", "count" -> index.numFmts.size.toString)(
            index.numFmts.sortBy(_._1).map { case (id, fmt) =>
              val code = fmt match
                case NumFmt.Custom(c) => c
                case _ => "General" // Shouldn't happen
              elem("numFmt", "numFmtId" -> id.toString, "formatCode" -> code)()
            }*
          )
        )
      else None

    // Fonts
    val fontsElem = elem("fonts", "count" -> index.fonts.size.toString)(
      index.fonts.map(fontToXml)*
    )

    /**
     * Default fills required by OOXML spec (ECMA-376 Part 1, §18.8.21)
     *
     * REQUIRES: None ENSURES:
     *   - Vector contains exactly 2 fills
     *   - Index 0: Fill.None (patternType="none")
     *   - Index 1: Fill.Pattern with Gray125 pattern
     *   - Gray125 uses black foreground (0xFF000000) and silver background (0xFFC0C0C0)
     * DETERMINISTIC: Yes (immutable constant) ERROR CASES: None (compile-time constant)
     */
    val defaultFills = Vector(
      Fill.None,
      Fill.Pattern(
        foreground = Color.Rgb(0xff000000), // Black foreground
        background = Color.Rgb(0xffc0c0c0), // Silver background
        pattern = PatternType.Gray125
      )
    )
    // Deterministic deduplication preserving first occurrence order
    // Manual tracking ensures determinism while avoiding duplicate defaults
    val allFills = {
      val builder = Vector.newBuilder[Fill]
      val seen = scala.collection.mutable.Set.empty[Fill]

      for (fill <- defaultFills ++ index.fills) {
        if (!seen.contains(fill)) {
          seen += fill
          builder += fill
        }
      }
      builder.result()
    }
    val fillsElem = elem("fills", "count" -> allFills.size.toString)(
      allFills.map(fillToXml)*
    )

    // Borders
    val bordersElem = elem("borders", "count" -> index.borders.size.toString)(
      index.borders.map(borderToXml)*
    )

    // CellXfs (cell format styles)
    // Pre-build lookup maps for O(1) access instead of O(n) indexOf
    val fontMap = index.fonts.zipWithIndex.toMap
    val fillMap = allFills.zipWithIndex.toMap
    val borderMap = index.borders.zipWithIndex.toMap

    val cellXfsElem = elem("cellXfs", "count" -> index.cellStyles.size.toString)(
      index.cellStyles.map { style =>
        // Use 0 (default style) as fallback - OOXML requires non-negative indices
        val fontIdx = fontMap.getOrElse(style.font, 0)
        val fillIdx = fillMap.getOrElse(style.fill, 0)
        val borderIdx = borderMap.getOrElse(style.border, 0)
        val numFmtId = style.numFmtId.getOrElse {
          // No raw ID → derive from NumFmt enum (programmatic creation)
          NumFmt
            .builtInId(style.numFmt)
            .getOrElse(
              index.numFmts.find(_._2 == style.numFmt).map(_._1).getOrElse(0)
            )
        }

        // Serialize alignment as child element if non-default
        val alignmentChild = alignmentToXml(style.align).toSeq
        val hasAlignment = alignmentChild.nonEmpty

        elem(
          "xf",
          "applyAlignment" -> (if hasAlignment then "1" else "0"),
          "borderId" -> borderIdx.toString,
          "fillId" -> fillIdx.toString,
          "fontId" -> fontIdx.toString,
          "numFmtId" -> numFmtId.toString,
          "xfId" -> "0"
        )(alignmentChild*)
      }*
    )

    // cellStyleXfs: Master formatting records (required per ECMA-376 §18.8.9)
    // At minimum, need one default entry that cellXfs can reference via xfId
    val cellStyleXfsElem = elem("cellStyleXfs", "count" -> "1")(
      elem(
        "xf",
        "numFmtId" -> "0",
        "fontId" -> "0",
        "fillId" -> "0",
        "borderId" -> "0"
      )()
    )

    // cellStyles: Named styles (required per ECMA-376 §18.8.8)
    // At minimum, need the default "Normal" style
    val cellStylesElem = elem("cellStyles", "count" -> "1")(
      elem(
        "cellStyle",
        "name" -> "Normal",
        "xfId" -> "0",
        "builtinId" -> "0"
      )()
    )

    // Assemble styles.xml with preserved namespaces and differential formats
    // OOXML order: numFmts, fonts, fills, borders, cellStyleXfs, cellXfs, cellStyles, dxfs
    val children = numFmtsElem.toList ++ Seq(
      fontsElem,
      fillsElem,
      bordersElem,
      cellStyleXfsElem,
      cellXfsElem,
      cellStylesElem
    ) ++ preservedDxfs.toList

    // Use preserved attributes if available, otherwise create minimal xmlns
    rootAttributes match
      case Some(attrs) =>
        // Use preserved attributes and scope from original
        Elem(null, "styleSheet", attrs, rootScope, minimizeEmpty = false, children*)
      case None =>
        // No preserved metadata - create minimal element
        elem("styleSheet", "xmlns" -> nsSpreadsheetML)(children*)

  def writeSax(writer: SaxWriter): Unit =
    writer.startDocument()
    writer.startElement("styleSheet")

    val attrs = rootAttributes.getOrElse(Null)
    val scope = Option(rootScope).getOrElse(defaultStylesScope)
    SaxWriter.withAttributes(
      writer,
      writer.namespaceAttributes(scope) ++ writer.metaDataAttributes(attrs)*
    ) {
      // numFmts
      if index.numFmts.nonEmpty then
        writer.startElement("numFmts")
        writer.writeAttribute("count", index.numFmts.size.toString)
        index.numFmts.sortBy(_._1).foreach { case (id, fmt) =>
          writer.startElement("numFmt")
          writer.writeAttribute("numFmtId", id.toString)
          writer.writeAttribute(
            "formatCode",
            fmt match
              case NumFmt.Custom(c) => c
              case _ => "General"
          )
          writer.endElement()
        }
        writer.endElement() // numFmts

      // Fonts
      writer.startElement("fonts")
      writer.writeAttribute("count", index.fonts.size.toString)
      index.fonts.foreach(writeFontSax(writer, _))
      writer.endElement()

      // Fills (include defaults)
      val defaultFills = Vector(
        Fill.None,
        Fill.Pattern(
          foreground = Color.Rgb(0xff000000),
          background = Color.Rgb(0xffc0c0c0),
          pattern = PatternType.Gray125
        )
      )
      val allFills = {
        val builder = Vector.newBuilder[Fill]
        val seen = scala.collection.mutable.Set.empty[Fill]
        for fill <- defaultFills ++ index.fills do
          if !seen.contains(fill) then
            seen += fill
            builder += fill
        builder.result()
      }

      writer.startElement("fills")
      writer.writeAttribute("count", allFills.size.toString)
      allFills.foreach(writeFillSax(writer, _))
      writer.endElement()

      // Borders
      writer.startElement("borders")
      writer.writeAttribute("count", index.borders.size.toString)
      index.borders.foreach(writeBorderSax(writer, _))
      writer.endElement()

      // cellStyleXfs: Master formatting records (required per ECMA-376 §18.8.9)
      writer.startElement("cellStyleXfs")
      writer.writeAttribute("count", "1")
      writer.startElement("xf")
      writer.writeAttribute("numFmtId", "0")
      writer.writeAttribute("fontId", "0")
      writer.writeAttribute("fillId", "0")
      writer.writeAttribute("borderId", "0")
      writer.endElement() // xf
      writer.endElement() // cellStyleXfs

      // CellXfs
      val fontMap = index.fonts.zipWithIndex.toMap
      val fillMap = allFills.zipWithIndex.toMap
      val borderMap = index.borders.zipWithIndex.toMap

      writer.startElement("cellXfs")
      writer.writeAttribute("count", index.cellStyles.size.toString)
      index.cellStyles.foreach { style =>
        // Use 0 (default style) as fallback - OOXML requires non-negative indices
        val fontIdx = fontMap.getOrElse(style.font, 0)
        val fillIdx = fillMap.getOrElse(style.fill, 0)
        val borderIdx = borderMap.getOrElse(style.border, 0)
        val numFmtId = style.numFmtId.getOrElse {
          NumFmt
            .builtInId(style.numFmt)
            .getOrElse(index.numFmts.find(_._2 == style.numFmt).map(_._1).getOrElse(0))
        }

        writer.startElement("xf")
        val alignmentChild = alignmentToSax(writer, style.align)
        val hasAlignment = alignmentChild.isDefined
        writer.writeAttribute("applyAlignment", if hasAlignment then "1" else "0")
        writer.writeAttribute("borderId", borderIdx.toString)
        writer.writeAttribute("fillId", fillIdx.toString)
        writer.writeAttribute("fontId", fontIdx.toString)
        writer.writeAttribute("numFmtId", numFmtId.toString)
        writer.writeAttribute("xfId", "0")
        alignmentChild.foreach(_.apply())
        writer.endElement()
      }
      writer.endElement() // cellXfs

      // cellStyles: Named styles (required per ECMA-376 §18.8.8)
      writer.startElement("cellStyles")
      writer.writeAttribute("count", "1")
      writer.startElement("cellStyle")
      writer.writeAttribute("name", "Normal")
      writer.writeAttribute("xfId", "0")
      writer.writeAttribute("builtinId", "0")
      writer.endElement() // cellStyle
      writer.endElement() // cellStyles

      // dxfs if preserved
      preservedDxfs.foreach(writer.writeElem)
    }

    writer.endElement() // styleSheet
    writer.endDocument()
    writer.flush()

  private def fontToXml(font: Font): Elem =
    val children = Vector(
      Some(elem("name", "val" -> font.name)()),
      Some(elem("sz", "val" -> font.sizePt.toString)()),
      if font.bold then Some(elem("b")()) else None,
      if font.italic then Some(elem("i")()) else None,
      if font.underline then Some(elem("u")()) else None,
      font.color.map(colorToXml)
    ).flatten

    elem("font")(children*)

  private def fillToXml(fill: Fill): Elem =
    fill match
      case Fill.None =>
        elem("fill")(elem("patternFill", "patternType" -> "none")())

      case Fill.Solid(color) =>
        elem("fill")(
          elem("patternFill", "patternType" -> "solid")(
            colorToXml(color).copy(label = "fgColor") // Use colorToXml to preserve theme colors
          )
        )

      case Fill.Pattern(fg, bg, patternType) =>
        elem("fill")(
          elem("patternFill", "patternType" -> patternType.toString.toLowerCase)(
            colorToXml(fg).copy(label = "fgColor"), // Use colorToXml to preserve theme colors
            colorToXml(bg).copy(label = "bgColor") // Use colorToXml to preserve theme colors
          )
        )

  private def borderToXml(border: Border): Elem =
    elem("border")(
      borderSideToXml("left", border.left),
      borderSideToXml("right", border.right),
      borderSideToXml("top", border.top),
      borderSideToXml("bottom", border.bottom)
    )

  private def borderSideToXml(side: String, borderSide: BorderSide): Elem =
    if borderSide.style == BorderStyle.None then elem(side)()
    else
      val children = borderSide.color.map { color =>
        colorToXml(color) // Use colorToXml to preserve theme colors in borders
      }.toList
      elem(side, "style" -> borderSide.style.toString.toLowerCase)(children*)

  private def colorToXml(color: Color): Elem =
    color match
      case Color.Rgb(argb) =>
        elem("color", "rgb" -> f"$argb%08X")()
      case Color.Theme(slot, tint) =>
        val slotIdx = slot.ordinal
        elem("color", "theme" -> slotIdx.toString, "tint" -> tint.toString)()

  /**
   * Serialize alignment to XML element if non-default (ECMA-376 Part 1, §18.8.1)
   *
   * Only emits <alignment> if at least one property differs from default. Omits default properties
   * for minimal XML output.
   *
   * REQUIRES: align is valid Align instance ENSURES:
   *   - Returns None if align == Align.default (optimization)
   *   - Returns Some(<alignment .../>) if any property differs from default
   *   - Only emits non-default attributes (horizontal, vertical, wrapText, indent)
   *   - Attribute values match OOXML spec enum names (lowercase)
   *   - VAlign.Middle serializes as "center" per OOXML spec
   *   - wrapText serializes as "1" (true) or "0" (false)
   * DETERMINISTIC: Yes (pure transformation based on Align equality) ERROR CASES: None (total
   * function)
   *
   * @param align
   *   Alignment settings to serialize
   * @return
   *   Some(<alignment .../>) if non-default, None if all properties match Align.default
   */
  private def alignmentToXml(align: Align): Option[Elem] =
    // Don't emit alignment if it matches default exactly
    if align == Align.default then None
    else
      val attrs = Seq.newBuilder[(String, String)]

      // Only include non-default properties
      if align.horizontal != Align.default.horizontal then
        val hAlignStr = align.horizontal match
          case HAlign.CenterContinuous => "centerContinuous" // OOXML requires camelCase
          case other => other.toString.toLowerCase(java.util.Locale.ROOT)
        attrs += ("horizontal" -> hAlignStr)

      if align.vertical != Align.default.vertical then
        // VAlign.Middle serializes as "center" per OOXML spec
        val vAlignStr = align.vertical match
          case VAlign.Middle => "center"
          case other => other.toString.toLowerCase(java.util.Locale.ROOT)
        attrs += ("vertical" -> vAlignStr)

      if align.wrapText != Align.default.wrapText then
        attrs += ("wrapText" -> (if align.wrapText then "1" else "0"))

      if align.indent != Align.default.indent then attrs += ("indent" -> align.indent.toString)

      val attrSeq = attrs.result()
      // If no attributes, don't emit element (though this shouldn't happen since align != default)
      if attrSeq.isEmpty then None
      else Some(elem("alignment", attrSeq*)())

  private def alignmentToSax(writer: SaxWriter, align: Align): Option[() => Unit] =
    if align == Align.default then None
    else
      Some { () =>
        val attrs = Seq.newBuilder[(String, String)]

        if align.horizontal != Align.default.horizontal then
          val hAlignStr = align.horizontal match
            case HAlign.CenterContinuous => "centerContinuous"
            case other => other.toString.toLowerCase(java.util.Locale.ROOT)
          attrs += ("horizontal" -> hAlignStr)

        if align.vertical != Align.default.vertical then
          val vAlignStr = align.vertical match
            case VAlign.Middle => "center"
            case other => other.toString.toLowerCase(java.util.Locale.ROOT)
          attrs += ("vertical" -> vAlignStr)

        if align.wrapText != Align.default.wrapText then
          attrs += ("wrapText" -> (if align.wrapText then "1" else "0"))

        if align.indent != Align.default.indent then attrs += ("indent" -> align.indent.toString)

        SaxWriter.withAttributes(writer, attrs.result()*) {
          ()
        }
      }

  private def writeFontSax(writer: SaxWriter, font: Font): Unit =
    writer.startElement("font")
    writer.startElement("name")
    writer.writeAttribute("val", font.name)
    writer.endElement()

    writer.startElement("sz")
    writer.writeAttribute("val", font.sizePt.toString)
    writer.endElement()

    if font.bold then
      writer.startElement("b"); writer.endElement()
    if font.italic then
      writer.startElement("i"); writer.endElement()
    if font.underline then
      writer.startElement("u"); writer.endElement()

    font.color.foreach { c =>
      writer.startElement("color")
      writeColorAttributes(writer, c)
      writer.endElement()
    }

    writer.endElement() // font

  private def writeFillSax(writer: SaxWriter, fill: Fill): Unit =
    writer.startElement("fill")
    fill match
      case Fill.None =>
        writer.startElement("patternFill")
        writer.writeAttribute("patternType", "none")
        writer.endElement()
      case Fill.Solid(color) =>
        writer.startElement("patternFill")
        writer.writeAttribute("patternType", "solid")
        writer.startElement("fgColor")
        writeColorAttributes(writer, color)
        writer.endElement()
        writer.endElement()
      case Fill.Pattern(fg, bg, patternType) =>
        writer.startElement("patternFill")
        writer.writeAttribute("patternType", patternType.toString.toLowerCase)
        writer.startElement("fgColor")
        writeColorAttributes(writer, fg)
        writer.endElement()
        writer.startElement("bgColor")
        writeColorAttributes(writer, bg)
        writer.endElement()
        writer.endElement()
    writer.endElement()

  private def writeBorderSax(writer: SaxWriter, border: Border): Unit =
    writer.startElement("border")
    writeBorderSideSax(writer, "left", border.left)
    writeBorderSideSax(writer, "right", border.right)
    writeBorderSideSax(writer, "top", border.top)
    writeBorderSideSax(writer, "bottom", border.bottom)
    writer.endElement()

  private def writeBorderSideSax(writer: SaxWriter, side: String, borderSide: BorderSide): Unit =
    if borderSide.style == BorderStyle.None then
      writer.startElement(side); writer.endElement()
    else
      writer.startElement(side)
      writer.writeAttribute("style", borderSide.style.toString.toLowerCase)
      borderSide.color.foreach { color =>
        writer.startElement("color")
        writeColorAttributes(writer, color)
        writer.endElement()
      }
      writer.endElement()

  private def writeColorAttributes(writer: SaxWriter, color: Color): Unit =
    color match
      case Color.Rgb(argb) =>
        writer.writeAttribute("rgb", f"$argb%08X")
      case Color.Theme(slot, tint) =>
        writer.writeAttribute("theme", slot.ordinal.toString)
        writer.writeAttribute("tint", tint.toString)

object OoxmlStyles:
  /** Create minimal styles (default only) */
  def minimal: OoxmlStyles =
    val defaultIndex = StyleIndex(
      fonts = Vector(Font.default),
      fills = Vector(Fill.None),
      borders = Vector(Border.none),
      numFmts = Vector.empty,
      cellStyles = Vector(CellStyle.default),
      styleToIndex = Map(CellStyle.default.canonicalKey -> StyleId(0))
    )
    OoxmlStyles(defaultIndex)

  /** Create from workbook (discards remapping tables - only for simple use cases) */
  def fromWorkbook(wb: Workbook): OoxmlStyles =
    val (styleIndex, _) = StyleIndex.fromWorkbook(wb)
    OoxmlStyles(styleIndex)

// ========== Workbook Styles Parsing ==========

/**
 * Parsed workbook-level styles with complete OOXML structure.
 *
 * Stores both domain model (cellStyles) and raw OOXML vectors (fonts, fills, borders) for
 * byte-perfect preservation during surgical writes.
 */
final case class WorkbookStyles(
  cellStyles: Vector[CellStyle],
  fonts: Vector[Font],
  fills: Vector[Fill],
  borders: Vector[Border],
  customNumFmts: Vector[(Int, NumFmt)]
):
  def styleAt(index: Int): Option[CellStyle] = cellStyles.lift(index)

object WorkbookStyles:
  val default: WorkbookStyles = WorkbookStyles(
    cellStyles = Vector(CellStyle.default),
    fonts = Vector(Font.default),
    fills = Vector(Fill.default),
    borders = Vector(Border.none),
    customNumFmts = Vector.empty
  )

  def fromXml(elem: Elem): Either[String, WorkbookStyles] =
    val numFmts = parseNumFmts(elem)
    val fonts = parseFonts(elem)
    val fills = parseFills(elem)
    val borders = parseBorders(elem)
    val cellStyles = parseCellXfs(elem, fonts, fills, borders, numFmts)
    Right(
      WorkbookStyles(
        cellStyles = cellStyles,
        fonts = fonts,
        fills = fills,
        borders = borders,
        customNumFmts = numFmts.toVector.sortBy(_._1)
      )
    )

  private def parseNumFmts(root: Elem): Map[Int, NumFmt] =
    (root \ "numFmts").headOption match
      case Some(numFmtsElem: Elem) =>
        getChildren(numFmtsElem, "numFmt").flatMap { fmtElem =>
          val idOpt = fmtElem.attribute("numFmtId").flatMap(attr => attr.text.toIntOption)
          val codeOpt = fmtElem.attribute("formatCode").map(_.text)
          (idOpt, codeOpt) match
            case (Some(id), Some(code)) => Some(id -> NumFmt.parse(code))
            case _ => None
        }.toMap
      case Some(_) => Map.empty
      case None => Map.empty

  private def parseFonts(root: Elem): Vector[Font] =
    (root \ "fonts").headOption match
      case Some(fontsElem: Elem) =>
        val parsed = getChildren(fontsElem, "font").map(parseFont).toVector
        if parsed.nonEmpty then parsed else Vector(Font.default)
      case Some(_) => Vector(Font.default)
      case None => Vector(Font.default)

  private def parseFont(fontElem: Elem): Font =
    val name = (fontElem \ "name").headOption
      .flatMap(_.attribute("val"))
      .map(_.text)
      .filter(_.nonEmpty)
      .getOrElse(Font.default.name)
    val size = (fontElem \ "sz").headOption
      .flatMap(_.attribute("val"))
      .flatMap(v => v.text.toDoubleOption)
      .getOrElse(Font.default.sizePt)
    val bold = (fontElem \ "b").nonEmpty
    val italic = (fontElem \ "i").nonEmpty
    val underline = (fontElem \ "u").nonEmpty
    val color = (fontElem \ "color").headOption.collect { case e: Elem => e }.flatMap(parseColor)
    Font(name, size, bold, italic, underline, color)

  private def parseFills(root: Elem): Vector[Fill] =
    (root \ "fills").headOption match
      case Some(fillsElem: Elem) =>
        val parsed = getChildren(fillsElem, "fill").map(parseFill).toVector
        if parsed.nonEmpty then parsed else Vector(Fill.None)
      case Some(_) => Vector(Fill.None)
      case None => Vector(Fill.None)

  private def parseFill(fillElem: Elem): Fill =
    (fillElem \ "patternFill").headOption match
      case Some(patternElem: Elem) =>
        val patternType = patternElem
          .attribute("patternType")
          .flatMap(attr => parsePatternType(attr.text))
          .getOrElse(PatternType.None)
        patternType match
          case PatternType.None => Fill.None
          case PatternType.Solid =>
            val fg =
              (patternElem \ "fgColor").headOption.collect { case e: Elem => e }.flatMap(parseColor)
            fg.map(Fill.Solid.apply).getOrElse(Fill.None)
          case other =>
            val fg =
              (patternElem \ "fgColor").headOption.collect { case e: Elem => e }.flatMap(parseColor)
            val bg =
              (patternElem \ "bgColor").headOption.collect { case e: Elem => e }.flatMap(parseColor)
            (fg, bg) match
              case (Some(fgColor), Some(bgColor)) => Fill.Pattern(fgColor, bgColor, other)
              case _ => Fill.None
      case Some(_) => Fill.None
      case None => Fill.None

  private def parsePatternType(value: String): Option[PatternType] =
    val normalized = value.toLowerCase
    PatternType.values.find(_.toString.toLowerCase == normalized)

  private def parseBorders(root: Elem): Vector[Border] =
    (root \ "borders").headOption match
      case Some(bordersElem: Elem) =>
        val parsed = getChildren(bordersElem, "border").map(parseBorder).toVector
        if parsed.nonEmpty then parsed else Vector(Border.none)
      case Some(_) => Vector(Border.none)
      case None => Vector(Border.none)

  private def parseBorder(borderElem: Elem): Border =
    Border(
      left = parseBorderSide(borderElem, "left"),
      right = parseBorderSide(borderElem, "right"),
      top = parseBorderSide(borderElem, "top"),
      bottom = parseBorderSide(borderElem, "bottom")
    )

  private def parseBorderSide(borderElem: Elem, side: String): BorderSide =
    (borderElem \ side).headOption match
      case Some(sideElem: Elem) =>
        val style = sideElem
          .attribute("style")
          .flatMap(attr => parseBorderStyle(attr.text))
          .getOrElse(BorderStyle.None)
        val color =
          (sideElem \ "color").headOption.collect { case e: Elem => e }.flatMap(parseColor)
        BorderSide(style, color)
      case Some(_) => BorderSide.none
      case None => BorderSide.none

  private def parseBorderStyle(value: String): Option[BorderStyle] =
    value.toLowerCase match
      case "none" => Some(BorderStyle.None)
      case "thin" => Some(BorderStyle.Thin)
      case "medium" => Some(BorderStyle.Medium)
      case "thick" => Some(BorderStyle.Thick)
      case "dashed" => Some(BorderStyle.Dashed)
      case "dotted" => Some(BorderStyle.Dotted)
      case "double" => Some(BorderStyle.Double)
      case "hair" => Some(BorderStyle.Hair)
      case "dashdot" => Some(BorderStyle.DashDot)
      case "dashdotdot" => Some(BorderStyle.DashDotDot)
      case "slantdashdot" => Some(BorderStyle.SlantDashDot)
      case "mediumdashed" => Some(BorderStyle.MediumDashed)
      case "mediumdashdot" => Some(BorderStyle.MediumDashDot)
      case "mediumdashdotdot" => Some(BorderStyle.MediumDashDotDot)
      case _ => None

  private def parseCellXfs(
    root: Elem,
    fonts: Vector[Font],
    fills: Vector[Fill],
    borders: Vector[Border],
    numFmts: Map[Int, NumFmt]
  ): Vector[CellStyle] =
    (root \ "cellXfs").headOption match
      case Some(cellXfsElem: Elem) =>
        val parsed = getChildren(cellXfsElem, "xf").map { xfElem =>
          parseCellStyle(xfElem, fonts, fills, borders, numFmts)
        }.toVector
        if parsed.nonEmpty then parsed else Vector(CellStyle.default)
      case Some(_) => Vector(CellStyle.default)
      case None => Vector(CellStyle.default)

  private def parseCellStyle(
    xfElem: Elem,
    fonts: Vector[Font],
    fills: Vector[Fill],
    borders: Vector[Border],
    numFmts: Map[Int, NumFmt]
  ): CellStyle =
    val font = xfElem
      .attribute("fontId")
      .flatMap(attr => attr.text.toIntOption)
      .flatMap(fonts.lift)
      .getOrElse(Font.default)
    val fill = xfElem
      .attribute("fillId")
      .flatMap(attr => attr.text.toIntOption)
      .flatMap(fills.lift)
      .getOrElse(Fill.default)
    val border = xfElem
      .attribute("borderId")
      .flatMap(attr => attr.text.toIntOption)
      .flatMap(borders.lift)
      .getOrElse(Border.none)
    val numFmtIdOpt = xfElem.attribute("numFmtId").flatMap(attr => attr.text.toIntOption)
    val numFmt =
      numFmtIdOpt.flatMap(id => numFmts.get(id).orElse(NumFmt.fromId(id))).getOrElse(NumFmt.General)
    val align = parseAlignment(xfElem).getOrElse(Align.default)
    CellStyle(
      font = font,
      fill = fill,
      border = border,
      numFmt = numFmt,
      numFmtId = numFmtIdOpt,
      align = align
    )

  private def parseAlignment(xfElem: Elem): Option[Align] =
    (xfElem \ "alignment").headOption match
      case Some(alignElem: Elem) =>
        val horizontal = alignElem.attribute("horizontal").flatMap(attr => parseHAlign(attr.text))
        val vertical = alignElem.attribute("vertical").flatMap(attr => parseVAlign(attr.text))
        val wrap = alignElem
          .attribute("wrapText")
          .flatMap(attr => parseBool(attr.text))
          .getOrElse(Align.default.wrapText)
        val indent = alignElem
          .attribute("indent")
          .flatMap(attr => attr.text.toIntOption)
          .getOrElse(Align.default.indent)
        Some(
          Align(
            horizontal = horizontal.getOrElse(Align.default.horizontal),
            vertical = vertical.getOrElse(Align.default.vertical),
            wrapText = wrap,
            indent = indent
          )
        )
      case Some(_) => None
      case None => None

  private def parseHAlign(value: String): Option[HAlign] =
    value.toLowerCase match
      case "general" => Some(HAlign.General)
      case "left" => Some(HAlign.Left)
      case "center" => Some(HAlign.Center)
      case "right" => Some(HAlign.Right)
      case "justify" => Some(HAlign.Justify)
      case "fill" => Some(HAlign.Fill)
      case "centercontinuous" => Some(HAlign.CenterContinuous)
      case "distributed" => Some(HAlign.Distributed)
      case _ => None

  private def parseVAlign(value: String): Option[VAlign] =
    value.toLowerCase match
      case "top" => Some(VAlign.Top)
      case "center" => Some(VAlign.Middle)
      case "bottom" => Some(VAlign.Bottom)
      case "justify" => Some(VAlign.Justify)
      case "distributed" => Some(VAlign.Distributed)
      case _ => None

  private def parseBool(value: String): Option[Boolean] =
    value match
      case "1" | "true" | "TRUE" => Some(true)
      case "0" | "false" | "FALSE" => Some(false)
      case _ => None

  private def parseColor(colorElem: Elem): Option[Color] =
    colorElem
      .attribute("rgb")
      .flatMap(attr => parseRgb(attr.text))
      .orElse {
        for
          themeAttr <- colorElem.attribute("theme")
          idx <- themeAttr.text.toIntOption
          slot <- themeSlotFromIndex(idx)
        yield
          val tint =
            colorElem.attribute("tint").flatMap(attr => attr.text.toDoubleOption).getOrElse(0.0)
          Color.Theme(slot, tint)
      }

  private def parseRgb(value: String): Option[Color] =
    val cleaned = value.trim.stripPrefix("#")
    val normalized =
      if cleaned.length == 6 then Some("FF" + cleaned)
      else if cleaned.length == 8 then Some(cleaned)
      else None
    normalized.flatMap { hex =>
      try Some(Color.Rgb(java.lang.Long.parseUnsignedLong(hex, 16).toInt))
      catch case _: NumberFormatException => None
    }

  // OOXML theme color indices (ECMA-376 Part 1, 18.8.3):
  // 0=lt1, 1=dk1, 2=lt2, 3=dk2, 4-9=accent1-6, 10=hlink, 11=folHlink
  private def themeSlotFromIndex(idx: Int): Option[ThemeSlot] = idx match
    case 0 => Some(ThemeSlot.Light1)
    case 1 => Some(ThemeSlot.Dark1)
    case 2 => Some(ThemeSlot.Light2)
    case 3 => Some(ThemeSlot.Dark2)
    case 4 => Some(ThemeSlot.Accent1)
    case 5 => Some(ThemeSlot.Accent2)
    case 6 => Some(ThemeSlot.Accent3)
    case 7 => Some(ThemeSlot.Accent4)
    case 8 => Some(ThemeSlot.Accent5)
    case 9 => Some(ThemeSlot.Accent6)
    case _ => None
