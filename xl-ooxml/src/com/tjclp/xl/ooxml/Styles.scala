package com.tjclp.xl.ooxml

import scala.xml.*
import XmlUtil.*
import com.tjclp.xl.api.*
import com.tjclp.xl.style.{CellStyle, StyleRegistry}
import com.tjclp.xl.style.alignment.{Align, HAlign, VAlign}
import com.tjclp.xl.style.border.{Border, BorderSide, BorderStyle}
import com.tjclp.xl.style.color.{Color, ThemeSlot}
import com.tjclp.xl.style.fill.{Fill, PatternType}
import com.tjclp.xl.style.font.Font
import com.tjclp.xl.style.numfmt.NumFmt
import com.tjclp.xl.style.units.StyleId

/**
 * Style components and indexing for xl/styles.xml
 *
 * Styles are deduplicated by canonical keys to avoid Excel's 64k style limit. The StyleIndex builds
 * collections of unique fonts, fills, borders, and cellXfs.
 */

/** Index mapping for style components */
case class StyleIndex(
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
   * Build unified style index from a workbook and per-sheet remapping tables.
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
  def fromWorkbook(wb: Workbook): (StyleIndex, Map[Int, Map[Int, Int]]) =
    import scala.collection.mutable

    // Build unified style index by merging all sheet registries
    var unifiedStyles = Vector(CellStyle.default)
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
            unifiedStyles = unifiedStyles :+ style
            unifiedIndex = unifiedIndex + (key -> StyleId(nextIdx))
            remapping(localIdx) = nextIdx
            nextIdx += 1
      }

      sheetIdx -> remapping.toMap
    }.toMap

    // Deduplicate components using LinkedHashSet for O(1) deduplication (60-80% faster than .distinct)
    import scala.collection.mutable
    val uniqueFonts = {
      val seen = mutable.LinkedHashSet.empty[Font]
      unifiedStyles.foreach(style => seen += style.font)
      seen.toVector
    }
    val uniqueFills = {
      val seen = mutable.LinkedHashSet.empty[Fill]
      unifiedStyles.foreach(style => seen += style.fill)
      seen.toVector
    }
    val uniqueBorders = {
      val seen = mutable.LinkedHashSet.empty[Border]
      unifiedStyles.foreach(style => seen += style.border)
      seen.toVector
    }

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

/** Serializer for xl/styles.xml */
case class OoxmlStyles(
  index: StyleIndex
) extends XmlWritable:

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
        val fontIdx = fontMap.getOrElse(style.font, -1)
        val fillIdx = fillMap.getOrElse(style.fill, -1)
        val borderIdx = borderMap.getOrElse(style.border, -1)
        val numFmtId = NumFmt
          .builtInId(style.numFmt)
          .getOrElse(
            index.numFmts.find(_._2 == style.numFmt).map(_._1).getOrElse(0)
          )

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

    // Assemble styles.xml
    val children = numFmtsElem.toList ++ Seq(fontsElem, fillsElem, bordersElem, cellXfsElem)
    elem("styleSheet", "xmlns" -> nsSpreadsheetML)(children*)

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
            elem("fgColor", "rgb" -> f"${color.toArgb}%08X")()
          )
        )

      case Fill.Pattern(fg, bg, patternType) =>
        elem("fill")(
          elem("patternFill", "patternType" -> patternType.toString.toLowerCase)(
            elem("fgColor", "rgb" -> f"${fg.toArgb}%08X")(),
            elem("bgColor", "rgb" -> f"${bg.toArgb}%08X")()
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
        elem("color", "rgb" -> f"${color.toArgb}%08X")()
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
        val hAlignStr = align.horizontal.toString.toLowerCase(java.util.Locale.ROOT)
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

/** Parsed workbook-level styles mapped by cellXf index */
case class WorkbookStyles(cellStyles: Vector[CellStyle]):
  def styleAt(index: Int): Option[CellStyle] = cellStyles.lift(index)

object WorkbookStyles:
  val default: WorkbookStyles = WorkbookStyles(Vector(CellStyle.default))

  def fromXml(elem: Elem): Either[String, WorkbookStyles] =
    val numFmts = parseNumFmts(elem)
    val fonts = parseFonts(elem)
    val fills = parseFills(elem)
    val borders = parseBorders(elem)
    val cellStyles = parseCellXfs(elem, fonts, fills, borders, numFmts)
    Right(WorkbookStyles(cellStyles))

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
    val numFmtId = xfElem.attribute("numFmtId").flatMap(attr => attr.text.toIntOption)
    val numFmt =
      numFmtId.flatMap(id => numFmts.get(id).orElse(NumFmt.fromId(id))).getOrElse(NumFmt.General)
    val align = parseAlignment(xfElem).getOrElse(Align.default)
    CellStyle(font = font, fill = fill, border = border, numFmt = numFmt, align = align)

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

  private def themeSlotFromIndex(idx: Int): Option[ThemeSlot] = idx match
    case 0 => Some(ThemeSlot.Dark1)
    case 1 => Some(ThemeSlot.Light1)
    case 2 => Some(ThemeSlot.Dark2)
    case 3 => Some(ThemeSlot.Light2)
    case 4 => Some(ThemeSlot.Accent1)
    case 5 => Some(ThemeSlot.Accent2)
    case 6 => Some(ThemeSlot.Accent3)
    case 7 => Some(ThemeSlot.Accent4)
    case 8 => Some(ThemeSlot.Accent5)
    case 9 => Some(ThemeSlot.Accent6)
    case _ => None
