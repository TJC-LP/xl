package com.tjclp.xl.ooxml.style

import scala.xml.*

import com.tjclp.xl.api.Workbook
import com.tjclp.xl.ooxml.XmlUtil.{elem, nsSpreadsheetML}
import com.tjclp.xl.ooxml.{SaxSerializable, SaxWriter, XmlWritable}
import com.tjclp.xl.ooxml.SaxSupport.*
import com.tjclp.xl.styles.CellStyle
import com.tjclp.xl.styles.alignment.{Align, HAlign, VAlign}
import com.tjclp.xl.styles.border.{Border, BorderSide, BorderStyle}
import com.tjclp.xl.styles.color.Color
import com.tjclp.xl.styles.fill.{Fill, PatternType}
import com.tjclp.xl.styles.font.Font
import com.tjclp.xl.styles.numfmt.NumFmt
import com.tjclp.xl.styles.units.StyleId

private[ooxml] val defaultStylesScope: NamespaceBinding =
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
     * Default fills required by OOXML spec (ECMA-376 Part 1, section 18.8.21)
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

      for fill <- defaultFills ++ index.fills do
        if !seen.contains(fill) then
          seen += fill
          builder += fill
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
          // No raw ID -> derive from NumFmt enum (programmatic creation)
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

    // cellStyleXfs: Master formatting records (required per ECMA-376 section 18.8.9)
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

    // cellStyles: Named styles (required per ECMA-376 section 18.8.8)
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
    SaxWriter.withAttributes(writer, writer.combinedAttributes(scope, attrs)*) {
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

      // cellStyleXfs: Master formatting records (required per ECMA-376 section 18.8.9)
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

      // cellStyles: Named styles (required per ECMA-376 section 18.8.8)
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
   * Serialize alignment to XML element if non-default (ECMA-376 Part 1, section 18.8.1)
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

        val attrSeq = attrs.result()
        if attrSeq.nonEmpty then
          writer.startElement("alignment")
          SaxWriter.withAttributes(writer, attrSeq*) {
            ()
          }
          writer.endElement()
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
      writer.startElement("b")
      writer.endElement()
    if font.italic then
      writer.startElement("i")
      writer.endElement()
    if font.underline then
      writer.startElement("u")
      writer.endElement()

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
      writer.startElement(side)
      writer.endElement()
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
