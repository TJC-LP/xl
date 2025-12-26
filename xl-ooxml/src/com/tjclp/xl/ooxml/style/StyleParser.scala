package com.tjclp.xl.ooxml.style

import scala.xml.*

import com.tjclp.xl.ooxml.XmlUtil.getChildren
import com.tjclp.xl.styles.CellStyle
import com.tjclp.xl.styles.alignment.{Align, HAlign, VAlign}
import com.tjclp.xl.styles.border.{Border, BorderSide, BorderStyle}
import com.tjclp.xl.styles.color.Color
import com.tjclp.xl.styles.fill.{Fill, PatternType}
import com.tjclp.xl.styles.font.Font
import com.tjclp.xl.styles.numfmt.NumFmt

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
    // Try RGB first
    colorElem
      .attribute("rgb")
      .flatMap(attr => parseRgb(attr.text))
      .orElse {
        // Try theme color
        for
          themeAttr <- colorElem.attribute("theme")
          idx <- themeAttr.text.toIntOption
          slot <- ColorHelpers.themeSlotFromIndex(idx)
        yield
          val tint =
            colorElem.attribute("tint").flatMap(attr => attr.text.toDoubleOption).getOrElse(0.0)
          Color.Theme(slot, tint)
      }
      .orElse {
        // Try indexed color (standard Excel palette)
        colorElem
          .attribute("indexed")
          .flatMap(attr => attr.text.toIntOption)
          .flatMap(idx => ColorHelpers.indexedColorToRgb(idx))
      }

  private def parseRgb(value: String): Option[Color] =
    val cleaned = value.trim.stripPrefix("#")
    val normalized =
      if cleaned.length == 6 then Some("FF" + cleaned)
      else if cleaned.length == 8 then
        // Excel/openpyxl often write 00RRGGBB where 00 means opaque, not transparent
        // This is a quirk of the format - treat 00 alpha as fully opaque (FF)
        val alpha = cleaned.substring(0, 2)
        if alpha == "00" then Some("FF" + cleaned.substring(2))
        else Some(cleaned)
      else None
    normalized.flatMap { hex =>
      try Some(Color.Rgb(java.lang.Long.parseUnsignedLong(hex, 16).toInt))
      catch case _: NumberFormatException => None
    }
