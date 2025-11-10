package com.tjclp.xl.ooxml

import scala.xml.*
import XmlUtil.*
import com.tjclp.xl.*

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
  styleToIndex: Map[String, Int] // Canonical key → cellXf index
):
  /** Get style index for a CellStyle (returns 0 if not found - default style) */
  def indexOf(style: CellStyle): Int =
    styleToIndex.getOrElse(CellStyle.canonicalKey(style), 0)

object StyleIndex:
  /** Build style index from a workbook, collecting all unique styles */
  def fromWorkbook(wb: Workbook): StyleIndex =
    // Collect all styles from cells
    val allStyles = wb.sheets.flatMap { sheet =>
      sheet.cells.values.flatMap { cell =>
        // For now, cells don't have CellStyle, only styleId
        // This will be enhanced when we add full style support
        None // TODO: extract CellStyle from cells
      }
    }

    // Start with default style
    val defaultStyle = CellStyle.default
    val styles = (defaultStyle +: allStyles.toVector).distinct

    // Deduplicate components
    val uniqueFonts = styles.map(_.font).distinct
    val uniqueFills = styles.map(_.fill).distinct
    val uniqueBorders = styles.map(_.border).distinct

    // Collect custom number formats (built-ins don't need entries)
    val customNumFmts = styles
      .map(_.numFmt)
      .collect { case NumFmt.Custom(code) => code }
      .distinct
      .zipWithIndex
      .map { case (code, idx) => (164 + idx, NumFmt.Custom(code)) } // Start at 164 for custom

    // Build style index map
    val styleIndex = styles.zipWithIndex.map { case (style, idx) =>
      CellStyle.canonicalKey(style) -> idx
    }.toMap

    StyleIndex(uniqueFonts, uniqueFills, uniqueBorders, customNumFmts.toVector, styles, styleIndex)

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

    // Fills (Excel requires 2 default fills at indices 0-1)
    val defaultFills = Vector(Fill.None, Fill.Solid(Color.Rgb(0x00000000)))
    val allFills = (defaultFills ++ index.fills.filterNot(defaultFills.contains)).distinct
    val fillsElem = elem("fills", "count" -> allFills.size.toString)(
      allFills.map(fillToXml)*
    )

    // Borders
    val bordersElem = elem("borders", "count" -> index.borders.size.toString)(
      index.borders.map(borderToXml)*
    )

    // CellXfs (cell format styles)
    val cellXfsElem = elem("cellXfs", "count" -> index.cellStyles.size.toString)(
      index.cellStyles.map { style =>
        val fontIdx = index.fonts.indexOf(style.font)
        val fillIdx = allFills.indexOf(style.fill)
        val borderIdx = index.borders.indexOf(style.border)
        val numFmtId = NumFmt
          .builtInId(style.numFmt)
          .getOrElse(
            index.numFmts.find(_._2 == style.numFmt).map(_._1).getOrElse(0)
          )

        elem(
          "xf",
          "borderId" -> borderIdx.toString,
          "fillId" -> fillIdx.toString,
          "fontId" -> fontIdx.toString,
          "numFmtId" -> numFmtId.toString,
          "xfId" -> "0"
        )()
      }*
    )

    // Assemble styles.xml
    val children = numFmtsElem.toSeq ++ Seq(fontsElem, fillsElem, bordersElem, cellXfsElem)
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
      }.toSeq
      elem(side, "style" -> borderSide.style.toString.toLowerCase)(children*)

  private def colorToXml(color: Color): Elem =
    color match
      case Color.Rgb(argb) =>
        elem("color", "rgb" -> f"$argb%08X")()
      case Color.Theme(slot, tint) =>
        val slotIdx = slot.ordinal
        elem("color", "theme" -> slotIdx.toString, "tint" -> tint.toString)()

object OoxmlStyles:
  /** Create minimal styles (default only) */
  def minimal: OoxmlStyles =
    val defaultIndex = StyleIndex(
      fonts = Vector(Font.default),
      fills = Vector(Fill.None),
      borders = Vector(Border.none),
      numFmts = Vector.empty,
      cellStyles = Vector(CellStyle.default),
      styleToIndex = Map(CellStyle.canonicalKey(CellStyle.default) -> 0)
    )
    OoxmlStyles(defaultIndex)

  /** Create from workbook */
  def fromWorkbook(wb: Workbook): OoxmlStyles =
    OoxmlStyles(StyleIndex.fromWorkbook(wb))
