package com.tjclp.xl.io.streaming

import scala.xml.*
import com.tjclp.xl.styles.CellStyle
import com.tjclp.xl.styles.border.{Border, BorderSide, BorderStyle}
import com.tjclp.xl.styles.fill.{Fill, PatternType}
import com.tjclp.xl.styles.font.Font
import com.tjclp.xl.styles.numfmt.NumFmt
import com.tjclp.xl.styles.alignment.{Align, HAlign, VAlign}
import com.tjclp.xl.styles.color.Color
import com.tjclp.xl.ooxml.XmlUtil
import java.util.Locale

/**
 * Patches styles.xml to add new CellStyle entries with O(small) memory.
 *
 * styles.xml is typically <1MB so full parse is acceptable. The real memory savings come from not
 * loading the worksheet data.
 */
@SuppressWarnings(Array("org.wartremover.warts.AsInstanceOf"))
object StylePatcher:

  /**
   * Add a new style to styles.xml, returning updated XML and the new style ID.
   *
   * @param stylesXml
   *   The original styles.xml content
   * @param newStyle
   *   The CellStyle to add
   * @return
   *   (updated styles.xml content, new cellXf index)
   */
  def addStyle(stylesXml: String, newStyle: CellStyle): (String, Int) =
    val parsed = XML.loadString(stylesXml)
    addStyleToElem(parsed, newStyle)

  /**
   * Add multiple styles to styles.xml, returning updated XML and style ID mapping.
   *
   * @param stylesXml
   *   The original styles.xml content
   * @param newStyles
   *   Map of identifier to CellStyle to add
   * @return
   *   (updated styles.xml content, Map[identifier -> new cellXf index])
   */
  def addStyles[K](stylesXml: String, newStyles: Map[K, CellStyle]): (String, Map[K, Int]) =
    var parsed = XML.loadString(stylesXml)
    val mapping = scala.collection.mutable.Map[K, Int]()

    newStyles.foreach { case (key, style) =>
      val (updated, newId) = addStyleToElem(parsed, style)
      parsed = XML.loadString(updated)
      mapping(key) = newId
    }

    (XmlUtil.compact(parsed), mapping.toMap)

  /**
   * Get existing style by index from styles.xml.
   *
   * @param stylesXml
   *   The styles.xml content
   * @param styleId
   *   The cellXf index
   * @return
   *   CellStyle if found
   */
  def getStyle(stylesXml: String, styleId: Int): Option[CellStyle] =
    val parsed = XML.loadString(stylesXml)
    extractStyle(parsed, styleId)

  /**
   * Merge styles: combines existing style with new style properties.
   *
   * @param existing
   *   The existing CellStyle
   * @param overlay
   *   The new style properties to merge
   * @return
   *   Merged CellStyle
   */
  def mergeStyles(existing: CellStyle, overlay: CellStyle): CellStyle =
    // Merge font
    val mergedFont = Font(
      name =
        if overlay.font.name != Font.default.name then overlay.font.name else existing.font.name,
      sizePt =
        if overlay.font.sizePt != Font.default.sizePt then overlay.font.sizePt
        else existing.font.sizePt,
      bold = overlay.font.bold || existing.font.bold,
      italic = overlay.font.italic || existing.font.italic,
      underline = overlay.font.underline || existing.font.underline,
      color = overlay.font.color.orElse(existing.font.color)
    )

    // Merge fill
    val mergedFill = overlay.fill match
      case Fill.None if existing.fill != Fill.None => existing.fill
      case other => other

    // Merge border - prefer overlay if specified, else existing
    val mergedBorder = mergeBorders(existing.border, overlay.border)

    // Merge numFmt - prefer overlay if not General
    val mergedNumFmt = overlay.numFmt match
      case NumFmt.General => existing.numFmt
      case other => other

    // Merge alignment
    val mergedAlign =
      if overlay.align == Align.default then existing.align
      else
        // Merge individual components
        val h =
          if overlay.align.horizontal != HAlign.General then overlay.align.horizontal
          else existing.align.horizontal
        val v =
          if overlay.align.vertical != VAlign.Bottom then overlay.align.vertical
          else existing.align.vertical
        val wrap = overlay.align.wrapText || existing.align.wrapText
        val indent =
          if overlay.align.indent > 0 then overlay.align.indent else existing.align.indent
        Align(h, v, wrap, indent)

    CellStyle(mergedFont, mergedFill, mergedBorder, mergedNumFmt, align = mergedAlign)

  private def mergeBorders(existing: Border, overlay: Border): Border =
    def mergeSide(e: BorderSide, o: BorderSide): BorderSide =
      if o.style != BorderStyle.None then o
      else if e.style != BorderStyle.None then e
      else BorderSide.none

    Border(
      left = mergeSide(existing.left, overlay.left),
      right = mergeSide(existing.right, overlay.right),
      top = mergeSide(existing.top, overlay.top),
      bottom = mergeSide(existing.bottom, overlay.bottom)
    )

  private def addStyleToElem(root: Elem, newStyle: CellStyle): (String, Int) =
    // Extract existing component collections
    val fonts = (root \ "fonts").headOption.map(_.asInstanceOf[Elem]).getOrElse(<fonts count="0"/>)
    val fills = (root \ "fills").headOption.map(_.asInstanceOf[Elem]).getOrElse(<fills count="0"/>)
    val borders =
      (root \ "borders").headOption.map(_.asInstanceOf[Elem]).getOrElse(<borders count="0"/>)
    val numFmts = (root \ "numFmts").headOption.map(_.asInstanceOf[Elem])
    val cellXfs =
      (root \ "cellXfs").headOption.map(_.asInstanceOf[Elem]).getOrElse(<cellXfs count="0"/>)

    // Find or add font
    val (updatedFonts, fontId) = findOrAddFont(fonts, newStyle.font)

    // Find or add fill
    val (updatedFills, fillId) = findOrAddFill(fills, newStyle.fill)

    // Find or add border
    val (updatedBorders, borderId) = findOrAddBorder(borders, newStyle.border)

    // Find or add numFmt
    val (updatedNumFmts, numFmtId) = findOrAddNumFmt(numFmts, newStyle.numFmt)

    // Add new cellXf
    val existingXfs = (cellXfs \ "xf").toSeq
    val newXfId = existingXfs.size
    val newXf = createCellXf(fontId, fillId, borderId, numFmtId, newStyle.align)
    val newXfCount = existingXfs.size + 1
    val updatedCellXfs = Elem(
      null,
      "cellXfs",
      new UnprefixedAttribute("count", newXfCount.toString, Null),
      TopScope,
      minimizeEmpty = false,
      (existingXfs :+ newXf)*
    )

    // Rebuild styles.xml with updated components
    val newChildren = root.child.map {
      case e: Elem if e.label == "fonts" => updatedFonts
      case e: Elem if e.label == "fills" => updatedFills
      case e: Elem if e.label == "borders" => updatedBorders
      case e: Elem if e.label == "cellXfs" => updatedCellXfs
      case other => other
    }

    // Add numFmts if needed
    val finalChildren = updatedNumFmts match
      case Some(nf) =>
        val hasNumFmts = newChildren.exists {
          case e: Elem if e.label == "numFmts" => true
          case _ => false
        }
        if hasNumFmts then
          newChildren.map {
            case e: Elem if e.label == "numFmts" => nf
            case other => other
          }
        else
          // Insert numFmts before fonts (OOXML order)
          val insertIdx = newChildren.indexWhere {
            case e: Elem if e.label == "fonts" => true
            case _ => false
          }
          if insertIdx >= 0 then
            newChildren.take(insertIdx) ++ Seq(nf) ++ newChildren.drop(insertIdx)
          else newChildren :+ nf
      case None => newChildren

    val updatedRoot = Elem(
      root.prefix,
      root.label,
      root.attributes,
      root.scope,
      minimizeEmpty = false,
      finalChildren*
    )
    (XmlUtil.compact(updatedRoot), newXfId)

  private def findOrAddFont(fonts: Elem, font: Font): (Elem, Int) =
    val existingFonts = (fonts \ "font").toSeq
    val fontXml = fontToXml(font)

    // Check if font already exists (by comparing XML structure)
    val existing = existingFonts.zipWithIndex.find { case (f, _) =>
      XmlUtil.compact(f.asInstanceOf[Elem]) == XmlUtil.compact(fontXml)
    }

    existing match
      case Some((_, idx)) => (fonts, idx)
      case None =>
        val newCount = existingFonts.size + 1
        val updated = Elem(
          null,
          "fonts",
          new UnprefixedAttribute("count", newCount.toString, Null),
          TopScope,
          minimizeEmpty = false,
          (existingFonts :+ fontXml)*
        )
        (updated, existingFonts.size)

  private def findOrAddFill(fills: Elem, fill: Fill): (Elem, Int) =
    val existingFills = (fills \ "fill").toSeq
    val fillXml = fillToXml(fill)

    val existing = existingFills.zipWithIndex.find { case (f, _) =>
      XmlUtil.compact(f.asInstanceOf[Elem]) == XmlUtil.compact(fillXml)
    }

    existing match
      case Some((_, idx)) => (fills, idx)
      case None =>
        val newCount = existingFills.size + 1
        val updated = Elem(
          null,
          "fills",
          new UnprefixedAttribute("count", newCount.toString, Null),
          TopScope,
          minimizeEmpty = false,
          (existingFills :+ fillXml)*
        )
        (updated, existingFills.size)

  private def findOrAddBorder(borders: Elem, border: Border): (Elem, Int) =
    val existingBorders = (borders \ "border").toSeq
    val borderXml = borderToXml(border)

    val existing = existingBorders.zipWithIndex.find { case (b, _) =>
      XmlUtil.compact(b.asInstanceOf[Elem]) == XmlUtil.compact(borderXml)
    }

    existing match
      case Some((_, idx)) => (borders, idx)
      case None =>
        val newCount = existingBorders.size + 1
        val updated = Elem(
          null,
          "borders",
          new UnprefixedAttribute("count", newCount.toString, Null),
          TopScope,
          minimizeEmpty = false,
          (existingBorders :+ borderXml)*
        )
        (updated, existingBorders.size)

  private def findOrAddNumFmt(
    numFmts: Option[Elem],
    numFmt: NumFmt
  ): (Option[Elem], Int) =
    numFmt match
      case NumFmt.General => (numFmts, 0)
      case NumFmt.Text => (numFmts, 49)
      case NumFmt.Decimal => (numFmts, 4)
      case NumFmt.Percent => (numFmts, 10)
      case NumFmt.Currency => (numFmts, 44)
      case NumFmt.Date => (numFmts, 14)
      case NumFmt.DateTime => (numFmts, 22)
      case NumFmt.Time => (numFmts, 21)
      case NumFmt.Integer => (numFmts, 1)
      case NumFmt.ThousandsSeparator => (numFmts, 3)
      case NumFmt.ThousandsDecimal => (numFmts, 4)
      case NumFmt.PercentDecimal => (numFmts, 10)
      case NumFmt.Scientific => (numFmts, 11)
      case NumFmt.Fraction => (numFmts, 12)
      case NumFmt.Custom(code) =>
        val existing = numFmts.map(n => (n \ "numFmt").toSeq).getOrElse(Seq.empty)

        // Check if custom format already exists
        val found = existing.find(nf => (nf \ "@formatCode").text == code)
        found match
          case Some(nf) =>
            val id = (nf \ "@numFmtId").text.toIntOption.getOrElse(164)
            (numFmts, id)
          case None =>
            val nextId = existing
              .flatMap(nf => (nf \ "@numFmtId").text.toIntOption)
              .maxOption
              .map(_ + 1)
              .getOrElse(164)

            val newNumFmt =
              XmlUtil.elem("numFmt", "numFmtId" -> nextId.toString, "formatCode" -> code)()

            val updated = numFmts match
              case Some(elem) =>
                Elem(
                  null,
                  "numFmts",
                  new UnprefixedAttribute("count", (existing.size + 1).toString, Null),
                  TopScope,
                  minimizeEmpty = false,
                  (existing :+ newNumFmt)*
                )
              case None =>
                Elem(
                  null,
                  "numFmts",
                  new UnprefixedAttribute("count", "1", Null),
                  TopScope,
                  minimizeEmpty = false,
                  newNumFmt
                )

            (Some(updated), nextId)

  private def createCellXf(
    fontId: Int,
    fillId: Int,
    borderId: Int,
    numFmtId: Int,
    align: Align
  ): Elem =
    val baseAttrs = Seq(
      "fontId" -> fontId.toString,
      "fillId" -> fillId.toString,
      "borderId" -> borderId.toString,
      "numFmtId" -> numFmtId.toString
    )

    // Add applyX attributes if non-default
    val applyAttrs = scala.collection.mutable.ListBuffer[(String, String)]()
    if fontId > 0 then applyAttrs += "applyFont" -> "1"
    if fillId > 0 then applyAttrs += "applyFill" -> "1"
    if borderId > 0 then applyAttrs += "applyBorder" -> "1"
    if numFmtId > 0 then applyAttrs += "applyNumberFormat" -> "1"
    if align != Align.default then applyAttrs += "applyAlignment" -> "1"

    val allAttrs = (baseAttrs ++ applyAttrs.toSeq).foldRight(Null: MetaData) { case ((k, v), acc) =>
      new UnprefixedAttribute(k, v, acc)
    }

    if align != Align.default then
      Elem(null, "xf", allAttrs, TopScope, minimizeEmpty = false, alignmentToXml(align))
    else Elem(null, "xf", allAttrs, TopScope, minimizeEmpty = true)

  private def fontToXml(font: Font): Elem =
    val children = scala.collection.mutable.ListBuffer[Elem]()
    if font.bold then children += <b/>
    if font.italic then children += <i/>
    if font.underline then children += <u/>
    children += <sz val={font.sizePt.toString}/>
    font.color.foreach {
      case Color.Rgb(argb) => children += <color rgb={f"$argb%08X"}/>
      case Color.Theme(slot, tint) =>
        children += <color theme={slot.ordinal.toString} tint={tint.toString}/>
    }
    children += <name val={font.name}/>
    Elem(null, "font", Null, TopScope, minimizeEmpty = false, children.toSeq*)

  private def fillToXml(fill: Fill): Elem =
    fill match
      case Fill.None => <fill><patternFill patternType="none"/></fill>
      case Fill.Solid(color) =>
        val colorElem: Elem = color match
          case Color.Rgb(argb) => <fgColor rgb={f"$argb%08X"}/>
          case Color.Theme(slot, tint) =>
            <fgColor theme={slot.ordinal.toString} tint={tint.toString}/>
        val pfElem = Elem(
          null,
          "patternFill",
          new UnprefixedAttribute("patternType", "solid", Null),
          TopScope,
          minimizeEmpty = false,
          colorElem
        )
        Elem(null, "fill", Null, TopScope, minimizeEmpty = false, pfElem)
      case Fill.Pattern(fgColor, bgColor, patternType) =>
        val pf = scala.collection.mutable.ListBuffer[Elem]()
        fgColor match
          case Color.Rgb(argb) => pf += <fgColor rgb={f"$argb%08X"}/>
          case Color.Theme(slot, tint) =>
            pf += <fgColor theme={slot.ordinal.toString} tint={tint.toString}/>
        bgColor match
          case Color.Rgb(argb) => pf += <bgColor rgb={f"$argb%08X"}/>
          case Color.Theme(slot, tint) =>
            pf += <bgColor theme={slot.ordinal.toString} tint={tint.toString}/>
        val ptStr = patternType.toString.toLowerCase(Locale.ROOT)
        val pfAttrs = new UnprefixedAttribute("patternType", ptStr, Null)
        val pfElem =
          Elem(null, "patternFill", pfAttrs, TopScope, minimizeEmpty = pf.isEmpty, pf.toSeq*)
        <fill>{pfElem}</fill>

  private def borderToXml(border: Border): Elem =
    def sideToXml(label: String, side: BorderSide): Elem =
      if side.style == BorderStyle.None then Elem(null, label, Null, TopScope, minimizeEmpty = true)
      else
        val styleStr = side.style.toString.toLowerCase(Locale.ROOT)
        val styleAttr = new UnprefixedAttribute("style", styleStr, Null)
        val colorElem = side.color.map {
          case Color.Rgb(argb) => <color rgb={f"$argb%08X"}/>
          case Color.Theme(slot, tint) =>
            <color theme={slot.ordinal.toString} tint={tint.toString}/>
        }
        Elem(null, label, styleAttr, TopScope, minimizeEmpty = colorElem.isEmpty, colorElem.toSeq*)

    val sides = Seq(
      sideToXml("left", border.left),
      sideToXml("right", border.right),
      sideToXml("top", border.top),
      sideToXml("bottom", border.bottom),
      sideToXml("diagonal", BorderSide.none) // OOXML requires diagonal element
    )

    <border>{sides}</border>

  private def alignmentToXml(align: Align): Elem =
    val attrs = scala.collection.mutable.ListBuffer[(String, String)]()
    if align.horizontal != HAlign.General then
      val hAlignStr = align.horizontal match
        case HAlign.CenterContinuous => "centerContinuous"
        case other => other.toString.toLowerCase(Locale.ROOT)
      attrs += "horizontal" -> hAlignStr
    if align.vertical != VAlign.Bottom then
      val vAlignStr = align.vertical.toString.toLowerCase(Locale.ROOT)
      attrs += "vertical" -> vAlignStr
    if align.wrapText then attrs += "wrapText" -> "1"
    if align.indent > 0 then attrs += "indent" -> align.indent.toString

    val attrMeta = attrs.foldRight(Null: MetaData) { case ((k, v), acc) =>
      new UnprefixedAttribute(k, v, acc)
    }
    Elem(null, "alignment", attrMeta, TopScope, minimizeEmpty = true)

  private def extractStyle(root: Elem, styleId: Int): Option[CellStyle] =
    val cellXfs = (root \ "cellXfs" \ "xf")
    cellXfs.lift(styleId).map { xf =>
      val fontId = (xf \ "@fontId").text.toIntOption.getOrElse(0)
      val fillId = (xf \ "@fillId").text.toIntOption.getOrElse(0)
      val borderId = (xf \ "@borderId").text.toIntOption.getOrElse(0)
      val numFmtId = (xf \ "@numFmtId").text.toIntOption.getOrElse(0)

      val font = extractFont(root, fontId)
      val fill = extractFill(root, fillId)
      val border = extractBorder(root, borderId)
      val numFmt = extractNumFmt(root, numFmtId)
      val align = extractAlignment(xf)

      CellStyle(font, fill, border, numFmt, align = align)
    }

  private def extractFont(root: Elem, fontId: Int): Font =
    val fonts = (root \ "fonts" \ "font")
    fonts
      .lift(fontId)
      .map { f =>
        val bold = (f \ "b").nonEmpty
        val italic = (f \ "i").nonEmpty
        val underline = (f \ "u").nonEmpty
        val size = (f \ "sz" \ "@val").text.toDoubleOption.getOrElse(11.0)
        val name = (f \ "name" \ "@val").text match
          case "" => "Calibri"
          case n => n
        val color = extractColor(f \ "color")
        Font(name, size, bold, italic, underline, color)
      }
      .getOrElse(Font.default)

  private def extractFill(root: Elem, fillId: Int): Fill =
    val fills = (root \ "fills" \ "fill")
    fills
      .lift(fillId)
      .flatMap { f =>
        val pf = (f \ "patternFill")
        if pf.isEmpty then None
        else
          val patternTypeStr = (pf \ "@patternType").text
          patternTypeStr match
            case "none" | "" => Some(Fill.None)
            case "solid" =>
              extractColor(pf \ "fgColor") match
                case Some(fg) => Some(Fill.Solid(fg))
                case None => Some(Fill.None)
            case other =>
              val patternType = other match
                case "gray125" => PatternType.Gray125
                case _ => PatternType.Solid
              (extractColor(pf \ "fgColor"), extractColor(pf \ "bgColor")) match
                case (Some(fg), Some(bg)) => Some(Fill.Pattern(fg, bg, patternType))
                case _ => Some(Fill.None)
      }
      .getOrElse(Fill.None)

  private def extractBorder(root: Elem, borderId: Int): Border =
    val borders = (root \ "borders" \ "border")
    borders
      .lift(borderId)
      .map { b =>
        def extractSide(name: String): BorderSide =
          val side = (b \ name)
          if side.isEmpty then BorderSide.none
          else
            val style = (side \ "@style").text match
              case "thin" => BorderStyle.Thin
              case "medium" => BorderStyle.Medium
              case "thick" => BorderStyle.Thick
              case "dashed" => BorderStyle.Dashed
              case "dotted" => BorderStyle.Dotted
              case "double" => BorderStyle.Double
              case _ => BorderStyle.None
            val color = extractColor(side \ "color")
            BorderSide(style, color)
        Border(
          left = extractSide("left"),
          right = extractSide("right"),
          top = extractSide("top"),
          bottom = extractSide("bottom")
        )
      }
      .getOrElse(Border.none)

  private def extractNumFmt(root: Elem, numFmtId: Int): NumFmt =
    numFmtId match
      case 0 => NumFmt.General
      case 49 => NumFmt.Text
      case 4 => NumFmt.Decimal
      case 10 => NumFmt.Percent
      case 44 => NumFmt.Currency
      case 14 => NumFmt.Date
      case 22 => NumFmt.DateTime
      case 21 => NumFmt.Time
      case id =>
        // Look up custom format
        val numFmts = (root \ "numFmts" \ "numFmt")
        numFmts
          .find(nf => (nf \ "@numFmtId").text.toIntOption.contains(id))
          .map(nf => NumFmt.Custom((nf \ "@formatCode").text))
          .getOrElse(NumFmt.General)

  private def extractAlignment(xf: Node): Align =
    val alignment = (xf \ "alignment")
    if alignment.isEmpty then Align.default
    else
      val h = (alignment \ "@horizontal").text match
        case "left" => HAlign.Left
        case "center" => HAlign.Center
        case "right" => HAlign.Right
        case "justify" => HAlign.Justify
        case "fill" => HAlign.Fill
        case _ => HAlign.General
      val v = (alignment \ "@vertical").text match
        case "top" => VAlign.Top
        case "center" => VAlign.Middle
        case "bottom" => VAlign.Bottom
        case "justify" => VAlign.Justify
        case "distributed" => VAlign.Distributed
        case _ => VAlign.Bottom
      val wrap = (alignment \ "@wrapText").text == "1"
      val indent = (alignment \ "@indent").text.toIntOption.getOrElse(0)
      Align(h, v, wrap, indent)

  private def extractColor(nodes: NodeSeq): Option[Color] =
    nodes.headOption.flatMap { c =>
      val rgb = (c \ "@rgb").text
      val theme = (c \ "@theme").text
      val tint = (c \ "@tint").text.toDoubleOption.getOrElse(0.0)

      if rgb.nonEmpty then
        // Parse ARGB hex string
        try
          val argb = java.lang.Long.parseLong(rgb, 16).toInt
          Some(Color.Rgb(argb))
        catch case _: Exception => None
      else if theme.nonEmpty then
        theme.toIntOption.flatMap { t =>
          import com.tjclp.xl.styles.color.ThemeSlot
          ThemeSlot.values.lift(t).map(slot => Color.Theme(slot, tint))
        }
      else None
    }
