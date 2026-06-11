package com.tjclp.xl.ooxml.style

import scala.xml.*

import com.tjclp.xl.ooxml.XmlUtil.{elem, elemOrdered}
import com.tjclp.xl.styles.{Dxf, DxfFont}
import com.tjclp.xl.styles.border.{Border, BorderSide, BorderStyle}
import com.tjclp.xl.styles.color.Color
import com.tjclp.xl.styles.fill.Fill
import com.tjclp.xl.styles.numfmt.NumFmt

/**
 * Strict-or-None codec for differential formats (styles.xml `<dxf>`, GH-136).
 *
 * `parse` returns Some only when the dxf is FULLY understood by the typed model — any unknown
 * child/attr, alignment, protection, extLst, gradient fill, texture pattern, font name/size, or
 * unmapped color degrades the whole dxf to None so the referencing cfRule rides through Preserved
 * ("typed = fully understood"). `toXml` emits the Excel-native dialect; the parser additionally
 * accepts the openpyxl dialect (patternType="solid" with fgColor/bgColor).
 *
 * Law: `parse(toXml(d)) == Some(d)` for every representable Dxf whose colors/numFmt codes are
 * stable under the house parse (alpha-FF colors; format codes that NumFmt.parse maps back to the
 * same case — every case except Currency).
 */
object DxfCodec:

  private def childElems(e: Elem): Either[Unit, Vector[Elem]] =
    val nonWs = e.child.filterNot {
      case Text(t) => t.forall(_.isWhitespace)
      case _ => false
    }
    val elems = nonWs.collect { case c: Elem => c }.toVector
    if elems.sizeIs == nonWs.size then Right(elems) else Left(())

  private def attrKeys(e: Elem): Set[String] = e.attributes.asAttrMap.keySet

  // ===== parse =====

  /** Strict parse: Some only when every child is inside the modeled subset. */
  def parse(e: Elem): Option[Dxf] =
    childElems(e).toOption.flatMap { children =>
      if attrKeys(e).nonEmpty then None
      else if children.exists(c => !Set("font", "numFmt", "fill", "border").contains(c.label))
      then None
      else if children.map(_.label).distinct.sizeIs != children.size then None
      else
        def one(label: String): Option[Elem] = children.find(_.label == label)
        for
          font <- traverseOpt(one("font"))(parseFont)
          numFmt <- traverseOpt(one("numFmt"))(parseNumFmt)
          fill <- traverseOpt(one("fill"))(parseFill)
          border <- traverseOpt(one("border"))(parseBorder)
        yield Dxf(font = font, fill = fill, border = border, numFmt = numFmt)
    }

  /** Option-of-Option traversal: absent → Some(None); present-but-unparseable → None. */
  private def traverseOpt[A](o: Option[Elem])(f: Elem => Option[A]): Option[Option[A]] =
    o match
      case None => Some(None)
      case Some(e) => f(e).map(Some.apply)

  /** Boolean delta child: `<b/>`/val="1"/"true" → true, val="0"/"false" → false, else degrade. */
  private def parseFlag(e: Elem): Option[Boolean] =
    if attrKeys(e).exists(_ != "val") then None
    else
      e.attribute("val").map(_.text) match
        case None => Some(true)
        case Some("1") | Some("true") => Some(true)
        case Some("0") | Some("false") => Some(false)
        case Some(_) => None

  /** Underline delta: `<u/>`/single → true, none → false; other vals (double, ...) degrade. */
  private def parseUnderline(e: Elem): Option[Boolean] =
    if attrKeys(e).exists(_ != "val") then None
    else
      e.attribute("val").map(_.text) match
        case None | Some("single") => Some(true)
        case Some("none") => Some(false)
        case Some(_) => None

  private def parseFont(e: Elem): Option[DxfFont] =
    childElems(e).toOption.flatMap { children =>
      if attrKeys(e).nonEmpty then None
      else if children.exists(c => !Set("b", "i", "strike", "u", "color").contains(c.label))
      then None
      else if children.map(_.label).distinct.sizeIs != children.size then None
      else
        def one(label: String): Option[Elem] = children.find(_.label == label)
        for
          bold <- traverseOpt(one("b"))(parseFlag)
          italic <- traverseOpt(one("i"))(parseFlag)
          strike <- traverseOpt(one("strike"))(parseFlag)
          underline <- traverseOpt(one("u"))(parseUnderline)
          color <- traverseOpt(one("color"))(parseColor)
        yield DxfFont(bold, italic, strike, underline, color)
    }

  /**
   * Strict color: exactly one of rgb / theme(+tint) / indexed-with-palette-mapping. `auto`, unknown
   * attrs, or an unmapped indexed value degrade. A 00 alpha byte is canonicalized to FF (the house
   * parser rule — Excel/openpyxl write 00RRGGBB for opaque).
   */
  private[ooxml] def parseColor(e: Elem): Option[Color] =
    childElems(e).toOption.filter(_.isEmpty).flatMap { _ =>
      val attrs = e.attributes.asAttrMap
      attrs.keySet match
        case s if s == Set("rgb") => attrs.get("rgb").flatMap(parseRgb)
        case s if s == Set("theme") || s == Set("theme", "tint") =>
          for
            idx <- attrs.get("theme").flatMap(_.toIntOption)
            slot <- ColorHelpers.themeSlotFromIndex(idx)
            tint <- attrs.get("tint").fold(Option(0.0))(_.toDoubleOption)
          yield Color.Theme(slot, tint)
        case s if s == Set("indexed") =>
          attrs.get("indexed").flatMap(_.toIntOption).flatMap(ColorHelpers.indexedColorToRgb)
        case _ => None
    }

  private def parseRgb(value: String): Option[Color] =
    val cleaned = value.trim.stripPrefix("#")
    val normalized =
      if cleaned.length == 6 then Some("FF" + cleaned)
      else if cleaned.length == 8 then
        if cleaned.substring(0, 2) == "00" then Some("FF" + cleaned.substring(2)) else Some(cleaned)
      else None
    normalized.flatMap { hex =>
      try Some(Color.Rgb(java.lang.Long.parseUnsignedLong(hex, 16).toInt))
      catch case _: NumberFormatException => None
    }

  /**
   * Both wild fill dialects, normalized to Fill.Solid: Excel-native
   * `<patternFill><bgColor/></patternFill>` and openpyxl
   * `<patternFill patternType="solid"><fgColor/><bgColor/></patternFill>` (prefer bgColor).
   */
  private def parseFill(e: Elem): Option[Fill] =
    childElems(e).toOption.flatMap {
      case Vector(pf) if pf.label == "patternFill" && attrKeys(e).isEmpty =>
        childElems(pf).toOption.flatMap { colorChildren =>
          val patternOk = attrKeys(pf) match
            case s if s.isEmpty => true
            case s if s == Set("patternType") =>
              pf.attribute("patternType").map(_.text).contains("solid")
            case _ => false
          if !patternOk then None
          else if colorChildren.exists(c => !Set("fgColor", "bgColor").contains(c.label)) then None
          else
            for
              fg <- traverseOpt(colorChildren.find(_.label == "fgColor"))(parseColor)
              bg <- traverseOpt(colorChildren.find(_.label == "bgColor"))(parseColor)
              chosen <- bg.orElse(fg) match
                case Some(c) => Some(c)
                case None => None
            yield Fill.Solid(chosen)
        }
      case _ => None
    }

  private val borderStyleByToken: Map[String, BorderStyle] =
    BorderStyle.values.map(s => OoxmlStyles.borderStyleToken(s) -> s).toMap

  private def parseBorder(e: Elem): Option[Border] =
    childElems(e).toOption.flatMap { children =>
      val allowed = Set("left", "right", "top", "bottom")
      if attrKeys(e).nonEmpty then None
      else if children.exists(c => !allowed.contains(c.label)) then None
      else if children.map(_.label).distinct.sizeIs != children.size then None
      else
        def side(label: String): Option[BorderSide] =
          children.find(_.label == label) match
            case None => Some(BorderSide.none)
            case Some(s) => parseBorderSide(s)
        for
          left <- side("left")
          right <- side("right")
          top <- side("top")
          bottom <- side("bottom")
        yield Border(left, right, top, bottom)
    }

  private def parseBorderSide(e: Elem): Option[BorderSide] =
    childElems(e).toOption.flatMap { children =>
      if attrKeys(e).exists(_ != "style") then None
      else if children.exists(_.label != "color") || children.sizeIs > 1 then None
      else
        val styleOpt = e.attribute("style").map(_.text) match
          case None => Some(BorderStyle.None)
          case Some(token) => borderStyleByToken.get(token)
        styleOpt.flatMap { style =>
          children.headOption match
            case None => Some(BorderSide(style, None))
            case Some(c) => parseColor(c).map(col => BorderSide(style, Some(col)))
        }
    }

  private def parseNumFmt(e: Elem): Option[NumFmt] =
    childElems(e).toOption.filter(_.isEmpty).flatMap { _ =>
      if attrKeys(e).exists(k => k != "numFmtId" && k != "formatCode") then None
      else e.attribute("formatCode").map(a => NumFmt.parse(a.text))
    }

  // ===== emission =====

  /** Canonical Excel-native emission, children in CT_Dxf order: font → numFmt → fill → border. */
  def toXml(dxf: Dxf): Elem =
    val children = Vector(
      dxf.font.map(fontToXml),
      dxf.numFmt.map(numFmtToXml),
      dxf.fill.map(fillToXml),
      dxf.border.map(borderToXml)
    ).flatten
    elem("dxf")(children*)

  private def flagToXml(label: String, value: Boolean): Elem =
    if value then elem(label)() else elem(label, "val" -> "0")()

  private def fontToXml(f: DxfFont): Elem =
    val children = Vector(
      f.bold.map(flagToXml("b", _)),
      f.italic.map(flagToXml("i", _)),
      f.strike.map(flagToXml("strike", _)),
      f.underline.map(u => if u then elem("u")() else elem("u", "val" -> "none")()),
      f.color.map(colorToXml)
    ).flatten
    elem("font")(children*)

  private[ooxml] def colorToXml(color: Color): Elem = color match
    case Color.Rgb(argb) => elem("color", "rgb" -> f"$argb%08X")()
    case Color.Theme(slot, tint) =>
      elem(
        "color",
        "theme" -> ColorHelpers.themeSlotToIndex(slot).toString,
        "tint" -> tint.toString
      )()

  /**
   * Excel-native dxf fill: solid via bare patternFill + bgColor (matches Excel's own cf writer).
   * Fill.None and Fill.Pattern keep total emission via the standard dialect; they are outside the
   * typed parse subset and exist only for direct model construction.
   */
  private def fillToXml(fill: Fill): Elem = fill match
    case Fill.Solid(color) =>
      elem("fill")(elem("patternFill")(colorToXml(color).copy(label = "bgColor")))
    case Fill.None =>
      elem("fill")(elem("patternFill", "patternType" -> "none")())
    case Fill.Pattern(fg, bg, pattern) =>
      elem("fill")(
        elem("patternFill", "patternType" -> OoxmlStyles.patternTypeToken(pattern))(
          colorToXml(fg).copy(label = "fgColor"),
          colorToXml(bg).copy(label = "bgColor")
        )
      )

  private def borderToXml(border: Border): Elem =
    def sideToXml(label: String, side: BorderSide): Elem =
      if side.style == BorderStyle.None then elem(label)()
      else
        elem(label, "style" -> OoxmlStyles.borderStyleToken(side.style))(
          side.color.map(colorToXml).toList*
        )
    elem("border")(
      sideToXml("left", border.left),
      sideToXml("right", border.right),
      sideToXml("top", border.top),
      sideToXml("bottom", border.bottom)
    )

  /**
   * Dxf numFmts are self-contained (id + code inline, not a numFmts-table reference). Custom codes
   * use the conventional first user id (164); consumers read the code, the id is vestigial.
   */
  private def numFmtToXml(fmt: NumFmt): Elem =
    val code = formatCode(fmt)
    val id = NumFmt.builtInId(fmt).getOrElse(164)
    elemOrdered("numFmt", "numFmtId" -> id.toString, "formatCode" -> code)()

  /** The format-code string for emission (the NumFmt.parse retraction where one exists). */
  private def formatCode(fmt: NumFmt): String = fmt match
    case NumFmt.General => "General"
    case NumFmt.Integer => "0"
    case NumFmt.Decimal => "0.00"
    case NumFmt.ThousandsSeparator => "#,##0"
    case NumFmt.ThousandsDecimal => "#,##0.00"
    case NumFmt.Currency => "$#,##0.00" // NB: parses back as Custom — documented one-way case
    case NumFmt.Percent => "0%"
    case NumFmt.PercentDecimal => "0.00%"
    case NumFmt.Scientific => "0.00E+00"
    case NumFmt.Fraction => "# ?/?"
    case NumFmt.Date => "m/d/yy"
    case NumFmt.DateTime => "m/d/yy h:mm"
    case NumFmt.Time => "h:mm:ss"
    case NumFmt.Text => "@"
    case NumFmt.Custom(code) => code
