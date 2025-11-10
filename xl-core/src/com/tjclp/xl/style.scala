package com.tjclp.xl

import cats.Monoid

/** Style system for Excel cells with units, colors, fonts, fills, borders, and number formats.
  *
  * All types are pure values with deterministic canonicalization for deduplication.
  */

// ========== Units ==========

/** Point (1/72 inch) - Excel's primary unit for fonts and dimensions */
opaque type Pt = Double

object Pt:
  def apply(value: Double): Pt = value

  extension (pt: Pt)
    def value: Double = pt

    /** Convert to pixels (assumes 96 DPI) */
    def toPx: Px = Px(pt * 96.0 / 72.0)

    /** Convert to EMUs (English Metric Units) */
    def toEmu: Emu = Emu((pt * 914400.0 / 72.0).toLong)

/** Pixel - screen unit (assumes 96 DPI for conversions) */
opaque type Px = Double

object Px:
  def apply(value: Double): Px = value

  extension (px: Px)
    def value: Double = px

    /** Convert to points */
    def toPt: Pt = Pt(px * 72.0 / 96.0)

    /** Convert to EMUs */
    def toEmu: Emu = Emu((px * 914400.0 / 96.0).toLong)

/** English Metric Unit - OOXML's precise unit (1/914400 inch) */
opaque type Emu = Long

object Emu:
  def apply(value: Long): Emu = value

  extension (emu: Emu)
    def value: Long = emu

    /** Convert to points */
    def toPt: Pt = Pt((emu * 72.0) / 914400.0)

    /** Convert to pixels */
    def toPx: Px = Px((emu * 96.0) / 914400.0)

// ========== Colors ==========

/** Theme color slots from Office theme */
enum ThemeSlot:
  case Dark1, Light1, Dark2, Light2
  case Accent1, Accent2, Accent3, Accent4, Accent5, Accent6

/** Color representation: either RGB or theme-based with tint */
enum Color:
  /** RGB color with alpha channel (ARGB format: 0xAARRGGBB) */
  case Rgb(argb: Int)

  /** Theme color with optional tint (-1.0 to 1.0, where 0 is theme default) */
  case Theme(slot: ThemeSlot, tint: Double)

object Color:
  /** Create RGB color from components (0-255) */
  def fromRgb(r: Int, g: Int, b: Int, a: Int = 255): Color =
    Rgb((a << 24) | (r << 16) | (g << 8) | b)

  /** Create color from hex string (#RRGGBB or #AARRGGBB) */
  def fromHex(hex: String): Either[String, Color] =
    val cleaned = if hex.startsWith("#") then hex.substring(1) else hex
    cleaned.length match
      case 6 => // #RRGGBB
        try Right(Rgb(0xFF000000 | Integer.parseInt(cleaned, 16)))
        catch case _: NumberFormatException => Left(s"Invalid hex color: $hex")
      case 8 => // #AARRGGBB
        try Right(Rgb(Integer.parseUnsignedInt(cleaned, 16)))
        catch case _: NumberFormatException => Left(s"Invalid hex color: $hex")
      case _ => Left(s"Invalid hex color length: $hex")

  /** Validate tint is in valid range */
  def validTint(tint: Double): Either[String, Double] =
    if tint >= -1.0 && tint <= 1.0 then Right(tint)
    else Left(s"Tint must be in [-1.0, 1.0], got: $tint")

  extension (color: Color)
    /** Convert to ARGB integer */
    def toArgb: Int = color match
      case Rgb(argb) => argb
      case Theme(_, _) => 0 // Needs theme resolution at IO boundary

    /** Get hex string representation (#AARRGGBB) */
    def toHex: String = color match
      case Rgb(argb) => f"#$argb%08X"
      case Theme(slot, tint) => s"Theme($slot, $tint)"

// ========== Number Formats ==========

/** Number format for cell values */
enum NumFmt:
  case General
  case Integer              // 0
  case Decimal              // 0.00
  case ThousandsSeparator   // #,##0
  case ThousandsDecimal     // #,##0.00
  case Currency             // $#,##0.00
  case Percent              // 0%
  case PercentDecimal       // 0.00%
  case Scientific           // 0.00E+00
  case Fraction             // # ?/?
  case Date                 // m/d/yy
  case DateTime             // m/d/yy h:mm
  case Time                 // h:mm:ss
  case Text                 // @
  case Custom(code: String) // User-defined format

object NumFmt:
  /** Built-in format ID mapping for OOXML */
  def builtInId(fmt: NumFmt): Option[Int] = fmt match
    case General => Some(0)
    case Integer => Some(1)
    case Decimal => Some(2)
    case ThousandsSeparator => Some(3)
    case ThousandsDecimal => Some(4)
    case Percent => Some(9)
    case PercentDecimal => Some(10)
    case Scientific => Some(11)
    case Fraction => Some(12)
    case Currency => Some(7) // Simplified; actual format varies by locale
    case Date => Some(14)
    case Time => Some(21)
    case DateTime => Some(22)
    case Text => Some(49)
    case Custom(_) => None

  /** Parse format code to NumFmt */
  def parse(code: String): NumFmt = code match
    case "General" => General
    case "0" => Integer
    case "0.00" => Decimal
    case "#,##0" => ThousandsSeparator
    case "#,##0.00" => ThousandsDecimal
    case "0%" => Percent
    case "0.00%" => PercentDecimal
    case "0.00E+00" => Scientific
    case "# ?/?" => Fraction
    case "m/d/yy" => Date
    case "m/d/yy h:mm" => DateTime
    case "h:mm:ss" => Time
    case "@" => Text
    case other => Custom(other)

// ========== Font ==========

/** Font styling for cell text */
case class Font(
  name: String = "Calibri",
  sizePt: Double = 11.0,
  bold: Boolean = false,
  italic: Boolean = false,
  underline: Boolean = false,
  color: Option[Color] = None
):
  require(sizePt > 0, s"Font size must be positive, got: $sizePt")
  require(name.nonEmpty, "Font name cannot be empty")

  def withName(n: String): Font = copy(name = n)
  def withSize(size: Double): Font = copy(sizePt = size)
  def withBold(b: Boolean = true): Font = copy(bold = b)
  def withItalic(i: Boolean = true): Font = copy(italic = i)
  def withUnderline(u: Boolean = true): Font = copy(underline = u)
  def withColor(c: Color): Font = copy(color = Some(c))
  def clearColor: Font = copy(color = None)

object Font:
  val default: Font = Font()

// ========== Fill ==========

/** Fill pattern types */
enum PatternType:
  case None, Solid, Gray125, Gray0625
  case DarkGray, MediumGray, LightGray
  case DarkHorizontal, DarkVertical, DarkDown, DarkUp
  case DarkGrid, DarkTrellis
  case LightHorizontal, LightVertical, LightDown, LightUp
  case LightGrid, LightTrellis

/** Cell background fill */
enum Fill:
  case None
  case Solid(color: Color)
  case Pattern(foreground: Color, background: Color, pattern: PatternType)

object Fill:
  val default: Fill = None

// ========== Border ==========

/** Border line style */
enum BorderStyle:
  case None, Thin, Medium, Thick
  case Dashed, Dotted, Double
  case Hair, DashDot, DashDotDot
  case SlantDashDot, MediumDashed, MediumDashDot, MediumDashDotDot

/** Single border side */
case class BorderSide(
  style: BorderStyle = BorderStyle.None,
  color: Option[Color] = None
)

object BorderSide:
  val none: BorderSide = BorderSide()
  def apply(style: BorderStyle): BorderSide = BorderSide(style, None)
  def apply(style: BorderStyle, color: Color): BorderSide = BorderSide(style, Some(color))

/** Cell borders (all four sides) */
case class Border(
  left: BorderSide = BorderSide.none,
  right: BorderSide = BorderSide.none,
  top: BorderSide = BorderSide.none,
  bottom: BorderSide = BorderSide.none
):
  def withLeft(side: BorderSide): Border = copy(left = side)
  def withRight(side: BorderSide): Border = copy(right = side)
  def withTop(side: BorderSide): Border = copy(top = side)
  def withBottom(side: BorderSide): Border = copy(bottom = side)

object Border:
  val none: Border = Border()

  /** Create uniform border on all sides */
  def all(style: BorderStyle, color: Option[Color] = None): Border =
    val side = BorderSide(style, color)
    Border(side, side, side, side)

// ========== Alignment ==========

/** Horizontal alignment */
enum HAlign:
  case Left, Center, Right, Justify, Fill, CenterContinuous, Distributed

/** Vertical alignment */
enum VAlign:
  case Top, Middle, Bottom, Justify, Distributed

/** Cell alignment settings */
case class Align(
  horizontal: HAlign = HAlign.Left,
  vertical: VAlign = VAlign.Bottom,
  wrapText: Boolean = false,
  indent: Int = 0
):
  require(indent >= 0, s"Indent must be non-negative, got: $indent")

  def withHAlign(h: HAlign): Align = copy(horizontal = h)
  def withVAlign(v: VAlign): Align = copy(vertical = v)
  def withWrap(w: Boolean = true): Align = copy(wrapText = w)
  def withIndent(i: Int): Align = copy(indent = i)

object Align:
  val default: Align = Align()

// ========== CellStyle ==========

/** Complete cell style combining all formatting aspects */
case class CellStyle(
  font: Font = Font.default,
  fill: Fill = Fill.default,
  border: Border = Border.none,
  numFmt: NumFmt = NumFmt.General,
  align: Align = Align.default
):
  def withFont(f: Font): CellStyle = copy(font = f)
  def withFill(f: Fill): CellStyle = copy(fill = f)
  def withBorder(b: Border): CellStyle = copy(border = b)
  def withNumFmt(n: NumFmt): CellStyle = copy(numFmt = n)
  def withAlign(a: Align): CellStyle = copy(align = a)

object CellStyle:
  val default: CellStyle = CellStyle()

  /** Generate canonical key for style deduplication.
    *
    * Two styles with the same key are structurally equivalent and
    * should map to the same style index in styles.xml.
    */
  def canonicalKey(style: CellStyle): String =
    val fontKey = s"F:${style.font.name},${style.font.sizePt},${style.font.bold},${style.font.italic},${style.font.underline},${style.font.color}"
    val fillKey = s"FL:${style.fill}"
    val borderKey = s"B:${style.border.left},${style.border.right},${style.border.top},${style.border.bottom}"
    val numFmtKey = s"N:${style.numFmt}"
    val alignKey = s"A:${style.align.horizontal},${style.align.vertical},${style.align.wrapText},${style.align.indent}"
    s"$fontKey|$fillKey|$borderKey|$numFmtKey|$alignKey"

// ========== StylePatch Monoid ==========

/** Patch operations for CellStyle with Monoid semantics */
enum StylePatch:
  case SetFont(font: Font)
  case SetFill(fill: Fill)
  case SetBorder(border: Border)
  case SetNumFmt(numFmt: NumFmt)
  case SetAlign(align: Align)
  case Batch(patches: Vector[StylePatch])

object StylePatch:
  val empty: StylePatch = Batch(Vector.empty)

  def combine(p1: StylePatch, p2: StylePatch): StylePatch = (p1, p2) match
    case (Batch(ps1), Batch(ps2)) => Batch(ps1 ++ ps2)
    case (Batch(ps1), p2) => Batch(ps1 :+ p2)
    case (p1, Batch(ps2)) => Batch(p1 +: ps2)
    case (p1, p2) => Batch(Vector(p1, p2))

  given Monoid[StylePatch] with
    def empty: StylePatch = StylePatch.empty
    def combine(x: StylePatch, y: StylePatch): StylePatch = StylePatch.combine(x, y)

  /** Apply a style patch to create a new style */
  def applyPatch(style: CellStyle, patch: StylePatch): CellStyle = patch match
    case SetFont(font) => style.withFont(font)
    case SetFill(fill) => style.withFill(fill)
    case SetBorder(border) => style.withBorder(border)
    case SetNumFmt(numFmt) => style.withNumFmt(numFmt)
    case SetAlign(align) => style.withAlign(align)
    case Batch(patches) =>
      patches.foldLeft(style)((s, p) => applyPatch(s, p))

  /** Apply multiple patches in sequence */
  def applyPatches(style: CellStyle, patches: Iterable[StylePatch]): CellStyle =
    applyPatch(style, Batch(patches.toVector))

  extension (style: CellStyle)
    @annotation.targetName("applyStylePatchExt")
    def applyPatch(patch: StylePatch): CellStyle =
      StylePatch.applyPatch(style, patch)

    @annotation.targetName("applyStylePatchesExt")
    def applyPatches(patches: StylePatch*): CellStyle =
      StylePatch.applyPatches(style, patches)
