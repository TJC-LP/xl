package com.tjclp.xl.ooxml

/**
 * Generated default Office theme part (GH-387).
 *
 * A scratch-built workbook whose styles reference `Color.Theme` slots serializes `<color
 * theme="N"/>` records; without a theme part nothing in the package can resolve them (Excel
 * silently substitutes its built-in Office theme, strict consumers reject the package). This is the
 * writer's counterpart to [[ThemeParser]] — parse-only until now — emitting the standard Office
 * palette for exactly that case. Source-backed writes never use it: a preserved
 * `xl/theme/theme1.xml` is copied verbatim.
 *
 * INVARIANTS:
 *   - `<a:clrScheme>` children are in OOXML document order (dk1, lt1, dk2, lt2, accent1-6, hlink,
 *     folHlink) — the SLOT-INDEX mapping consumers apply is `ColorHelpers.themeSlotToIndex` (lt1=0,
 *     dk1=1, lt2=2, dk2=3, accent1-6=4-9), which is NOT document order; only the element NAMES
 *     carry identity here.
 *   - Palette values match [[ThemeParser]]'s per-slot fallback defaults byte-for-byte, so a
 *     resolved color is identical whether a consumer reads this part or defaults it (the accent
 *     slots share the [[DefaultTheme.accents]] constants on both sides — GH-407).
 *   - Tint stays on the style color (`<color theme="N" tint="..."/>`) — the theme part carries only
 *     the base palette.
 *   - Schema-complete `CT_BaseStyles`: clrScheme, fontScheme (latin+ea+cs), and a minimal fmtScheme
 *     with the required three entries per style list.
 */
private[ooxml] object DefaultTheme:

  val path: String = "xl/theme/theme1.xml"

  /**
   * The modern Office accent palette (accent1-6) as RGB ints — the single source of truth for the
   * emitted theme part below (its `<a:accentN>` entries derive from this vector, so string and
   * constants cannot drift), [[ThemeParser]]'s per-slot fallback defaults, and the chart writer's
   * default series-fill cycle (GH-407). NOT the 2007-era `ThemePalette.office` palette.
   */
  val accents: Vector[Int] =
    Vector(0x4472c4, 0xed7d31, 0xa5a5a5, 0xffc000, 0x5b9bd5, 0x70ad47)

  /** Accent color for 0-based index `i` (cycling) as 6-digit uppercase hex (`4472C4`). */
  def accentHex(i: Int): String = hex6(accents(math.floorMod(i, accents.size)))

  /** Accent color for 0-based index `i` (cycling) as an opaque ARGB int. */
  def accentArgb(i: Int): Int = 0xff000000 | accents(math.floorMod(i, accents.size))

  /** 6-digit uppercase RRGGBB of an RGB/ARGB int (alpha dropped). */
  def hex6(rgb: Int): String = f"${rgb & 0xffffff}%06X"

  private val accentEntries: String =
    accents.zipWithIndex
      .map((rgb, i) => s"""<a:accent${i + 1}><a:srgbClr val="${hex6(rgb)}"/></a:accent${i + 1}>""")
      .mkString

  val xml: String =
    s"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme"><a:themeElements><a:clrScheme name="Office"><a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="44546A"/></a:dk2><a:lt2><a:srgbClr val="E7E6E6"/></a:lt2>$accentEntries<a:hlink><a:srgbClr val="0563C1"/></a:hlink><a:folHlink><a:srgbClr val="954F72"/></a:folHlink></a:clrScheme><a:fontScheme name="Office"><a:majorFont><a:latin typeface="Calibri Light" panose="020F0302020204030204"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont><a:minorFont><a:latin typeface="Calibri" panose="020F0502020204030204"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont></a:fontScheme><a:fmtScheme name="Office"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst><a:lnStyleLst><a:ln w="6350" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln><a:ln w="12700" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln><a:ln w="19050" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst></a:fmtScheme></a:themeElements><a:objectDefaults/><a:extraClrSchemeLst/></a:theme>"""
