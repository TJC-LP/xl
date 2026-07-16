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
 *     resolved color is identical whether a consumer reads this part or defaults it.
 *   - Tint stays on the style color (`<color theme="N" tint="..."/>`) — the theme part carries only
 *     the base palette.
 *   - Schema-complete `CT_BaseStyles`: clrScheme, fontScheme (latin+ea+cs), and a minimal fmtScheme
 *     with the required three entries per style list.
 */
private[ooxml] object DefaultTheme:

  val path: String = "xl/theme/theme1.xml"

  val xml: String =
    """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme"><a:themeElements><a:clrScheme name="Office"><a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="44546A"/></a:dk2><a:lt2><a:srgbClr val="E7E6E6"/></a:lt2><a:accent1><a:srgbClr val="4472C4"/></a:accent1><a:accent2><a:srgbClr val="ED7D31"/></a:accent2><a:accent3><a:srgbClr val="A5A5A5"/></a:accent3><a:accent4><a:srgbClr val="FFC000"/></a:accent4><a:accent5><a:srgbClr val="5B9BD5"/></a:accent5><a:accent6><a:srgbClr val="70AD47"/></a:accent6><a:hlink><a:srgbClr val="0563C1"/></a:hlink><a:folHlink><a:srgbClr val="954F72"/></a:folHlink></a:clrScheme><a:fontScheme name="Office"><a:majorFont><a:latin typeface="Calibri Light" panose="020F0302020204030204"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont><a:minorFont><a:latin typeface="Calibri" panose="020F0502020204030204"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont></a:fontScheme><a:fmtScheme name="Office"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst><a:lnStyleLst><a:ln w="6350" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln><a:ln w="12700" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln><a:ln w="19050" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst></a:fmtScheme></a:themeElements><a:objectDefaults/><a:extraClrSchemeLst/></a:theme>"""
