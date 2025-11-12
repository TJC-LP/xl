package com.tjclp.xl

/**
 * Pure Scala 3.7 Excel (OOXML) Library.
 *
 * Import `com.tjclp.xl.api.*` for the full domain model (cells, sheets, workbooks, styles, etc.)
 * and `com.tjclp.xl.syntax.*` for helpers like `col`, `row`, `ref`, and `String` parsing
 * extensions.
 *
 * Example:
 * {{{
 * import com.tjclp.xl.api.*
 * import com.tjclp.xl.syntax.*
 *
 * val sheet = Sheet("Demo").map { s =>
 *   s.put(col(0) -> row(0), CellValue.Text("Hello"))
 * }
 * }}}
 */
object api:
  // Error types
  export error.{XLError, XLResult}

  // Cell types
  export cell.{Cell, CellValue, CellError}

  // Addressing types
  export addressing.{Column, Row, SheetName, ARef, CellRange, RefType}

  // Rich text types
  export richtext.{TextRun, RichText}

  // Sheet types
  export sheet.{Sheet, ColumnProperties, RowProperties}

  // Patch types
  export patch.Patch

  // Workbook types
  export workbook.{Workbook, WorkbookMetadata}

  // Optics types
  export optics.{Lens, Optional, Optics}

  // Formatted type
  export formatted.Formatted

  // Codec types
  export codec.{CellCodec, CellReader, CellWriter, CodecError}

  // Style types - core
  export style.{CellStyle, StyleRegistry}

  // Style types - alignment
  export style.alignment.{HAlign, VAlign, Align}

  // Style types - border
  export style.border.{BorderStyle, BorderSide, Border}

  // Style types - color
  export style.color.{ThemeSlot, Color}

  // Style types - fill
  export style.fill.{PatternType, Fill}

  // Style types - font
  export style.font.Font

  // Style types - numfmt
  export style.numfmt.NumFmt

  // Style types - patch
  export style.patch.StylePatch

  // Style types - units
  export style.units.{Pt, Px, Emu, StyleId}

export api.*
