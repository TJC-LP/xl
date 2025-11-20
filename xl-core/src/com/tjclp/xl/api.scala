package com.tjclp.xl

/**
 * Pure Scala 3.7 Excel (OOXML) Library.
 *
 * Import `com.tjclp.xl.api.*` for the full domain model (cells, sheets, workbooks, styles, etc.)
 * and `com.tjclp.xl.syntax.*` for helpers like `col`, `row`, `ref`, and `String` parsing extensions
 * (requires `xl-macros`). Macro-free builds can use `com.tjclp.xl.coreSyntax.*`.
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
  export cells.{Cell, CellValue, CellError}

  // Addressing types
  export addressing.{Column, Row, SheetName, ARef, CellRange, RefType}

  // Rich text types
  export richtext.{TextRun, RichText}

  // Sheet types
  export sheets.{Sheet, ColumnProperties, RowProperties}

  // Patch types
  export patch.Patch

  // Workbook types
  export workbooks.{Workbook, WorkbookMetadata}

  // Optics types
  export optics.{Lens, Optional, Optics}

  // Formatted type
  export formatted.Formatted

  // Codec types
  export codec.{CellCodec, CellReader, CellWriter, CodecError}

  // Style types - core
  export styles.{CellStyle, StyleRegistry}

  // Style types - alignment
  export styles.alignment.{HAlign, VAlign, Align}

  // Style types - border
  export styles.border.{BorderStyle, BorderSide, Border}

  // Style types - color
  export styles.color.{ThemeSlot, Color}

  // Style types - fill
  export styles.fill.{PatternType, Fill}

  // Style types - font
  export styles.font.Font

  // Style types - numfmt
  export styles.numfmt.NumFmt

  // Style types - patch
  export styles.patch.StylePatch

  // Style types - units
  export styles.units.{Pt, Px, Emu, StyleId}

  // Surface Column helpers at the root package for single-import ergonomics
  export addressing.Column.toLetter

  // ========== String Parsing Helpers ==========
  extension (s: String)
    /** Parse string as cell reference */
    def asCell: XLResult[ARef] =
      ARef.parse(s).left.map(err => XLError.InvalidCellRef(s, err))

    /** Parse string as range */
    def asRange: XLResult[CellRange] =
      CellRange.parse(s).left.map(err => XLError.InvalidRange(s, err))

    /** Parse string as sheet name */
    def asSheetName: XLResult[SheetName] =
      SheetName(s).left.map(err => XLError.InvalidSheetName(s, err))

export api.*
