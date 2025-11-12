package com.tjclp.xl.style

/**
 * Package object for style types.
 *
 * Re-exports all types from subdirectories for convenient imports. This allows both:
 *   - `import com.tjclp.xl.style.Font` (convenient)
 *   - `import com.tjclp.xl.style.font.Font` (explicit)
 */

// Alignment types
export alignment.{HAlign, VAlign, Align}

// Border types
export border.{BorderStyle, BorderSide, Border}

// Color types
export color.{ThemeSlot, Color}

// Fill types
export fill.{PatternType, Fill}

// Font type
export font.Font

// Number format type
export numfmt.NumFmt

// Style patch type
export patch.StylePatch

// Unit types
export units.{Pt, Px, Emu, StyleId}
