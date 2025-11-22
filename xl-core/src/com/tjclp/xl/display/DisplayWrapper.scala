package com.tjclp.xl.display

/**
 * Wrapper for Excel cell values that provides formatted toString representation.
 *
 * This type enables automatic formatted display in string interpolation:
 * {{{
 * import com.tjclp.xl.display.{*, given}
 * given Sheet = mySheet
 * println(s"Revenue: \${ref"A1"}")  // Automatic formatting!
 * }}}
 *
 * The formatted string respects the cell's NumFmt (Currency, Percent, Date, etc.) and matches what
 * Excel would display.
 *
 * @param formatted
 *   The pre-formatted display string matching Excel conventions
 */
final case class DisplayWrapper(formatted: String):
  /** Returns the formatted string when used in string interpolation or toString. */
  override def toString: String = formatted
