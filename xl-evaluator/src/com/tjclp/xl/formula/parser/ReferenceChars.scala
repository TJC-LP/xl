package com.tjclp.xl.formula.parser

/**
 * The character classes the parser reads reference tokens with, in one place. The `:` printer
 * (`ReferenceOperators.renderRange`) and the dynamic-cell pre-filter
 * (`ReferenceOperators.mayContainComputedRange`) restate parts of the parser's grammar over printed
 * text, so they read tokens with these same predicates: widening one here widens all three, and
 * `ReferenceOperatorLawsSpec` checks the printer against the parser itself.
 */
private[formula] object ReferenceChars:

  /**
   * A character of a bare token — a cell (`$A$1`), a defined or LET name (`Sales.Total`, `_rate`,
   * `\path`, GH-394), a function name: letters, digits, `$`, `_`, `.` and `\`.
   */
  def isToken(c: Char): Boolean =
    c.isLetterOrDigit || c == '$' || c == '_' || c == '.' || c == '\\'

  /**
   * A character of a function name as a call spells it (`LOG10`, `_xlfn.XLOOKUP`): letters, digits,
   * `_` and `.` — never `$` or `\`, which no function name holds.
   */
  def isFunctionName(c: Char): Boolean = c.isLetterOrDigit || c == '_' || c == '.'

  /**
   * What the sheet-qualified reader takes after a qualifier's `!`: token characters and `:`, so it
   * reads a range (`Sheet1!A1:B2`) — and runs on through any `:` that follows.
   */
  def isQualifiedRef(c: Char): Boolean = isToken(c) || c == ':'

  /**
   * What the reader of a lexical range takes as its end token after `A1:`: every token character
   * ([[isToken]], so a defined name that begins like a cell — `B2_x`, `C3.Total`, `D5\y`, `UK9\f` —
   * is read whole and classified whole, as the first token and a sheet-qualified end are: #841) and
   * `!` (a qualified end such as `A1:Sheet2!B2` is read whole, then refused). `CellRange.parse`
   * then decides whether the token is a cell, column or row; any other token is the range
   * operator's right operand (`FormulaParser.rangeFallback`).
   */
  def isRangeEnd(c: Char): Boolean = isToken(c) || c == '!'
