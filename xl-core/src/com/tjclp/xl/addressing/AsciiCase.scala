package com.tjclp.xl.addressing

/**
 * ASCII-only uppercase fold for column-letter parsing (ARef/Column), dropping the
 * `java.util.Locale` dependency from the addressing hot path.
 *
 * Column letters are structurally ASCII (`A`-`XFD`): everything outside `a`-`z` passes through
 * unchanged and is then rejected by the letter validation, exactly as before — including exotic
 * code points like U+0131 whose full-Unicode uppercase happens to land in A-Z, which the previous
 * `toUpperCase(Locale.ROOT)` fold smuggled in. Scoped to addressing ON PURPOSE: defined-name
 * folding (evaluator DependencyGraph, ooxml WorkbookLint) handles non-ASCII Unicode names and must
 * keep the locale-invariant full-Unicode `toUpperCase(Locale.ROOT)`.
 */
private[addressing] object AsciiCase:

  /** Uppercase `a`-`z` only; every other character is returned unchanged. */
  def upper(s: String): String =
    if s.forall(c => c < 'a' || c > 'z') then s
    else s.map(c => if c >= 'a' && c <= 'z' then (c - 32).toChar else c)
