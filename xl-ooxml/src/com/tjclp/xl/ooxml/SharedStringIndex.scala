package com.tjclp.xl.ooxml

/**
 * The table position a `t="s"` cell value names (ECMA-376 Part 1, §18.3.1.4 `c`, `t` value `s`;
 * §18.4.9 `sst`: a zero-based index into the table's `<si>` items).
 *
 * Every reader and the writer's reference scan read the value with this one rule, so a position the
 * writer keeps is the position the readers resolve. The value names position `n` when it is, in
 * order:
 *   - optional XML whitespace (space, tab, CR, LF);
 *   - an optional `+`;
 *   - one or more ASCII digits `0`-`9` (leading zeros allowed) whose value is at most
 *     `Int.MaxValue`;
 *   - optionally a `.` followed by zero or more `0` (an integral decimal such as `3.0`);
 *   - optional XML whitespace.
 *
 * Anything else names no position: an empty value, a sign other than one leading `+`, a fraction
 * with a non-zero digit, an exponent, and digits outside ASCII (Arabic-Indic `٣`, fullwidth `３`).
 * LibreOffice reads `+3`, `03`, ` 3` and `3.0` as position 3 and the non-ASCII digits as 0, so
 * those spellings are accepted and the non-ASCII digits are not.
 *
 * [[Accumulator]] applies the rule to a value delivered in pieces (SAX `characters` calls) without
 * keeping its characters: it holds one state and one running value, so a value of any length,
 * including one padded with any number of leading zeros, costs the same memory.
 */
private[xl] object SharedStringIndex:

  /** The position `value` names, or None when it names none (see the object doc). */
  def parse(value: CharSequence): Option[Int] =
    val acc = new Accumulator
    acc.append(value)
    acc.result

  /** One value's parse state, fed piece by piece. Reusable after [[reset]]. */
  @SuppressWarnings(Array("org.wartremover.warts.Var", "org.wartremover.warts.While"))
  final class Accumulator:
    // 0 leading whitespace, 1 after '+', 2 digits, 3 '.' or fraction zeros, 4 trailing
    // whitespace, 5 names no position
    private var state = 0
    private var value = 0L

    def reset(): Unit =
      state = 0
      value = 0L

    def append(text: CharSequence): Unit =
      var i = 0
      while i < text.length do
        step(text.charAt(i))
        i += 1

    def append(ch: Array[Char], start: Int, length: Int): Unit =
      var i = start
      while i < start + length do
        step(ch(i))
        i += 1

    /** The position read so far, or None when the value read so far names none. */
    def result: Option[Int] =
      if state == 2 || state == 3 || state == 4 then Some(value.toInt) else None

    private def isXmlSpace(c: Char): Boolean =
      c == ' ' || c == '\t' || c == '\n' || c == '\r'

    private def step(c: Char): Unit =
      val digit = c >= '0' && c <= '9'
      state = state match
        case 0 =>
          if isXmlSpace(c) then 0
          else if c == '+' then 1
          else if digit then addDigit(c)
          else 5
        case 1 => if digit then addDigit(c) else 5
        case 2 =>
          if digit then addDigit(c)
          else if c == '.' then 3
          else if isXmlSpace(c) then 4
          else 5
        case 3 =>
          if c == '0' then 3
          else if isXmlSpace(c) then 4
          else 5
        case 4 => if isXmlSpace(c) then 4 else 5
        case _ => 5

    private def addDigit(c: Char): Int =
      value = value * 10 + (c - '0')
      if value > Int.MaxValue then 5 else 2
