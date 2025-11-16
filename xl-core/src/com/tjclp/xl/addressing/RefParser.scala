package com.tjclp.xl.addressing

import scala.util.boundary, boundary.break

/** Unified reference parser shared by macros and runtime code. */
object RefParser:

  enum ParsedRef:
    case Cell(sheet: Option[String], col0: Int, row0: Int)
    case Range(sheet: Option[String], cs: Int, rs: Int, ce: Int, re: Int)

  /** Parse any Excel reference string into an unescaped structure. */
  def parse(input: String): Either[String, ParsedRef] =
    if input.isEmpty then Left("Empty reference")
    else
      try Right(parseUnsafe(input))
      catch case e: IllegalArgumentException => Left(e.getMessage)

  private def parseUnsafe(s: String): ParsedRef =
    val bangIdx = findUnquotedBang(s)
    if bangIdx < 0 then
      if s.contains(':') then
        val ((cs, rs), (ce, re)) = parseRangeLit(s)
        ParsedRef.Range(None, cs, rs, ce, re)
      else
        val (c0, r0) = parseCellLit(s)
        ParsedRef.Cell(None, c0, r0)
    else
      val sheetPart = s.substring(0, bangIdx)
      val refPart = s.substring(bangIdx + 1)
      if refPart.isEmpty then fail(s"Missing reference after '!' in: $s")
      val sheetName = parseSheetName(sheetPart)
      if refPart.contains(':') then
        val ((cs, rs), (ce, re)) = parseRangeLit(refPart)
        ParsedRef.Range(Some(sheetName), cs, rs, ce, re)
      else
        val (c0, r0) = parseCellLit(refPart)
        ParsedRef.Cell(Some(sheetName), c0, r0)

  private def parseSheetName(part: String): String =
    if part.startsWith("'") then
      if !part.endsWith("'") then
        fail(s"Unbalanced quotes in sheet name: $part (missing closing quote)")
      val quoted = part.substring(1, part.length - 1)
      if quoted.isEmpty then fail("Empty sheet name in quotes")
      val unescaped = quoted.replace("''", "'")
      validateSheetName(unescaped)
      unescaped
    else
      if part.contains("'") then
        fail(s"Misplaced quote in sheet name: $part (quotes must wrap entire name)")
      validateSheetName(part)
      part

  @SuppressWarnings(Array("org.wartremover.warts.Var", "org.wartremover.warts.While"))
  private def findUnquotedBang(s: String): Int =
    boundary:
      var i = 0
      var inQuote = false
      while i < s.length do
        val c = s.charAt(i)
        if c == '\'' then inQuote = !inQuote
        else if c == '!' && !inQuote then break(i)
        i += 1
      -1

  private def validateSheetName(name: String): Unit =
    if name.isEmpty then fail("sheet name cannot be empty")
    if name.length > 31 then fail("sheet name max length is 31 chars")
    val invalid = Set(':', '\\', '/', '?', '*', '[', ']')
    name.foreach { c =>
      if invalid.contains(c) then fail(s"sheet name cannot contain: $c")
    }

  @SuppressWarnings(Array("org.wartremover.warts.Var", "org.wartremover.warts.While"))
  private def parseCellLit(s: String): (Int, Int) =
    var i = 0; val n = s.length
    if n == 0 then fail("empty")
    var col = 0
    var parsing = true
    while i < n && parsing do
      val ch = s.charAt(i)
      if ch >= 'A' && ch <= 'Z' then { col = col * 26 + (ch - 'A' + 1); i += 1 }
      else if ch >= 'a' && ch <= 'z' then { col = col * 26 + (ch - 'a' + 1); i += 1 }
      else parsing = false
    if col == 0 then fail("missing column letters")
    var row = 0; var sawDigit = false
    while i < n && s.charAt(i).isDigit do
      row = row * 10 + (s.charAt(i) - '0'); i += 1; sawDigit = true
    if !sawDigit then fail("missing row digits")
    if i != n then fail("trailing junk")
    if row < 1 then fail("row must be ≥ 1")
    (col - 1, row - 1)

  private def parseRangeLit(s: String): ((Int, Int), (Int, Int)) =
    val i = s.indexOf(':')
    if i <= 0 || i >= s.length - 1 then fail("use A1:B2 form")
    val (a, b) = (s.substring(0, i), s.substring(i + 1))
    val (c1, r1) = parseCellLit(a)
    val (c2, r2) = parseCellLit(b)
    val (cs, rs) = (math.min(c1, c2), math.min(r1, r2))
    val (ce, re) = (math.max(c1, c2), math.max(r1, r2))
    ((cs, rs), (ce, re))

  private def fail(msg: String): Nothing = throw new IllegalArgumentException(msg)

end RefParser
