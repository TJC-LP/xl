package com.tjclp.xl.formula

import com.tjclp.xl.addressing.{ARef, Column, Row, SheetName}
import com.tjclp.xl.cells.CellError
import org.scalacheck.Gen

/**
 * Text-level formula generators shared by FormulaParserSpec (GH-612) and FormulaGrammarSpec
 * (GH-653). They produce SOURCE text, never ASTs: the parser decides the tree and the printer
 * decides every paren, so a property over them exercises the grammar as Excel writes it rather than
 * the shapes the AST already models.
 */
private[formula] object FormulaTextGens:

  // ==================== References and ranges (hoisted from FormulaParserSpec, GH-612) ====================

  val genColumnLetters: Gen[String] =
    Gen
      .frequency(8 -> Gen.choose(0, 30), 2 -> Gen.choose(0, Column.MaxIndex0))
      .map(i => Column.from0(i).toLetter)

  val genRowNumber: Gen[String] =
    Gen
      .frequency(8 -> Gen.choose(1, 40), 2 -> Gen.choose(1, Row.MaxIndex0 + 1))
      .map(_.toString)

  val genDollar: Gen[String] = Gen.oneOf("", "$")

  /** Two column letters in grid order (the parser normalizes, so the source must already be). */
  val genFullColumnText: Gen[String] =
    for
      a <- genColumnLetters
      b <- genColumnLetters
      d1 <- genDollar
      d2 <- genDollar
    yield
      val (lo, hi) =
        if Column.fromLetter(a).exists(x => Column.fromLetter(b).exists(y => x.index0 <= y.index0))
        then (a, b)
        else (b, a)
      s"$d1$lo:$d2$hi"

  val genFullRowText: Gen[String] =
    for
      a <- genRowNumber
      b <- genRowNumber
      d1 <- genDollar
      d2 <- genDollar
    yield
      val (lo, hi) = if a.toInt <= b.toInt then (a, b) else (b, a)
      s"$d1$lo:$d2$hi"

  val genBoundedRangeText: Gen[String] =
    for
      c1 <- Gen.choose(0, 30)
      c2 <- Gen.choose(0, 30)
      r1 <- Gen.choose(1, 40)
      r2 <- Gen.choose(1, 40)
      d1 <- genDollar
      d2 <- genDollar
      d3 <- genDollar
      d4 <- genDollar
    yield
      val (cLo, cHi) = (math.min(c1, c2), math.max(c1, c2))
      val (rLo, rHi) = (math.min(r1, r2), math.max(r1, r2))
      s"$d1${Column.from0(cLo).toLetter}$d2$rLo:$d3${Column.from0(cHi).toLetter}$d4$rHi"

  /** A range in one of Excel's three spellings: whole columns, whole rows, or two corners. */
  val genRangeText: Gen[String] =
    Gen.oneOf(genFullColumnText, genFullRowText, genBoundedRangeText)

  /** One cell reference under each of the four anchors (`A1`, `$A1`, `A$1`, `$A$1`). */
  val genRefText: Gen[String] =
    for
      col <- genColumnLetters
      row <- genRowNumber
      d1 <- genDollar
      d2 <- genDollar
    yield s"$d1$col$d2$row"

  // ==================== Sheet qualifiers at the quoting boundary (GH-263) ====================

  private val genWordChar: Gen[Char] =
    Gen.frequency(8 -> Gen.alphaNumChar, 1 -> Gen.const('_'))

  /** Characters a sheet name may carry (Excel forbids only `: \ / ? * [ ]`). */
  private val genSheetChar: Gen[Char] =
    Gen.frequency(
      6 -> Gen.alphaNumChar,
      2 -> Gen.const(' '),
      1 -> Gen.oneOf('-', '&', '.', '(', ')', '+', ',', '%', '#', '@', '=', '<', '>', '!', '"'),
      1 -> Gen.const('\''),
      1 -> Gen.const('_'),
      1 -> Gen.oneOf('é', 'ß', '日', 'Ω', 'ø')
    )

  private def wordOf(first: Gen[Char], maxLen: Int): Gen[String] =
    for
      head <- first
      len <- Gen.choose(0, maxLen - 1)
      tail <- Gen.listOfN(len, genWordChar)
    yield (head :: tail).mkString

  /**
   * A valid sheet name (`SheetName.apply` accepts it) drawn from both sides of
   * `SheetName.needsQuoting`: plain identifiers (bare), function-named sheets (bare), names with
   * spaces, punctuation, apostrophes and non-ASCII letters, digit-first names, A1- and R1C1-shaped
   * names, `TRUE`/`FALSE`, and names at the 31-character limit.
   */
  val genSheetNameText: Gen[String] =
    val identifier = wordOf(Gen.alphaChar, 12)
    val functionNamed = Gen.oneOf("SUM", "NOT", "IF", "LET", "AND", "OR", "TRUE1", "N", "T")
    val refShaped =
      for
        col <- genColumnLetters
        row <- genRowNumber
      yield s"$col$row"
    val r1c1Shaped = Gen.oneOf("R1C1", "RC", "r2c", "R10C", "RC3")
    val boolNamed = Gen.oneOf("TRUE", "FALSE", "true", "False")
    val digitFirst =
      for
        d <- Gen.numChar
        rest <- wordOf(Gen.alphaNumChar, 8)
      yield s"$d$rest"
    val withSpecials =
      for
        len <- Gen.choose(1, 31)
        chars <- Gen.listOfN(len, genSheetChar)
      yield chars.mkString
    val apostrophes = Gen.oneOf("O'Brien", "Bob's Data", "'", "''", "Q1 'Actuals'")
    val edges = Gen.oneOf(" Lead", "Trail ", "A" * 31, "x", "_", "a b", "Sales & Marketing")
    Gen
      .frequency(
        4 -> identifier,
        1 -> functionNamed,
        2 -> refShaped,
        1 -> r1c1Shaped,
        1 -> boolNamed,
        1 -> digitFirst,
        3 -> withSpecials,
        1 -> apostrophes,
        1 -> edges
      )
      // `history` is the one reserved sheet name; every other product of the alphabet above
      // is a valid SheetName (no `: \ / ? * [ ]`, 1..31 chars)
      .map(name => if name.equalsIgnoreCase("history") then name + "2" else name)

  /** `Sheet1!`, `'Q1 Data'!`, `'O''Brien'!` — a sheet name rendered as the printer renders it. */
  val genSheetQualifier: Gen[String] =
    genSheetNameText.map(name => s"${SheetName.quoteForFormula(name)}!")

  /**
   * The name of an external workbook's sheet (or the workbook file itself): identifiers, dotted
   * file names, digit-first names, and names that force the quoted form.
   */
  val genExternalNameText: Gen[String] =
    Gen.frequency(
      3 -> wordOf(Gen.alphaChar, 10),
      2 -> wordOf(Gen.alphaChar, 8).map(_ + ".xlsx"),
      1 -> Gen.oneOf("2024", "Q1.Data.xlsx", "A1", "Book1", "Sheet 1", "Sales & Marketing.xlsx"),
      1 -> Gen.oneOf("O'Brien.xlsx", "Q1 Data", "Consolidation (final).xlsx", "Données"),
      1 -> (for
        a <- wordOf(Gen.alphaChar, 6)
        b <- wordOf(Gen.alphaChar, 6)
      yield s"$a $b")
    )

  /**
   * `[2]Book1!` or `'[3]Sheet Name'!` — the external-workbook qualifier, quoted exactly when the
   * printer quotes it (FormulaPrinter.formatExternalSheet: any character outside letters, digits,
   * `_` and `.`), with the bracket INSIDE the quotes and apostrophes doubled.
   */
  val genExternalQualifier: Gen[String] =
    for
      index <- Gen.frequency(8 -> Gen.choose(1, 9), 2 -> Gen.choose(10, 999))
      name <- genExternalNameText
    yield
      val quoted = name.exists(c => !c.isLetterOrDigit && c != '_' && c != '.')
      if quoted then s"'[$index]${name.replace("'", "''")}'!" else s"[$index]$name!"

  /** No qualifier (local), a sheet qualifier, or an external-workbook qualifier. */
  val genQualifier: Gen[String] =
    Gen.frequency(4 -> Gen.const(""), 4 -> genSheetQualifier, 2 -> genExternalQualifier)

  /** No qualifier or a sheet qualifier — the two a defined name may carry. */
  val genLocalOrSheetQualifier: Gen[String] =
    Gen.frequency(1 -> Gen.const(""), 1 -> genSheetQualifier)

  // ==================== Error literals ====================

  private def randomCase(text: String): Gen[String] =
    Gen.listOfN(text.length, Gen.oneOf(true, false)).map { flips =>
      text.zip(flips).map((c, upper) => if upper then c.toUpper else c.toLower).mkString
    }

  /**
   * Every Excel error code as (written, canonical): the written form varies the letter case (the
   * parser matches codes case-insensitively), the canonical form is `CellError.toExcel`.
   */
  val genErrorLitText: Gen[(String, String)] =
    for
      error <- Gen.oneOf(CellError.values.toIndexedSeq)
      canonical = error.toExcel
      written <- Gen.frequency(2 -> Gen.const(canonical), 1 -> randomCase(canonical))
    yield (written, canonical)

  // ==================== Defined names ====================

  /** The identifiers the parser reads as literals or operator keywords, never as names. */
  val ReservedNames: Set[String] = Set("TRUE", "FALSE", "AND", "OR", "NOT")

  /**
   * A defined-name identifier: starts with a letter, `_` or `\`, continues with letters, digits,
   * `_`, `.` and `\` (Excel's name characters, GH-394), is not cell-ref shaped and is not a
   * reserved word. Constructive: a ref-shaped or reserved product gets a `_n` suffix rather than
   * being discarded, so the generator never starves.
   */
  val genNameText: Gen[String] =
    val first = Gen.frequency(6 -> Gen.alphaChar, 2 -> Gen.const('_'), 1 -> Gen.const('\\'))
    val rest =
      Gen.frequency(
        6 -> Gen.alphaNumChar,
        1 -> Gen.const('_'),
        1 -> Gen.const('.'),
        1 -> Gen.const('\\')
      )
    val fresh =
      for
        head <- first
        len <- Gen.choose(0, 12)
        tail <- Gen.listOfN(len, rest)
      yield (head :: tail).mkString
    val idiomatic = Gen.oneOf(
      "case",
      "entry_mult",
      "ltm_ebitda",
      "Sales.Total",
      "_tmp",
      "\\name",
      "R1C1",
      "RC",
      "ABCD1",
      "rev_range",
      "x",
      "sum",
      "Max",
      "let",
      "note"
    )
    Gen.frequency(4 -> fresh, 1 -> idiomatic).map(nameFrom)

  /** A name-shaped identifier that is a cell reference or a reserved word is made a name. */
  def nameFrom(identifier: String): String =
    if ARef.parse(identifier).isRight || ReservedNames.contains(identifier.toUpperCase) then
      s"${identifier}_n"
    else identifier

  // ==================== String literals ====================

  private val genStringChar: Gen[Char] =
    Gen.frequency(
      8 -> Gen.alphaNumChar,
      3 -> Gen.const(' '),
      2 -> Gen.const('"'),
      1 -> Gen.const('\''),
      2 -> Gen.oneOf(',', '(', ')', '!', '#', '@', '[', ']', '{', '}', ';', ':', '$', '%'),
      1 -> Gen.oneOf('&', '=', '<', '>', '+', '-', '*', '/', '^', '\\', '.'),
      1 -> Gen.oneOf('é', 'ß', '日', 'Ω', '€', '\n', '\t')
    )

  /**
   * A string literal's VALUE, over an alphabet that includes the quote (doubled on the way out),
   * the argument separator, parens, the formula operators, whitespace and non-ASCII letters.
   */
  val genStringValue: Gen[String] =
    Gen.frequency(
      6 -> Gen.choose(0, 12).flatMap(n => Gen.listOfN(n, genStringChar).map(_.mkString)),
      1 -> Gen.oneOf(
        "",
        "\"",
        "\"\"",
        "a\"b",
        "Sheet1!A1",
        "SUM(A1)",
        "{1,2}",
        "Table1[Col]",
        "#N/A"
      )
    )

  /** The source spelling of a string literal: `"` + value with every `"` doubled + `"`. */
  def stringLiteral(value: String): String = s""""${value.replace("\"", "\"\"")}""""

  val genStringLitText: Gen[String] = genStringValue.map(stringLiteral)

  // ==================== Numbers ====================

  /**
   * A numeric literal in the spellings the parser accepts: integers, decimals, a leading dot,
   * trailing zeros and exponents with and without a sign. Only integers are canonical (the printer
   * re-spells the rest via BigDecimal.toString), so text-identity properties use `genIntText`.
   */
  val genNumberText: Gen[String] =
    Gen.frequency(
      5 -> Gen.choose(0, 999).map(_.toString),
      2 -> Gen.choose(0, 99999).map(n => s"${n / 100}.${"%02d".format(n % 100)}"),
      1 -> Gen.oneOf(".5", ".25", "0.10", "1.50", "5."),
      1 -> Gen.oneOf("1.5E10", "2.3E-7", "1e3", "1E+2", "4E0", "1.2e+1"),
      1 -> Gen.oneOf("0", "1048576", "16384", "9007199254740993", "0.000001")
    )

  val genIntText: Gen[String] = Gen.choose(0, 999).map(_.toString)
