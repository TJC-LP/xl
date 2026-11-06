package com.tjclp.xl.formula.functions

/**
 * GH-713: the names of the functions that return a reference, as plain constants.
 *
 * A leaf object on purpose: it references nothing, so the parser can consult it (every `:`, union
 * and intersection asks whether an operand is reference-shaped) without forcing `FunctionRegistry`
 * or `FunctionSpecs` to initialize, and no object-initialization cycle can reach it. The drift law
 * in `ReferenceOperatorLawsSpec` ties [[computing]] to the registry: it must equal the names of
 * every registered `FunctionSpec.Referencing`, so a new reference-computing spec fails that test
 * until it is added here.
 */
private[formula] object ReferenceFunctionNames:

  /**
   * Excel's functions that return the reference they select. They evaluate their branches lazily
   * rather than through `FunctionSpec.referencing`, so no spec marks them.
   */
  val selecting: Set[String] = Set("IF", "IFS", "CHOOSE", "SWITCH")

  /** The registered functions that compute a reference: every `FunctionSpec.Referencing`. */
  val computing: Set[String] = Set("OFFSET", "INDIRECT", "INDEX", "XLOOKUP")

  /**
   * GH-723: a lexical pre-filter for calls carried in computed range slots. Ignore literals, quoted
   * sheets and bracket payloads; a substring inside them is no call. Use original-text indices:
   * Unicode upper-casing can change the length of a token.
   */
  def mayContainComputingCall(text: String, upperCased: String): Boolean =
    @annotation.tailrec
    def quoted(i: Int, quote: Char): Int =
      if i >= text.length then i
      else if text.charAt(i) == quote then
        if text.lift(i + 1).contains(quote) then quoted(i + 2, quote) else i + 1
      else quoted(i + 1, quote)
    @annotation.tailrec
    def bracket(i: Int, depth: Int): Int =
      if i >= text.length || depth == 0 then i
      else
        text.charAt(i) match
          case '\'' => bracket(Math.min(i + 2, text.length), depth)
          case '[' => bracket(i + 1, depth + 1)
          case ']' => bracket(i + 1, depth - 1)
          case _ => bracket(i + 1, depth)
    @annotation.tailrec
    def beforeSpace(i: Int): Int =
      if i > 0 && text.charAt(i - 1).isWhitespace then beforeSpace(i - 1) else i
    @annotation.tailrec
    def beforeName(i: Int): Int =
      if i > 0 && com.tjclp.xl.formula.parser.ReferenceChars.isFunctionName(text.charAt(i - 1))
      then beforeName(i - 1)
      else i
    @annotation.tailrec
    def scan(i: Int): Boolean =
      if i >= text.length then false
      else
        text.charAt(i) match
          case '"' => scan(quoted(i + 1, '"'))
          case '\'' => scan(quoted(i + 1, '\''))
          case '[' => scan(bracket(i + 1, 1))
          case '(' =>
            val end = beforeSpace(i)
            val name = text
              .substring(beforeName(end), end)
              .toUpperCase(java.util.Locale.ROOT)
              .stripPrefix("_XLFN.")
              .stripPrefix("_XLWS.")
            computing.contains(name) || scan(i + 1)
          case _ => scan(i + 1)
    computing.exists(upperCased.contains) && scan(0)
