package com.tjclp.xl.ooxml

/**
 * GH-556: the storage form of formula text — how a cell formula is spelled inside `<f>`.
 *
 * Excel stores functions introduced after Excel 2007 with a `_xlfn.` prefix (`_xlfn.MAXIFS(...)`,
 * `_xlfn.XLOOKUP(...)`), and the dynamic-array pair FILTER/SORT under `_xlfn._xlws.`. A bare
 * `MAXIFS(` in the file is an undefined name to desktop Excel and LibreOffice: the cell shows
 * `#NAME?` as soon as it recalculates, even though the cached value was correct. The model keeps
 * the user-facing bare spelling (what Excel's own formula bar shows); this object maps between the
 * two at the `<f>` boundary in both directions:
 *
 *   - [[toStored]] — model → file: strip the display form's leading '=' (GH-456) and prefix every
 *     call to a function in [[FutureFunctions]] that is not already prefixed.
 *   - [[fromStored]] — file → model: strip `_xlfn.` / `_xlfn._xlws.` from calls to functions in
 *     [[FutureFunctions]]. Prefixes on functions outside the list are kept verbatim so a round trip
 *     never loses a prefix the writer would not restore.
 *   - [[bareFunctionName]] — the parser's rule: any `_xlfn.` / `_xlws.` prefix is dropped before
 *     the registry lookup, so inherited formulas parse whether or not the reader canonicalized
 *     them.
 *
 * The scanner is grammar-light on purpose: it skips string literals (`"..."` with `""` escapes),
 * quoted sheet names (`'...'` with `''` escapes) and bracketed structured/external references
 * (`Table1[...]`, `[1]Sheet!A1`), and rewrites only identifier tokens directly followed by `(`.
 * Everything else is copied byte for byte, so formulas the evaluator cannot parse still get the
 * right storage form.
 */
object FormulaStorage:

  /** The prefix Excel stores in front of post-2007 functions. */
  val XlfnPrefix: String = "_xlfn."

  /** The extra prefix Excel stores in front of the worksheet-scoped dynamic-array functions. */
  val XlwsPrefix: String = "_xlws."

  /**
   * Functions Excel stores as `_xlfn._xlws.NAME` (the rest of [[FutureFunctions]] take `_xlfn.`).
   */
  val WorksheetScoped: Set[String] = Set("FILTER", "SORT")

  /**
   * Upper-case names of every function Excel stores with a `_xlfn.` prefix: Microsoft's published
   * future-function list (Excel 2010 through the 2024 dynamic-array and REGEX additions).
   */
  val FutureFunctions: Set[String] = Set(
    // Excel 2010
    "AGGREGATE",
    "BETA.DIST",
    "BETA.INV",
    "BINOM.DIST",
    "BINOM.INV",
    "CEILING.PRECISE",
    "CHISQ.DIST",
    "CHISQ.DIST.RT",
    "CHISQ.INV",
    "CHISQ.INV.RT",
    "CHISQ.TEST",
    "CONFIDENCE.NORM",
    "CONFIDENCE.T",
    "COVARIANCE.P",
    "COVARIANCE.S",
    "ERF.PRECISE",
    "ERFC.PRECISE",
    "EXPON.DIST",
    "F.DIST",
    "F.DIST.RT",
    "F.INV",
    "F.INV.RT",
    "F.TEST",
    "FLOOR.PRECISE",
    "GAMMA.DIST",
    "GAMMA.INV",
    "GAMMALN.PRECISE",
    "HYPGEOM.DIST",
    "ISO.CEILING",
    "LOGNORM.DIST",
    "LOGNORM.INV",
    "MODE.MULT",
    "MODE.SNGL",
    "NEGBINOM.DIST",
    "NETWORKDAYS.INTL",
    "NORM.DIST",
    "NORM.INV",
    "NORM.S.DIST",
    "NORM.S.INV",
    "PERCENTILE.EXC",
    "PERCENTILE.INC",
    "PERCENTRANK.EXC",
    "PERCENTRANK.INC",
    "POISSON.DIST",
    "QUARTILE.EXC",
    "QUARTILE.INC",
    "RANK.AVG",
    "RANK.EQ",
    "STDEV.P",
    "STDEV.S",
    "T.DIST",
    "T.DIST.2T",
    "T.DIST.RT",
    "T.INV",
    "T.INV.2T",
    "T.TEST",
    "VAR.P",
    "VAR.S",
    "WEIBULL.DIST",
    "WORKDAY.INTL",
    "Z.TEST",
    // Excel 2013
    "ACOT",
    "ACOTH",
    "ARABIC",
    "BASE",
    "BINOM.DIST.RANGE",
    "BITAND",
    "BITLSHIFT",
    "BITOR",
    "BITRSHIFT",
    "BITXOR",
    "CEILING.MATH",
    "COMBINA",
    "COT",
    "COTH",
    "CSC",
    "CSCH",
    "DAYS",
    "DECIMAL",
    "ENCODEURL",
    "FILTERXML",
    "FLOOR.MATH",
    "FORMULATEXT",
    "GAMMA",
    "GAUSS",
    "IFNA",
    "IMCOSH",
    "IMCOT",
    "IMCSC",
    "IMCSCH",
    "IMSEC",
    "IMSECH",
    "IMSINH",
    "IMTAN",
    "ISFORMULA",
    "ISOWEEKNUM",
    "MUNIT",
    "NUMBERVALUE",
    "PDURATION",
    "PERMUTATIONA",
    "PHI",
    "RRI",
    "SEC",
    "SECH",
    "SHEET",
    "SHEETS",
    "SKEW.P",
    "UNICHAR",
    "UNICODE",
    "WEBSERVICE",
    "XOR",
    // Excel 2016
    "CONCAT",
    "FORECAST.ETS",
    "FORECAST.ETS.CONFINT",
    "FORECAST.ETS.SEASONALITY",
    "FORECAST.ETS.STAT",
    "FORECAST.LINEAR",
    "IFS",
    "MAXIFS",
    "MINIFS",
    "SWITCH",
    "TEXTJOIN",
    // Excel 2019 / 365 (dynamic arrays, LET/LAMBDA, XLOOKUP)
    "ANCHORARRAY",
    "BYCOL",
    "BYROW",
    "FIELDVALUE",
    "FILTER",
    "ISOMITTED",
    "LAMBDA",
    "LET",
    "MAKEARRAY",
    "MAP",
    "RANDARRAY",
    "REDUCE",
    "SCAN",
    "SEQUENCE",
    "SINGLE",
    "SORT",
    "SORTBY",
    "STOCKHISTORY",
    "UNIQUE",
    "XLOOKUP",
    "XMATCH",
    // Excel 365 2022+ (text and array-shaping functions, IMAGE)
    "ARRAYTOTEXT",
    "CHOOSECOLS",
    "CHOOSEROWS",
    "DROP",
    "EXPAND",
    "HSTACK",
    "IMAGE",
    "TAKE",
    "TEXTAFTER",
    "TEXTBEFORE",
    "TEXTSPLIT",
    "TOCOL",
    "TOROW",
    "VALUETOTEXT",
    "VSTACK",
    "WRAPCOLS",
    "WRAPROWS",
    // Excel 365 2024+
    "DETECTLANGUAGE",
    "GROUPBY",
    "PERCENTOF",
    "PIVOTBY",
    "PY",
    "REGEXEXTRACT",
    "REGEXREPLACE",
    "REGEXTEST",
    "TRANSLATE",
    "TRIMRANGE"
  )

  /**
   * The storage form of a model formula: no leading '=' (GH-456) and every future-function call
   * carrying Excel's `_xlfn.` (or `_xlfn._xlws.`) prefix. Calls already carrying `_xlfn.` are left
   * alone, so the mapping is idempotent; a call spelled with `_xlws.` alone (which Excel does not
   * resolve) gains the `_xlfn.` in front of it.
   *
   * Byte-identical re-serialization holds for Excel-authored books. A third-party writer that
   * spells FILTER/SORT as plain `_xlfn.FILTER(` (no `_xlws.`) reads back bare and is re-written in
   * Excel's own `_xlfn._xlws.FILTER(` form.
   */
  def toStored(expr: String): String =
    rewriteCalls(expr.stripPrefix("=")) { token =>
      if startsWithIgnoreCase(token, XlfnPrefix) then token
      else if startsWithIgnoreCase(token, XlwsPrefix) then s"$XlfnPrefix$token"
      else if isIn(WorksheetScoped, token) then s"$XlfnPrefix$XlwsPrefix$token"
      else if isIn(FutureFunctions, token) then s"$XlfnPrefix$token"
      else token
    }

  /**
   * The model form of a stored formula: `_xlfn.` / `_xlfn._xlws.` stripped from calls to functions
   * in [[FutureFunctions]]. A prefix on any other function is kept verbatim (the writer would not
   * restore it), so `fromStored` then `toStored` reproduces the file's text.
   */
  def fromStored(text: String): String =
    rewriteCalls(text) { token =>
      if !hasStoragePrefix(token) then token
      else
        val bare = bareFunctionName(token)
        if isIn(FutureFunctions, bare) then bare else token
    }

  /**
   * Drop any `_xlfn.` and `_xlws.` prefixes from a function identifier (case-insensitive). The
   * parser applies this before the registry lookup so `_xlfn.XLOOKUP(` resolves to XLOOKUP. A name
   * without a prefix is returned as-is with no allocation — this sits on the parse hot path.
   */
  def bareFunctionName(name: String): String =
    @annotation.tailrec
    def loop(s: String): String =
      if startsWithIgnoreCase(s, XlfnPrefix) then loop(s.substring(XlfnPrefix.length))
      else if startsWithIgnoreCase(s, XlwsPrefix) then loop(s.substring(XlwsPrefix.length))
      else s
    if hasStoragePrefix(name) then loop(name) else name

  /** Case-insensitive prefix test without allocating (both prefixes are ASCII). */
  private def startsWithIgnoreCase(s: String, prefix: String): Boolean =
    s.regionMatches(true, 0, prefix, 0, prefix.length)

  private def hasStoragePrefix(s: String): Boolean =
    startsWithIgnoreCase(s, XlfnPrefix) || startsWithIgnoreCase(s, XlwsPrefix)

  /**
   * Membership of a token spelled in the author's case. The common all-caps spelling hits (or
   * misses) without allocating; only a token that actually carries lower-case letters is
   * upper-cased for a second look.
   */
  private def isIn(names: Set[String], token: String): Boolean =
    names.contains(token) || (hasLowerCase(token) && names.contains(token.toUpperCase))

  @SuppressWarnings(Array("org.wartremover.warts.Var", "org.wartremover.warts.While"))
  private def hasLowerCase(s: String): Boolean =
    var i = 0
    var found = false
    while !found && i < s.length do
      found = s.charAt(i).isLower
      i += 1
    found

  private def isIdentChar(c: Char): Boolean =
    c.isLetterOrDigit || c == '_' || c == '.'

  private def isIdentStart(c: Char): Boolean =
    c.isLetter || c == '_'

  /**
   * Copy `text`, applying `f` to every identifier token that is directly followed (modulo
   * whitespace) by `(` — i.e. every function call — outside string literals, quoted sheet names,
   * and bracketed references. Inside brackets nothing is interpreted: a structured-reference column
   * name escapes its specials with a single quote (`Table1['#Sales]`), so quote handling there
   * would swallow the rest of the formula.
   */
  @SuppressWarnings(Array("org.wartremover.warts.Var", "org.wartremover.warts.While"))
  private def rewriteCalls(text: String)(f: String => String): String =
    val n = text.length
    val sb = new StringBuilder(n + 16)
    var i = 0
    var bracketDepth = 0
    while i < n do
      val c = text.charAt(i)
      if bracketDepth == 0 && c == '"' then
        // String literal: copy through the closing quote, honoring "" escapes
        val end = closingQuote(text, i, '"')
        sb.append(text.substring(i, end))
        i = end
      else if bracketDepth == 0 && c == '\'' then
        // Quoted sheet name: copy through the closing quote, honoring '' escapes
        val end = closingQuote(text, i, '\'')
        sb.append(text.substring(i, end))
        i = end
      else if c == '[' then
        bracketDepth += 1
        sb.append(c)
        i += 1
      else if c == ']' then
        if bracketDepth > 0 then bracketDepth -= 1
        sb.append(c)
        i += 1
      else if bracketDepth == 0 && isIdentChar(c) then
        var end = i + 1
        while end < n && isIdentChar(text.charAt(end)) do end += 1
        val token = text.substring(i, end)
        var probe = end
        while probe < n && text.charAt(probe).isWhitespace do probe += 1
        val isCall = probe < n && text.charAt(probe) == '(' && isIdentStart(c)
        sb.append(if isCall then f(token) else token)
        i = end
      else
        sb.append(c)
        i += 1
    sb.toString

  /** Index just past the quote closing the literal that opens at `start` (or `n` if unclosed). */
  @SuppressWarnings(Array("org.wartremover.warts.Var", "org.wartremover.warts.While"))
  private def closingQuote(text: String, start: Int, quote: Char): Int =
    val n = text.length
    var i = start + 1
    var end = -1
    while end < 0 && i < n do
      if text.charAt(i) == quote then
        if i + 1 < n && text.charAt(i + 1) == quote then i += 2
        else end = i + 1
      else i += 1
    if end < 0 then n else end
