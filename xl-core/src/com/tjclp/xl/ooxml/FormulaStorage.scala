package com.tjclp.xl.ooxml

import java.util.Locale

/**
 * GH-556: the storage form of formula text — how a cell formula is spelled inside `<f>`.
 *
 * Excel stores functions introduced after Excel 2007 with a `_xlfn.` prefix (`_xlfn.MAXIFS(...)`,
 * `_xlfn.XLOOKUP(...)`), the dynamic-array pair FILTER/SORT under `_xlfn._xlws.`, and the parameter
 * names of LET and LAMBDA — each declaration and each reference — under `_xlpm.`
 * (`_xlfn.LET(_xlpm.x,1,_xlpm.x+1)`). A bare `MAXIFS(` in the file is an undefined name to desktop
 * Excel and LibreOffice: the cell shows `#NAME?` as soon as it recalculates, even though the cached
 * value was correct; a LET whose names lack `_xlpm.` is reported as unreadable content on open. The
 * model keeps the user-facing bare spelling (what Excel's own formula bar shows); this object maps
 * between the two at the `<f>` boundary in both directions:
 *
 *   - [[toStored]] — model → file: strip the display form's leading '=' (GH-456), prefix every call
 *     to a function in [[FutureFunctions]] that is not already prefixed, and prefix every
 *     LET/LAMBDA parameter (declaration or reference) that is not already prefixed.
 *   - [[fromStored]] — file → model: strip `_xlfn.` / `_xlfn._xlws.` from calls to functions in
 *     [[FutureFunctions]] and `_xlpm.` from LET/LAMBDA parameters. A prefix the writer would not
 *     restore — on a function outside the list, or `_xlpm.` on a token outside a recognized
 *     parameter position — is kept verbatim, so a round trip never loses one.
 *   - [[bareFunctionName]] — the parser's rule: any `_xlfn.` / `_xlws.` prefix is dropped before
 *     the registry lookup, so inherited formulas parse whether or not the reader canonicalized
 *     them.
 *   - [[bareFutureCalls]] — the lint's rule (`xlfn-missing`, GH-577): the calls in a stored formula
 *     that [[toStored]] would still prefix, observed by the same scanner.
 *
 * The scanner is grammar-light on purpose: it skips string literals (`"..."` with `""` escapes),
 * quoted sheet names (`'...'` with `''` escapes), bracketed structured/external references
 * (`Table1[...]`, `[1]Sheet!A1`, with their `'`-escaped specials), array constants (`{...}`) and
 * error literals (`#N/A`), and rewrites only identifier tokens directly followed by `(` — plus,
 * inside a LET or LAMBDA call, the identifiers at parameter positions and the references to them.
 * Everything else is copied byte for byte, so formulas the evaluator cannot parse still get the
 * right storage form.
 *
 * Both mappings return their input untouched, without scanning or allocating, on the common miss: a
 * formula without a call for [[toStored]], a formula without `_xl` anywhere for [[fromStored]].
 */
object FormulaStorage:

  /** The prefix Excel stores in front of post-2007 functions. */
  val XlfnPrefix: String = "_xlfn."

  /** The extra prefix Excel stores in front of the worksheet-scoped dynamic-array functions. */
  val XlwsPrefix: String = "_xlws."

  /** The prefix Excel stores in front of every LET / LAMBDA parameter name and reference. */
  val XlpmPrefix: String = "_xlpm."

  /**
   * Functions Excel stores as `_xlfn._xlws.NAME` (the rest of [[FutureFunctions]] take `_xlfn.`).
   */
  val WorksheetScoped: Set[String] = Set("FILTER", "SORT")

  /**
   * Functions whose parameter names Excel stores as `_xlpm.NAME`: LET (a name at every even
   * argument position, each followed by its value, the last argument being the calculation) and
   * LAMBDA (every argument but the last, an optional one in brackets: `[y]`). A reference to a
   * declared name anywhere inside the call carries the prefix too.
   */
  val ParameterScoped: Set[String] = Set("LET", "LAMBDA")

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
   * The storage form of a model formula: no leading '=' (GH-456), every future-function call
   * carrying Excel's `_xlfn.` (or `_xlfn._xlws.`) prefix, and every LET/LAMBDA parameter carrying
   * `_xlpm.`. Tokens already carrying their prefix are left alone, so the mapping is idempotent; a
   * call spelled with `_xlws.` alone (which Excel does not resolve) gains the `_xlfn.` in front of
   * it, and an openpyxl-style `_xlfn.LET(x,1,x+1)` gains the `_xlpm.` on its parameters.
   *
   * Byte-identical re-serialization holds for Excel-authored books. A third-party writer that
   * spells FILTER/SORT as plain `_xlfn.FILTER(` (no `_xlws.`) reads back bare and is re-written in
   * Excel's own `_xlfn._xlws.FILTER(` form.
   */
  def toStored(expr: String): String =
    val bare = expr.stripPrefix("=")
    // Fast path: without a '(' there is no call and no LET/LAMBDA scope, and without an '@' no
    // implicit intersection (GH-604) — nothing to prefix
    if bare.indexOf('(') < 0 && bare.indexOf('@') < 0 then bare
    else
      rewriteCalls(wrapIntersections(bare, () => ()))(storedCall, (token, _) => storedParam(token))

  /**
   * GH-577: the upper-case bare names of the calls in a stored formula whose spelling [[toStored]]
   * would change — a [[FutureFunctions]] call without `_xlfn.` (or with `_xlws.` alone, which Excel
   * does not resolve), a LET / LAMBDA whose parameters lack `_xlpm.` (the openpyxl-style
   * `_xlfn.LET(x,1,x+1)`), and the token `@` for an implicit intersection stored bare (GH-604) —
   * distinct, in order of first appearance. This runs the SAME scanner as [[toStored]] with
   * recording callbacks in place of the rewriting ones, so `bareFutureCalls(text).isEmpty` holds
   * exactly when `toStored(text)` returns `text` (minus a leading '=') unchanged: a lint built on
   * it cannot disagree with the writer. Excel and LibreOffice show `#NAME?` for a bare future
   * function on the first recalculation and report a LET / LAMBDA without `_xlpm.` as unreadable
   * content on open.
   */
  @SuppressWarnings(Array("org.wartremover.warts.Var"))
  def bareFutureCalls(text: String): Vector[String] =
    // Same fast path as toStored: without a '(' or an '@' there is nothing the writer would change
    if text.indexOf('(') < 0 && text.indexOf('@') < 0 then Vector.empty
    else
      var found = Vector.empty[String]
      def record(name: String): Unit = if !found.contains(name) then found = found :+ name
      // GH-604/GH-654: a bare `@x` in a file (Excel stores `_xlfn.SINGLE(x)`) is recorded as the
      // token the file holds, `@` — the lint names what the user can find, not a call that is
      // not there
      val wrapped = wrapIntersections(text, () => record("@"))
      rewriteCalls(wrapped)(
        token =>
          if storedCall(token) != token then
            record(bareFunctionName(token).toUpperCase(Locale.ROOT))
          token
        ,
        (token, isLet) =>
          if storedParam(token) != token then record(if isLet then "LET" else "LAMBDA")
          token
      )
      found

  /**
   * The model form of a stored formula: `_xlfn.` / `_xlfn._xlws.` stripped from calls to functions
   * in [[FutureFunctions]], `_xlpm.` stripped from LET/LAMBDA parameters. A prefix on any other
   * function, or a `_xlpm.` outside a recognized parameter position, is kept verbatim (the writer
   * would not restore it), so `fromStored` then `toStored` reproduces the file's text.
   */
  @SuppressWarnings(Array("org.wartremover.warts.Var"))
  def fromStored(text: String): String =
    // Fast path: every storage prefix starts with "_xl"; a formula without it is already bare
    if !containsStoragePrefix(text) then text
    else
      // GH-654: the unwrap pass runs only when the call scan met a SINGLE — every other prefixed
      // formula (a 1M-row book of _xlfn.XLOOKUP cells on the SAX read path) pays one scan, not two
      var sawSingle = false
      val bare = rewriteCalls(text)(
        token =>
          if !sawSingle && bareFunctionName(token).equalsIgnoreCase("SINGLE") then sawSingle = true
          modelCall(token)
        ,
        (token, _) => modelParam(token)
      )
      if sawSingle then unwrapIntersections(bare) else bare

  // ===== GH-604: the implicit-intersection operator =====
  //
  // Excel 365 stores `@x` (implicit intersection — the formula bar's spelling) as
  // `_xlfn.SINGLE(x)`. The model keeps the formula-bar form: [[toStored]] wraps every `@` operand
  // in `_xlfn.SINGLE(...)` and [[fromStored]] unwraps every one-argument `SINGLE(...)` call back to
  // `@`. The operand of `@` is one primary, exactly what the parser's `@` arm accepts: a reference
  // or name (optionally sheet-qualified, quoted or external), a function call, a structured
  // reference, a string / number / boolean / error literal, another `@`, or a parenthesized
  // expression, optionally after whitespace — so `@acq`, `@'M&A'!I12:I40`, `@INDEX(rng,,2)`,
  // `@"x"`, `@@A1:A3` and `@(A1:A3*2)` all wrap, and the parens of the last are the call's own:
  // `_xlfn.SINGLE(A1:A3*2)`. Both directions rescan the operand, so a nested `@` / `SINGLE(`
  // inside it (`@INDEX(@A1:A3,1)`) is rewritten too (GH-654). An `@` inside a structured reference
  // (`Table1[@Col]`), a string, or one with no operand after it is copied verbatim.

  /**
   * Wrap every `@operand` as `_xlfn.SINGLE(operand)`, calling `onWrap` once per rewrite. The
   * operand is rescanned rather than copied, so a nested `@` wraps too; whitespace between `@` and
   * its operand is dropped (`@ A1` stores as `_xlfn.SINGLE(A1)`, the parser's own reading).
   */
  @SuppressWarnings(Array("org.wartremover.warts.Var", "org.wartremover.warts.While"))
  private def wrapIntersections(text: String, onWrap: () => Unit): String =
    if text.indexOf('@') < 0 then text
    else
      val n = text.length
      val sb = new java.lang.StringBuilder(n + 16)
      var i = 0
      var bracketDepth = 0
      // indices at which a wrapping `)` is due — innermost operand first (nested ends never cross)
      var pendingClose: List[Int] = Nil
      def closeDue(): Unit =
        while pendingClose.headOption.contains(i) do
          sb.append(')')
          pendingClose = pendingClose.drop(1)
      while i < n do
        closeDue()
        val c = text.charAt(i)
        if bracketDepth == 0 && (c == '"' || c == '\'') then
          val end = closingQuote(text, i, c)
          sb.append(text, i, end)
          i = end
        else if bracketDepth > 0 && c == '\'' then
          val end = math.min(i + 2, n)
          sb.append(text, i, end)
          i = end
        else if c == '[' then
          bracketDepth += 1
          sb.append(c)
          i += 1
        else if c == ']' then
          if bracketDepth > 0 then bracketDepth -= 1
          sb.append(c)
          i += 1
        else if bracketDepth == 0 && c == '@' then
          val start = skipSpaces(text, i + 1)
          val end = intersectionOperandEnd(text, start)
          if end > start then
            onWrap()
            sb.append(XlfnPrefix).append("SINGLE(")
            if text.charAt(start) == '(' then
              // `@(expr)` — the parens become the call's own: skip the '(' and let the ')' at
              // end - 1 close the call as the scan reaches it
              i = start + 1
            else
              pendingClose = end :: pendingClose
              i = start
          else
            sb.append(c)
            i += 1
        else
          sb.append(c)
          i += 1
      // every wrap opened is closed here at the latest — balanced output is a property of this
      // loop, not of the operand scanner (closeDue() has already fired for i == n)
      while pendingClose.nonEmpty do
        sb.append(')')
        pendingClose = pendingClose.drop(1)
      sb.toString

  /**
   * Rewrite every one-argument `SINGLE(x)` call (any storage prefix already stripped) to `@x`. The
   * argument is rescanned rather than copied, so a nested `SINGLE(` unwraps too.
   */
  @SuppressWarnings(Array("org.wartremover.warts.Var", "org.wartremover.warts.While"))
  private def unwrapIntersections(text: String): String =
    val n = text.length
    val sb = new java.lang.StringBuilder(n)
    var i = 0
    var bracketDepth = 0
    // indices of the `)` closing an unwrapped call whose operand needs no parens — dropped as the
    // scan reaches them (innermost first; two calls never close on the same character)
    var dropClose: List[Int] = Nil
    while i < n do
      val c = text.charAt(i)
      if dropClose.headOption.contains(i) then
        dropClose = dropClose.drop(1)
        i += 1
      else if bracketDepth == 0 && (c == '"' || c == '\'') then
        val end = closingQuote(text, i, c)
        sb.append(text, i, end)
        i = end
      else if bracketDepth > 0 && c == '\'' then
        val end = math.min(i + 2, n)
        sb.append(text, i, end)
        i = end
      else if c == '[' then
        bracketDepth += 1
        sb.append(c)
        i += 1
      else if c == ']' then
        if bracketDepth > 0 then bracketDepth -= 1
        sb.append(c)
        i += 1
      else if bracketDepth == 0 && isIdentStart(c) && (i == 0 || !isIdentChar(text.charAt(i - 1)))
        && text.regionMatches(true, i, "SINGLE(", 0, 7)
      then
        val open = i + 6
        val close = closingParen(text, open)
        val inner = if close < 0 then "" else text.substring(open + 1, close - 1)
        if close > 0 && singleArgument(inner) then
          val simple = inner.nonEmpty && intersectionOperandEnd(inner, 0) == inner.length &&
            inner.charAt(0) != '('
          sb.append('@')
          // the operand is rescanned from open + 1: a simple operand sheds the call's parens (the
          // `)` at close - 1 is dropped on arrival), any other keeps them as its grouping
          if simple then dropClose = (close - 1) :: dropClose else sb.append('(')
          i = open + 1
        else
          sb.append(text, i, open + 1)
          i = open + 1
      else
        sb.append(c)
        i += 1
    sb.toString

  /**
   * Index just past the operand of an `@` starting at `start`: a parenthesized expression, a string
   * or error literal, another `@` and its operand, or a reference / name / number (quoted-sheet,
   * unquoted-sheet or external-workbook qualified) optionally followed by a call's argument list or
   * a structured reference's brackets. `start` when nothing operand-shaped follows.
   */
  @SuppressWarnings(Array("org.wartremover.warts.Var", "org.wartremover.warts.While"))
  private def intersectionOperandEnd(text: String, start: Int): Int =
    val n = text.length
    if start >= n then start
    else
      val c = text.charAt(start)
      def orStart(end: Int): Int = if end < 0 then start else end
      if c == '(' then orStart(closingParen(text, start))
      else if c == '"' then closingQuote(text, start, '"')
      else if c == '#' then
        val end = errorLiteralEnd(text, start)
        if end > start + 1 then end else start
      else if c == '@' then
        // GH-654: `@@A1:A3` — the outer operand ends where the inner one does
        val innerStart = skipSpaces(text, start + 1)
        val end = intersectionOperandEnd(text, innerStart)
        if end > innerStart then end else start
      else
        // qualifier: 'quoted sheet'! or [n]Sheet!
        val afterQualifier =
          if c == '\'' then
            val q = closingQuote(text, start, '\'')
            if q < n && text.charAt(q) == '!' then q + 1 else q
          else if c == '[' then closingBracket(text, start)
          else start
        if afterQualifier < 0 then start
        else
          // the reference / name body (Sheet1!A1:B2, $A$1, name.part, A:A)
          var i = afterQualifier
          while i < n && (isIdentChar(text.charAt(i)) || text.charAt(i) == '!' ||
              text.charAt(i) == ':')
          do i += 1
          if i == start then start
          else if i < n && text.charAt(i) == '(' then orStart(closingParen(text, i))
          else if i < n && text.charAt(i) == '[' then orStart(closingBracket(text, i))
          else i

  /** True when `inner` (a call's argument text) holds exactly one argument: no top-level comma. */
  @SuppressWarnings(Array("org.wartremover.warts.Var", "org.wartremover.warts.While"))
  private def singleArgument(inner: String): Boolean =
    val n = inner.length
    var i = 0
    var depth = 0
    var bracketDepth = 0
    var single = true
    while single && i < n do
      val c = inner.charAt(i)
      if bracketDepth == 0 && (c == '"' || c == '\'') then i = closingQuote(inner, i, c)
      else
        if c == '[' then bracketDepth += 1
        else if c == ']' then { if bracketDepth > 0 then bracketDepth -= 1 }
        else if bracketDepth == 0 && (c == '(' || c == '{') then depth += 1
        else if bracketDepth == 0 && (c == ')' || c == '}') then depth -= 1
        else if bracketDepth == 0 && depth == 0 && c == ',' then single = false
        i += 1
    single

  /**
   * Index just past the `)` matching the `(` at `open` (strings and quoted names skipped); -1 when
   * unbalanced.
   */
  @SuppressWarnings(Array("org.wartremover.warts.Var", "org.wartremover.warts.While"))
  private def closingParen(text: String, open: Int): Int =
    val n = text.length
    var i = open + 1
    var depth = 1
    var bracketDepth = 0
    while depth > 0 && i < n do
      val c = text.charAt(i)
      if bracketDepth == 0 && (c == '"' || c == '\'') then i = closingQuote(text, i, c)
      else
        if c == '[' then bracketDepth += 1
        else if c == ']' then { if bracketDepth > 0 then bracketDepth -= 1 }
        else if bracketDepth == 0 && c == '(' then depth += 1
        else if bracketDepth == 0 && c == ')' then depth -= 1
        i += 1
    if depth == 0 then i else -1

  /**
   * Index just past the `]` matching the `[` at `open` (`'`-escaped specials skipped); -1 when
   * unbalanced.
   */
  @SuppressWarnings(Array("org.wartremover.warts.Var", "org.wartremover.warts.While"))
  private def closingBracket(text: String, open: Int): Int =
    val n = text.length
    var i = open + 1
    var depth = 1
    while depth > 0 && i < n do
      val c = text.charAt(i)
      if c == '\'' then i += 2
      else
        if c == '[' then depth += 1
        else if c == ']' then depth -= 1
        i += 1
    if depth == 0 then math.min(i, n) else -1

  private def storedCall(token: String): String =
    if startsWithIgnoreCase(token, XlfnPrefix) then token
    else if startsWithIgnoreCase(token, XlwsPrefix) then s"$XlfnPrefix$token"
    else if isIn(WorksheetScoped, token) then s"$XlfnPrefix$XlwsPrefix$token"
    else if isIn(FutureFunctions, token) then s"$XlfnPrefix$token"
    else token

  private def storedParam(token: String): String =
    if startsWithIgnoreCase(token, XlpmPrefix) then token else s"$XlpmPrefix$token"

  private def modelCall(token: String): String =
    if !hasStoragePrefix(token) then token
    else
      val bare = bareFunctionName(token)
      if isIn(FutureFunctions, bare) then bare else token

  private def modelParam(token: String): String = bareParameterName(token)

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

  /**
   * Drop the `_xlpm.` prefix from a parameter name (case-insensitive); no allocation without it.
   */
  private def bareParameterName(name: String): String =
    if startsWithIgnoreCase(name, XlpmPrefix) then name.substring(XlpmPrefix.length) else name

  /** Case-insensitive prefix test without allocating (all prefixes are ASCII). */
  private def startsWithIgnoreCase(s: String, prefix: String): Boolean =
    s.regionMatches(true, 0, prefix, 0, prefix.length)

  private def hasStoragePrefix(s: String): Boolean =
    startsWithIgnoreCase(s, XlfnPrefix) || startsWithIgnoreCase(s, XlwsPrefix)

  /** True when `text` contains `_xl` in any letter case — the stem every storage prefix shares. */
  @SuppressWarnings(Array("org.wartremover.warts.Var", "org.wartremover.warts.While"))
  private def containsStoragePrefix(text: String): Boolean =
    var i = text.indexOf('_')
    var found = false
    while !found && i >= 0 && i + 2 < text.length do
      val x = text.charAt(i + 1)
      val l = text.charAt(i + 2)
      found = (x == 'x' || x == 'X') && (l == 'l' || l == 'L')
      if !found then i = text.indexOf('_', i + 1)
    found

  /**
   * Membership of a token spelled in the author's case. The common all-caps spelling hits (or
   * misses) without allocating; only a token that actually carries lower-case letters is
   * upper-cased (Locale.ROOT: a Turkish default locale must not turn `ifs` into `İFS`) for a second
   * look.
   */
  private def isIn(names: Set[String], token: String): Boolean =
    names.contains(token) ||
      (hasLowerCase(token) && names.contains(token.toUpperCase(Locale.ROOT)))

  @SuppressWarnings(Array("org.wartremover.warts.Var", "org.wartremover.warts.While"))
  private def hasLowerCase(s: String): Boolean =
    var i = 0
    var found = false
    while !found && i < s.length do
      found = s.charAt(i).isLower
      i += 1
    found

  /**
   * Identifier body: names, references (`$A$1` is one token, so `$A$1` never matches a name `A`).
   */
  private def isIdentChar(c: Char): Boolean =
    c.isLetterOrDigit || c == '_' || c == '.' || c == '$'

  private def isIdentStart(c: Char): Boolean =
    c.isLetter || c == '_'

  /**
   * One open LET / LAMBDA call: the paren depth of its argument list, the index of the argument
   * being scanned, and the (upper-cased) names it has declared so far.
   */
  @SuppressWarnings(Array("org.wartremover.warts.Var"))
  private final class ParameterScope(val depth: Int, val isLet: Boolean):
    var argIndex: Int = 0
    var names: Set[String] = Set.empty

  /**
   * Copy `text`, applying `onCall` to every identifier token that is directly followed (modulo
   * whitespace) by `(` — i.e. every function call — and `onParam` to every LET/LAMBDA parameter (a
   * declaration at a parameter position, or a later reference to a declared name; the second
   * argument is true inside a LET, false inside a LAMBDA), outside string literals, quoted sheet
   * names, bracketed references, array constants and error literals. Inside brackets nothing is
   * interpreted except the `'` escape (`Table1['[Total]`), which keeps the bracket depth honest;
   * quote handling there would swallow the rest of the formula. Every other byte is copied in
   * place, so the result equals `text` exactly when both callbacks are identities on every token
   * they see — the property [[bareFutureCalls]] relies on.
   */
  @SuppressWarnings(Array("org.wartremover.warts.Var", "org.wartremover.warts.While"))
  private def rewriteCalls(text: String)(
    onCall: String => String,
    onParam: (String, Boolean) => String
  ): String =
    val n = text.length
    val sb = new java.lang.StringBuilder(n + 16)
    var i = 0
    var bracketDepth = 0
    var braceDepth = 0
    var parenDepth = 0
    var scopes: List[ParameterScope] = Nil
    // 1: a LET call token was just emitted, 2: a LAMBDA — its '(' opens a parameter scope
    var pendingScope = 0
    while i < n do
      val c = text.charAt(i)
      if bracketDepth == 0 && c == '"' then
        // String literal: copy through the closing quote, honoring "" escapes
        val end = closingQuote(text, i, '"')
        sb.append(text, i, end)
        i = end
      else if bracketDepth == 0 && c == '\'' then
        // Quoted sheet name: copy through the closing quote, honoring '' escapes
        val end = closingQuote(text, i, '\'')
        sb.append(text, i, end)
        i = end
      else if bracketDepth > 0 && c == '\'' then
        // Structured-reference escape: the quote and the special it escapes are one unit
        val end = math.min(i + 2, n)
        sb.append(text, i, end)
        i = end
      else if c == '[' then
        scopes match
          case head :: _ if bracketDepth == 0 && !head.isLet && head.depth == parenDepth =>
            // LAMBDA optional parameter `[y]`: a declaration like any other
            val end = optionalParameterEnd(text, i)
            if end > 0 then
              val token = text.substring(i + 1, end - 1).trim
              head.names += bareParameterName(token).toUpperCase(Locale.ROOT)
              sb.append('[').append(onParam(token, head.isLet)).append(']')
              i = end
            else
              bracketDepth += 1
              sb.append(c)
              i += 1
          case _ =>
            bracketDepth += 1
            sb.append(c)
            i += 1
      else if c == ']' then
        if bracketDepth > 0 then bracketDepth -= 1
        sb.append(c)
        i += 1
      else if bracketDepth == 0 && c == '{' then
        braceDepth += 1
        sb.append(c)
        i += 1
      else if bracketDepth == 0 && c == '}' then
        if braceDepth > 0 then braceDepth -= 1
        sb.append(c)
        i += 1
      else if bracketDepth == 0 && braceDepth == 0 && c == '(' then
        parenDepth += 1
        if pendingScope != 0 then
          scopes = new ParameterScope(parenDepth, pendingScope == 1) :: scopes
          pendingScope = 0
        sb.append(c)
        i += 1
      else if bracketDepth == 0 && braceDepth == 0 && c == ')' then
        scopes match
          case head :: tail if head.depth == parenDepth => scopes = tail
          case _ => ()
        if parenDepth > 0 then parenDepth -= 1
        sb.append(c)
        i += 1
      else if bracketDepth == 0 && braceDepth == 0 && c == ',' then
        scopes match
          case head :: _ if head.depth == parenDepth => head.argIndex += 1
          case _ => ()
        sb.append(c)
        i += 1
      else if bracketDepth == 0 && braceDepth == 0 && c == '#' then
        // Error literal (#N/A, #DIV/0!, #NAME?, ...): opaque, so a parameter named N is not
        // mistaken for the N in #N/A. A spill reference's trailing '#' consumes nothing.
        val end = errorLiteralEnd(text, i)
        sb.append(text, i, end)
        i = end
      else if bracketDepth == 0 && braceDepth == 0 && isIdentChar(c) then
        var end = i + 1
        while end < n && isIdentChar(text.charAt(end)) do end += 1
        var probe = end
        while probe < n && text.charAt(probe).isWhitespace do probe += 1
        val next = if probe < n then text.charAt(probe) else ' '
        if isIdentStart(c) && next == '(' then
          val token = text.substring(i, end)
          if isIn(ParameterScoped, bareFunctionName(token)) then
            pendingScope = if bareFunctionName(token).equalsIgnoreCase("LET") then 1 else 2
          sb.append(onCall(token))
        else if isIdentStart(c) && scopes.nonEmpty then
          val token = text.substring(i, end)
          val key = bareParameterName(token).toUpperCase(Locale.ROOT)
          // `A:A` beside a parameter named A is still a column reference
          val partOfRange = next == ':' || (i > 0 && text.charAt(i - 1) == ':')
          val declares = scopes match
            case head :: _ =>
              !partOfRange && head.depth == parenDepth && next == ',' &&
              (!head.isLet || head.argIndex % 2 == 0)
            case Nil => false
          scopes match
            case head :: _ if declares =>
              head.names += key
              sb.append(onParam(token, head.isLet))
            case _ =>
              // A reference resolves to the innermost scope declaring the name
              scopes.find(_.names.contains(key)) match
                case Some(scope) if !partOfRange => sb.append(onParam(token, scope.isLet))
                case _ => sb.append(token)
        else sb.append(text, i, end)
        i = end
      else
        sb.append(c)
        i += 1
    sb.toString

  /** Index of the first non-whitespace character at or after `start` (`n` when none). */
  @SuppressWarnings(Array("org.wartremover.warts.Var", "org.wartremover.warts.While"))
  private def skipSpaces(text: String, start: Int): Int =
    var i = start
    while i < text.length && text.charAt(i).isWhitespace do i += 1
    i

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

  /**
   * For a `[` at `start` inside a LAMBDA argument list: the index just past the matching `]` when
   * the brackets enclose exactly one identifier (an optional parameter, `[y]`), else -1.
   */
  @SuppressWarnings(Array("org.wartremover.warts.Var", "org.wartremover.warts.While"))
  private def optionalParameterEnd(text: String, start: Int): Int =
    val n = text.length
    var i = start + 1
    while i < n && text.charAt(i).isWhitespace do i += 1
    if i < n && isIdentStart(text.charAt(i)) then
      while i < n && isIdentChar(text.charAt(i)) do i += 1
      while i < n && text.charAt(i).isWhitespace do i += 1
      if i < n && text.charAt(i) == ']' then i + 1 else -1
    else -1

  /**
   * Index just past the error literal opening with the `#` at `start`: letters (and `_`), then an
   * optional `/` segment (`#DIV/0!`, `#N/A`), then an optional `!` or `?`. A `#` followed by
   * nothing letter-like (a spill reference `A1#`) yields `start + 1`.
   */
  @SuppressWarnings(Array("org.wartremover.warts.Var", "org.wartremover.warts.While"))
  private def errorLiteralEnd(text: String, start: Int): Int =
    val n = text.length
    var i = start + 1
    while i < n && (text.charAt(i).isLetter || text.charAt(i) == '_') do i += 1
    if i == start + 1 then i
    else
      if i < n && text.charAt(i) == '/' then
        i += 1
        while i < n && text.charAt(i).isLetterOrDigit do i += 1
      if i < n && (text.charAt(i) == '!' || text.charAt(i) == '?') then i += 1
      i
