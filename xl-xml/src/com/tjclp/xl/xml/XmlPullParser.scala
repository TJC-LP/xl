package com.tjclp.xl.xml

import com.tjclp.xl.xml.internal.{CharDecoder, Entities}

/**
 * Portable, OOXML-tuned XML pull parser (ADR-016 wave A2).
 *
 * Security posture (ports xl-ooxml's XmlSecurity):
 *   - ONLY the five predefined entities (`amp lt gt apos quot`) plus numeric character references
 *     validated against the XML 1.0 Char production — general/parameter entities are structurally
 *     impossible.
 *   - DOCTYPE declarations are skipped benignly (quote/bracket-depth state machine, comments and
 *     PIs inside the internal subset skipped wholesale — the GH-350 corruption cases); NOTHING
 *     declared inside is ever honored, and any later reference to an undeclared entity is an error.
 *     Bounded by `maxTotalChars` (retires the 16KB prolog-scan truncation of the stream stripper).
 *   - [[XmlLimits]] enforced: maxDepth at start-tag push, maxAttrsPerElement in the attribute loop,
 *     maxNameLength in the name scanner, maxTotalChars in the decoder.
 *   - UTF-8 and UTF-16 only (BOM + encoding-declaration detection); anything else fails closed with
 *     a named error. Malformed byte sequences are positioned errors, never replacement characters.
 *
 * Semantics:
 *   - TOTAL: [[next]] never throws. A `Left` is TERMINAL and sticky (every later call returns the
 *     same `Left`); after `Right(EndDocument)`, sticky `Right(EndDocument)`.
 *   - Line-end normalization (CRLF | CR -> LF) happens before tokenization; `&#xD;` survives.
 *   - Attribute values: raw tab/LF/CR become spaces (XML 1.0 §3.3.3); character-reference-produced
 *     characters are exempt.
 *   - `<e/>` yields StartElement followed by a synthesized EndElement.
 *   - Comments and processing instructions are consumed silently and do not split text runs;
 *     adjacent text is coalesced into one maximal run; CDATA is its own token.
 *   - Whitespace-only text is always delivered (no DTD means nothing is "ignorable").
 *   - Namespace-strict like JAXP: an unbound prefix on an element or attribute name is an error
 *     (checked only when the name contains ':').
 *
 * Cursor accessors ([[name]], [[attrs]], [[text]], ...) are valid ONLY until the next [[next]] call
 * (SAX-style aliasing contract). Not thread-safe.
 */
@SuppressWarnings(
  Array("org.wartremover.warts.Var", "org.wartremover.warts.While")
)
final class XmlPullParser private (in: ByteInput, limits: XmlLimits):
  import XmlPullParser.*
  import CharDecoder.{Eof, Err}

  private val dec = new CharDecoder(in, limits.maxTotalChars)

  // Cached results: zero per-token allocation on the hot path.
  private val RStart: Either[XmlParseError, XmlToken] = Right(XmlToken.StartElement)
  private val REnd: Either[XmlParseError, XmlToken] = Right(XmlToken.EndElement)
  private val RText: Either[XmlParseError, XmlToken] = Right(XmlToken.Text)
  private val RCData: Either[XmlParseError, XmlToken] = Right(XmlToken.CData)
  private val REndDoc: Either[XmlParseError, XmlToken] = Right(XmlToken.EndDocument)

  // --- lookahead ring over decoded code points (max construct discriminator is 9 chars) ---------
  private val la = new Array[Int](16)
  private val laLine = new Array[Int](16)
  private val laCol = new Array[Int](16)
  private var laStart = 0
  private var laLen = 0

  private def fillTo(n: Int): Unit =
    while laLen <= n do
      val idx = (laStart + laLen) & 15
      la(idx) = dec.cp
      laLine(idx) = dec.line
      laCol(idx) = dec.col
      if dec.cp >= 0 then dec.advance()
      laLen += 1

  private def cur: Int =
    fillTo(0)
    la(laStart)

  private def peek(i: Int): Int =
    fillTo(i)
    la((laStart + i) & 15)

  private def curLine: Int =
    fillTo(0)
    laLine(laStart)

  private def curCol: Int =
    fillTo(0)
    laCol(laStart)

  private def bump(): Unit =
    fillTo(0)
    laStart = (laStart + 1) & 15
    laLen -= 1

  /** True when the next code points spell the ASCII string `s` (length <= 9). */
  private def lookingAt(s: String): Boolean =
    var i = 0
    var ok = true
    while ok && i < s.length do
      if peek(i) != s.charAt(i).toInt then ok = false
      i += 1
    ok

  // --- parser state ------------------------------------------------------------------------------
  private var terminal: Option[Either[XmlParseError, XmlToken]] = None
  private var state = StProlog
  private var atStart = true
  private var seenDoctype = false
  private var pendingEnd = false

  // open-element stack (parallel arrays)
  private var elemQNames = new Array[String](16)
  private var elemPrefixes = new Array[String](16)
  private var elemLocals = new Array[String](16)
  private var elemNsMarks = new Array[Int](16)
  private var depthV = 0

  // namespace bindings: flat parallel arrays; per-element marks restore on pop
  private var nsPrefixes = new Array[String](16)
  private var nsUris = new Array[String](16)
  private var nsCount = 0

  // current-token payload
  private var curQName = ""
  private var curPrefixV = ""
  private var curLocalV = ""
  private val attrList = new AttrList()
  private var textBuf = new Array[Char](64)
  private var textLen = 0
  private var textStr = ""
  private var textStrValid = true
  private var tokLine = 1
  private var tokCol = 1

  // scratch
  private val nameSb = new java.lang.StringBuilder(32)
  private var splitPrefix = ""
  private var splitLocal = ""
  private var attrChars = new Array[Char](64)
  private var attrLen = 0

  // ================================================================================== public API

  /**
   * Advance to the next token. TOTAL — never throws. Terminal results (any `Left`, and
   * `Right(EndDocument)`) are sticky.
   */
  def next(): Either[XmlParseError, XmlToken] =
    terminal match
      case Some(t) => t
      case None =>
        val r = step()
        r match
          case Left(_) => terminal = Some(r)
          case Right(XmlToken.EndDocument) => terminal = Some(r)
          case _ => ()
        r

  /** qName as written (`"c"`, `"x14ac:dyDescent"`). Valid for Start/EndElement. */
  def name: String = curQName

  /** Local part of the current element name. */
  def localName: String = curLocalV

  /** Prefix of the current element name; "" when unprefixed. */
  def prefix: String = curPrefixV

  /** Namespace of the current element (default namespace for unprefixed names). Lazy. */
  def elementNs: Option[String] = resolveNs(curPrefixV)

  /**
   * Resolve `prefix` against the live namespace stack ("" = default namespace). `"xml"` is
   * hardwired; an `xmlns=""` undeclare yields None.
   */
  def resolveNs(prefix: String): Option[String] =
    if prefix == "xml" then Some(XmlNamespaceUri)
    else
      var i = nsCount - 1
      var res: Option[String] = None
      var found = false
      while !found && i >= 0 do
        if nsPrefixes(i) == prefix then
          found = true
          if nsUris(i).nonEmpty then res = Some(nsUris(i))
        i -= 1
      res

  /**
   * Attributes of the current StartElement (namespace declarations included verbatim). Single
   * instance, reused across elements — valid only until the next [[next]] call.
   */
  def attrs: AttrList = attrList

  /** Decoded character data of the current Text/CData token; materialized on demand. */
  def text: String =
    if !textStrValid then
      textStr = new String(textBuf, 0, textLen)
      textStrValid = true
    textStr

  /** Append the current Text/CData content to `sb` without materializing a String. */
  def appendTextTo(sb: java.lang.StringBuilder): Unit =
    val _ = sb.append(textBuf, 0, textLen)

  /** True when the current Text/CData token is entirely XML whitespace. */
  def isWhitespaceText: Boolean =
    var i = 0
    var ws = true
    while ws && i < textLen do
      val c = textBuf(i)
      if c != ' ' && c != '\t' && c != '\n' && c != '\r' then ws = false
      i += 1
    ws

  /** 1-based line of the current token's first character (SAXParseException convention). */
  def line: Int = tokLine

  /** 1-based column of the current token's first character. */
  def col: Int = tokCol

  /** Open-element depth after the current token (StartElement of the root = 1). */
  def depth: Int = depthV

  // ============================================================================== token stepping

  private def step(): Either[XmlParseError, XmlToken] =
    if pendingEnd then
      pendingEnd = false
      popElement()
      if depthV == 0 then state = StEpilog
      REnd
    else
      state match
        case StProlog => prologStep()
        case StContent => contentStep()
        case _ => epilogStep()

  private def hereErr(msg: String): XmlParseError = XmlParseError(curLine, curCol, msg)

  private def decoderErr(): XmlParseError = XmlParseError(curLine, curCol, dec.errorMsg)

  // --- prolog ------------------------------------------------------------------------------------

  private def prologStep(): Either[XmlParseError, XmlToken] =
    var res: Option[Either[XmlParseError, XmlToken]] = None
    while res.isEmpty do
      val c = cur
      if c == Err then res = Some(Left(decoderErr()))
      else if c == Eof then res = Some(Left(hereErr("premature end of file: no root element")))
      else if isWs(c) then
        atStart = false
        bump()
      else if c == '<' then
        val c1 = peek(1)
        if c1 == '?' then
          handlePi() match
            case Some(e) => res = Some(Left(e))
            case None => ()
        else if c1 == '!' then
          if peek(2) == '-' && peek(3) == '-' then
            atStart = false
            skipComment() match
              case Some(e) => res = Some(Left(e))
              case None => ()
          else if lookingAt("<!DOCTYPE") then
            atStart = false
            if seenDoctype then
              res = Some(Left(hereErr("multiple DOCTYPE declarations are not allowed")))
            else
              skipDoctype() match
                case Some(e) => res = Some(Left(e))
                case None => seenDoctype = true
          else if lookingAt("<![CDATA[") then
            res = Some(Left(hereErr("CDATA section is not allowed in the prolog")))
          else res = Some(Left(hereErr("markup not recognized in the prolog")))
        else if c1 == '/' then res = Some(Left(hereErr("end tag is not allowed in the prolog")))
        else
          state = StContent
          res = Some(parseStartTag())
      else res = Some(Left(hereErr("content is not allowed in the prolog")))
    res.getOrElse(Left(hereErr("internal parser error")))

  // --- content -----------------------------------------------------------------------------------

  private def contentStep(): Either[XmlParseError, XmlToken] =
    resetText()
    var res: Option[Either[XmlParseError, XmlToken]] = None
    while res.isEmpty do
      val c = cur
      if c == Err then res = Some(Left(decoderErr()))
      else if c == Eof then
        val open = if depthV > 0 then elemQNames(depthV - 1) else ""
        res = Some(Left(hereErr(s"premature end of file: element '$open' is not closed")))
      else if c == '<' then
        val c1 = peek(1)
        if c1 == '!' && peek(2) == '-' && peek(3) == '-' then
          // comments do not split text runs: skip and keep coalescing
          skipComment() match
            case Some(e) => res = Some(Left(e))
            case None => ()
        else if c1 == '?' then
          handlePi() match
            case Some(e) => res = Some(Left(e))
            case None => ()
        else if lookingAt("<![CDATA[") then
          if textLen > 0 then res = Some(RText)
          else res = Some(parseCData())
        else if c1 == '/' then
          if textLen > 0 then res = Some(RText)
          else res = Some(parseEndTag())
        else if c1 == '!' then res = Some(Left(hereErr("markup not recognized in content")))
        else if textLen > 0 then res = Some(RText)
        else res = Some(parseStartTag())
      else if c == '&' then
        if textLen == 0 then
          tokLine = curLine
          tokCol = curCol
        decodeReference(attrMode = false) match
          case Some(e) => res = Some(Left(e))
          case None => ()
      else if c == ']' && peek(1) == ']' && peek(2) == '>' then
        res = Some(Left(hereErr("']]>' is not allowed in character data")))
      else
        if textLen == 0 then
          tokLine = curLine
          tokCol = curCol
        appendTextCp(c)
        bump()
    res.getOrElse(Left(hereErr("internal parser error")))

  // --- epilog ------------------------------------------------------------------------------------

  private def epilogStep(): Either[XmlParseError, XmlToken] =
    var res: Option[Either[XmlParseError, XmlToken]] = None
    while res.isEmpty do
      val c = cur
      if c == Err then res = Some(Left(decoderErr()))
      else if c == Eof then
        tokLine = curLine
        tokCol = curCol
        state = StDone
        res = Some(REndDoc)
      else if isWs(c) then bump()
      else if c == '<' then
        val c1 = peek(1)
        if c1 == '?' then
          handlePi() match
            case Some(e) => res = Some(Left(e))
            case None => ()
        else if c1 == '!' && peek(2) == '-' && peek(3) == '-' then
          skipComment() match
            case Some(e) => res = Some(Left(e))
            case None => ()
        else if lookingAt("<!DOCTYPE") then
          res = Some(Left(hereErr("DOCTYPE is not allowed after the root element")))
        else if c1 == '/' then
          res = Some(Left(hereErr("unexpected end tag after the root element")))
        else res = Some(Left(hereErr("document must contain exactly one root element")))
      else res = Some(Left(hereErr("content is not allowed after the root element")))
    res.getOrElse(Left(hereErr("internal parser error")))

  // --- start / end tags ----------------------------------------------------------------------------

  private def parseStartTag(): Either[XmlParseError, XmlToken] =
    tokLine = curLine
    tokCol = curCol
    atStart = false
    bump() // '<'
    scanQName() match
      case Left(e) => Left(e)
      case Right(q) =>
        val ePrefix = splitPrefix
        val eLocal = splitLocal
        if depthV + 1 > limits.maxDepth then
          Left(
            XmlParseError(
              tokLine,
              tokCol,
              s"maxDepth limit exceeded: element nesting deeper than ${limits.maxDepth}"
            )
          )
        else
          val nsMark = nsCount
          attrList.clear()
          var selfClosing = false
          var err: Option[XmlParseError] = None
          var done = false
          while err.isEmpty && !done do
            val hadWs = skipWs()
            val c = cur
            if c == Err then err = Some(decoderErr())
            else if c == Eof then err = Some(hereErr("premature end of file in start tag"))
            else if c == '>' then
              bump()
              done = true
            else if c == '/' then
              if peek(1) == '>' then
                bump()
                bump()
                selfClosing = true
                done = true
              else err = Some(hereErr("malformed start tag: expected '/>'"))
            else if !hadWs then err = Some(hereErr("whitespace required between attributes"))
            else err = parseAttribute()
          err.orElse(checkUnboundPrefixes(ePrefix, q)) match
            case Some(e) =>
              nsCount = nsMark
              Left(e)
            case None =>
              pushElement(q, ePrefix, eLocal, nsMark)
              curQName = q
              curPrefixV = ePrefix
              curLocalV = eLocal
              if selfClosing then pendingEnd = true
              RStart

  /** JAXP-strict: unbound prefix on the element or any attribute name is an error. */
  private def checkUnboundPrefixes(ePrefix: String, q: String): Option[XmlParseError] =
    if ePrefix.nonEmpty && ePrefix != "xml" && resolveNs(ePrefix).isEmpty then
      Some(XmlParseError(tokLine, tokCol, s"unbound namespace prefix '$ePrefix' on element '$q'"))
    else
      var i = 0
      var bad: Option[XmlParseError] = None
      while bad.isEmpty && i < attrList.size do
        val an = attrList.name(i)
        val colon = an.indexOf(':')
        if colon > 0 && an != "xmlns" && !an.startsWith("xmlns:") then
          val p = an.substring(0, colon)
          if p != "xml" && resolveNs(p).isEmpty then
            bad = Some(
              XmlParseError(tokLine, tokCol, s"unbound namespace prefix '$p' on attribute '$an'")
            )
        i += 1
      bad

  private def parseAttribute(): Option[XmlParseError] =
    val aLine = curLine
    val aCol = curCol
    scanQName() match
      case Left(e) => Some(e)
      case Right(aName) =>
        val _ = skipWs()
        if cur != '=' then Some(hereErr(s"expected '=' after attribute name '$aName'"))
        else
          bump()
          val _ = skipWs()
          val qc = cur
          if qc != '"' && qc != '\'' then Some(hereErr("attribute value must be quoted"))
          else
            bump()
            attrLen = 0
            var err: Option[XmlParseError] = None
            var done = false
            while err.isEmpty && !done do
              val c = cur
              if c == Err then err = Some(decoderErr())
              else if c == Eof then err = Some(hereErr("unterminated attribute value"))
              else if c == qc then
                bump()
                done = true
              else if c == '<' then err = Some(hereErr("'<' is not allowed in attribute value"))
              else if c == '&' then err = decodeReference(attrMode = true)
              else if c == '\t' || c == '\n' then
                // §3.3.3 attribute-value normalization (raw CR already normalized to LF)
                appendAttrCp(' ')
                bump()
              else
                appendAttrCp(c)
                bump()
            err.orElse {
              val v = new String(attrChars, 0, attrLen)
              if attrList.indexOf(aName) >= 0 then
                Some(XmlParseError(aLine, aCol, s"duplicate attribute '$aName'"))
              else if attrList.size + 1 > limits.maxAttrsPerElement then
                Some(
                  XmlParseError(
                    aLine,
                    aCol,
                    s"maxAttrsPerElement limit exceeded: more than ${limits.maxAttrsPerElement} attributes on one element"
                  )
                )
              else
                val nsIssue = recordNamespaceDecl(aName, v, aLine, aCol)
                if nsIssue.isEmpty then attrList.add(aName, v)
                nsIssue
            }

  /** Feed `xmlns` / `xmlns:*` attributes into the namespace stack (they stay in the AttrList). */
  private def recordNamespaceDecl(
    aName: String,
    v: String,
    aLine: Int,
    aCol: Int
  ): Option[XmlParseError] =
    if aName == "xmlns" then
      bindNs("", v) // v == "" undeclares the default namespace
      None
    else if aName.startsWith("xmlns:") then
      val p = aName.substring(6)
      if p == "xmlns" then Some(XmlParseError(aLine, aCol, "prefix 'xmlns' cannot be declared"))
      else if p == "xml" && v != XmlNamespaceUri then
        Some(XmlParseError(aLine, aCol, "prefix 'xml' can only bind the XML namespace"))
      else if v.isEmpty then
        Some(XmlParseError(aLine, aCol, s"namespace prefix '$p' cannot be undeclared"))
      else
        bindNs(p, v)
        None
    else None

  private def parseEndTag(): Either[XmlParseError, XmlToken] =
    tokLine = curLine
    tokCol = curCol
    bump() // '<'
    bump() // '/'
    scanQName() match
      case Left(e) => Left(e)
      case Right(q) =>
        val _ = skipWs()
        if cur != '>' then Left(hereErr("malformed end tag"))
        else
          bump()
          if depthV == 0 then Left(XmlParseError(tokLine, tokCol, "end tag without open element"))
          else if elemQNames(depthV - 1) != q then
            Left(
              XmlParseError(
                tokLine,
                tokCol,
                s"mismatched end tag: expected '</${elemQNames(depthV - 1)}>', found '</$q>'"
              )
            )
          else
            popElement()
            if depthV == 0 then state = StEpilog
            REnd

  // --- CDATA -------------------------------------------------------------------------------------

  private def parseCData(): Either[XmlParseError, XmlToken] =
    tokLine = curLine
    tokCol = curCol
    var i = 0
    while i < 9 do // "<![CDATA["
      bump()
      i += 1
    var res: Option[Either[XmlParseError, XmlToken]] = None
    while res.isEmpty do
      val c = cur
      if c == Err then res = Some(Left(decoderErr()))
      else if c == Eof then
        res = Some(Left(XmlParseError(tokLine, tokCol, "unterminated CDATA section")))
      else if c == ']' && peek(1) == ']' && peek(2) == '>' then
        bump()
        bump()
        bump()
        res = Some(RCData)
      else
        appendTextCp(c)
        bump()
    res.getOrElse(Left(hereErr("internal parser error")))

  // --- comments, PIs, XML declaration --------------------------------------------------------------

  private def skipComment(): Option[XmlParseError] =
    val cl = curLine
    val cc = curCol
    var i = 0
    while i < 4 do // "<!--"
      bump()
      i += 1
    var res: Option[Option[XmlParseError]] = None
    while res.isEmpty do
      val c = cur
      if c == Err then res = Some(Some(decoderErr()))
      else if c == Eof then res = Some(Some(XmlParseError(cl, cc, "unterminated comment")))
      else if c == '-' && peek(1) == '-' then
        if peek(2) == '>' then
          bump()
          bump()
          bump()
          res = Some(None)
        else res = Some(Some(hereErr("'--' is not allowed within a comment")))
      else bump()
    res.getOrElse(None)

  /** Handle `<?...?>`: the XML declaration at the very start, a skipped PI anywhere else. */
  private def handlePi(): Option[XmlParseError] =
    val declCandidate = atStart && state == StProlog
    val piLine = curLine
    val piCol = curCol
    atStart = false
    bump() // '<'
    bump() // '?'
    scanQName() match
      case Left(e) => Some(e)
      case Right(target) =>
        if splitPrefix.nonEmpty then
          Some(
            XmlParseError(piLine, piCol, "processing instruction target cannot contain ':'")
          )
        else if target == "xml" && declCandidate then parseXmlDecl()
        else if asciiEqualsIgnoreCase(target, "xml") then
          Some(XmlParseError(piLine, piCol, "processing instruction target 'xml' is reserved"))
        else if cur == '?' && peek(1) == '>' then
          bump()
          bump()
          None
        else if isWs(cur) then
          var res: Option[Option[XmlParseError]] = None
          while res.isEmpty do
            val c = cur
            if c == Err then res = Some(Some(decoderErr()))
            else if c == Eof then
              res = Some(Some(XmlParseError(piLine, piCol, "unterminated processing instruction")))
            else if c == '?' && peek(1) == '>' then
              bump()
              bump()
              res = Some(None)
            else bump()
          res.getOrElse(None)
        else Some(hereErr("expected whitespace after processing instruction target"))

  private def parseXmlDecl(): Option[XmlParseError] =
    val r =
      for
        _ <- requireWs("XML declaration")
        _ <- declPseudoAttr("version").flatMap { v =>
          if v == "1.0" then Right(())
          else Left(hereErr(s"unsupported XML version '$v': only 1.0 is supported"))
        }
        _ <- declTail(allowEncoding = true)
      yield ()
    r.left.toOption

  /** Optional `encoding` then optional `standalone`, then `?>`. */
  private def declTail(allowEncoding: Boolean): Either[XmlParseError, Unit] =
    val hadWs = skipWs()
    if cur == '?' then expectPiClose()
    else if !hadWs then Left(hereErr("malformed XML declaration"))
    else
      scanQName().flatMap {
        case "encoding" if allowEncoding =>
          declEq()
            .flatMap(_ => declQuotedValue())
            .flatMap(validateEncodingDecl)
            .flatMap(_ => declTail(allowEncoding = false))
        case "standalone" =>
          declEq().flatMap(_ => declQuotedValue()).flatMap { v =>
            if v == "yes" || v == "no" then
              val _ = skipWs()
              expectPiClose()
            else Left(hereErr("standalone must be 'yes' or 'no'"))
          }
        case other => Left(hereErr(s"unexpected '$other' in XML declaration"))
      }

  private def declPseudoAttr(expected: String): Either[XmlParseError, String] =
    scanQName().flatMap { n =>
      if n != expected then Left(hereErr(s"expected '$expected' in XML declaration, found '$n'"))
      else declEq().flatMap(_ => declQuotedValue())
    }

  private def declEq(): Either[XmlParseError, Unit] =
    val _ = skipWs()
    if cur != '=' then Left(hereErr("expected '=' in XML declaration"))
    else
      bump()
      val _ = skipWs()
      Right(())

  private def declQuotedValue(): Either[XmlParseError, String] =
    val qc = cur
    if qc != '"' && qc != '\'' then Left(hereErr("XML declaration value must be quoted"))
    else
      bump()
      nameSb.setLength(0)
      var res: Option[Either[XmlParseError, String]] = None
      while res.isEmpty do
        val c = cur
        if c == Err then res = Some(Left(decoderErr()))
        else if c == Eof then res = Some(Left(hereErr("unterminated XML declaration")))
        else if c == qc then
          bump()
          res = Some(Right(nameSb.toString))
        else
          appendCpToSb(nameSb, c)
          bump()
      res.getOrElse(Left(hereErr("internal parser error")))

  private def validateEncodingDecl(value: String): Either[XmlParseError, Unit] =
    val norm = asciiLower(value)
    val detected = dec.encodingName
    val consistent = detected match
      case "UTF-16LE" => norm == "utf-16" || norm == "utf-16le"
      case "UTF-16BE" => norm == "utf-16" || norm == "utf-16be"
      case _ => norm == "utf-8" || norm == "us-ascii"
    if consistent then Right(())
    else if norm == "utf-8" || norm == "us-ascii" || norm == "utf-16" || norm == "utf-16le" ||
      norm == "utf-16be"
    then Left(hereErr(s"encoding declaration '$value' conflicts with detected encoding $detected"))
    else Left(hereErr(s"unsupported encoding '$value': only UTF-8 and UTF-16 are supported"))

  private def expectPiClose(): Either[XmlParseError, Unit] =
    if cur == '?' && peek(1) == '>' then
      bump()
      bump()
      Right(())
    else Left(hereErr("malformed XML declaration"))

  private def requireWs(where: String): Either[XmlParseError, Unit] =
    if skipWs() then Right(()) else Left(hereErr(s"whitespace required in $where"))

  // --- DOCTYPE (skip-benign, GH-350) ----------------------------------------------------------------

  private def skipDoctype(): Option[XmlParseError] =
    val dl = curLine
    val dc = curCol
    var i = 0
    while i < 9 do // "<!DOCTYPE"
      bump()
      i += 1
    var quote = 0
    var bracketDepth = 0
    var res: Option[Option[XmlParseError]] = None
    while res.isEmpty do
      val c = cur
      if c == Err then res = Some(Some(decoderErr()))
      else if c == Eof then
        res = Some(Some(XmlParseError(dl, dc, "unterminated DOCTYPE declaration")))
      else if quote != 0 then
        if c == quote then quote = 0
        bump()
      else if c == '<' && peek(1) == '!' && peek(2) == '-' && peek(3) == '-' then
        skipComment() match
          case Some(e) => res = Some(Some(e))
          case None => ()
      else if c == '<' && peek(1) == '?' then
        // PI spans inside the internal subset are skipped wholesale (GH-350)
        bump()
        bump()
        var piDone = false
        while res.isEmpty && !piDone do
          val pc = cur
          if pc == Err then res = Some(Some(decoderErr()))
          else if pc == Eof then
            res = Some(Some(XmlParseError(dl, dc, "unterminated DOCTYPE declaration")))
          else if pc == '?' && peek(1) == '>' then
            bump()
            bump()
            piDone = true
          else bump()
      else if c == '"' || c == '\'' then
        quote = c
        bump()
      else if c == '[' then
        bracketDepth += 1
        bump()
      else if c == ']' then
        bracketDepth -= 1
        bump()
      else if c == '>' && bracketDepth == 0 then
        bump()
        res = Some(None)
      else bump()
    res.getOrElse(None)

  // --- entities & character references ---------------------------------------------------------------

  /**
   * Decode `&...;` at the cursor into the text or attribute buffer. Charref-produced characters are
   * exempt from attribute-value normalization.
   */
  private def decodeReference(attrMode: Boolean): Option[XmlParseError] =
    val rLine = curLine
    val rCol = curCol
    bump() // '&'
    if cur == '#' then
      bump()
      var hex = false
      if cur == 'x' then
        hex = true
        bump()
      var value = 0
      var digits = 0
      var doneD = false
      while !doneD do
        val d = digitVal(cur, hex)
        if d < 0 then doneD = true
        else
          digits += 1
          if value <= 0x10ffff then
            value = value * (if hex then 16 else 10) + d
            if value > 0x10ffff then value = 0x110000
          bump()
      if cur == Err then Some(decoderErr())
      else if digits == 0 then Some(XmlParseError(rLine, rCol, "malformed character reference"))
      else if cur != ';' then
        Some(XmlParseError(rLine, rCol, "character reference must end with ';'"))
      else
        bump()
        if !Entities.isXmlChar(value) then
          Some(
            XmlParseError(
              rLine,
              rCol,
              "invalid character reference: not an XML 1.0 Char (surrogates and values beyond U+10FFFF are not allowed)"
            )
          )
        else
          appendRefCp(value, attrMode)
          None
    else
      nameSb.setLength(0)
      var n = 0
      while n < 32 && isNameChar(cur) do
        appendCpToSb(nameSb, cur)
        bump()
        n += 1
      if cur == Err then Some(decoderErr())
      else if nameSb.length == 0 then
        Some(
          XmlParseError(
            rLine,
            rCol,
            "'&' must start an entity or character reference (use '&amp;' for a literal '&')"
          )
        )
      else if cur != ';' then Some(XmlParseError(rLine, rCol, "entity reference must end with ';'"))
      else
        val nm = nameSb.toString
        bump()
        val ch = Entities.predefined(nm)
        if ch < 0 then
          Some(
            XmlParseError(
              rLine,
              rCol,
              s"undeclared entity '&$nm;' (only the five predefined XML entities are supported)"
            )
          )
        else
          appendRefCp(ch, attrMode)
          None

  private def appendRefCp(cp: Int, attrMode: Boolean): Unit =
    if attrMode then appendAttrCp(cp) else appendTextCp(cp)

  private def digitVal(c: Int, hex: Boolean): Int =
    if c >= '0' && c <= '9' then c - '0'
    else if hex && c >= 'a' && c <= 'f' then c - 'a' + 10
    else if hex && c >= 'A' && c <= 'F' then c - 'A' + 10
    else -1

  // --- names -----------------------------------------------------------------------------------------

  /**
   * Scan a Name at the cursor: XML 1.0 NameStartChar/NameChar productions, maxNameLength, and the
   * namespace-strict single-colon shape. Sets [[splitPrefix]]/[[splitLocal]].
   */
  private def scanQName(): Either[XmlParseError, String] =
    val l = curLine
    val c0 = curCol
    val first = cur
    if first == Err then Left(decoderErr())
    else if first == Eof then Left(XmlParseError(l, c0, "premature end of file: name expected"))
    else if !isNameStart(first) then Left(XmlParseError(l, c0, "invalid name start character"))
    else
      nameSb.setLength(0)
      appendCpToSb(nameSb, first)
      bump()
      var colon = -1
      var count = 1
      var err: Option[XmlParseError] = None
      var doneScan = false
      while err.isEmpty && !doneScan do
        val ch = cur
        if ch == ':' then
          if colon >= 0 then
            err = Some(XmlParseError(curLine, curCol, "invalid qualified name: more than one ':'"))
          else
            colon = nameSb.length
            nameSb.append(':')
            bump()
            count += 1
            if !isNameStart(cur) then
              err = Some(XmlParseError(curLine, curCol, "invalid qualified name"))
        else if isNameChar(ch) then
          appendCpToSb(nameSb, ch)
          bump()
          count += 1
        else doneScan = true
        if err.isEmpty && count > limits.maxNameLength then
          err = Some(
            XmlParseError(
              l,
              c0,
              s"maxNameLength limit exceeded: name longer than ${limits.maxNameLength} characters"
            )
          )
      err match
        case Some(e) => Left(e)
        case None =>
          val q = nameSb.toString
          if colon >= 0 then
            splitPrefix = q.substring(0, colon)
            splitLocal = q.substring(colon + 1)
          else
            splitPrefix = ""
            splitLocal = q
          Right(q)

  // XML 1.0 (5th ed.) NameStartChar, minus ':' (namespace-strict qualified names)
  private def isNameStart(c: Int): Boolean =
    (c >= 'a' && c <= 'z') || (c >= 'A' && c <= 'Z') || c == '_' ||
      (c >= 0xc0 && c <= 0xd6) || (c >= 0xd8 && c <= 0xf6) || (c >= 0xf8 && c <= 0x2ff) ||
      (c >= 0x370 && c <= 0x37d) || (c >= 0x37f && c <= 0x1fff) ||
      (c >= 0x200c && c <= 0x200d) || (c >= 0x2070 && c <= 0x218f) ||
      (c >= 0x2c00 && c <= 0x2fef) || (c >= 0x3001 && c <= 0xd7ff) ||
      (c >= 0xf900 && c <= 0xfdcf) || (c >= 0xfdf0 && c <= 0xfffd) ||
      (c >= 0x10000 && c <= 0xeffff)

  private def isNameChar(c: Int): Boolean =
    isNameStart(c) || c == '-' || c == '.' || (c >= '0' && c <= '9') || c == 0xb7 ||
      (c >= 0x300 && c <= 0x36f) || (c >= 0x203f && c <= 0x2040)

  // --- small helpers ----------------------------------------------------------------------------------

  private def isWs(c: Int): Boolean = c == ' ' || c == '\t' || c == '\n'

  /** Skip XML whitespace; true when at least one character was consumed. */
  private def skipWs(): Boolean =
    var any = false
    while isWs(cur) do
      any = true
      bump()
    any

  private def resetText(): Unit =
    textLen = 0
    textStr = ""
    textStrValid = false

  private def appendTextCp(cp: Int): Unit =
    if textLen + 2 > textBuf.length then
      val grown = new Array[Char](textBuf.length * 2)
      System.arraycopy(textBuf, 0, grown, 0, textLen)
      textBuf = grown
    if cp > 0xffff then
      textBuf(textLen) = Character.highSurrogate(cp)
      textBuf(textLen + 1) = Character.lowSurrogate(cp)
      textLen += 2
    else
      textBuf(textLen) = cp.toChar
      textLen += 1

  private def appendAttrCp(cp: Int): Unit =
    if attrLen + 2 > attrChars.length then
      val grown = new Array[Char](attrChars.length * 2)
      System.arraycopy(attrChars, 0, grown, 0, attrLen)
      attrChars = grown
    if cp > 0xffff then
      attrChars(attrLen) = Character.highSurrogate(cp)
      attrChars(attrLen + 1) = Character.lowSurrogate(cp)
      attrLen += 2
    else
      attrChars(attrLen) = cp.toChar
      attrLen += 1

  private def appendCpToSb(sb: java.lang.StringBuilder, cp: Int): Unit =
    val _ = if cp > 0xffff then sb.appendCodePoint(cp) else sb.append(cp.toChar)

  private def bindNs(prefix: String, uri: String): Unit =
    if nsCount == nsPrefixes.length then
      val grownP = new Array[String](nsPrefixes.length * 2)
      val grownU = new Array[String](nsUris.length * 2)
      System.arraycopy(nsPrefixes, 0, grownP, 0, nsCount)
      System.arraycopy(nsUris, 0, grownU, 0, nsCount)
      nsPrefixes = grownP
      nsUris = grownU
    nsPrefixes(nsCount) = prefix
    nsUris(nsCount) = uri
    nsCount += 1

  private def pushElement(q: String, p: String, l: String, nsMark: Int): Unit =
    if depthV == elemQNames.length then
      val n = elemQNames.length * 2
      val gq = new Array[String](n)
      val gp = new Array[String](n)
      val gl = new Array[String](n)
      val gm = new Array[Int](n)
      System.arraycopy(elemQNames, 0, gq, 0, depthV)
      System.arraycopy(elemPrefixes, 0, gp, 0, depthV)
      System.arraycopy(elemLocals, 0, gl, 0, depthV)
      System.arraycopy(elemNsMarks, 0, gm, 0, depthV)
      elemQNames = gq
      elemPrefixes = gp
      elemLocals = gl
      elemNsMarks = gm
    elemQNames(depthV) = q
    elemPrefixes(depthV) = p
    elemLocals(depthV) = l
    elemNsMarks(depthV) = nsMark
    depthV += 1

  private def popElement(): Unit =
    depthV -= 1
    curQName = elemQNames(depthV)
    curPrefixV = elemPrefixes(depthV)
    curLocalV = elemLocals(depthV)
    nsCount = elemNsMarks(depthV)

  private def asciiLower(s: String): String =
    val sb = new java.lang.StringBuilder(s.length)
    var i = 0
    while i < s.length do
      val c = s.charAt(i)
      val _ = sb.append(if c >= 'A' && c <= 'Z' then (c + 32).toChar else c)
      i += 1
    sb.toString

  private def asciiEqualsIgnoreCase(a: String, b: String): Boolean =
    asciiLower(a) == asciiLower(b)

object XmlPullParser:
  private val XmlNamespaceUri = "http://www.w3.org/XML/1998/namespace"
  private final val StProlog = 0
  private final val StContent = 1
  private final val StEpilog = 2
  private final val StDone = 3

  def apply(in: ByteInput, limits: XmlLimits = XmlLimits.default): XmlPullParser =
    new XmlPullParser(in, limits)

  def fromArray(bytes: Array[Byte], limits: XmlLimits = XmlLimits.default): XmlPullParser =
    apply(ByteInput.fromArray(bytes), limits)

  def fromString(xml: String, limits: XmlLimits = XmlLimits.default): XmlPullParser =
    apply(ByteInput.fromString(xml), limits)
