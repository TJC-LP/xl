package com.tjclp.xl.sheets

import com.tjclp.xl.addressing.{ARef, CellRange}

/**
 * Textual `sqref` rewrite for reader-constructed Preserved payloads under structural edits
 * (GH-429).
 *
 * `DataValidation.Preserved` / `ConditionalFormat.Preserved` carry the scope-self-contained
 * canonical XML of one whole element — as produced by the reader's `CfCodec.preservedXml`, i.e.
 * `XmlUtil.compact`: an XML declaration, a newline, then the element. Without this rewrite they are
 * structurally inert and detach from the cells they governed on every row/column insert or delete.
 * Rather than parse XML in xl-core, the rewrite exploits the payload's canonicality (serialized by
 * scala.xml: attributes are double-quoted, and `"` / `>` are escaped EVERYWHERE, so a raw `"`
 * occurs only as an attribute delimiter and the first `>` past the prologue ends the start tag) to
 * rewrite exactly the root element's `sqref` attribute, byte-surgically.
 *
 * Guards are layered and all-or-nothing — any failure returns the payload unchanged (today's
 * behavior; never a partial rewrite):
 *   - an optional leading `<?…?>` prologue plus following whitespace is skipped (and spliced back
 *     verbatim on rewrite); every guard below applies to the element region after it,
 *   - the element region must start with `<` + rootLabel (label ended by whitespace, `/` or `>`),
 *   - the match region is the start tag only (up to the first `>` in the element region),
 *   - exactly ONE whitespace-preceded `sqref="…"` may match there (the whitespace guard rejects
 *     prefixed lookalikes like `xr:sqref`),
 *   - every whitespace-split token must parse as a range or single cell (the `CfCodec.parseSqref`
 *     grammar).
 *
 * Tokens whose shifted range equals their parse keep their ORIGINAL text (full-column tokens are
 * byte-identical fixed points under row inserts post-GH-428). All tokens collapsed → None (the
 * caller drops the entry, mirroring the typed-envelope algebra).
 *
 * Semantics note: relative refs in un-rewritten Preserved FORMULAS self-heal (Excel stores them
 * sqref-relative); absolute refs below the edit go stale — documented, and strictly better than
 * full inertness.
 */
private[xl] object SqrefShift:

  private val SqrefAttr = """(^|\s)sqref="([^"]*)"""".r

  private[xl] def shiftPayload(
    xml: String,
    rootLabel: String,
    shift: CellRange => Option[CellRange]
  ): Option[String] =
    // Reader payloads open with the XmlUtil.compact declaration; skip one optional `<?…?>`
    // prologue (plus whitespace) so the guards see the element region. A `<?` with no `?>`
    // keeps bodyStart at 0 and fails the root-label guard below — unchanged, like any other
    // guard failure.
    val bodyStart =
      if !xml.startsWith("<?") then 0
      else
        val close = xml.indexOf("?>")
        if close < 0 then 0
        else
          val afterDecl = close + 2
          afterDecl + xml.segmentLength(_.isWhitespace, afterDecl)
    val open = "<" + rootLabel
    val boundaryOk = xml.startsWith(open, bodyStart) && {
      val next = xml.lift(bodyStart + open.length)
      next.exists(c => c.isWhitespace || c == '/' || c == '>')
    }
    val gt = xml.indexOf('>', bodyStart)
    if !boundaryOk || gt < 0 then Some(xml)
    else
      val region = xml.substring(bodyStart, gt)
      SqrefAttr.findAllMatchIn(region).toList match
        case m :: Nil =>
          val tokens = m.group(2).trim.split("\\s+").toVector.filter(_.nonEmpty)
          val parsed = tokens.map { token =>
            CellRange
              .parse(token)
              .toOption
              .orElse(ARef.parse(token).toOption.map(r => CellRange(r, r)))
              .map(token -> _)
          }
          if tokens.isEmpty || parsed.exists(_.isEmpty) then Some(xml)
          else
            val shifted = parsed.flatten.flatMap { (text, range) =>
              shift(range).map(nr => if nr == range then text else toSqrefToken(nr))
            }
            if shifted.isEmpty then None // every token collapsed -> drop the entry
            else
              val newSqref = shifted.mkString(" ")
              if newSqref == m.group(2) then Some(xml)
              else
                Some(
                  xml.substring(0, bodyStart + m.start(2)) + newSqref +
                    xml.substring(bodyStart + m.end(2))
                )
        case _ => Some(xml)

  /**
   * Print one shifted sqref token in the shape Excel uses: full-column `A:C`, full-row `1:5`,
   * single cell `A1`, otherwise `A1:B2`. Only reached for tokens the shift CHANGED — identity
   * tokens keep their original text.
   */
  private def toSqrefToken(r: CellRange): String =
    if r.start == r.end then r.start.toA1
    else if r.isFullColumn then s"${r.start.col.toLetter}:${r.end.col.toLetter}"
    else if r.isFullRow then s"${r.start.row.index1}:${r.end.row.index1}"
    else r.toA1
