package com.tjclp.xl.formula.printer

import java.util.regex.Pattern

import com.tjclp.xl.addressing.{ARef, SheetName}
import com.tjclp.xl.error.{XLError, XLResult}
import com.tjclp.xl.formula.ast.TExpr
import com.tjclp.xl.formula.parser.{FormulaParser, ParseError}
import com.tjclp.xl.ops.FormulaSupport

/**
 * String-in, string-out formula rewriting (ADR-017 §2.9): parse → transform → reprint, total.
 *
 * Every path that rewrites formula TEXT — structural edits, `putf --from`, `copy`, `batch`, the
 * sheet renamer — is the same three steps; this object is the one place they live. Conventions:
 *   - The caller's leading `=` is preserved: `=A1` shifts to `=B2`, bare `A1` (the model's
 *     equals-free cell text, GH-427; CF/DV/defined-name text) shifts to bare `B2`.
 *   - Output is Excel's FILE form (`FormulaPrinter.printFileForm`: bare `,` separators), the form
 *     every rewrite path writes back into cell text (GH-484).
 *   - Text that needs no rewriting is returned byte-identical — never parsed and reprinted — so a
 *     non-participant keeps its spacing and casing exactly.
 */
object FormulaOps:

  /**
   * Rewrite every reference to sheet `from` so it names `to` (case-insensitive match, quoted on
   * output when the new name needs it — `'Q1 Data'!A1`). Text that does not mention `from` —
   * including a string literal that spells it, an external-workbook reference (`[2]Sheet1!A1`) or a
   * sibling whose name merely contains it (`Sheet10!A1`) — comes back byte-identical, parseable or
   * not. `Left(FormulaError)` only when text that DOES mention the sheet cannot be parsed: an
   * unknown reference expression must not survive a rename with silently changed meaning. The
   * reason carries the parser's diagnostic (`ParseError.describe`), the text `eval` and `putf`
   * print for the same formula.
   */
  def renameSheet(text: String, from: SheetName, to: SheetName): XLResult[String] =
    if from == to || !mentionsSheet(text, from) then Right(text)
    else
      FormulaParser.parse(text) match
        case Left(err) =>
          Left(
            XLError.FormulaError(
              text,
              s"Cannot rewrite its reference to sheet '${from.value}': ${ParseError.describe(err)}"
            )
          )
        case Right(expr) if !FormulaShifter.mentionsSheet(expr, from.value) => Right(text)
        case Right(expr) => Right(reprint(text, FormulaShifter.renameSheet(expr, from, to)))

  /**
   * GH-936: `text` with every sheet qualifier spelled as the sheet's own name among `names`, as
   * Excel stores a formula when it is entered (`=sheet1!A1` becomes `=Sheet1!A1`, `='my data'!A1`
   * becomes `='My Data'!A1`, `=data!rate` becomes `=Data!rate`). A qualifier resolves by
   * [[SheetName.indexIn]]'s rule (#752); one that matches no sheet, or folds onto two tabs
   * differing only by case, stays as typed, and an external-workbook reference (`[1]sheet1!A1`) is
   * never touched. Total: text that needs no respelling — including text that does not parse, which
   * the caller's own gate reports — comes back byte-identical, so a formula already in canonical
   * form keeps its spacing and casing exactly. For newly authored formulas only: a formula read
   * from a file is never passed through here.
   */
  def canonicalSheets(text: String, names: IndexedSeq[SheetName]): String =
    if names.isEmpty || !text.contains('!') then text
    else
      FormulaParser.parse(text) match
        case Left(_) => text
        case Right(expr) =>
          val canonical = FormulaShifter.canonicalSheets(expr, names)
          if canonical == expr then text
          else
            // Excel changes only the qualifier's spelling: rewrite those spans, keep every other
            // byte, and keep the text as typed when the scanner or the parser disagree
            FormulaSpans
              .scan(text)
              .map { scan =>
                val edits = scan.qualifiers.flatMap { q =>
                  SheetName(q.name).toOption
                    .flatMap(typed => SheetName.indexIn(names, typed).toOption.map(names(_)))
                    .filter(_.value != q.name)
                    .map(own => (q.start, q.end, SheetName.quoteForFormula(own.value)))
                }
                FormulaSpans.replace(text, edits)
              }
              .filter(rewritten => sameFormula(rewritten, canonical))
              .getOrElse(text)

  /** Whether `text` parses to a formula that prints as `expected` does. */
  private def sameFormula(text: String, expected: TExpr[?]): Boolean =
    FormulaParser
      .parse(text)
      .exists(e => FormulaPrinter.printFileForm(e) == FormulaPrinter.printFileForm(expected))

  /**
   * GH-935: `text` with every reference to the single cell `sheet!from` respelled as `sheet!to`
   * ([[FormulaShifter.relocateCell]]), as Excel rewrites a cut cell's readers on paste. `host` is
   * the sheet the formula lives on (`None` for a defined name: only qualified references count).
   * Total: text that names no such reference, or does not parse, comes back byte-identical.
   */
  def relocateCell(
    text: String,
    host: Option[SheetName],
    sheet: SheetName,
    from: ARef,
    to: ARef
  ): String =
    if from == to || !text.replace("$", "").toUpperCase(java.util.Locale.ROOT).contains(from.toA1)
    then text
    else
      FormulaParser.parse(text) match
        case Left(_) => text
        case Right(expr) =>
          val moved = FormulaShifter.relocateCell(expr, host, sheet, from, to)
          if moved == expr then text
          else
            // Excel's paste rewrites only the reference: replace those spans, keep every other
            // byte; when the scanner and the parser disagree, the parser's reprint is the answer
            def onSheet(q: Option[FormulaSpans.Qualifier]): Boolean =
              q.fold(host.exists(_.value.equalsIgnoreCase(sheet.value)))(
                _.name.equalsIgnoreCase(sheet.value)
              )
            FormulaSpans
              .scan(text)
              .map { scan =>
                val edits = scan.references.filter(r => onSheet(r.qualifier)).flatMap {
                  case FormulaSpans.Reference(_, first, None) if first.ref == from =>
                    Vector((first.start, first.end, first.spelledAt(to)))
                  case FormulaSpans.Reference(_, first, Some(second))
                      if first.ref == from && second.ref == from =>
                    Vector(
                      (first.start, first.end, first.spelledAt(to)),
                      (second.start, second.end, second.spelledAt(to))
                    )
                  case _ => Vector.empty
                }
                FormulaSpans.replace(text, edits)
              }
              .filter(rewritten => sameFormula(rewritten, moved))
              .getOrElse(reprint(text, moved))

  /**
   * True when `text` carries a sheet qualifier naming `sheet` — bare (`Sheet1!`), quoted (`'Q1
   * Data'!`, apostrophes doubled), case-insensitive, outside string literals. A qualifier preceded
   * by an identifier character, `]` or `'` is not a match: `MySheet1!A1` is another sheet and
   * `[2]Sheet1!A1` another workbook. Either end of a 3-D range counts too — `Sheet1:Sheet3!A1`,
   * `'Q1 Data':'Q3 Data'!A1`, `'Sheet1:Sheet 3'!A1` — so a rename of either end reaches the parser
   * (which has no 3-D support) and refuses instead of leaving one end stale. This is the text gate
   * the structural refusal (`StructuralEditor`) and the renamer share; the AST decides the rest.
   */
  def mentionsSheet(text: String, sheet: SheetName): Boolean =
    qualifierPattern(sheet).matcher(maskStringLiterals(text)).find()

  /**
   * `text` with the contents of its string literals blanked (same length, quotes kept), so the
   * qualifier scan never matches inside one. A `"` inside a single-quoted sheet name is part of the
   * name, not a literal boundary (GH-653: splitting the text on `"` cut `'Q1 "Final"'!A1` in half
   * and a rename of that sheet silently left the reference stale); a `'` inside a string literal
   * likewise opens no quoted name.
   */
  private def maskStringLiterals(text: String): String =
    val (masked, _, _) =
      text.foldLeft((new StringBuilder(text.length), false, false)) {
        case ((sb, inString, inQuotedName), c) =>
          if inString then (sb.append(if c == '"' then c else ' '), c != '"', inQuotedName)
          else if inQuotedName then (sb.append(c), inString, c != '\'')
          else (sb.append(c), c == '"', c == '\'')
      }
    masked.toString

  /**
   * Shift every relative reference by (`colDelta`, `rowDelta`) the way a fill-drag does — anchors
   * respected, whole-column / whole-row references moving only along their own axis (GH-612), and a
   * reference that would leave the grid written as `#REF!` exactly as `FormulaShifter.shift` writes
   * it. `Left(FormulaError)` when the text cannot be parsed.
   */
  def shift(text: String, colDelta: Int, rowDelta: Int): XLResult[String] =
    shiftReporting(text, colDelta, rowDelta).map(_.formula)

  /**
   * GH-628: [[shift]], also reporting the references the shift carried off the grid and wrote as
   * `#REF!`, spelled as they were in `text` — the `FormulaSupport.shiftReporting` behind fill and
   * copy, and the `putf`/batch drag's source for the CLI's `OFF_GRID_REF` warning.
   */
  def shiftReporting(text: String, colDelta: Int, rowDelta: Int): XLResult[FormulaSupport.Shifted] =
    FormulaParser.parse(text) match
      case Left(err) =>
        Left(XLError.FormulaError(text, s"Cannot shift references: ${ParseError.describe(err)}"))
      case Right(expr) =>
        val shifted = FormulaShifter.shiftReporting(expr, colDelta, rowDelta)
        Right(FormulaSupport.Shifted(reprint(text, shifted.expr), shifted.voided))

  /** File-form reprint that keeps the caller's leading-'=' convention. */
  private def reprint(original: String, expr: TExpr[?]): String =
    val body = FormulaPrinter.printFileForm(expr)
    if original.startsWith("=") then s"=$body" else body

  private def qualifierPattern(sheet: SheetName): Pattern =
    val bare = Pattern.quote(sheet.value)
    val inner = Pattern.quote(sheet.value.replace("'", "''"))
    val quoted = s"'$inner'"
    // The other end of a 3-D range: a bare identifier-shaped name or a quoted one.
    val otherBare = "[\\p{L}\\p{N}_.]+"
    val otherQuoted = "'(?:[^']|'')*'"
    val notPreceded = "(?<![\\p{L}\\p{N}_.\\\\'\\]])"
    Pattern.compile(
      s"(?iu)$notPreceded(?:" +
        // Sheet1!A1, 'Q1 Data'!A1 — and the unquoted END of a 3-D range (Sheet1:Sheet3!A1)
        s"(?:$bare|$quoted)\\s*!" +
        // the START of a 3-D range, each end bare or quoted: Sheet1:Sheet3!A1, 'Q1 Data':'Q3 Data'!A1
        s"|(?:$bare|$quoted)\\s*:\\s*(?:$otherBare|$otherQuoted)\\s*!" +
        // a 3-D range quoted as a whole, this sheet at the start or the end: 'Sheet1:Sheet 3'!A1
        s"|'$inner:(?:[^']|'')*'\\s*!" +
        s"|'(?:[^']|'')*:$inner'\\s*!" +
        ")"
    )
