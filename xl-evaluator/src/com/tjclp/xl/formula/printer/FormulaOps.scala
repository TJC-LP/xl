package com.tjclp.xl.formula.printer

import java.util.regex.Pattern

import com.tjclp.xl.addressing.SheetName
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
   * True when `text` carries a sheet qualifier naming `sheet` — bare (`Sheet1!`), quoted (`'Q1
   * Data'!`, apostrophes doubled), case-insensitive, outside string literals. A qualifier preceded
   * by an identifier character, `]` or `'` is not a match: `MySheet1!A1` is another sheet and
   * `[2]Sheet1!A1` another workbook. Either end of a 3-D range counts too — `Sheet1:Sheet3!A1`,
   * `'Q1 Data':'Q3 Data'!A1`, `'Sheet1:Sheet 3'!A1` — so a rename of either end reaches the parser
   * (which has no 3-D support) and refuses instead of leaving one end stale. This is the text gate
   * the structural refusal (`StructuralEditor`) and the renamer share; the AST decides the rest.
   */
  def mentionsSheet(text: String, sheet: SheetName): Boolean =
    val qualifier = qualifierPattern(sheet)
    text.split("\"", -1).iterator.zipWithIndex.exists { (part, index) =>
      index % 2 == 0 && qualifier.matcher(part).find()
    }

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
