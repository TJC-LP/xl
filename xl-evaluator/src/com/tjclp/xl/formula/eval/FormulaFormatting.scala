package com.tjclp.xl.formula.eval

import com.tjclp.xl.formula.ast.TExpr
import com.tjclp.xl.formula.functions.ArgValue

import com.tjclp.xl.addressing.{ARef, CellRange, SheetName}
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.styles.numfmt.NumFmt
import com.tjclp.xl.workbooks.Workbook
import scala.annotation.{nowarn, tailrec}

/**
 * Number-format inheritance for formula cells (GH-184).
 *
 * Excel automatically formats a formula entered into a General cell using the number format of the
 * cells it references (`=B2-B3` over two currency cells displays as currency). xl keeps
 * `put`/`putf` format-neutral by default; this object provides the opt-in inference used by
 * [[SheetEvaluator.putFormulaInheriting]].
 *
 * Inheritance rules (simplified from Excel's observed first-reference behavior — Excel's full rules
 * are per-operator and undocumented; we implement the stable subset):
 *   - References with General (or no) format are ignored.
 *   - All remaining formats identical → that format.
 *   - Same category (currency / percent / date-time / plain numeric / text) but different precision
 *     → the first referenced format in formula source order wins.
 *   - Mixed categories (e.g. currency vs percent) → no inheritance (General).
 *
 * Range references are enumerated in row-major order, bounded by the target sheet's used range
 * (cells outside it carry no styles), so full-column references like `SUM(A:A)` stay cheap.
 */
object FormulaFormatting:

  /** Broad format families used to decide whether two formats are compatible for inheritance. */
  private enum FormatCategory derives CanEqual:
    case Currency, Percent, Temporal, Numeric, Text

  /**
   * Infer the number format a formula cell should inherit from the cells it references.
   *
   * @param expr
   *   Parsed formula AST (see [[com.tjclp.xl.formula.parser.FormulaParser.parse]])
   * @param sheet
   *   Sheet providing the context for unqualified references
   * @param workbook
   *   Workbook context for cross-sheet references (`=Data!B2`); without it, cross-sheet references
   *   simply contribute no format
   * @return
   *   `Some(fmt)` when the referenced formats agree per the rules above, `None` for General
   */
  def inferFormatFromReferences(
    expr: TExpr[?],
    sheet: Sheet,
    workbook: Option[Workbook] = None
  ): Option[NumFmt] =
    val formats = referencedFormats(expr, sheet, workbook)
    formats.headOption.flatMap { case (firstFmt, firstCat) =>
      if formats.forall { case (_, cat) => cat == firstCat } then Some(firstFmt)
      else None
    }

  // ===== Reference collection (ordered) =====

  /**
   * Resolved (format, category) pairs for each referenced cell, in formula source order.
   * General-formatted, uncategorizable, unstyled, and unresolvable references are skipped.
   */
  private def referencedFormats(
    expr: TExpr[?],
    sheet: Sheet,
    workbook: Option[Workbook]
  ): Vector[(NumFmt, FormatCategory)] =
    orderedRefs(expr, sheet, workbook).flatMap { case (target, ref) =>
      target.cells
        .get(ref)
        .flatMap(_.styleId)
        .flatMap(target.styleRegistry.get)
        .map(_.numFmt)
        .flatMap(fmt => category(fmt).map(fmt -> _))
    }

  /**
   * Referenced cells in formula source order (left-to-right traversal; ranges expand row-major).
   *
   * Order is load-bearing for the first-reference-wins rule, which is why this does not reuse
   * [[com.tjclp.xl.formula.graph.DependencyGraph.extractDependencies]] (Set-based, unordered, and
   * intra-sheet only). Bounding mirrors `extractDependenciesBounded`: ranges are clipped to the
   * target sheet's used range so full column/row references never materialize 1M+ cells.
   */
  // nowarn: compiler incorrectly reports PolyRef as unreachable (same as DependencyGraph)
  @nowarn("msg=Unreachable case")
  private def orderedRefs(
    expr: TExpr[?],
    sheet: Sheet,
    workbook: Option[Workbook]
  ): Vector[(Sheet, ARef)] =
    def resolve(name: SheetName): Option[Sheet] =
      if name == sheet.name then Some(sheet)
      else workbook.flatMap(wb => wb(name).toOption)

    def boundedCells(target: Sheet, range: CellRange): Vector[(Sheet, ARef)] =
      target.usedRange match
        case Some(bounds) =>
          range.intersect(bounds) match
            case Some(clipped) => clipped.cells.map(target -> _).toVector
            case None => Vector.empty
        case None => Vector.empty // empty sheet: no styled cells to inherit from

    def locationRefs(location: TExpr.RangeLocation): Vector[(Sheet, ARef)] =
      location match
        case TExpr.RangeLocation.Local(range) => boundedCells(sheet, range)
        case TExpr.RangeLocation.CrossSheet(name, range) =>
          resolve(name).map(boundedCells(_, range)).getOrElse(Vector.empty)

    def loop(e: TExpr[?]): Vector[(Sheet, ARef)] =
      e match
        // Single cell references (never bounded, mirroring DependencyGraph)
        case TExpr.Ref(at, _, _) => Vector((sheet, at))
        case TExpr.PolyRef(at, _) => Vector((sheet, at))
        case TExpr.SheetRef(name, at, _, _) =>
          resolve(name).map(target => Vector(target -> at)).getOrElse(Vector.empty)
        case TExpr.SheetPolyRef(name, at, _) =>
          resolve(name).map(target => Vector(target -> at)).getOrElse(Vector.empty)

        // Range references
        case TExpr.RangeRef(range) => boundedCells(sheet, range)
        case TExpr.SheetRange(name, range) =>
          resolve(name).map(boundedCells(_, range)).getOrElse(Vector.empty)
        case TExpr.Aggregate(_, location) => locationRefs(location)

        // Function calls: arguments in declaration order
        case call: TExpr.Call[?] =>
          call.spec.argSpec
            .toValues(call.args)
            .foldLeft(Vector.empty[(Sheet, ARef)]) { (acc, value) =>
              value match
                case ArgValue.Expr(inner) => acc ++ loop(inner)
                case ArgValue.Range(location) => acc ++ locationRefs(location)
                case ArgValue.Cells(range) => acc ++ boundedCells(sheet, range)
            }

        // Binary operators: left before right
        case TExpr.Add(l, r) => loop(l) ++ loop(r)
        case TExpr.Sub(l, r) => loop(l) ++ loop(r)
        case TExpr.Mul(l, r) => loop(l) ++ loop(r)
        case TExpr.Div(l, r) => loop(l) ++ loop(r)
        case TExpr.Pow(l, r) => loop(l) ++ loop(r)
        case TExpr.Concat(l, r) => loop(l) ++ loop(r)
        case TExpr.Eq(l, r) => loop(l) ++ loop(r)
        case TExpr.Neq(l, r) => loop(l) ++ loop(r)
        case TExpr.Lt(l, r) => loop(l) ++ loop(r)
        case TExpr.Lte(l, r) => loop(l) ++ loop(r)
        case TExpr.Gt(l, r) => loop(l) ++ loop(r)
        case TExpr.Gte(l, r) => loop(l) ++ loop(r)

        // Unary wrappers
        case TExpr.ToInt(inner) => loop(inner)
        case TExpr.DateToSerial(inner) => loop(inner)
        case TExpr.DateTimeToSerial(inner) => loop(inner)

        // Literals: no references
        case TExpr.Lit(_) => Vector.empty

    loop(expr)

  // ===== Format categorization =====

  private val currencySymbols = Set('$', '€', '£', '¥')
  private val temporalTokens = Set('y', 'm', 'd', 'h', 's')

  private def category(fmt: NumFmt): Option[FormatCategory] =
    fmt match
      case NumFmt.General => None
      case NumFmt.Currency => Some(FormatCategory.Currency)
      case NumFmt.Percent | NumFmt.PercentDecimal => Some(FormatCategory.Percent)
      case NumFmt.Date | NumFmt.DateTime | NumFmt.Time => Some(FormatCategory.Temporal)
      case NumFmt.Integer | NumFmt.Decimal | NumFmt.ThousandsSeparator | NumFmt.ThousandsDecimal |
          NumFmt.Scientific | NumFmt.Fraction =>
        Some(FormatCategory.Numeric)
      case NumFmt.Text => Some(FormatCategory.Text)
      case NumFmt.Custom(code) => categorizeCode(code)

  /**
   * Categorize a custom format code by inspecting its positive (first) section.
   *
   * Only the first section determines the category: the trailing text section of four-section
   * formats (`_(@_)` in accounting codes) must not classify a numeric format as Text. Quoted
   * literals, escapes, padding (`_x`), fill (`*x`), and bracket sections are stripped before token
   * analysis; currency symbols are detected on the unstripped section because accounting formats
   * quote them (`"$"`). A bare locale prefix `[$-404]` is not a currency marker, while `[$€-407]`
   * is.
   */
  private def categorizeCode(rawCode: String): Option[FormatCategory] =
    if rawCode.trim.equalsIgnoreCase("general") then None
    else
      val section = firstSection(rawCode).replaceAll("""\[\$-[^\]]*\]""", "")
      val tokens = section
        .replaceAll("\"[^\"]*\"", "")
        .replaceAll("""\\.""", "")
        .replaceAll("_.", "")
        .replaceAll("""\*.""", "")
        .replaceAll("""\[[^\]]*\]""", "")
      if tokens.contains("%") then Some(FormatCategory.Percent)
      else if section.exists(currencySymbols.contains) then Some(FormatCategory.Currency)
      else if tokens.exists(c => temporalTokens.contains(c.toLower)) then
        Some(FormatCategory.Temporal)
      else if tokens.contains("@") then Some(FormatCategory.Text)
      else Some(FormatCategory.Numeric)

  /** First `;`-separated section of a format code, ignoring `;` inside quotes/brackets/escapes. */
  private def firstSection(code: String): String =
    @tailrec
    def scan(i: Int, inQuote: Boolean, inBracket: Boolean): Int =
      if i >= code.length then i
      else
        code.charAt(i) match
          case '"' if !inBracket => scan(i + 1, !inQuote, inBracket)
          case '[' if !inQuote => scan(i + 1, inQuote, true)
          case ']' if !inQuote => scan(i + 1, inQuote, false)
          case '\\' if !inQuote && !inBracket => scan(i + 2, inQuote, inBracket)
          case ';' if !inQuote && !inBracket => i
          case _ => scan(i + 1, inQuote, inBracket)
    code.take(scan(0, false, false))
