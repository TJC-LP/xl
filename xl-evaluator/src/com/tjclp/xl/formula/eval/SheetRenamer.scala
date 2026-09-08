package com.tjclp.xl.formula.eval

import cats.syntax.all.*

import com.tjclp.xl.addressing.{ARef, SheetName}
import com.tjclp.xl.cells.{Cell, CellValue, FormulaKind}
import com.tjclp.xl.cf.{CfPoint, CfRule, Cfvo, ConditionalFormat}
import com.tjclp.xl.error.{XLError, XLResult}
import com.tjclp.xl.formula.graph.DependencyGraph.QualifiedRef
import com.tjclp.xl.formula.printer.FormulaOps
import com.tjclp.xl.sheets.{DataValidation, DvKind, Sheet}
import com.tjclp.xl.workbooks.{DefinedName, Workbook}

/**
 * GH-559: rename a sheet AND every reference to it (ADR-017 §2.9).
 *
 * `Workbook.rename` is formula-blind by design (xl-core has no parser): it changes the tab, remaps
 * typed chart references and the source tracker, and leaves `Sheet1!A1` in every other sheet's
 * formulas — a file that lints clean and opens in Excel as `#REF!`. This layer adds what xl-core
 * cannot: the sheet qualifier is rewritten in cell formulas on every sheet (the renamed sheet's own
 * self-qualified references included), in defined names (segment by segment through top-level comma
 * unions), and in typed conditional-format and data-validation formulas — quoted when the new name
 * needs it. Cached values and formula record kinds are preserved: a rename changes no value.
 *
 * Refuse-before-mutate: every text that mentions the old name is rewritten first, purely; if any
 * cannot be parsed the whole rename is `Left` and the workbook is returned untouched. An unknown
 * reference expression must not survive with silently changed meaning. [[renameLocated]] says WHERE
 * the unrewritable text lives ([[Site]]) so a caller can point at the cell (GH-608); [[rename]] is
 * the same decision as a plain `XLResult`.
 *
 * ADR-017 invariant 1: every changed sheet is written back through `Workbook.put`, never
 * `copy(sheets = …)`, so `SourceContext.modifiedSheets` names exactly the sheets whose XML must be
 * regenerated — the surgical writer copies every other sheet byte-for-byte from the source zip, and
 * a rename applied only in memory would leave `='Old'!A1` on disk. Pinned through a real file in
 * SheetRenamerSpec.
 *
 * Not rewritten (documented in LIMITATIONS.md): `Preserved` conditional-format, data-validation and
 * chart payloads (opaque XML re-emitted verbatim) and hyperlink locations.
 */
object SheetRenamer:

  /**
   * Where a refused rename's unrewritable text lives, named as the caller's file knows it: sheets
   * by their PRE-rename names.
   */
  enum Site derives CanEqual:
    /** A cell formula. */
    case Cell(sheet: SheetName, ref: ARef)

    /** A typed conditional-format formula on `sheet`. */
    case ConditionalFormat(sheet: SheetName)

    /** A typed data-validation formula on `sheet`. */
    case DataValidation(sheet: SheetName)

    /** A workbook defined name's `refersTo`. */
    case Name(name: String)

    /**
     * The site as the refusal message spells it: `Summary!I23` / `'Q1 Data'!C3` (a cell, quoted the
     * way a formula would), `Sheet1!conditional format`, `Sheet1!data validation`, `defined name
     * 'Total'`.
     */
    def describe: String = this match
      case Cell(sheet, ref) => s"${SheetName.quoteForFormula(sheet.value)}!${ref.toA1}"
      case ConditionalFormat(sheet) =>
        s"${SheetName.quoteForFormula(sheet.value)}!conditional format"
      case DataValidation(sheet) => s"${SheetName.quoteForFormula(sheet.value)}!data validation"
      case Name(name) => s"defined name '$name'"

  /**
   * A refused rename. `error` is what [[rename]] reports — a `FormulaError` whose message already
   * names the site, or one of `Workbook.rename`'s own refusals; `site` is where the unrewritable
   * text lives, `None` for the workbook-level refusals (`SheetNotFound`, `DuplicateSheet`).
   */
  final case class Refusal(site: Option[Site], error: XLError) derives CanEqual

  private type Refused[A] = Either[Refusal, A]

  /**
   * `Workbook.rename(from, to)` plus the reference rewrite described above. `Left` with the same
   * errors `Workbook.rename` reports (`SheetNotFound`, `DuplicateSheet`) or a `FormulaError` naming
   * the first text that mentions `from` but cannot be parsed — in every `Left` case the input
   * workbook is untouched. [[renameLocated]] with the [[Site]] dropped.
   */
  def rename(wb: Workbook, from: SheetName, to: SheetName): XLResult[Workbook] =
    renameLocated(wb, from, to).left.map(_.error)

  /**
   * [[rename]], with every refusal carrying WHERE it fired: the cell, the sheet whose conditional
   * format or data validation, or the defined name whose text cannot be rewritten (GH-608).
   */
  def renameLocated(wb: Workbook, from: SheetName, to: SheetName): Either[Refusal, Workbook] =
    if from == to then wb.rename(from, to).left.map(Refusal(None, _))
    else
      for
        // `Workbook.rename` FIRST: it is pure, so a later `Left` simply discards `renamed` and the
        // caller still holds `wb` untouched, and its SheetNotFound/DuplicateSheet refusals take
        // precedence over a formula refusal. The rewrite then runs over the RENAMED sheets, so the
        // GH-222 typed-chart remap `Workbook.rename` performed on every sheet is carried forward
        // (rewriting the pre-rename sheets and putting them back would overwrite it).
        renamed <- wb.rename(from, to).left.map(Refusal(None, _))
        rewrittenSheets <- renamed.sheets
          .zip(wb.sheets)
          .traverse((sheet, original) => rewriteSheet(sheet, original.name, from, to))
        rewrittenNames <- renamed.metadata.definedNames.traverse(name =>
          rewriteName(name, from, to)
        )
      yield
        // Invariant 1: every changed sheet goes back through `put`, which marks it modified. The
        // sheet already carries the name its tab now has, so `put` replaces it in place.
        val withSheets = rewrittenSheets.foldLeft(renamed) {
          case (acc, Some(rewritten)) => acc.put(rewritten)
          case (acc, None) => acc
        }
        if rewrittenNames == renamed.metadata.definedNames then withSheets
        else
          withSheets.copy(
            metadata = withSheets.metadata.copy(definedNames = rewrittenNames),
            sourceContext = withSheets.sourceContext.map(_.markMetadataModified)
          )

  /**
   * The cells whose formula text mentions `sheet` (the text gate `FormulaOps.mentionsSheet`; string
   * literals and external-workbook references do not count), in sheet order then row-major — what
   * `rename` would rewrite, for `deps`/`describe`-style reporting.
   */
  def references(wb: Workbook, sheet: SheetName): Vector[QualifiedRef] =
    wb.sheets.flatMap { s =>
      s.cells.iterator
        .collect {
          case (ref, cell) if mentions(cell.value, sheet) => QualifiedRef(s.name, ref)
        }
        .toVector
        .sortBy(q => (q.ref.row.index0, q.ref.col.index0))
    }

  private def mentions(value: CellValue, sheet: SheetName): Boolean = value match
    case CellValue.Formula(_, _, _: FormulaKind.DataTable) => false
    case CellValue.Formula(text, _, _) => FormulaOps.mentionsSheet(text, sheet)
    case _ => false

  /**
   * `Some(rewritten)` when anything on the sheet changed, `None` when it rides untouched.
   * `labelName` is the sheet's PRE-rename name: refusals must name the sheet as the caller's file
   * knows it.
   */
  private def rewriteSheet(
    sheet: Sheet,
    labelName: SheetName,
    from: SheetName,
    to: SheetName
  ): Refused[Option[Sheet]] =
    def rewriteText(site: Site)(text: String): Refused[String] =
      FormulaOps.renameSheet(text, from, to).left.map(located(site, _))
    for
      changedCells <- sheet.cells.toVector
        .sortBy((ref, _) => (ref.row.index0, ref.col.index0))
        .traverse((ref, cell) => rewriteCell(ref, cell, rewriteText(Site.Cell(labelName, ref))))
        .map(_.flatten)
      cfs <- sheet.conditionalFormats.traverse(
        rewriteCf(_, rewriteText(Site.ConditionalFormat(labelName)))
      )
      dvs <- sheet.dataValidations.traverse(
        rewriteDv(_, rewriteText(Site.DataValidation(labelName)))
      )
    yield
      val cfChanged = cfs != sheet.conditionalFormats
      val dvChanged = dvs != sheet.dataValidations
      if changedCells.isEmpty && !cfChanged && !dvChanged then None
      else
        Some(
          sheet.copy(
            cells = sheet.cells ++ changedCells,
            conditionalFormats = if cfChanged then cfs else sheet.conditionalFormats,
            dataValidations = if dvChanged then dvs else sheet.dataValidations
          )
        )

  /** A formula cell whose text changed, with its cache and record kind carried unchanged. */
  private def rewriteCell(
    ref: ARef,
    cell: Cell,
    rewrite: String => Refused[String]
  ): Refused[Option[(ARef, Cell)]] =
    cell.value match
      // GH-430: a data-table record's TABLE(...) text is derived from local geometry, never parsed
      case CellValue.Formula(_, _, _: FormulaKind.DataTable) => Right(None)
      case CellValue.Formula(text, cached, kind) =>
        rewrite(text).map { rewritten =>
          if rewritten == text then None
          else Some(ref -> cell.copy(value = CellValue.Formula(rewritten, cached, kind)))
        }
      case _ => Right(None)

  private def rewriteName(
    name: DefinedName,
    from: SheetName,
    to: SheetName
  ): Refused[DefinedName] =
    // A refersTo may be a top-level comma union the parser rejects as a whole; each segment is a
    // formula of its own (the StructuralEditor convention).
    StructuralEditor
      .splitTopLevelCommas(name.formula)
      .traverse(segment => FormulaOps.renameSheet(segment, from, to))
      .map(segments => segments.mkString(","))
      .map(rewritten => if rewritten == name.formula then name else name.copy(formula = rewritten))
      .left
      .map(located(Site.Name(name.name), _))

  /**
   * Typed CF formulas follow the rename (CellIs.formula1/formula2, Expression.formula, Cfvo.Formula
   * in color-scale points and data-bar bounds). Text-family rules carry text, not references; Top10
   * has no formula; Preserved payloads are opaque XML and ride verbatim.
   */
  private def rewriteCf(
    cf: ConditionalFormat,
    rewrite: String => Refused[String]
  ): Refused[ConditionalFormat] =
    def cfvo(v: Cfvo): Refused[Cfvo] = v match
      case Cfvo.Formula(f) => rewrite(f).map(Cfvo.Formula.apply)
      case other => Right(other)
    def point(p: CfPoint): Refused[CfPoint] = cfvo(p.cfvo).map(c => p.copy(cfvo = c))
    cf match
      case ConditionalFormat.Rules(ranges, rules, pivot) =>
        rules
          .traverse[Refused, CfRule] {
            case r: CfRule.CellIs =>
              for
                f1 <- rewrite(r.formula1)
                f2 <- r.formula2.traverse(rewrite)
              yield r.copy(formula1 = f1, formula2 = f2)
            case r: CfRule.Expression => rewrite(r.formula).map(f => r.copy(formula = f))
            case r: CfRule.ColorScale =>
              for
                min <- point(r.min)
                mid <- r.mid.traverse(point)
                max <- point(r.max)
              yield r.copy(min = min, mid = mid, max = max)
            case r: CfRule.DataBar =>
              for
                min <- cfvo(r.min)
                max <- cfvo(r.max)
              yield r.copy(min = min, max = max)
            case other @ (_: CfRule.Top10 | _: CfRule.Text | _: CfRule.Preserved) => Right(other)
          }
          .map(rewritten => ConditionalFormat.Rules(ranges, rewritten, pivot))
      case preserved: ConditionalFormat.Preserved => Right(preserved)

  /**
   * Typed DV formulas follow the rename (List/Custom/Bounded). An inline list literal (`"yes,no"`)
   * never mentions a sheet and rides verbatim; Preserved payloads are opaque XML.
   */
  private def rewriteDv(
    dv: DataValidation,
    rewrite: String => Refused[String]
  ): Refused[DataValidation] =
    dv match
      case rules: DataValidation.Rules =>
        val kind: Refused[DvKind] = rules.kind match
          case DvKind.List(f) => rewrite(f).map(DvKind.List.apply)
          case DvKind.Custom(f) => rewrite(f).map(DvKind.Custom.apply)
          case DvKind.AnyValue => Right(DvKind.AnyValue)
          case DvKind.Bounded(t, op, f1, f2) =>
            for
              r1 <- rewrite(f1)
              r2 <- f2.traverse(rewrite)
            yield DvKind.Bounded(t, op, r1, r2)
        kind.map(k => rules.copy(kind = k))
      case preserved: DataValidation.Preserved => Right(preserved)

  /**
   * Attach the site to a refusal, and prefix the message with it so a text reader sees where the
   * offending text lives (`Summary!I23: Cannot rewrite …`); the formula text itself is kept.
   */
  private def located(site: Site, error: XLError): Refusal = error match
    case XLError.FormulaError(formula, reason) =>
      Refusal(Some(site), XLError.FormulaError(formula, s"${site.describe}: $reason"))
    case other => Refusal(Some(site), other)
