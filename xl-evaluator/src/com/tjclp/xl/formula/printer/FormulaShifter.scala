package com.tjclp.xl.formula.printer

import com.tjclp.xl.formula.ast.{RangeForm, TExpr}
import com.tjclp.xl.formula.functions.{FunctionSpec, FunctionSpecs}

import scala.annotation.nowarn

import com.tjclp.xl.{Anchor, CellRange}
import com.tjclp.xl.addressing.{ARef, Column, Row, SheetName}
import com.tjclp.xl.cells.CellError
import TExpr.RangeLocation

/**
 * AST transformer that shifts cell references for formula dragging.
 *
 * When a formula is applied across a range (e.g., B2:B100), references adjust based on their anchor
 * mode:
 *   - Relative (A1): Both column and row adjust
 *   - AbsCol ($A1): Column fixed, row adjusts
 *   - AbsRow (A$1): Column adjusts, row fixed
 *   - Absolute ($A$1): Both fixed
 *
 * Example: =A1*$B$1 applied from B2 to B5 becomes:
 *   - B2: =A1*$B$1 (no shift, delta = (0,0))
 *   - B3: =A2*$B$1 (A1→A2 shifted, $B$1 fixed)
 *   - B4: =A3*$B$1
 *   - B5: =A4*$B$1
 *
 * GH-612: a whole-column reference (`E:E`, [[RangeForm.Columns]]) moves only along columns and a
 * whole-row reference (`3:3`) only along rows — the other axis is "all of them", not a coordinate.
 * A reference that a shift would carry off the grid (before A1, past XFD1048576) becomes the error
 * literal `#REF!`, per reference, exactly as Excel writes when copying `=A1` from B2 to B1; a range
 * in a range-typed argument slot becomes [[RangeLocation.Error]] the same way (`SUM(A1:A2)` →
 * `SUM(#REF!)`).
 *
 * GH-628: [[shiftReporting]] / [[shiftStructuralReporting]] also return the references a shift
 * voided, spelled as they were, so a caller can warn about a `#REF!` it produced without grepping a
 * text that may legitimately contain one.
 *
 * Laws:
 *   - Identity: shift(expr, 0, 0) == expr
 *   - Commutativity: shift(shift(expr, c1, r1), c2, r2) == shift(expr, c1+c2, r1+r2) while every
 *     reference stays on the grid and no relative corner overtakes an anchored one (the normalised
 *     print of a crossed range swaps the `$`, which is not invertible — Excel has the same
 *     asymmetry; pinned in RangeFormSpec)
 *   - Anchor preservation: Anchor of shifted ref equals original anchor
 */
object FormulaShifter:

  /**
   * GH-628: the result of a reporting shift — the rewritten expression and the references the shift
   * VOIDED to `#REF!`, each spelled as it was before the shift (`A1`, `$A:$A`, `Data!B2:B4`). A
   * formula may legitimately contain `#REF!` already, so callers that want to say "this drag
   * produced one" read `voided` instead of grepping the text.
   */
  final case class Shifted[A](expr: TExpr[A], voided: Vector[String]) derives CanEqual:
    def anyVoided: Boolean = voided.nonEmpty

  /** The voided references of one shift, appended as the traversal meets them (document order). */
  private final class VoidedLog:
    private val buffer = scala.collection.mutable.ArrayBuffer.empty[String]
    def record(reference: String): Unit = buffer += reference
    def result: Vector[String] = buffer.toVector

  /** GH-612: what an off-grid reference becomes, in expression and in range-slot position. */
  private val refError: TExpr[Nothing] = TExpr.ErrorLit(CellError.Ref)
  private val refErrorLocation: RangeLocation = RangeLocation.Error(CellError.Ref)

  // The spellings a voided reference is reported under — the printer's own, so a report reads
  // exactly as the formula did
  private def refText(at: ARef, anchor: Anchor): String = FormulaPrinter.formatARef(at, anchor)
  private def sheetRefText(sheet: SheetName, at: ARef, anchor: Anchor): String =
    s"${FormulaPrinter.formatSheetName(sheet)}!${refText(at, anchor)}"
  private def externalRefText(index: Int, name: String, at: ARef, anchor: Anchor): String =
    s"${FormulaPrinter.formatExternalSheet(index, name)}!${refText(at, anchor)}"
  private def rangeText(range: CellRange, form: RangeForm): String =
    FormulaPrinter.formatRange(range, form)
  private def sheetRangeText(sheet: SheetName, range: CellRange, form: RangeForm): String =
    s"${FormulaPrinter.formatSheetName(sheet)}!${rangeText(range, form)}"
  private def externalRangeText(index: Int, name: String, range: CellRange, form: RangeForm) =
    s"${FormulaPrinter.formatExternalSheet(index, name)}!${rangeText(range, form)}"
  private def locationText(location: RangeLocation): String =
    FormulaPrinter.formatLocation(location)

  /**
   * Shift all cell references in the expression by the given deltas.
   *
   * @param expr
   *   The formula AST to transform
   * @param colDelta
   *   Number of columns to shift (positive = right)
   * @param rowDelta
   *   Number of rows to shift (positive = down)
   * @return
   *   Transformed AST with shifted references; a reference the shift carries off the grid is the
   *   error literal `#REF!` (GH-612). [[shiftReporting]] also says which ones.
   */
  def shift[A](expr: TExpr[A], colDelta: Int, rowDelta: Int): TExpr[A] =
    shiftReporting(expr, colDelta, rowDelta).expr

  /** GH-628: [[shift]], reporting the references that left the grid and were written as `#REF!`. */
  def shiftReporting[A](expr: TExpr[A], colDelta: Int, rowDelta: Int): Shifted[A] =
    if colDelta == 0 && rowDelta == 0 then Shifted(expr, Vector.empty)
    else
      val voidedLog = new VoidedLog
      val shifted = shiftInternal(expr, colDelta, rowDelta, voidedLog)
      Shifted(shifted, voidedLog.result)

  // Var: the deprecated bare-CellRange argument slot (ArgSpec.cellRange, no in-repo consumers) has
  // no error form, so an off-grid range there voids the call through a local var exactly as
  // shiftStructuralInternal flags a deleted one
  @SuppressWarnings(Array("org.wartremover.warts.AsInstanceOf", "org.wartremover.warts.Var"))
  @nowarn(
    "msg=Unreachable case"
  ) // PolyRef extends TExpr[Nothing], reachable via asInstanceOf casts
  private def shiftInternal[A](
    expr: TExpr[A],
    colDelta: Int,
    rowDelta: Int,
    voidedLog: VoidedLog
  ): TExpr[A] =
    import TExpr.*

    def go[B](e: TExpr[B]): TExpr[B] = shiftInternal(e, colDelta, rowDelta, voidedLog)

    // GH-612: an off-grid reference becomes #REF! — the whole reference, never a clamped or
    // non-existent address; GH-628: and is reported under its original spelling
    def orRefError(shifted: Option[TExpr[?]], original: => String): TExpr[A] =
      shifted match
        case Some(e) => e.asInstanceOf[TExpr[A]]
        case None =>
          voidedLog.record(original)
          refError.asInstanceOf[TExpr[A]]

    def goLocation(location: RangeLocation): RangeLocation =
      shiftLocation(location, colDelta, rowDelta, voidedLog)

    expr match
      // Cell references - apply anchor-aware shifting
      case Ref(at, anchor, decode) =>
        orRefError(
          shiftARef(at, anchor, colDelta, rowDelta).map(Ref(_, anchor, decode)),
          refText(at, anchor)
        )

      case PolyRef(at, anchor) =>
        orRefError(
          shiftARef(at, anchor, colDelta, rowDelta).map(PolyRef(_, anchor)),
          refText(at, anchor)
        )

      // Sheet-qualified references - shift the cell ref but keep the sheet
      case SheetRef(sheet, at, anchor, decode) =>
        orRefError(
          shiftARef(at, anchor, colDelta, rowDelta).map(SheetRef(sheet, _, anchor, decode)),
          sheetRefText(sheet, at, anchor)
        )

      case SheetPolyRef(sheet, at, anchor) =>
        orRefError(
          shiftARef(at, anchor, colDelta, rowDelta).map(SheetPolyRef(sheet, _, anchor)),
          sheetRefText(sheet, at, anchor)
        )

      case RangeRef(range, form) =>
        orRefError(
          shiftRange(range, form, colDelta, rowDelta).map(RangeRef(_, form)),
          rangeText(range, form)
        )

      case SheetRange(sheet, range, form) =>
        orRefError(
          shiftRange(range, form, colDelta, rowDelta).map(SheetRange(sheet, _, form)),
          sheetRangeText(sheet, range, form)
        )

      // GH-353: external-workbook references shift anchor-aware like sheet-qualified ones —
      // the workbook/sheet qualifier is fixed, the cell coordinates drag
      case ExternalRef(index, name, at, anchor) =>
        orRefError(
          shiftARef(at, anchor, colDelta, rowDelta).map(ExternalRef(index, name, _, anchor)),
          externalRefText(index, name, at, anchor)
        )

      case ExternalRange(index, name, range, form) =>
        orRefError(
          shiftRange(range, form, colDelta, rowDelta).map(ExternalRange(index, name, _, form)),
          externalRangeText(index, name, range, form)
        )

      // Literals - unchanged
      case lit: Lit[?] => lit.asInstanceOf[TExpr[A]]
      // GH-612: an error literal has no coordinates; GH-603: neither has an omitted argument
      case err: ErrorLit => err.asInstanceOf[TExpr[A]]
      case Missing => expr

      // Arithmetic operators
      case Add(x, y) => Add(go(x), go(y)).asInstanceOf[TExpr[A]]
      case Sub(x, y) => Sub(go(x), go(y)).asInstanceOf[TExpr[A]]
      case Mul(x, y) => Mul(go(x), go(y)).asInstanceOf[TExpr[A]]
      case Div(x, y) => Div(go(x), go(y)).asInstanceOf[TExpr[A]]
      case Pow(x, y) => Pow(go(x), go(y)).asInstanceOf[TExpr[A]]

      // String operators
      case Concat(x, y) => Concat(go(x), go(y)).asInstanceOf[TExpr[A]]

      // Comparison operators
      case Eq(x, y) => Eq(go(x), go(y)).asInstanceOf[TExpr[A]]
      case Neq(x, y) => Neq(go(x), go(y)).asInstanceOf[TExpr[A]]
      case Lt(x, y) => Lt(go(x), go(y)).asInstanceOf[TExpr[A]]
      case Lte(x, y) => Lte(go(x), go(y)).asInstanceOf[TExpr[A]]
      case Gt(x, y) => Gt(go(x), go(y)).asInstanceOf[TExpr[A]]
      case Gte(x, y) => Gte(go(x), go(y)).asInstanceOf[TExpr[A]]

      // Type conversion
      case ToInt(e) => ToInt(go(e)).asInstanceOf[TExpr[A]]

      // GH-374: unary plus is transparent — recurse so refs under it drag
      case UnaryPlus(e) => UnaryPlus(go(e))

      // GH-355: postfix percent — recurse so refs under it drag
      case Percent(e) => Percent(go(e)).asInstanceOf[TExpr[A]]

      // Arithmetic range functions (now using RangeLocation). GH-612: a range slot that falls off
      // the grid becomes RangeLocation.Error — SUM(#REF!), as Excel writes it
      case Aggregate(aggregatorId, location) =>
        Aggregate(aggregatorId, goLocation(location)).asInstanceOf[TExpr[A]]

      case call: Call[?] =>
        var voided = false
        val shifted =
          call.spec.argSpec.map(call.args)(
            expr => go(expr),
            goLocation,
            range =>
              shiftRange(range, RangeForm.Cells, colDelta, rowDelta).getOrElse {
                voided = true
                voidedLog.record(rangeText(range, RangeForm.Cells))
                range
              }
          )
        if voided then refError.asInstanceOf[TExpr[A]]
        else Call(call.spec, shifted).asInstanceOf[TExpr[A]]

      // Date-to-serial converters - shift inner expression
      case DateToSerial(dateExpr) => DateToSerial(go(dateExpr)).asInstanceOf[TExpr[A]]
      case DateTimeToSerial(dtExpr) => DateTimeToSerial(go(dtExpr)).asInstanceOf[TExpr[A]]

      // GH-193: LET — shift refs in binding values and body; binding names are not cell refs
      case Let(bindings, body) =>
        Let(
          bindings.map((name, value) => (name, go(value.asInstanceOf[TExpr[Any]]))),
          go(body)
        ).asInstanceOf[TExpr[A]]
      case bref: BindingRef => bref.asInstanceOf[TExpr[A]]
      // GH-384: defined names are identifiers, not coordinates — they never shift on drag
      case nref: NameRef => nref.asInstanceOf[TExpr[A]]
      // GH-394: sheet-qualified names are identifiers too
      case snref: SheetNameRef => snref.asInstanceOf[TExpr[A]]
      case cbref: CoercedBindingRef[?] => cbref.asInstanceOf[TExpr[A]]

      // GH-306: runtime coercion wrapper — shift the wrapped expression, preserve the wrapper
      case Coerced(inner, target) => Coerced(go(inner), target).asInstanceOf[TExpr[A]]

  /**
   * Shift one 0-based coordinate unless anchored; None when it leaves the grid (GH-612: Excel
   * writes `#REF!`, never a clamped or non-existent address). Long arithmetic so a pathological
   * delta cannot wrap.
   */
  private def shiftIndex(index0: Int, anchored: Boolean, delta: Int, max: Int): Option[Int] =
    if anchored then Some(index0)
    else
      val shifted = index0.toLong + delta
      Option.when(shifted >= 0 && shifted <= max)(shifted.toInt)

  /**
   * Shift a cell reference based on its anchor mode.
   *
   * @param ref
   *   The cell reference to shift
   * @param anchor
   *   The anchor mode determining which dimensions are fixed
   * @param colDelta
   *   Column shift amount
   * @param rowDelta
   *   Row shift amount
   * @return
   *   Shifted cell reference; None when either moving axis leaves the grid
   */
  private def shiftARef(
    cellRef: ARef,
    anchor: Anchor,
    colDelta: Int,
    rowDelta: Int
  ): Option[ARef] =
    for
      newCol <- shiftIndex(
        Column.index0(cellRef.col),
        anchor.isColAbsolute,
        colDelta,
        Column.MaxIndex0
      )
      newRow <- shiftIndex(Row.index0(cellRef.row), anchor.isRowAbsolute, rowDelta, Row.MaxIndex0)
    yield ARef.from0(newCol, newRow)

  /**
   * Shift a cell range by the given deltas, respecting per-endpoint anchors and (GH-612) the form
   * the range was written in.
   *
   * Each endpoint shifts according to its own anchor mode:
   *   - `$A$1:B10` → start ($A$1) is Absolute (fixed), end (B10) is Relative (shifts)
   *   - `$A1:B$10` → start has AbsCol (col fixed), end has AbsRow (row fixed)
   *
   * A whole-column range (`E:E`) moves only along columns and a whole-row range (`3:3`) only along
   * rows: the other axis spans the sheet and is not a coordinate. None when an endpoint leaves the
   * grid — the reference becomes `#REF!`.
   */
  private def shiftRange(
    range: CellRange,
    form: RangeForm,
    colDelta: Int,
    rowDelta: Int
  ): Option[CellRange] =
    // `actualFor`: a hand-built `Columns` form on a range that does not span every row is a corner
    // range and moves on both axes — the same guard the printer applies
    val (cd, rd) = form.actualFor(range) match
      case RangeForm.Cells | RangeForm.Cell => (colDelta, rowDelta)
      case RangeForm.Columns => (colDelta, 0)
      case RangeForm.Rows => (0, rowDelta)
    for
      newStart <- shiftARef(range.start, range.startAnchor, cd, rd)
      newEnd <- shiftARef(range.end, range.endAnchor, cd, rd)
    // The normalising constructor: when a relative corner overtakes an anchored one (`E:$E` dragged
    // right is `$E:F`, `A1:$B$1` is `$B$1:C1`) the corners swap and each anchor follows its corner,
    // as Excel prints. A DIAGONAL swap of mixed anchors (`A$1:$B2` overtaken on one axis only) still
    // swaps the anchors wholesale where Excel recombines them per axis.
    yield CellRange(newStart, newEnd, range.startAnchor, range.endAnchor)

  /**
   * Shift a RangeLocation by the given deltas.
   *
   * Handles both Local and CrossSheet locations, shifting the underlying CellRange. A range that
   * leaves the grid becomes [[RangeLocation.Error]] `#REF!` (GH-612), exactly as Excel writes
   * `SUM(#REF!)`, and is reported to `voidedLog` (GH-628).
   */
  private def shiftLocation(
    location: RangeLocation,
    colDelta: Int,
    rowDelta: Int,
    voidedLog: VoidedLog
  ): RangeLocation =
    def voided: RangeLocation =
      voidedLog.record(locationText(location))
      refErrorLocation
    location match
      case RangeLocation.Local(range, form) =>
        shiftRange(range, form, colDelta, rowDelta).fold(voided)(RangeLocation.Local(_, form))
      case RangeLocation.CrossSheet(sheet, range, form) =>
        shiftRange(range, form, colDelta, rowDelta)
          .fold(voided)(RangeLocation.CrossSheet(sheet, _, form))
      // GH-353: external-workbook range args drag anchor-aware like the TExpr.ExternalRange node
      case RangeLocation.External(index, name, range, form) =>
        shiftRange(range, form, colDelta, rowDelta)
          .fold(voided)(RangeLocation.External(index, name, _, form))
      // GH-394: a defined name is an identifier, not coordinates — shifting is a no-op
      // (its refersTo text lives in workbook metadata, not in this formula)
      case name @ RangeLocation.Name(_, _) => name
      // GH-612: an error has no coordinates
      case error @ RangeLocation.Error(_) => error

  // ============================================================================
  // GH-128 / GH-129: structural shifting for row/column insert & delete.
  //
  // Unlike `shift` (uniform fill-drag, anchor-aware), this is POSITION-CONDITIONAL and
  // ANCHOR-INDEPENDENT — structural edits move absolute refs too:
  //   - delta > 0: insert `delta` rows/cols at index `at`; refs at index >= at move +delta.
  //   - delta < 0: delete `-delta` rows/cols starting at `at`; refs after the band move by
  //     +delta; a ref/range that lands entirely inside the deleted band becomes the error
  //     literal `#REF!` IN PLACE (GH-629: `=A1+B3` with row 1 deleted is `=#REF!+B2`, as Excel
  //     writes it — the rest of the formula survives). Partially-overlapped ranges shrink to the
  //     surviving span.
  //
  // Only references that point at the EDITED sheet move: local refs move iff `shiftLocal`
  // (the formula lives on the edited sheet); sheet-qualified refs move iff their sheet name
  // equals `editedSheet` (case-insensitive).
  // ============================================================================

  /**
   * GH-455: true when `expr` contains a sheet-qualified reference to `editedSheet`
   * (case-insensitive, matching [[shiftStructural]]'s convention). Nodes the structural shift
   * treats as identity — local refs on a non-edited sheet, external refs, defined names (including
   * sheet-qualified ones) — do not count: a formula for which this is false is provably untouched
   * by a structural edit of `editedSheet` from another sheet, so callers can skip the
   * parse→shift→reprint cycle and keep its text and cached value byte-identical.
   */
  @SuppressWarnings(Array("org.wartremover.warts.Var"))
  def referencesSheet(expr: TExpr[?], editedSheet: String): Boolean =
    import TExpr.*
    def matches(sheet: com.tjclp.xl.SheetName): Boolean =
      sheet.value.equalsIgnoreCase(editedSheet)
    def goLoc(location: RangeLocation): Boolean = location match
      case RangeLocation.CrossSheet(sheet, _, _) => matches(sheet)
      case RangeLocation.Local(_, _) | RangeLocation.External(_, _, _, _) |
          RangeLocation.Name(_, _) | RangeLocation.Error(_) =>
        false
    def go(e: TExpr[?]): Boolean = e match
      case SheetRef(sheet, _, _, _) => matches(sheet)
      case SheetPolyRef(sheet, _, _) => matches(sheet)
      case SheetRange(sheet, _, _) => matches(sheet)
      case Aggregate(_, location) => goLoc(location)
      case call: Call[?] =>
        var found = false
        call.spec.argSpec.map(call.args)(
          arg => { if go(arg) then found = true; arg },
          loc => { if goLoc(loc) then found = true; loc },
          range => range
        )
        found
      case Add(x, y) => go(x) || go(y)
      case Sub(x, y) => go(x) || go(y)
      case Mul(x, y) => go(x) || go(y)
      case Div(x, y) => go(x) || go(y)
      case Pow(x, y) => go(x) || go(y)
      case Concat(x, y) => go(x) || go(y)
      case Eq(x, y) => go(x) || go(y)
      case Neq(x, y) => go(x) || go(y)
      case Lt(x, y) => go(x) || go(y)
      case Lte(x, y) => go(x) || go(y)
      case Gt(x, y) => go(x) || go(y)
      case Gte(x, y) => go(x) || go(y)
      case ToInt(inner) => go(inner)
      case UnaryPlus(inner) => go(inner)
      case Percent(inner) => go(inner)
      case DateToSerial(inner) => go(inner)
      case DateTimeToSerial(inner) => go(inner)
      case Coerced(inner, _) => go(inner)
      case Let(bindings, body) => bindings.exists((_, value) => go(value)) || go(body)
      case Lit(_) | ErrorLit(_) | Missing | Ref(_, _, _) | PolyRef(_, _) | RangeRef(_, _) |
          ExternalRef(_, _, _, _) | ExternalRange(_, _, _, _) | BindingRef(_) | NameRef(_) |
          SheetNameRef(_, _) | CoercedBindingRef(_, _) =>
        false
    go(expr)

  /**
   * GH-559: true when `expr` mentions `sheet` anywhere a rename must follow — every node
   * [[referencesSheet]] counts PLUS sheet-qualified defined names (`Sheet!name`,
   * [[TExpr.SheetNameRef]]). `referencesSheet` deliberately returns false for `SheetNameRef` (a
   * structural edit never moves an identifier) and `StructuralEditor` depends on that; a rename
   * changes the qualifier itself, so it needs this wider gate. External-workbook references never
   * count: `[2]Sheet1!A1` names a sheet in ANOTHER workbook.
   */
  def mentionsSheet(expr: TExpr[?], sheet: String): Boolean =
    referencesSheet(expr, sheet) || mentionsSheetName(expr, sheet)

  /** The `SheetNameRef` half of [[mentionsSheet]]: does any sheet-qualified NAME target `sheet`? */
  @SuppressWarnings(Array("org.wartremover.warts.Var"))
  private def mentionsSheetName(expr: TExpr[?], sheet: String): Boolean =
    import TExpr.*
    // GH-394: a sheet-qualified name can also sit in a range slot (`SUMIF(Model!rev_range, …)`)
    def goLoc(location: RangeLocation): Boolean = location match
      case RangeLocation.Name(_, Some(scope)) => scope.value.equalsIgnoreCase(sheet)
      case RangeLocation.Name(_, None) | RangeLocation.Local(_, _) |
          RangeLocation.CrossSheet(_, _, _) | RangeLocation.External(_, _, _, _) |
          RangeLocation.Error(_) =>
        false
    def go(e: TExpr[?]): Boolean = e match
      case SheetNameRef(qualifier, _) => qualifier.value.equalsIgnoreCase(sheet)
      case Aggregate(_, location) => goLoc(location)
      case call: Call[?] =>
        var found = false
        call.spec.argSpec.map(call.args)(
          arg => { if go(arg) then found = true; arg },
          loc => { if goLoc(loc) then found = true; loc },
          range => range
        )
        found
      case Add(x, y) => go(x) || go(y)
      case Sub(x, y) => go(x) || go(y)
      case Mul(x, y) => go(x) || go(y)
      case Div(x, y) => go(x) || go(y)
      case Pow(x, y) => go(x) || go(y)
      case Concat(x, y) => go(x) || go(y)
      case Eq(x, y) => go(x) || go(y)
      case Neq(x, y) => go(x) || go(y)
      case Lt(x, y) => go(x) || go(y)
      case Lte(x, y) => go(x) || go(y)
      case Gt(x, y) => go(x) || go(y)
      case Gte(x, y) => go(x) || go(y)
      case ToInt(inner) => go(inner)
      case UnaryPlus(inner) => go(inner)
      case Percent(inner) => go(inner)
      case DateToSerial(inner) => go(inner)
      case DateTimeToSerial(inner) => go(inner)
      case Coerced(inner, _) => go(inner)
      case Let(bindings, body) => bindings.exists((_, value) => go(value)) || go(body)
      case Lit(_) | ErrorLit(_) | Missing | Ref(_, _, _) | PolyRef(_, _) | RangeRef(_, _) |
          SheetRef(_, _, _, _) | SheetPolyRef(_, _, _) | SheetRange(_, _, _) |
          ExternalRef(_, _, _, _) | ExternalRange(_, _, _, _) | BindingRef(_) | NameRef(_) |
          CoercedBindingRef(_, _) =>
        false
    go(expr)

  /**
   * GH-559: rewrite every reference to sheet `from` (case-insensitive, matching
   * [[shiftStructural]]'s convention) so it names `to` instead — `SheetRef`, `SheetPolyRef`,
   * `SheetRange`, `RangeLocation.CrossSheet` in aggregate and function arguments, and
   * sheet-qualified defined names (`SheetNameRef`). Coordinates, anchors, decoders, local
   * references, defined names and external-workbook references (`[2]Sheet1!A1` lives in another
   * workbook) are untouched, so the result prints with the new qualifier and nothing else changed.
   *
   * Laws (pinned in SheetRenamerSpec): `renameSheet(renameSheet(e, a, b), b, a) == e` when `e`
   * never mentions `b`; `!mentionsSheet(renameSheet(e, a, b), a)` for `a != b`.
   */
  def renameSheet[A](expr: TExpr[A], from: SheetName, to: SheetName): TExpr[A] =
    if from == to then expr else renameSheetInternal(expr, from.value, to)

  @SuppressWarnings(Array("org.wartremover.warts.AsInstanceOf"))
  @nowarn("msg=Unreachable case")
  private def renameSheetInternal[A](expr: TExpr[A], from: String, to: SheetName): TExpr[A] =
    import TExpr.*
    def target(sheet: SheetName): SheetName =
      if sheet.value.equalsIgnoreCase(from) then to else sheet
    def go[B](e: TExpr[B]): TExpr[B] = renameSheetInternal(e, from, to)
    def goLocation(location: RangeLocation): RangeLocation = location match
      case RangeLocation.CrossSheet(sheet, range, form) =>
        RangeLocation.CrossSheet(target(sheet), range, form)
      // GH-394: a sheet-qualified name in a range slot (`SUMIF(Model!rev_range, …)`) follows too
      case RangeLocation.Name(name, Some(scope)) => RangeLocation.Name(name, Some(target(scope)))
      case other @ (RangeLocation.Local(_, _) | RangeLocation.External(_, _, _, _) |
          RangeLocation.Name(_, None) | RangeLocation.Error(_)) =>
        other

    expr match
      case SheetRef(sheet, at, anchor, decode) => SheetRef(target(sheet), at, anchor, decode)
      case SheetPolyRef(sheet, at, anchor) =>
        SheetPolyRef(target(sheet), at, anchor).asInstanceOf[TExpr[A]]
      case SheetRange(sheet, range, form) =>
        SheetRange(target(sheet), range, form).asInstanceOf[TExpr[A]]
      case SheetNameRef(sheet, name) => SheetNameRef(target(sheet), name).asInstanceOf[TExpr[A]]
      // Nothing to rename: local refs, literals, identifiers, external-workbook refs
      case _: Ref[?] | _: PolyRef | _: RangeRef | _: ExternalRef | _: ExternalRange | _: Lit[?] |
          _: ErrorLit | Missing | _: BindingRef | _: NameRef | _: CoercedBindingRef[?] =>
        expr
      case Add(x, y) => Add(go(x), go(y)).asInstanceOf[TExpr[A]]
      case Sub(x, y) => Sub(go(x), go(y)).asInstanceOf[TExpr[A]]
      case Mul(x, y) => Mul(go(x), go(y)).asInstanceOf[TExpr[A]]
      case Div(x, y) => Div(go(x), go(y)).asInstanceOf[TExpr[A]]
      case Pow(x, y) => Pow(go(x), go(y)).asInstanceOf[TExpr[A]]
      case Concat(x, y) => Concat(go(x), go(y)).asInstanceOf[TExpr[A]]
      case Eq(x, y) => Eq(go(x), go(y)).asInstanceOf[TExpr[A]]
      case Neq(x, y) => Neq(go(x), go(y)).asInstanceOf[TExpr[A]]
      case Lt(x, y) => Lt(go(x), go(y)).asInstanceOf[TExpr[A]]
      case Lte(x, y) => Lte(go(x), go(y)).asInstanceOf[TExpr[A]]
      case Gt(x, y) => Gt(go(x), go(y)).asInstanceOf[TExpr[A]]
      case Gte(x, y) => Gte(go(x), go(y)).asInstanceOf[TExpr[A]]
      case ToInt(e) => ToInt(go(e)).asInstanceOf[TExpr[A]]
      case UnaryPlus(e) => UnaryPlus(go(e))
      case Percent(e) => Percent(go(e)).asInstanceOf[TExpr[A]]
      case Aggregate(aggId, location) =>
        Aggregate(aggId, goLocation(location)).asInstanceOf[TExpr[A]]
      case call: Call[?] =>
        val renamed = call.spec.argSpec.map(call.args)(
          e => go(e),
          goLocation,
          range => range
        )
        Call(call.spec, renamed).asInstanceOf[TExpr[A]]
      case DateToSerial(e) => DateToSerial(go(e)).asInstanceOf[TExpr[A]]
      case DateTimeToSerial(e) => DateTimeToSerial(go(e)).asInstanceOf[TExpr[A]]
      case Let(bindings, body) =>
        Let(
          bindings.map((name, value) => (name, go(value.asInstanceOf[TExpr[Any]]))),
          go(body)
        ).asInstanceOf[TExpr[A]]
      case Coerced(inner, coercion) => Coerced(go(inner), coercion).asInstanceOf[TExpr[A]]

  /**
   * Structurally shift references in `expr` for a row/column insert (delta > 0) or delete (delta <
   * 0). A reference to a fully-deleted cell or range — or one an insert pushes past the sheet edge
   * (GH-428) — becomes the error literal `#REF!` in place, the rest of the formula intact (GH-629);
   * [[shiftStructuralReporting]] also says which references were voided.
   */
  def shiftStructural[A](
    expr: TExpr[A],
    shiftLocal: Boolean,
    editedSheet: String,
    isRow: Boolean,
    at: Int,
    delta: Int
  ): TExpr[A] =
    shiftStructuralReporting(expr, shiftLocal, editedSheet, isRow, at, delta).expr

  /** [[shiftStructural]] with the voided references reported under their original spelling. */
  def shiftStructuralReporting[A](
    expr: TExpr[A],
    shiftLocal: Boolean,
    editedSheet: String,
    isRow: Boolean,
    at: Int,
    delta: Int
  ): Shifted[A] =
    val voidedLog = new VoidedLog
    val shifted =
      shiftStructuralInternal(expr, shiftLocal, editedSheet, isRow, at, delta, voidedLog)
    Shifted(shifted, voidedLog.result)

  /**
   * Map a 0-based position through a structural edit. None = the position was deleted — or pushed
   * past the axis maximum `max` by an insert (GH-428): a ref past the sheet edge has no home, so
   * the formula degrades to `#REF!` (the `SharedFormula.shiftedIndex` bound). Long intermediate
   * math so a pathological delta cannot overflow.
   */
  private def shiftPos(p: Int, at: Int, delta: Int, max: Int): Option[Int] =
    if delta >= 0 then
      val np = if p >= at then p.toLong + delta else p.toLong
      Option.when(np <= max)(np.toInt)
    else
      val n = -delta
      if p < at then Some(p)
      else if p < at + n then None
      else Some(p - n)

  /**
   * Map an inclusive [s,e] position range through a structural edit. None = fully deleted — or
   * start pushed past the axis maximum by an insert (GH-428); an end past the maximum clamps to it
   * (Excel's behavior for full-height/width ranges).
   */
  private def shiftRangePos(s: Int, e: Int, at: Int, delta: Int, max: Int): Option[(Int, Int)] =
    if delta >= 0 then
      val ns = if s >= at then s.toLong + delta else s.toLong
      val ne = if e >= at then e.toLong + delta else e.toLong
      if ns > max then None else Some((ns.toInt, math.min(ne, max.toLong).toInt))
    else
      val n = -delta
      val ns = if s < at then s else if s < at + n then at else s - n
      val ne = if e < at then e else if e < at + n then at - 1 else e - n
      if ns > ne then None else Some((ns, ne))

  private def shiftARefStructural(
    cellRef: ARef,
    isRow: Boolean,
    at: Int,
    delta: Int
  ): Option[ARef] =
    val colIdx = Column.index0(cellRef.col)
    val rowIdx = Row.index0(cellRef.row)
    if isRow then shiftPos(rowIdx, at, delta, Row.MaxIndex0).map(nr => ARef.from0(colIdx, nr))
    else shiftPos(colIdx, at, delta, Column.MaxIndex0).map(nc => ARef.from0(nc, rowIdx))

  /**
   * GH-612: a whole-column range (`E:E`) has no row coordinates, so a ROW insert or delete leaves
   * it as it is (it still spans every row); a whole-row range is likewise untouched by a column
   * edit. Along its own axis it moves, widens, narrows or voids exactly like a corner range.
   */
  private def shiftRangeStructural(
    range: CellRange,
    form: RangeForm,
    isRow: Boolean,
    at: Int,
    delta: Int
  ): Option[CellRange] =
    val spansEditedAxis = form.actualFor(range) match
      case RangeForm.Columns => isRow
      case RangeForm.Rows => !isRow
      case RangeForm.Cells | RangeForm.Cell => false
    if spansEditedAxis then Some(range) else shiftCornersStructural(range, isRow, at, delta)

  private def shiftCornersStructural(
    range: CellRange,
    isRow: Boolean,
    at: Int,
    delta: Int
  ): Option[CellRange] =
    val (sPos, ePos) =
      if isRow then (Row.index0(range.start.row), Row.index0(range.end.row))
      else (Column.index0(range.start.col), Column.index0(range.end.col))
    val axisMax = if isRow then Row.MaxIndex0 else Column.MaxIndex0
    shiftRangePos(sPos, ePos, at, delta, axisMax).map { case (ns, ne) =>
      if isRow then
        new CellRange(
          ARef.from0(Column.index0(range.start.col), ns),
          ARef.from0(Column.index0(range.end.col), ne),
          range.startAnchor,
          range.endAnchor
        )
      else
        new CellRange(
          ARef.from0(ns, Row.index0(range.start.row)),
          ARef.from0(ne, Row.index0(range.end.row)),
          range.startAnchor,
          range.endAnchor
        )
    }

  /**
   * A range slot through a structural edit: moved, shrunk, or — fully deleted — the error
   * [[RangeLocation.Error]] `#REF!` in place (GH-629), reported to `voidedLog`.
   */
  private def shiftLocationStructural(
    location: RangeLocation,
    shiftLocal: Boolean,
    editedSheet: String,
    isRow: Boolean,
    at: Int,
    delta: Int,
    voidedLog: VoidedLog
  ): RangeLocation =
    def voided: RangeLocation =
      voidedLog.record(locationText(location))
      refErrorLocation
    location match
      case RangeLocation.Local(range, form) =>
        if shiftLocal then
          shiftRangeStructural(range, form, isRow, at, delta)
            .fold(voided)(RangeLocation.Local(_, form))
        else location
      case RangeLocation.CrossSheet(sheet, range, form) =>
        if sheet.value.equalsIgnoreCase(editedSheet) then
          shiftRangeStructural(range, form, isRow, at, delta)
            .fold(voided)(r => RangeLocation.CrossSheet(sheet, r, form))
        else location
      // GH-353: external-workbook ranges point into ANOTHER workbook — structural edits here
      // never move or void them
      case RangeLocation.External(_, _, _, _) => location
      // GH-394: a defined name is an identifier — structural edits never move or void it
      // (its refersTo text lives in workbook metadata, not in this formula)
      case RangeLocation.Name(_, _) => location
      // GH-612: an error has no coordinates
      case RangeLocation.Error(_) => location

  @SuppressWarnings(
    Array("org.wartremover.warts.AsInstanceOf", "org.wartremover.warts.Var")
  )
  @nowarn("msg=Unreachable case")
  private def shiftStructuralInternal[A](
    expr: TExpr[A],
    shiftLocal: Boolean,
    editedSheet: String,
    isRow: Boolean,
    at: Int,
    delta: Int,
    voidedLog: VoidedLog
  ): TExpr[A] =
    import TExpr.*
    def go[B](e: TExpr[B]): TExpr[B] =
      shiftStructuralInternal(e, shiftLocal, editedSheet, isRow, at, delta, voidedLog)
    def goLocation(location: RangeLocation): RangeLocation =
      shiftLocationStructural(location, shiftLocal, editedSheet, isRow, at, delta, voidedLog)
    def edited(sheet: SheetName): Boolean = sheet.value.equalsIgnoreCase(editedSheet)

    // GH-629: a deleted reference becomes #REF! in place — the whole reference, the rest of the
    // formula intact — and is reported under its original spelling (GH-628)
    def orRefError(shifted: Option[TExpr[?]], original: => String): TExpr[A] =
      shifted match
        case Some(e) => e.asInstanceOf[TExpr[A]]
        case None =>
          voidedLog.record(original)
          refError.asInstanceOf[TExpr[A]]

    expr match
      case Ref(at0, anchor, decode) =>
        if shiftLocal then
          orRefError(
            shiftARefStructural(at0, isRow, at, delta).map(r => Ref(r, anchor, decode)),
            refText(at0, anchor)
          )
        else expr
      case PolyRef(at0, anchor) =>
        if shiftLocal then
          orRefError(
            shiftARefStructural(at0, isRow, at, delta).map(r => PolyRef(r, anchor)),
            refText(at0, anchor)
          )
        else expr
      case SheetRef(sheet, at0, anchor, decode) =>
        if edited(sheet) then
          orRefError(
            shiftARefStructural(at0, isRow, at, delta).map(r => SheetRef(sheet, r, anchor, decode)),
            sheetRefText(sheet, at0, anchor)
          )
        else expr
      case SheetPolyRef(sheet, at0, anchor) =>
        if edited(sheet) then
          orRefError(
            shiftARefStructural(at0, isRow, at, delta).map(r => SheetPolyRef(sheet, r, anchor)),
            sheetRefText(sheet, at0, anchor)
          )
        else expr
      case RangeRef(range, form) =>
        if shiftLocal then
          orRefError(
            shiftRangeStructural(range, form, isRow, at, delta).map(r => RangeRef(r, form)),
            rangeText(range, form)
          )
        else expr
      case SheetRange(sheet, range, form) =>
        if edited(sheet) then
          orRefError(
            shiftRangeStructural(range, form, isRow, at, delta).map(r =>
              SheetRange(sheet, r, form)
            ),
            sheetRangeText(sheet, range, form)
          )
        else expr
      // GH-353: external-workbook refs point into ANOTHER workbook — structural edits here
      // never move or void them
      case _: ExternalRef | _: ExternalRange => expr
      case lit: Lit[?] => lit.asInstanceOf[TExpr[A]]
      // GH-612: an error literal has no coordinates; GH-603: neither has an omitted argument
      case _: ErrorLit => expr
      case Missing => expr
      case Add(x, y) => Add(go(x), go(y)).asInstanceOf[TExpr[A]]
      case Sub(x, y) => Sub(go(x), go(y)).asInstanceOf[TExpr[A]]
      case Mul(x, y) => Mul(go(x), go(y)).asInstanceOf[TExpr[A]]
      case Div(x, y) => Div(go(x), go(y)).asInstanceOf[TExpr[A]]
      case Pow(x, y) => Pow(go(x), go(y)).asInstanceOf[TExpr[A]]
      case Concat(x, y) => Concat(go(x), go(y)).asInstanceOf[TExpr[A]]
      case Eq(x, y) => Eq(go(x), go(y)).asInstanceOf[TExpr[A]]
      case Neq(x, y) => Neq(go(x), go(y)).asInstanceOf[TExpr[A]]
      case Lt(x, y) => Lt(go(x), go(y)).asInstanceOf[TExpr[A]]
      case Lte(x, y) => Lte(go(x), go(y)).asInstanceOf[TExpr[A]]
      case Gt(x, y) => Gt(go(x), go(y)).asInstanceOf[TExpr[A]]
      case Gte(x, y) => Gte(go(x), go(y)).asInstanceOf[TExpr[A]]
      case ToInt(e) => ToInt(go(e)).asInstanceOf[TExpr[A]]
      // GH-374: unary plus is transparent — recurse so refs under it move/void structurally
      case UnaryPlus(e) => UnaryPlus(go(e))
      // GH-355: postfix percent — recurse so refs under it move/void structurally
      case Percent(e) => Percent(go(e)).asInstanceOf[TExpr[A]]
      case Aggregate(aggId, location) =>
        Aggregate(aggId, goLocation(location)).asInstanceOf[TExpr[A]]
      case call: Call[?] =>
        // The deprecated bare-CellRange slot has no error form: a deleted range there still voids
        // the whole call (the pre-GH-629 behaviour, now confined to that slot)
        var voided = false
        val shifted = call.spec.argSpec.map(call.args)(
          e => go(e),
          goLocation,
          range =>
            if shiftLocal then
              shiftCornersStructural(range, isRow, at, delta).getOrElse {
                voided = true
                voidedLog.record(rangeText(range, RangeForm.Cells))
                range
              }
            else range
        )
        if voided then refError.asInstanceOf[TExpr[A]]
        else Call(call.spec, shifted).asInstanceOf[TExpr[A]]
      case DateToSerial(e) => DateToSerial(go(e)).asInstanceOf[TExpr[A]]
      case DateTimeToSerial(e) => DateTimeToSerial(go(e)).asInstanceOf[TExpr[A]]
      // GH-193: LET — a deleted ref in a binding value or the body is #REF! in place, like anywhere
      case Let(bindings, body) =>
        Let(
          bindings.map((name, value) => (name, go(value.asInstanceOf[TExpr[Any]]))),
          go(body)
        ).asInstanceOf[TExpr[A]]
      case _: BindingRef => expr
      // GH-384: defined names are identifiers — structural edits never move or void them
      // (their refersTo text lives in workbook metadata, not in this formula)
      case _: NameRef => expr
      case _: SheetNameRef => expr
      case _: CoercedBindingRef[?] => expr

      // GH-306: runtime coercion wrapper — shift the wrapped expression, preserve the wrapper
      case Coerced(inner, target) => Coerced(go(inner), target).asInstanceOf[TExpr[A]]
