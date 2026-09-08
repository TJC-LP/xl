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
 * in a range-typed argument slot that falls off the grid voids the enclosing call (`SUM(A1:A2)` →
 * `#REF!`, the same value Excel's `SUM(#REF!)` evaluates to).
 *
 * Laws:
 *   - Identity: shift(expr, 0, 0) == expr
 *   - Commutativity: shift(shift(expr, c1, r1), c2, r2) == shift(expr, c1+c2, r1+r2) while every
 *     reference stays on the grid
 *   - Anchor preservation: Anchor of shifted ref equals original anchor
 */
object FormulaShifter:

  /** GH-612: what an off-grid reference becomes. */
  private val refError: TExpr[Nothing] = TExpr.ErrorLit(CellError.Ref)

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
   *   Transformed AST with shifted references
   */
  def shift[A](expr: TExpr[A], colDelta: Int, rowDelta: Int): TExpr[A] =
    if colDelta == 0 && rowDelta == 0 then expr
    else shiftInternal(expr, colDelta, rowDelta)

  // Var: the argSpec.map traversal cannot return a different node shape, so a voided range slot
  // is flagged through a local var exactly as shiftStructuralInternal does
  @SuppressWarnings(Array("org.wartremover.warts.AsInstanceOf", "org.wartremover.warts.Var"))
  @nowarn(
    "msg=Unreachable case"
  ) // PolyRef extends TExpr[Nothing], reachable via asInstanceOf casts
  private def shiftInternal[A](expr: TExpr[A], colDelta: Int, rowDelta: Int): TExpr[A] =
    import TExpr.*

    // GH-612: an off-grid reference becomes #REF! — the whole reference, never a clamped or
    // non-existent address
    def orRefError(shifted: Option[TExpr[?]]): TExpr[A] =
      shifted.getOrElse(refError).asInstanceOf[TExpr[A]]

    expr match
      // Cell references - apply anchor-aware shifting
      case Ref(at, anchor, decode) =>
        orRefError(shiftARef(at, anchor, colDelta, rowDelta).map(Ref(_, anchor, decode)))

      case PolyRef(at, anchor) =>
        orRefError(shiftARef(at, anchor, colDelta, rowDelta).map(PolyRef(_, anchor)))

      // Sheet-qualified references - shift the cell ref but keep the sheet
      case SheetRef(sheet, at, anchor, decode) =>
        orRefError(
          shiftARef(at, anchor, colDelta, rowDelta).map(SheetRef(sheet, _, anchor, decode))
        )

      case SheetPolyRef(sheet, at, anchor) =>
        orRefError(shiftARef(at, anchor, colDelta, rowDelta).map(SheetPolyRef(sheet, _, anchor)))

      case RangeRef(range, form) =>
        orRefError(shiftRange(range, form, colDelta, rowDelta).map(RangeRef(_, form)))

      case SheetRange(sheet, range, form) =>
        orRefError(shiftRange(range, form, colDelta, rowDelta).map(SheetRange(sheet, _, form)))

      // GH-353: external-workbook references shift anchor-aware like sheet-qualified ones —
      // the workbook/sheet qualifier is fixed, the cell coordinates drag
      case ExternalRef(index, name, at, anchor) =>
        orRefError(
          shiftARef(at, anchor, colDelta, rowDelta).map(ExternalRef(index, name, _, anchor))
        )

      case ExternalRange(index, name, range, form) =>
        orRefError(
          shiftRange(range, form, colDelta, rowDelta).map(ExternalRange(index, name, _, form))
        )

      // Literals - unchanged
      case lit: Lit[?] => lit.asInstanceOf[TExpr[A]]
      // GH-612: an error literal has no coordinates
      case err: ErrorLit => err.asInstanceOf[TExpr[A]]

      // Arithmetic operators
      case Add(x, y) =>
        Add(shiftInternal(x, colDelta, rowDelta), shiftInternal(y, colDelta, rowDelta))
          .asInstanceOf[TExpr[A]]
      case Sub(x, y) =>
        Sub(shiftInternal(x, colDelta, rowDelta), shiftInternal(y, colDelta, rowDelta))
          .asInstanceOf[TExpr[A]]
      case Mul(x, y) =>
        Mul(shiftInternal(x, colDelta, rowDelta), shiftInternal(y, colDelta, rowDelta))
          .asInstanceOf[TExpr[A]]
      case Div(x, y) =>
        Div(shiftInternal(x, colDelta, rowDelta), shiftInternal(y, colDelta, rowDelta))
          .asInstanceOf[TExpr[A]]
      case Pow(x, y) =>
        Pow(shiftInternal(x, colDelta, rowDelta), shiftInternal(y, colDelta, rowDelta))
          .asInstanceOf[TExpr[A]]

      // String operators
      case Concat(x, y) =>
        Concat(shiftInternal(x, colDelta, rowDelta), shiftInternal(y, colDelta, rowDelta))
          .asInstanceOf[TExpr[A]]

      // Comparison operators
      case Eq(x, y) =>
        Eq(shiftInternal(x, colDelta, rowDelta), shiftInternal(y, colDelta, rowDelta))
          .asInstanceOf[TExpr[A]]
      case Neq(x, y) =>
        Neq(shiftInternal(x, colDelta, rowDelta), shiftInternal(y, colDelta, rowDelta))
          .asInstanceOf[TExpr[A]]
      case Lt(x, y) =>
        Lt(shiftInternal(x, colDelta, rowDelta), shiftInternal(y, colDelta, rowDelta))
          .asInstanceOf[TExpr[A]]
      case Lte(x, y) =>
        Lte(shiftInternal(x, colDelta, rowDelta), shiftInternal(y, colDelta, rowDelta))
          .asInstanceOf[TExpr[A]]
      case Gt(x, y) =>
        Gt(shiftInternal(x, colDelta, rowDelta), shiftInternal(y, colDelta, rowDelta))
          .asInstanceOf[TExpr[A]]
      case Gte(x, y) =>
        Gte(shiftInternal(x, colDelta, rowDelta), shiftInternal(y, colDelta, rowDelta))
          .asInstanceOf[TExpr[A]]

      // Type conversion
      case ToInt(e) =>
        ToInt(shiftInternal(e, colDelta, rowDelta)).asInstanceOf[TExpr[A]]

      // GH-374: unary plus is transparent — recurse so refs under it drag
      case UnaryPlus(e) =>
        UnaryPlus(shiftInternal(e, colDelta, rowDelta))

      // GH-355: postfix percent — recurse so refs under it drag
      case Percent(e) =>
        Percent(shiftInternal(e, colDelta, rowDelta)).asInstanceOf[TExpr[A]]

      // Arithmetic range functions (now using RangeLocation). GH-612: a range slot that falls off
      // the grid voids the call — the value Excel's `SUM(#REF!)` evaluates to
      case Aggregate(aggregatorId, location) =>
        orRefError(shiftLocation(location, colDelta, rowDelta).map(Aggregate(aggregatorId, _)))

      case call: Call[?] =>
        var voided = false
        val shifted =
          call.spec.argSpec.map(call.args)(
            expr => shiftInternal(expr, colDelta, rowDelta),
            loc => shiftLocation(loc, colDelta, rowDelta).getOrElse { voided = true; loc },
            range =>
              shiftRange(range, RangeForm.Cells, colDelta, rowDelta).getOrElse {
                voided = true; range
              }
          )
        if voided then refError.asInstanceOf[TExpr[A]]
        else Call(call.spec, shifted).asInstanceOf[TExpr[A]]

      // Date-to-serial converters - shift inner expression
      case DateToSerial(dateExpr) =>
        DateToSerial(shiftInternal(dateExpr, colDelta, rowDelta)).asInstanceOf[TExpr[A]]
      case DateTimeToSerial(dtExpr) =>
        DateTimeToSerial(shiftInternal(dtExpr, colDelta, rowDelta)).asInstanceOf[TExpr[A]]

      // GH-193: LET — shift refs in binding values and body; binding names are not cell refs
      case Let(bindings, body) =>
        Let(
          bindings.map((name, value) => (name, shiftWildcard(value, colDelta, rowDelta))),
          shiftInternal(body, colDelta, rowDelta)
        ).asInstanceOf[TExpr[A]]
      case bref: BindingRef => bref.asInstanceOf[TExpr[A]]
      // GH-384: defined names are identifiers, not coordinates — they never shift on drag
      case nref: NameRef => nref.asInstanceOf[TExpr[A]]
      // GH-394: sheet-qualified names are identifiers too
      case snref: SheetNameRef => snref.asInstanceOf[TExpr[A]]
      case cbref: CoercedBindingRef[?] => cbref.asInstanceOf[TExpr[A]]

      // GH-306: runtime coercion wrapper — shift the wrapped expression, preserve the wrapper
      case Coerced(inner, target) =>
        Coerced(shiftInternal(inner, colDelta, rowDelta), target).asInstanceOf[TExpr[A]]

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
    val (cd, rd) = form match
      case RangeForm.Cells => (colDelta, rowDelta)
      case RangeForm.Columns => (colDelta, 0)
      case RangeForm.Rows => (0, rowDelta)
    for
      newStart <- shiftARef(range.start, range.startAnchor, cd, rd)
      newEnd <- shiftARef(range.end, range.endAnchor, cd, rd)
    yield new CellRange(newStart, newEnd, range.startAnchor, range.endAnchor)

  /**
   * Shift a RangeLocation by the given deltas.
   *
   * Handles both Local and CrossSheet locations, shifting the underlying CellRange. None when the
   * range leaves the grid (GH-612).
   */
  private def shiftLocation(
    location: RangeLocation,
    colDelta: Int,
    rowDelta: Int
  ): Option[RangeLocation] =
    location match
      case RangeLocation.Local(range, form) =>
        shiftRange(range, form, colDelta, rowDelta).map(RangeLocation.Local(_, form))
      case RangeLocation.CrossSheet(sheet, range, form) =>
        shiftRange(range, form, colDelta, rowDelta).map(RangeLocation.CrossSheet(sheet, _, form))
      // GH-353: external-workbook range args drag anchor-aware like the TExpr.ExternalRange node
      case RangeLocation.External(index, name, range, form) =>
        shiftRange(range, form, colDelta, rowDelta)
          .map(RangeLocation.External(index, name, _, form))
      // GH-394: a defined name is an identifier, not coordinates — shifting is a no-op
      // (its refersTo text lives in workbook metadata, not in this formula)
      case name @ RangeLocation.Name(_, _) => Some(name)

  /**
   * Helper to shift TExpr[?] (wildcard type).
   */
  @SuppressWarnings(Array("org.wartremover.warts.AsInstanceOf"))
  private def shiftWildcard(expr: TExpr[?], colDelta: Int, rowDelta: Int): TExpr[?] =
    shiftInternal(expr.asInstanceOf[TExpr[Any]], colDelta, rowDelta)

  // ============================================================================
  // GH-128 / GH-129: structural shifting for row/column insert & delete.
  //
  // Unlike `shift` (uniform fill-drag, anchor-aware), this is POSITION-CONDITIONAL and
  // ANCHOR-INDEPENDENT — structural edits move absolute refs too:
  //   - delta > 0: insert `delta` rows/cols at index `at`; refs at index >= at move +delta.
  //   - delta < 0: delete `-delta` rows/cols starting at `at`; refs after the band move by
  //     +delta; a ref/range that lands entirely inside the deleted band makes the whole
  //     formula `#REF!` (the traversal returns None so the caller can replace the cell with
  //     CellValue.Error(Ref)). Partially-overlapped ranges shrink to the surviving span.
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
          RangeLocation.Name(_, _) =>
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
      case Lit(_) | ErrorLit(_) | Ref(_, _, _) | PolyRef(_, _) | RangeRef(_, _) |
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
          RangeLocation.CrossSheet(_, _, _) | RangeLocation.External(_, _, _, _) =>
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
      case Lit(_) | ErrorLit(_) | Ref(_, _, _) | PolyRef(_, _) | RangeRef(_, _) |
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
          RangeLocation.Name(_, None)) =>
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
          _: ErrorLit | _: BindingRef | _: NameRef | _: CoercedBindingRef[?] =>
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
   * 0). Returns None when the formula references a fully-deleted cell or range.
   */
  def shiftStructural[A](
    expr: TExpr[A],
    shiftLocal: Boolean,
    editedSheet: String,
    isRow: Boolean,
    at: Int,
    delta: Int
  ): Option[TExpr[A]] =
    shiftStructuralInternal(expr, shiftLocal, editedSheet, isRow, at, delta)

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
    val spansEditedAxis = form match
      case RangeForm.Columns => isRow
      case RangeForm.Rows => !isRow
      case RangeForm.Cells => false
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

  private def shiftLocationStructural(
    location: RangeLocation,
    shiftLocal: Boolean,
    editedSheet: String,
    isRow: Boolean,
    at: Int,
    delta: Int
  ): Option[RangeLocation] =
    location match
      case RangeLocation.Local(range, form) =>
        if shiftLocal then
          shiftRangeStructural(range, form, isRow, at, delta).map(RangeLocation.Local(_, form))
        else Some(location)
      case RangeLocation.CrossSheet(sheet, range, form) =>
        if sheet.value.equalsIgnoreCase(editedSheet) then
          shiftRangeStructural(range, form, isRow, at, delta)
            .map(r => RangeLocation.CrossSheet(sheet, r, form))
        else Some(location)
      // GH-353: external-workbook ranges point into ANOTHER workbook — structural edits here
      // never move or void them
      case RangeLocation.External(_, _, _, _) => Some(location)
      // GH-394: a defined name is an identifier — structural edits never move or void it
      // (its refersTo text lives in workbook metadata, not in this formula)
      case RangeLocation.Name(_, _) => Some(location)

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
    delta: Int
  ): Option[TExpr[A]] =
    import TExpr.*
    def go[B](e: TExpr[B]): Option[TExpr[B]] =
      shiftStructuralInternal(e, shiftLocal, editedSheet, isRow, at, delta)

    expr match
      case Ref(at0, anchor, decode) =>
        if shiftLocal then
          shiftARefStructural(at0, isRow, at, delta).map(r => Ref(r, anchor, decode))
        else Some(expr)
      case PolyRef(at0, anchor) =>
        if shiftLocal then
          shiftARefStructural(at0, isRow, at, delta).map(r =>
            PolyRef(r, anchor).asInstanceOf[TExpr[A]]
          )
        else Some(expr)
      case SheetRef(sheet, at0, anchor, decode) =>
        if sheet.value.equalsIgnoreCase(editedSheet) then
          shiftARefStructural(at0, isRow, at, delta).map(r => SheetRef(sheet, r, anchor, decode))
        else Some(expr)
      case SheetPolyRef(sheet, at0, anchor) =>
        if sheet.value.equalsIgnoreCase(editedSheet) then
          shiftARefStructural(at0, isRow, at, delta).map(r =>
            SheetPolyRef(sheet, r, anchor).asInstanceOf[TExpr[A]]
          )
        else Some(expr)
      case RangeRef(range, form) =>
        if shiftLocal then
          shiftRangeStructural(range, form, isRow, at, delta).map(r =>
            RangeRef(r, form).asInstanceOf[TExpr[A]]
          )
        else Some(expr)
      case SheetRange(sheet, range, form) =>
        if sheet.value.equalsIgnoreCase(editedSheet) then
          shiftRangeStructural(range, form, isRow, at, delta).map(r =>
            SheetRange(sheet, r, form).asInstanceOf[TExpr[A]]
          )
        else Some(expr)
      // GH-353: external-workbook refs point into ANOTHER workbook — structural edits here
      // never move or void them
      case _: ExternalRef | _: ExternalRange => Some(expr)
      case lit: Lit[?] => Some(lit.asInstanceOf[TExpr[A]])
      // GH-612: an error literal has no coordinates
      case _: ErrorLit => Some(expr)
      case Add(x, y) => for sx <- go(x); sy <- go(y) yield Add(sx, sy).asInstanceOf[TExpr[A]]
      case Sub(x, y) => for sx <- go(x); sy <- go(y) yield Sub(sx, sy).asInstanceOf[TExpr[A]]
      case Mul(x, y) => for sx <- go(x); sy <- go(y) yield Mul(sx, sy).asInstanceOf[TExpr[A]]
      case Div(x, y) => for sx <- go(x); sy <- go(y) yield Div(sx, sy).asInstanceOf[TExpr[A]]
      case Pow(x, y) => for sx <- go(x); sy <- go(y) yield Pow(sx, sy).asInstanceOf[TExpr[A]]
      case Concat(x, y) => for sx <- go(x); sy <- go(y) yield Concat(sx, sy).asInstanceOf[TExpr[A]]
      case Eq(x, y) => for sx <- go(x); sy <- go(y) yield Eq(sx, sy).asInstanceOf[TExpr[A]]
      case Neq(x, y) => for sx <- go(x); sy <- go(y) yield Neq(sx, sy).asInstanceOf[TExpr[A]]
      case Lt(x, y) => for sx <- go(x); sy <- go(y) yield Lt(sx, sy).asInstanceOf[TExpr[A]]
      case Lte(x, y) => for sx <- go(x); sy <- go(y) yield Lte(sx, sy).asInstanceOf[TExpr[A]]
      case Gt(x, y) => for sx <- go(x); sy <- go(y) yield Gt(sx, sy).asInstanceOf[TExpr[A]]
      case Gte(x, y) => for sx <- go(x); sy <- go(y) yield Gte(sx, sy).asInstanceOf[TExpr[A]]
      case ToInt(e) => go(e).map(se => ToInt(se).asInstanceOf[TExpr[A]])
      // GH-374: unary plus is transparent — recurse so refs under it move/void structurally
      case UnaryPlus(e) => go(e).map(UnaryPlus.apply)
      // GH-355: postfix percent — recurse so refs under it move/void structurally
      case Percent(e) => go(e).map(se => Percent(se).asInstanceOf[TExpr[A]])
      case Aggregate(aggId, location) =>
        shiftLocationStructural(location, shiftLocal, editedSheet, isRow, at, delta)
          .map(l => Aggregate(aggId, l).asInstanceOf[TExpr[A]])
      case call: Call[?] =>
        var deleted = false
        val shifted = call.spec.argSpec.map(call.args)(
          e => go(e).getOrElse { deleted = true; e },
          loc =>
            shiftLocationStructural(loc, shiftLocal, editedSheet, isRow, at, delta).getOrElse {
              deleted = true; loc
            },
          range =>
            if shiftLocal then
              shiftCornersStructural(range, isRow, at, delta).getOrElse { deleted = true; range }
            else range
        )
        if deleted then None else Some(Call(call.spec, shifted).asInstanceOf[TExpr[A]])
      case DateToSerial(e) => go(e).map(se => DateToSerial(se).asInstanceOf[TExpr[A]])
      case DateTimeToSerial(e) => go(e).map(se => DateTimeToSerial(se).asInstanceOf[TExpr[A]])
      // GH-193: LET — a fully-deleted ref in any binding value or the body voids the formula
      case Let(bindings, body) =>
        val shiftedBindings =
          bindings.foldLeft[Option[List[(String, TExpr[?])]]](Some(Nil)) {
            case (None, _) => None
            case (Some(acc), (name, value)) =>
              go(value.asInstanceOf[TExpr[Any]]).map(sv => (name, sv) :: acc)
          }
        for
          bs <- shiftedBindings
          sb <- go(body)
        yield Let(bs.reverse, sb).asInstanceOf[TExpr[A]]
      case _: BindingRef => Some(expr)
      // GH-384: defined names are identifiers — structural edits never move or void them
      // (their refersTo text lives in workbook metadata, not in this formula)
      case _: NameRef => Some(expr)
      case _: SheetNameRef => Some(expr)
      case _: CoercedBindingRef[?] => Some(expr)

      // GH-306: runtime coercion wrapper — shift the wrapped expression, preserve the wrapper
      case Coerced(inner, target) =>
        go(inner).map(si => Coerced(si, target).asInstanceOf[TExpr[A]])
