package com.tjclp.xl.formula.functions

import com.tjclp.xl.formula.ast.{BindingCoercion, ExprValue, RangeForm, TExpr}
import com.tjclp.xl.formula.eval.{EvalError, Evaluator}
import com.tjclp.xl.formula.parser.ParseError
import com.tjclp.xl.formula.{Clock, Arity, Rng}

import com.tjclp.xl.{Anchor, CellRange}
import com.tjclp.xl.addressing.ARef
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.workbooks.Workbook

/**
 * Metadata flags for function behavior that impacts formatting or inference.
 */
final case class FunctionFlags(
  returnsDate: Boolean = false,
  returnsTime: Boolean = false,
  /**
   * True for functions returning `BigDecimal`. Lets coercion code (e.g. `asIntExpr`) wrap any
   * numeric-returning Call in `ToInt` without listing each function by name. Mutually exclusive
   * with `returnsDate` / `returnsTime` — a function returns one type.
   */
  returnsNumeric: Boolean = false,
  /**
   * GH-274: True for functions whose data dependencies are not statically knowable (e.g. INDIRECT,
   * which reads whatever cell its evaluated text names). The static dependency graph sees only the
   * function's *arguments*; cells bearing such calls are deferred to the end of recalculation order
   * (`DependencyGraph.deferDynamic`) and treated as always-dirty by targeted recalculation
   * (`DependentRecalculation`).
   */
  dynamicDeps: Boolean = false,
  /**
   * GH-588: True for functions whose value can change between two recalculations without any input
   * changing — Excel's volatile marking (TODAY, NOW, RAND, RANDBETWEEN). The flag is the single
   * source of truth: `FunctionRegistry.volatileFunctionNames` lists the flagged specs and
   * `WorkbookAudit.volatile` reports the cells that call one.
   */
  volatile: Boolean = false,
  /**
   * Excel array lifting: whether (and where) the function evaluates once per element when a scalar
   * parameter receives an array — `ABS(C2:C4+D2:D4)` is `{2;2;3}` in an array context. The flag is
   * the single source of truth; the evaluator implements the lift once, at the Call node, over the
   * slots [[ArgSpec.scalarSlots]] offers. Off keeps the typed top-left collapse (GH-302).
   */
  lift: ArrayLift = ArrayLift.Off
)

/**
 * Which scalar slots of a function lift over arrays (see [[FunctionFlags.lift]]).
 *
 * `slots` indexes the list [[ArgSpec.scalarSlots]] returns (None = every scalar slot), so IFERROR
 * lifts its value and never replicates its fallback. `rangeRefs = false` is the Analysis ToolPak
 * lineage (EDATE, EOMONTH, WORKDAY, NETWORKDAYS, YEARFRAC, MROUND): an array lifts, but a
 * multi-cell range REFERENCE in a scalar slot is `#VALUE!` — `EOMONTH(A2:A10,0)` is `#VALUE!` in
 * Excel while `EOMONTH(+A2:A10,0)` spills.
 */
enum ArrayLift derives CanEqual:
  case Off
  case On(slots: Option[Set[Int]], rangeRefs: Boolean)

object ArrayLift:
  /** Every scalar slot lifts; range references lift like arrays. */
  val all: ArrayLift = On(None, rangeRefs = true)

  /** Only the scalar slots at these positions (indices into [[ArgSpec.scalarSlots]]) lift. */
  def slots(positions: Int*): ArrayLift = On(Some(positions.toSet), rangeRefs = true)

  /** Every scalar slot lifts over arrays; a multi-cell range reference is `#VALUE!`. */
  val arraysOnly: ArrayLift = On(None, rangeRefs = false)

/**
 * How one element of a lifted array is handed to its slot — the conventions a cell reference gets
 * in that slot, so `f(range)[i]` equals `f(cell_i)` (except for numeric text and cached formula
 * cells, whose single-reference decoders predate lifting; ArrayLiftingLawsSpec pins both).
 */
enum LiftSlot derives CanEqual:
  /** A typed slot: the element coerces to the target (`Coerced(Lit(element), target)`). */
  case Typed(target: BindingCoercion)

  /** A raw CellValue slot (IS*, IFERROR, lookup values): the element as is — ISBLANK sees Empty. */
  case Cell

  /**
   * An Any-typed value slot (criteria, MATCH/TEXT values): a cell element resolves (blank is 0).
   */
  case Value

final case class ArgPrinter(
  expr: TExpr[?] => String,
  location: TExpr.RangeLocation => String,
  cellRange: CellRange => String,
  /**
   * GH-484: the argument separator to join rendered args with — the canonical human-facing ", " by
   * default; Excel's bare "," when printing the file form (FormulaPrinter.printFileForm).
   */
  separator: String = ", "
)

final case class EvalContext(
  sheet: Sheet,
  clock: Clock,
  workbook: Option[Workbook],
  evalExpr: [A] => TExpr[A] => Either[EvalError, A],
  /**
   * GH-197: Array-aware expression evaluator. Unlike `evalExpr`, this allows ArrayResult to be
   * returned (doesn't reject arrays). Used by SUMPRODUCT to evaluate array expressions like
   * `(A1:A3>15)*B1:B3`.
   */
  evalArrayExpr: TExpr[Any] => Either[EvalError, Any],
  /** Current cell being evaluated. Used by ROW() and COLUMN() with no arguments. */
  currentCell: Option[ARef] = None,
  /** Recursion depth for cross-sheet formula evaluation. */
  depth: Int = 0,
  /** GH-193: in-scope LET bindings (declared name → evaluated value). */
  bindings: Map[String, Any] = Map.empty,
  /** GH-115: randomness capability for RAND/RANDBETWEEN (Clock pattern). */
  rng: Rng = Rng.system,
  /**
   * GH-346: the pass's memo for recursively evaluated uncached formula cells, so functions that
   * read uncached cells (aggregates, array materialization) evaluate each precedent once per pass
   * instead of once per reference (see Evaluator.EvalMemo).
   */
  memo: Option[Evaluator.EvalMemo] = None,
  /**
   * GH-424: the workbook's saved location, if the embedder knows one (thread it via
   * `Evaluator.instance(workbookPath = ...)`). Read by CELL("filename"); None reproduces Excel's
   * pre-save behavior (empty string).
   */
  workbookPath: Option[String] = None,
  /**
   * One WorkbookEvaluator recalculation generation's single-range aggregate memo. Public/direct
   * evaluation leaves this absent, preserving its one-shot semantics and allocation profile.
   */
  aggregateMemo: Option[Evaluator.AggregateMemo] = None,
  /**
   * Excel's operand classes for this call (internal to the evaluator): see
   * [[EvalContext.Operands]].
   */
  private[formula] operands: EvalContext.Operands = EvalContext.Operands.array
):

  /**
   * Whether the formula evaluates as an array (an array formula, `evala`, SUMPRODUCT's arguments)
   * or as a plain cell's legacy formula, where a reference in a value position is implicitly
   * intersected with the formula's cell.
   */
  private[formula] def arrayMode: Boolean = operands.arrayMode

  /**
   * Whether this call's result feeds a reference position, so the value IF, IFS, CHOOSE and SWITCH
   * select is kept as a reference (Excel's IF and CHOOSE return references):
   * `SUM(IF(A1:A10>2,A1:A10,0))` in row 5 sums the whole range.
   */
  private[formula] def selectsReference: Boolean = operands.selectsReference

  /**
   * An argument in Excel's reference operand class (an aggregate's, AND's, OR's, ROWS'). A
   * reference — a cell or range, a name bound to one, a LET name bound to one, IF/CHOOSE's selected
   * reference, the reference OFFSET, INDIRECT or INDEX return — reaches the function whole, as a
   * [[com.tjclp.xl.formula.eval.RangeOperand]]; any other expression evaluates in the formula's own
   * mode, its result not collapsed (an array-returning call folds whole), so in a plain cell its
   * references are intersected: `SUM(A1:A10*B1:B10)` in row 5 is A5*B5, as Excel computes a legacy
   * formula. In array mode, [[evalArrayExpr]].
   */
  private[formula] def evalReferenceArg(expr: TExpr[Any]): Either[EvalError, Any] =
    if arrayMode then evalArrayExpr(expr) else evalReference(expr)

  /**
   * The reference an expression denotes, in either mode, as a
   * [[com.tjclp.xl.formula.eval.RangeOperand]] — for the positions that need the reference itself
   * whatever the formula's mode: the `@` operand, OFFSET's base, the reference IF, IFS, CHOOSE or
   * SWITCH select for a reference position. An expression that denotes no reference evaluates in
   * the formula's mode, its result not collapsed.
   */
  private[formula] def evalReference(expr: TExpr[Any]): Either[EvalError, Any] =
    operands.referenceArg.fold(evalArrayExpr(expr))(_(expr))

object EvalContext:
  /**
   * How a call's arguments evaluate. `arrayMode`: the formula evaluates as an array rather than as
   * a plain cell's legacy formula. `referenceArg`: how an argument evaluates in the reference class
   * (the evaluator sets it for every call; None only for a context built outside the evaluator).
   * `selectsReference`: the call feeds a reference position.
   */
  private[formula] final case class Operands(
    arrayMode: Boolean,
    referenceArg: Option[TExpr[Any] => Either[EvalError, Any]],
    selectsReference: Boolean
  )

  private[formula] object Operands:
    /** Array evaluation: every argument may be an array; references materialize. */
    val array: Operands = Operands(arrayMode = true, referenceArg = None, selectsReference = false)

sealed trait ArgValue
object ArgValue:
  final case class Expr(value: TExpr[?]) extends ArgValue
  final case class Range(value: TExpr.RangeLocation) extends ArgValue
  final case class Cells(value: CellRange) extends ArgValue

/**
 * Describes how to parse, render, and transform a function's arguments.
 */
trait ArgSpec[A]:
  def describeParts: List[String]
  final def describe: String = describeParts.mkString(", ")

  /**
   * The argument's scalar slots, in declaration order, with the kind each lifts as
   * ([[FunctionFlags.lift]]). Range, variadic-numeric and SUMPRODUCT slots take arrays whole and
   * are never offered. The default offers none — an ArgSpec that does not override it never lifts.
   */
  def scalarSlots(args: A): List[(TExpr[?], LiftSlot)] = Nil

  /**
   * Rebuild `args` with its scalar slots replaced, in [[scalarSlots]] order, consuming the
   * replacements it uses and returning the rest (like [[parse]]). Law:
   * `replaceScalarSlots(a, scalarSlots(a).map(_._1)) == (a, Nil)`.
   */
  def replaceScalarSlots(args: A, replacements: List[TExpr[?]]): (A, List[TExpr[?]]) =
    (args, replacements)

  def parse(
    args: List[TExpr[?]],
    pos: Int,
    fnName: String
  ): Either[ParseError, (A, List[TExpr[?]])]

  def toValues(args: A): List[ArgValue]

  def map(
    args: A
  )(
    mapExpr: TExpr[?] => TExpr[?],
    mapRange: TExpr.RangeLocation => TExpr.RangeLocation,
    mapCells: CellRange => CellRange
  ): A

  def render(args: A, printer: ArgPrinter): List[String] =
    toValues(args).map {
      case ArgValue.Expr(expr) => printer.expr(expr)
      case ArgValue.Range(location) => printer.location(location)
      case ArgValue.Cells(range) => printer.cellRange(range)
    }

  /**
   * GH-603: the rendered argument SLOTS in declaration order — `None` for an ABSENT optional
   * argument (fewer arguments than slots, or an empty range slot), `Some("")` for a slot written
   * empty (`TExpr.Missing`), so `PMT(r,n,pv,,1)` and `VLOOKUP(x,rng,2,)` keep their commas.
   * [[render]] drops absent slots; [[FunctionSpec.render]] trims the trailing absent ones.
   */
  def renderSlots(args: A, printer: ArgPrinter): List[Option[String]] =
    render(args, printer).map(Some(_))

trait FunctionSpec[A]:
  type Args
  def name: String
  def arity: Arity
  def argSpec: ArgSpec[Args]
  def eval(args: Args, ctx: EvalContext): Either[EvalError, A]
  def flags: FunctionFlags = FunctionFlags()

  /**
   * The reference the call returns, for a function that returns one (OFFSET, INDIRECT, INDEX): a
   * [[com.tjclp.xl.formula.eval.RangeOperand]], unread, or the value the call computes when it
   * names no reference (a `#REF!`). The evaluator asks for it where Excel keeps a function's result
   * a reference: an aggregate's argument, ROWS', the `@` operand, OFFSET's base, a LET binding.
   * None for every other function.
   */
  private[formula] def reference(args: Args, ctx: EvalContext): Option[Either[EvalError, Any]] =
    None

  def render(args: Args, printer: ArgPrinter): String =
    s"${name}(${FunctionSpec.joinSlots(argSpec.renderSlots(args, printer), printer.separator)})"

object FunctionSpec:
  /**
   * GH-603: join rendered argument slots. An absent optional argument (`None`) keeps its comma when
   * a later slot is present and is dropped when it trails. A slot the formula wrote EMPTY
   * (`TExpr.Missing`, rendered "") always keeps its comma, trailing or not: `LEFT("abc",)` must
   * re-parse with two arguments and `VLOOKUP(x,rng,2,)` must stay an exact match (GH-654). With the
   * human-facing ", " separator an empty slot contributes a bare "," so the text reads `pv,, 1`
   * rather than `pv, , 1`.
   */
  def joinSlots(slots: List[Option[String]], separator: String): String =
    val trimmed = slots.reverse.dropWhile(_.isEmpty).reverse
    val sb = new StringBuilder
    trimmed.zipWithIndex.foreach { case (slot, i) =>
      val text = slot.getOrElse("")
      if i > 0 then sb.append(if text.isEmpty then separator.trim else separator)
      sb.append(text)
    }
    sb.toString

  final case class Simple[A, A0](
    name: String,
    arity: Arity,
    argSpec: ArgSpec[A0],
    evalFn: (A0, EvalContext) => Either[EvalError, A],
    override val flags: FunctionFlags = FunctionFlags(),
    renderFn: Option[(A0, ArgPrinter) => String] = None
  ) extends FunctionSpec[A]:
    type Args = A0
    def eval(args: A0, ctx: EvalContext): Either[EvalError, A] = evalFn(args, ctx)
    override def render(args: A0, printer: ArgPrinter): String =
      renderFn.map(_(args, printer)).getOrElse(super.render(args, printer))

  def simple[A, A0](
    name: String,
    arity: Arity,
    flags: FunctionFlags = FunctionFlags(),
    renderFn: Option[(A0, ArgPrinter) => String] = None
  )(evalFn: (A0, EvalContext) => Either[EvalError, A])(using
    spec: ArgSpec[A0]
  ): FunctionSpec[A] { type Args = A0 } =
    Simple(name, arity, spec, evalFn, flags, renderFn)

  /** A function that returns a reference: [[FunctionSpec.reference]] is `referenceFn`. */
  private[formula] final case class Referencing[A, A0](
    name: String,
    arity: Arity,
    argSpec: ArgSpec[A0],
    referenceFn: (A0, EvalContext) => Either[EvalError, Any],
    evalFn: (A0, EvalContext) => Either[EvalError, A],
    override val flags: FunctionFlags = FunctionFlags(),
    renderFn: Option[(A0, ArgPrinter) => String] = None
  ) extends FunctionSpec[A]:
    type Args = A0
    def eval(args: A0, ctx: EvalContext): Either[EvalError, A] = evalFn(args, ctx)
    override private[formula] def reference(
      args: A0,
      ctx: EvalContext
    ): Option[Either[EvalError, Any]] =
      Some(referenceFn(args, ctx))
    override def render(args: A0, printer: ArgPrinter): String =
      renderFn.map(_(args, printer)).getOrElse(super.render(args, printer))

  private[formula] def referencing[A, A0](
    name: String,
    arity: Arity,
    flags: FunctionFlags = FunctionFlags(),
    renderFn: Option[(A0, ArgPrinter) => String] = None
  )(referenceFn: (A0, EvalContext) => Either[EvalError, Any])(
    evalFn: (A0, EvalContext) => Either[EvalError, A]
  )(using spec: ArgSpec[A0]): FunctionSpec[A] { type Args = A0 } =
    Referencing(name, arity, spec, referenceFn, evalFn, flags, renderFn)

trait ExprCoercer[A]:
  def label: String
  def coerce(expr: TExpr[?]): TExpr[A]

  /** How an array element enters a slot this coercer typed, when the function lifts. */
  def liftSlot: Option[LiftSlot] = None

object ExprCoercer:
  given numeric: ExprCoercer[BigDecimal] with
    val label = "number"
    def coerce(expr: TExpr[?]): TExpr[BigDecimal] = TExpr.asNumericExpr(expr)
    override def liftSlot: Option[LiftSlot] = Some(LiftSlot.Typed(BindingCoercion.Numeric))

  given boolean: ExprCoercer[Boolean] with
    val label = "boolean"
    def coerce(expr: TExpr[?]): TExpr[Boolean] = TExpr.asBooleanExpr(expr)
    override def liftSlot: Option[LiftSlot] = Some(LiftSlot.Typed(BindingCoercion.Bool))

  given text: ExprCoercer[String] with
    val label = "text"
    def coerce(expr: TExpr[?]): TExpr[String] = TExpr.asStringExpr(expr)
    override def liftSlot: Option[LiftSlot] = Some(LiftSlot.Typed(BindingCoercion.Text))

  given intExpr: ExprCoercer[Int] with
    val label = "integer"
    def coerce(expr: TExpr[?]): TExpr[Int] = TExpr.asIntExpr(expr)
    override def liftSlot: Option[LiftSlot] = Some(LiftSlot.Typed(BindingCoercion.Integer))

  given cellValue: ExprCoercer[CellValue] with
    val label = "cell"
    def coerce(expr: TExpr[?]): TExpr[CellValue] = TExpr.asCellValueExpr(expr)
    override def liftSlot: Option[LiftSlot] = Some(LiftSlot.Cell)

  given dateExpr: ExprCoercer[java.time.LocalDate] with
    val label = "date"
    def coerce(expr: TExpr[?]): TExpr[java.time.LocalDate] = TExpr.asDateExpr(expr)
    override def liftSlot: Option[LiftSlot] = Some(LiftSlot.Typed(BindingCoercion.Date))

@SuppressWarnings(Array("org.wartremover.warts.AsInstanceOf"))
object ArgSpec:
  def expr[A](using coercer: ExprCoercer[A]): ArgSpec[TExpr[A]] =
    new ArgSpec[TExpr[A]]:
      def describeParts: List[String] = List(coercer.label)

      def parse(
        args: List[TExpr[?]],
        pos: Int,
        fnName: String
      ): Either[ParseError, (TExpr[A], List[TExpr[?]])] =
        args match
          case head :: tail => Right((coercer.coerce(head), tail))
          case Nil =>
            Left(ParseError.InvalidArguments(fnName, pos, describe, "0 arguments"))

      def toValues(args: TExpr[A]): List[ArgValue] =
        List(ArgValue.Expr(args))

      override def scalarSlots(args: TExpr[A]): List[(TExpr[?], LiftSlot)] =
        coercer.liftSlot.map(slot => (args, slot)).toList

      // a replacement is the slot's own expression or a lifted element typed for the slot
      // (Coerced to its target, a CellValue literal, a resolved value) — never a mistyped node
      override def replaceScalarSlots(
        args: TExpr[A],
        replacements: List[TExpr[?]]
      ): (TExpr[A], List[TExpr[?]]) =
        (coercer.liftSlot, replacements) match
          case (Some(_), head :: rest) => (head.asInstanceOf[TExpr[A]], rest)
          case _ => (args, replacements)

      def map(
        args: TExpr[A]
      )(
        mapExpr: TExpr[?] => TExpr[?],
        mapRange: TExpr.RangeLocation => TExpr.RangeLocation,
        mapCells: CellRange => CellRange
      ): TExpr[A] =
        mapExpr(args).asInstanceOf[TExpr[A]]

  given rangeLocation: ArgSpec[TExpr.RangeLocation] with
    def describeParts: List[String] = List("range")

    def parse(
      args: List[TExpr[?]],
      pos: Int,
      fnName: String
    ): Either[ParseError, (TExpr.RangeLocation, List[TExpr[?]])] =
      args match
        case TExpr.RangeRef(range, form) :: tail =>
          Right((TExpr.RangeLocation.Local(range, form), tail))
        case TExpr.SheetRange(sheet, range, form) :: tail =>
          Right((TExpr.RangeLocation.CrossSheet(sheet, range, form), tail))
        // GH-353: external-workbook ranges parse (SUMIF([2]Book1!A1:A9, …)); the location can
        // never resolve at evaluation time, but the cell's Excel-written cache pins its value
        case TExpr.ExternalRange(index, name, range, form) :: tail =>
          Right((TExpr.RangeLocation.External(index, name, range, form), tail))
        // GH-394: defined names are accepted in range-typed argument positions —
        // =VLOOKUP(x, named_table, 2), =SUMIF(rev_range, ">1"), =SUMIF(Model!rev_range, …).
        // The target range resolves at evaluation (Evaluator.resolveRangeLocation); a name
        // bound to a non-range is a clean per-cell error there, never a parse failure.
        case TExpr.NameRef(name) :: tail =>
          Right((TExpr.RangeLocation.Name(name, None), tail))
        case TExpr.SheetNameRef(sheet, name) :: tail =>
          Right((TExpr.RangeLocation.Name(name, Some(sheet)), tail))
        // GH-612: SUM(#REF!) / COUNTIF(#REF!, x) — what Excel writes after a delete or an
        // off-grid drag; the slot carries the error and evaluation yields it
        case TExpr.ErrorLit(error) :: tail =>
          Right((TExpr.RangeLocation.Error(error), tail))
        // GH-631: a single cell where a range is expected — SUMIF(A1:A10, ">0", C1),
        // COUNTIF(A1, "x") — is the 1×1 range it addresses, as Excel reads it; RangeForm.Cell
        // keeps the spelling so it prints back as `C1`
        case TExpr.PolyRef(at, anchor) :: tail =>
          Right((TExpr.RangeLocation.Local(singleCell(at, anchor), RangeForm.Cell), tail))
        case TExpr.Ref(at, anchor, _) :: tail =>
          Right((TExpr.RangeLocation.Local(singleCell(at, anchor), RangeForm.Cell), tail))
        case TExpr.SheetPolyRef(sheet, at, anchor) :: tail =>
          Right(
            (TExpr.RangeLocation.CrossSheet(sheet, singleCell(at, anchor), RangeForm.Cell), tail)
          )
        case TExpr.SheetRef(sheet, at, anchor, _) :: tail =>
          Right(
            (TExpr.RangeLocation.CrossSheet(sheet, singleCell(at, anchor), RangeForm.Cell), tail)
          )
        case TExpr.ExternalRef(index, name, at, anchor) :: tail =>
          Right(
            (
              TExpr.RangeLocation.External(index, name, singleCell(at, anchor), RangeForm.Cell),
              tail
            )
          )
        case _ =>
          Left(ParseError.InvalidArguments(fnName, pos, describe, s"${args.length} arguments"))

    /** The 1×1 range a cell reference addresses, its anchor on both corners. */
    private def singleCell(at: ARef, anchor: Anchor): CellRange =
      CellRange(at, at, anchor, anchor)

    def toValues(args: TExpr.RangeLocation): List[ArgValue] =
      List(ArgValue.Range(args))

    def map(
      args: TExpr.RangeLocation
    )(
      mapExpr: TExpr[?] => TExpr[?],
      mapRange: TExpr.RangeLocation => TExpr.RangeLocation,
      mapCells: CellRange => CellRange
    ): TExpr.RangeLocation =
      mapRange(args)

  @deprecated(
    "Use ArgSpec.rangeLocation (TExpr.RangeLocation) instead: range-typed argument slots " +
      "resolve through Evaluator.resolveRangeLocation since GH-394, admitting sheet-qualified " +
      "ranges and defined names. This local-literal-only spec has no remaining in-repo " +
      "consumers and will be removed in the next breaking cycle.",
    "0.18.0"
  )
  given cellRange: ArgSpec[CellRange] with
    def describeParts: List[String] = List("range")

    def parse(
      args: List[TExpr[?]],
      pos: Int,
      fnName: String
    ): Either[ParseError, (CellRange, List[TExpr[?]])] =
      args match
        case TExpr.RangeRef(range, _) :: tail =>
          Right((range, tail))
        // GH-353: this slot requires a LOCAL literal range — name the unsupported construct
        // instead of the generic arity message
        case (_: TExpr.ExternalRef | _: TExpr.ExternalRange) :: _ =>
          Left(
            ParseError.InvalidArguments(
              fnName,
              pos,
              describe,
              "an external-workbook reference (not supported in this argument; " +
                "external workbooks are not loaded)"
            )
          )
        case _ =>
          Left(ParseError.InvalidArguments(fnName, pos, describe, s"${args.length} arguments"))

    def toValues(args: CellRange): List[ArgValue] =
      List(ArgValue.Cells(args))

    def map(
      args: CellRange
    )(
      mapExpr: TExpr[?] => TExpr[?],
      mapRange: TExpr.RangeLocation => TExpr.RangeLocation,
      mapCells: CellRange => CellRange
    ): CellRange =
      mapCells(args)

  given option[A](using inner: ArgSpec[A]): ArgSpec[Option[A]] with
    def describeParts: List[String] = List(s"optional ${inner.describe}")

    def parse(
      args: List[TExpr[?]],
      pos: Int,
      fnName: String
    ): Either[ParseError, (Option[A], List[TExpr[?]])] =
      args match
        case Nil => Right((None, Nil))
        // GH-603/GH-654: an EMPTY optional slot is a present blank VALUE where the slot can hold
        // one — Excel's `VLOOKUP(x,rng,2,)` is an exact match (range_lookup FALSE), `MATCH(x,rng,)`
        // exact (match_type 0), `LOG(10,)` has base 0 — and absent where it cannot (a range slot:
        // `SUMIF(rng,crit,)` sums `rng`). The dynamic-array functions that read the empty slot as
        // omitted opt out at evaluation (FunctionSpecsBase.unlessOmitted). Keeping the slot
        // present also keeps the formula's text: the trailing `,)` reprints, so a structural edit
        // can never turn Excel's exact match into an approximate one.
        case TExpr.Missing :: rest =>
          inner.parse(args, pos, fnName) match
            case Right((value, rest2)) => Right((Some(value), rest2))
            case Left(_) => Right((None, rest))
        case _ =>
          inner.parse(args, pos, fnName).map { case (value, rest) => (Some(value), rest) }

    def toValues(args: Option[A]): List[ArgValue] =
      args.toList.flatMap(inner.toValues)

    override def scalarSlots(args: Option[A]): List[(TExpr[?], LiftSlot)] =
      args.toList.flatMap(inner.scalarSlots)

    override def replaceScalarSlots(
      args: Option[A],
      replacements: List[TExpr[?]]
    ): (Option[A], List[TExpr[?]]) =
      args match
        case None => (None, replacements)
        case Some(value) =>
          val (replaced, rest) = inner.replaceScalarSlots(value, replacements)
          (Some(replaced), rest)

    override def renderSlots(args: Option[A], printer: ArgPrinter): List[Option[String]] =
      args match
        // one absent marker per slot the inner spec would render, so the comma count matches
        // whether or not the argument is present (every inner spec is single-slot today)
        case None => List.fill(inner.describeParts.size)(None)
        case Some(value) => inner.renderSlots(value, printer)

    def map(
      args: Option[A]
    )(
      mapExpr: TExpr[?] => TExpr[?],
      mapRange: TExpr.RangeLocation => TExpr.RangeLocation,
      mapCells: CellRange => CellRange
    ): Option[A] =
      args.map(inner.map(_)(mapExpr, mapRange, mapCells))

  given list[A](using inner: ArgSpec[A]): ArgSpec[List[A]] with
    def describeParts: List[String] = List(s"${inner.describe}...")

    def parse(
      args: List[TExpr[?]],
      pos: Int,
      fnName: String
    ): Either[ParseError, (List[A], List[TExpr[?]])] =
      @annotation.tailrec
      def loop(
        remaining: List[TExpr[?]],
        acc: List[A]
      ): Either[ParseError, (List[A], List[TExpr[?]])] =
        remaining match
          case Nil => Right((acc.reverse, Nil))
          case _ =>
            inner.parse(remaining, pos, fnName) match
              case Right((value, rest)) => loop(rest, value :: acc)
              case Left(err) => Left(err)
      loop(args, Nil)

    def toValues(args: List[A]): List[ArgValue] =
      args.flatMap(inner.toValues)

    override def scalarSlots(args: List[A]): List[(TExpr[?], LiftSlot)] =
      args.flatMap(inner.scalarSlots)

    override def replaceScalarSlots(
      args: List[A],
      replacements: List[TExpr[?]]
    ): (List[A], List[TExpr[?]]) =
      val (replacedReversed, rest) =
        args.foldLeft((List.empty[A], replacements)) { case ((acc, remaining), value) =>
          val (replaced, next) = inner.replaceScalarSlots(value, remaining)
          (replaced :: acc, next)
        }
      (replacedReversed.reverse, rest)

    override def renderSlots(args: List[A], printer: ArgPrinter): List[Option[String]] =
      args.flatMap(inner.renderSlots(_, printer))

    def map(
      args: List[A]
    )(
      mapExpr: TExpr[?] => TExpr[?],
      mapRange: TExpr.RangeLocation => TExpr.RangeLocation,
      mapCells: CellRange => CellRange
    ): List[A] =
      args.map(inner.map(_)(mapExpr, mapRange, mapCells))

  given emptyTuple: ArgSpec[EmptyTuple] with
    def describeParts: List[String] = Nil

    def parse(
      args: List[TExpr[?]],
      pos: Int,
      fnName: String
    ): Either[ParseError, (EmptyTuple, List[TExpr[?]])] =
      Right((EmptyTuple, args))

    def toValues(args: EmptyTuple): List[ArgValue] = Nil

    def map(
      args: EmptyTuple
    )(
      mapExpr: TExpr[?] => TExpr[?],
      mapRange: TExpr.RangeLocation => TExpr.RangeLocation,
      mapCells: CellRange => CellRange
    ): EmptyTuple =
      args

  given tuple[H, T <: Tuple](using head: ArgSpec[H], tail: ArgSpec[T]): ArgSpec[H *: T] with
    def describeParts: List[String] =
      head.describeParts ++ tail.describeParts

    def parse(
      args: List[TExpr[?]],
      pos: Int,
      fnName: String
    ): Either[ParseError, (H *: T, List[TExpr[?]])] =
      for
        (h, rest) <- head.parse(args, pos, fnName)
        (t, rest2) <- tail.parse(rest, pos, fnName)
      yield (h *: t, rest2)

    def toValues(args: H *: T): List[ArgValue] =
      head.toValues(args.head) ++ tail.toValues(args.tail)

    override def scalarSlots(args: H *: T): List[(TExpr[?], LiftSlot)] =
      head.scalarSlots(args.head) ++ tail.scalarSlots(args.tail)

    override def replaceScalarSlots(
      args: H *: T,
      replacements: List[TExpr[?]]
    ): (H *: T, List[TExpr[?]]) =
      val (h, rest) = head.replaceScalarSlots(args.head, replacements)
      val (t, rest2) = tail.replaceScalarSlots(args.tail, rest)
      (h *: t, rest2)

    override def renderSlots(args: H *: T, printer: ArgPrinter): List[Option[String]] =
      head.renderSlots(args.head, printer) ++ tail.renderSlots(args.tail, printer)

    def map(
      args: H *: T
    )(
      mapExpr: TExpr[?] => TExpr[?],
      mapRange: TExpr.RangeLocation => TExpr.RangeLocation,
      mapCells: CellRange => CellRange
    ): H *: T =
      head.map(args.head)(mapExpr, mapRange, mapCells) *: tail
        .map(args.tail)(mapExpr, mapRange, mapCells)

  /** Variadic numeric arg: either a range location or a numeric expression. */
  type NumericArg = Either[TExpr.RangeLocation, TExpr[BigDecimal]]

  given numericArg: ArgSpec[NumericArg] with
    def describeParts: List[String] = List("number or range")

    def parse(
      args: List[TExpr[?]],
      pos: Int,
      fnName: String
    ): Either[ParseError, (NumericArg, List[TExpr[?]])] =
      args match
        case TExpr.RangeRef(range, form) :: tail =>
          Right((Left(TExpr.RangeLocation.Local(range, form)), tail))
        case TExpr.SheetRange(sheet, range, form) :: tail =>
          Right((Left(TExpr.RangeLocation.CrossSheet(sheet, range, form)), tail))
        // GH-353: external-workbook ranges take the range branch (like the other two shapes)
        case TExpr.ExternalRange(index, name, range, form) :: tail =>
          Right((Left(TExpr.RangeLocation.External(index, name, range, form)), tail))
        // GH-612: an error literal takes the range branch so SUM(#REF!) round-trips to the same
        // AST the shifter writes for an off-grid range
        case TExpr.ErrorLit(error) :: tail =>
          Right((Left(TExpr.RangeLocation.Error(error)), tail))
        case head :: tail =>
          Right((Right(TExpr.asNumericExpr(head)), tail))
        case Nil =>
          Left(ParseError.InvalidArguments(fnName, pos, describe, "0 arguments"))

    def toValues(args: NumericArg): List[ArgValue] =
      args match
        case Left(loc) => List(ArgValue.Range(loc))
        case Right(expr) => List(ArgValue.Expr(expr))

    def map(
      args: NumericArg
    )(
      mapExpr: TExpr[?] => TExpr[?],
      mapRange: TExpr.RangeLocation => TExpr.RangeLocation,
      mapCells: CellRange => CellRange
    ): NumericArg =
      args match
        case Left(loc) => Left(mapRange(loc))
        case Right(expr) => Right(mapExpr(expr).asInstanceOf[TExpr[BigDecimal]])

  /**
   * GH-197: SUMPRODUCT arg that accepts either a range location OR an array-producing expression.
   *
   * Unlike NumericArg which coerces to TExpr[BigDecimal], this accepts TExpr[Any] to handle array
   * expressions like `(A1:A5="Yes")*B1:B5` which evaluate to ArrayResult.
   */
  type SumProductArg = Either[TExpr.RangeLocation, TExpr[Any]]

  @SuppressWarnings(Array("org.wartremover.warts.AsInstanceOf"))
  given sumProductArg: ArgSpec[SumProductArg] with
    def describeParts: List[String] = List("array or range")

    def parse(
      args: List[TExpr[?]],
      pos: Int,
      fnName: String
    ): Either[ParseError, (SumProductArg, List[TExpr[?]])] =
      args match
        case TExpr.RangeRef(range, form) :: tail =>
          Right((Left(TExpr.RangeLocation.Local(range, form)), tail))
        case TExpr.SheetRange(sheet, range, form) :: tail =>
          Right((Left(TExpr.RangeLocation.CrossSheet(sheet, range, form)), tail))
        // GH-353: external-workbook ranges take the range branch (like the other two shapes)
        case TExpr.ExternalRange(index, name, range, form) :: tail =>
          Right((Left(TExpr.RangeLocation.External(index, name, range, form)), tail))
        case TExpr.ErrorLit(error) :: tail =>
          Right((Left(TExpr.RangeLocation.Error(error)), tail))
        case head :: tail =>
          Right((Right(head.asInstanceOf[TExpr[Any]]), tail))
        case Nil =>
          Left(ParseError.InvalidArguments(fnName, pos, describe, "0 arguments"))

    def toValues(args: SumProductArg): List[ArgValue] =
      args match
        case Left(loc) => List(ArgValue.Range(loc))
        case Right(expr) => List(ArgValue.Expr(expr))

    def map(
      args: SumProductArg
    )(
      mapExpr: TExpr[?] => TExpr[?],
      mapRange: TExpr.RangeLocation => TExpr.RangeLocation,
      mapCells: CellRange => CellRange
    ): SumProductArg =
      args match
        case Left(loc) => Left(mapRange(loc))
        case Right(expr) => Right(mapExpr(expr).asInstanceOf[TExpr[Any]])
