package com.tjclp.xl.formula.functions

import com.tjclp.xl.formula.Arity
import com.tjclp.xl.formula.ast.{RangeForm, TExpr}
import com.tjclp.xl.formula.eval.{ArrayResult, EvalError, Evaluator, RangeOperand}
import com.tjclp.xl.formula.parser.{FormulaParser, ParseError}

import com.tjclp.xl.addressing.{CellRange, SheetName}
import com.tjclp.xl.cells.{CellError, CellValue}

/**
 * GH-669: Excel's reference operators — union `,` (inside parentheses: `SUM((A1,A2))`) and
 * intersection ` ` (a single space: `A1:C3 B2:D4`); GH-713: the range operator `:` between
 * reference-valued operands (`A1:INDEX(A:A,COUNTA(A:A))`, `INDEX(r,1):INDEX(r,2)`).
 *
 * All three are unregistered `TExpr.Call` specs, like `@x` (SINGLE) and `x#` (ANCHORARRAY): every
 * walker already reaches a call's operands through its ArgSpec, so dependency edges, shifting,
 * sheet renames and externality come for free, and their names are not identifiers, so `=UNION(…)`
 * stays an unknown function and `xl functions` never lists them. The object lives outside the
 * `FunctionSpecs*` traits for that reason (FunctionRegistryMacro collects every spec those
 * declare).
 *
 * No multi-area value ever enters the evaluator's value plane. A union's areas exist only inside
 * [[areas]], which the consumers that accept several areas call — the variadic aggregates, INDEX's
 * area_num, AREAS and the intersection itself. Everywhere else a union evaluates through its own
 * `eval`, Excel's `#VALUE!` for a multi-area reference in a value position. An intersection that
 * resolves to one area is an ordinary reference, so every reference position takes it unchanged.
 *
 * Initialization: this object reads `Evaluator` only inside defs, never in a val initializer, so it
 * cannot join an object-initialization cycle with `FunctionSpecs` (whose AREAS uses it).
 */
@SuppressWarnings(Array("org.wartremover.warts.AsInstanceOf"))
private[formula] object ReferenceOperators extends FunctionSpecsBase:

  /** The operators' spec names — never identifiers, so no formula can call them by name. */
  final val UnionName = "(union)"
  final val IntersectionName = "(intersection)"
  final val RangeName = "(range)"

  /** One operand: a location the parser recognized, or a reference-valued expression. */
  type Operand = Either[TExpr.RangeLocation, TExpr[Any]]

  /** A call to one of the two operators, looking through the parser's coercion wrapper. */
  object OperatorCall:
    def unapply(e: TExpr[?]): Option[TExpr.Call[?]] = e match
      case TExpr.Coerced(inner, _) => unapply(inner)
      case call: TExpr.Call[?]
          if call.spec.name == UnionName || call.spec.name == IntersectionName =>
        Some(call)
      case _ => None

  def isOperatorCall(e: TExpr[?]): Boolean = OperatorCall.unapply(e).isDefined

  /**
   * GH-713: a call to the range operator, looking through the parser's coercion wrapper. Kept out
   * of [[OperatorCall]]: a range is one area, so [[areas]] reaches it as any reference-returning
   * call (its own `reference`).
   */
  object RangeOperatorCall:
    def unapply(e: TExpr[?]): Option[TExpr.Call[?]] = e match
      case TExpr.Coerced(inner, _) => unapply(inner)
      case call: TExpr.Call[?] if call.spec.name == RangeName => Some(call)
      case _ => None

  /**
   * Can `e` be an operand of `,` or ` `? Every reference the parser can see: a cell, range, name,
   * LET name or error literal (`A1:B2 #REF!` is what a deletion leaves), a union or intersection,
   * and a call to a function that returns a reference (IF, IFS, CHOOSE, SWITCH, OFFSET, INDIRECT,
   * INDEX, XLOOKUP, the range operator). Literals, arithmetic, comparisons, `@x`, `x#`, array
   * constants and every other call are values — Excel refuses them too.
   */
  def isReferenceShape(e: TExpr[?]): Boolean = e match
    case TExpr.Coerced(inner, _) => isReferenceShape(inner)
    case _: TExpr.RangeRef | _: TExpr.SheetRange | _: TExpr.ExternalRange | _: TExpr.PolyRef |
        _: TExpr.Ref[?] | _: TExpr.SheetPolyRef | _: TExpr.SheetRef[?] | _: TExpr.ExternalRef |
        _: TExpr.NameRef | _: TExpr.SheetNameRef | _: TExpr.ErrorLit | _: TExpr.BindingRef |
        _: TExpr.CoercedBindingRef[?] =>
      true
    case call: TExpr.Call[?] =>
      isOperatorCall(call) || Evaluator.referenceFunctions.contains(call.spec.name)
    case _ => false

  /**
   * An operand slot: a location as `Left` (every shape `ArgSpec.rangeLocation` accepts — a single
   * cell is its 1×1 range), otherwise the expression as `Right` when `acceptRight` admits it, else
   * the parse error a range slot gives. An operand is never an array-lifted scalar slot.
   */
  def operandSpec(acceptRight: TExpr[?] => Boolean): ArgSpec[Operand] = new ArgSpec[Operand]:
    def describeParts: List[String] = List("reference")

    def parse(
      args: List[TExpr[?]],
      pos: Int,
      fnName: String
    ): Either[ParseError, (Operand, List[TExpr[?]])] =
      args match
        case head :: tail =>
          // the whole list, so a refusal reports the call's own argument count
          ArgSpec.rangeLocation.parse(args, pos, fnName) match
            case Right((location, rest)) => Right((Left(location), rest))
            case Left(_) if acceptRight(head) => Right((Right(head.asInstanceOf[TExpr[Any]]), tail))
            case Left(err) => Left(err)
        case Nil => Left(ParseError.InvalidArguments(fnName, pos, describe, "0 arguments"))

    def toValues(args: Operand): List[ArgValue] = args match
      case Left(location) => List(ArgValue.Range(location))
      case Right(expr) => List(ArgValue.Expr(expr))

    def map(args: Operand)(
      mapExpr: TExpr[?] => TExpr[?],
      mapRange: TExpr.RangeLocation => TExpr.RangeLocation,
      mapCells: com.tjclp.xl.CellRange => com.tjclp.xl.CellRange
    ): Operand = args match
      case Left(location) => Left(mapRange(location))
      case Right(expr) => Right(mapExpr(expr).asInstanceOf[TExpr[Any]])

  /** An operand of `,` or ` ` (and AREAS' argument): any reference shape. */
  val referenceOperand: ArgSpec[Operand] = operandSpec(isReferenceShape)

  /**
   * INDEX's first slot: a location exactly as before, an operator call (area_num picks among its
   * areas), a range-operator call (GH-713: `INDEX(A1:INDEX(A:A,5),2)`) or an array constant (INDEX
   * indexes the constant's values). Any other expression keeps the range slot's parse error.
   */
  val indexOperand: ArgSpec[Operand] = operandSpec(e =>
    isOperatorCall(e) || RangeOperatorCall.unapply(e).isDefined || (e match
      case TExpr.Lit(_: ArrayResult) => true
      case _ => false)
  )

  private def operandText(operand: Operand, printer: ArgPrinter): String = operand match
    case Left(location) => printer.location(location)
    case Right(expr) => printer.expr(expr)

  /**
   * `(a,b,…)`. Printed with a bare `,` in every form: a space is itself an operator here, so the
   * human form and the file form are the same text.
   */
  val union: FunctionSpec[CellValue] { type Args = List[Operand] } =
    FunctionSpec.simple[CellValue, List[Operand]](
      UnionName,
      Arity.AtLeast(2),
      renderFn =
        Some((operands, printer) => operands.map(operandText(_, printer)).mkString("(", ",", ")"))
    )((_, _) => Left(multiAreaValue))(using ArgSpec.list(using referenceOperand))

  /** Excel's answer for a multi-area reference in a value position: `=(A1,A2)`, `ABS((A1,A2))`. */
  private def multiAreaValue: EvalError =
    EvalError.ErrorValue(
      CellError.Value,
      Some("a multi-area reference (a union) is not a value here")
    )

  /**
   * `a b` — the cells two references share. One area is an ordinary reference (a plain cell reads
   * it through the implicit intersection, array mode whole); two or more (a union intersecting
   * several areas) are `#VALUE!` outside the consumers that take several; none is `#NULL!`.
   */
  val intersection: FunctionSpec[ArrayResult] { type Args = (Operand, Operand) } =
    FunctionSpec.referencing[ArrayResult, (Operand, Operand)](
      IntersectionName,
      Arity.two,
      renderFn = Some { case ((left, right), printer) =>
        val rightText = operandText(right, printer)
        // left-associative: a right operand that is itself an intersection keeps its parens
        val grouped = right match
          case Right(OperatorCall(call)) if call.spec.name == IntersectionName => s"($rightText)"
          case _ => rightText
        s"${operandText(left, printer)} $grouped"
      }
    )((args, ctx) => singleArea(args, ctx)) { (args, ctx) =>
      singleArea(args, ctx).flatMap { case RangeOperand(sheet, range) =>
        referenceResult(range, sheet, ctx)(
          extractRangeAsMatrixEval(range, sheet, ctx).map(ArrayResult(_))
        )
      }
    }(using ArgSpec.tuple(using referenceOperand, ArgSpec.tuple(using referenceOperand)))

  /**
   * GH-713: `a:b` — the smallest range holding every area of both operands, on their one sheet.
   * Like the intersection, one area is an ordinary reference: a plain cell reads it through the
   * implicit intersection, array mode whole, a reference position takes it unread.
   */
  val range: FunctionSpec[ArrayResult] { type Args = (Operand, Operand) } =
    FunctionSpec.referencing[ArrayResult, (Operand, Operand)](
      RangeName,
      Arity.two,
      renderFn = Some((args, printer) => renderRange(args, printer))
    )((args, ctx) => boundingArea(args, ctx)) { (args, ctx) =>
      boundingArea(args, ctx).flatMap { area =>
        referenceResult(area.range, area.sheet, ctx)(boundedReferenceValues(area, ctx))
      }
    }(using ArgSpec.tuple(using referenceOperand, ArgSpec.tuple(using referenceOperand)))

  /**
   * `l:r` with the fewest grouping parentheses that re-parse to this very tree. The `:` is both a
   * lexical range (`A1:B2`) and this operator, and a reference token before a `:` absorbs what
   * follows it (`INDEX(r,1):B5:C6` is `INDEX(r,1):(B5:C6)`), so whether an operand needs its
   * parentheses depends on the neighbouring text: the candidates are checked against the parser
   * itself, and the fully parenthesized form — always faithful — is the fallback. Every
   * unparenthesized text the parser accepts prints back byte-for-byte.
   */
  private def renderRange(args: (Operand, Operand), printer: ArgPrinter): String =
    val left = operandText(args._1, printer)
    val right = operandText(args._2, printer)
    val grouped = s"($left):($right)"
    FormulaParser.parse(grouped) match
      case Right(target) =>
        List(s"$left:$right", s"$left:($right)", s"($left):$right")
          .find(candidate => FormulaParser.parse(candidate) == Right(target))
          .getOrElse(grouped)
      case Left(_) => grouped

  /**
   * GH-713: the area `a:b` denotes — the bounding range of every area of both operands, left to
   * right, the first error winning (`#REF!`, `#N/A`, a missing sheet). An operand that evaluates to
   * a value is `#VALUE!` (not a reference), as are areas on two sheets (Excel; LibreOffice spans
   * them).
   */
  private[formula] def boundingArea(
    args: (Operand, Operand),
    ctx: EvalContext
  ): Either[EvalError, RangeOperand] =
    for
      left <- areas(args._1, ctx)
      right <- areas(args._2, ctx)
      all = left ++ right
      first <- all.headOption.toRight(
        EvalError.ErrorValue(CellError.Value, Some("a range needs a reference on each side"))
      )
      _ <-
        if all.forall(_.sheet.name == first.sheet.name) then Right(())
        else
          Left(
            EvalError.ErrorValue(CellError.Value, Some("a range's references must be on one sheet"))
          )
    yield RangeOperand(
      first.sheet,
      all.foldLeft(first.range)((box, area) => box.expand(area.range.start).expand(area.range.end))
    )

  private def singleArea(
    args: (Operand, Operand),
    ctx: EvalContext
  ): Either[EvalError, RangeOperand] =
    intersect(args._1, args._2, ctx).flatMap {
      case Vector(one) => Right(one)
      case _ =>
        Left(
          EvalError.ErrorValue(CellError.Value, Some("a multi-area reference is not accepted here"))
        )
    }

  /** Build `(a,b,…)` from parsed operands, each of which must be a reference. */
  def mkUnion(areas: List[TExpr[?]], pos: Int): Either[ParseError, TExpr[?]] =
    if areas.size < 2 then Left(ParseError.InvalidOperator(",", pos, "a union needs two areas"))
    else if !areas.forall(isReferenceShape) then
      Left(ParseError.InvalidOperator(",", pos, "every area of a union must be a reference"))
    else
      ArgSpec.list(using referenceOperand).parse(areas, pos, UnionName).map { case (operands, _) =>
        TExpr.Call(union, operands)
      }

  /** Build `l r`; both operands must be references. */
  def mkIntersection(l: TExpr[?], r: TExpr[?], pos: Int): Either[ParseError, TExpr[?]] =
    if !(isReferenceShape(l) && isReferenceShape(r)) then
      Left(ParseError.InvalidOperator(" ", pos, "intersection operands must be references"))
    else
      for
        (left, _) <- referenceOperand.parse(List(l), pos, IntersectionName)
        (right, _) <- referenceOperand.parse(List(r), pos, IntersectionName)
      yield TExpr.Call(intersection, (left, right))

  /**
   * GH-713: build `l:r`; both operands must be references. Two single cells on one sheet (`(A1):B2`
   * — the only text that reaches here with them) fold into the lexical range they spell.
   */
  def mkRange(l: TExpr[?], r: TExpr[?], pos: Int): Either[ParseError, TExpr[?]] =
    if !(isReferenceShape(l) && isReferenceShape(r)) then
      Left(ParseError.InvalidOperator(":", pos, "range operands must be references"))
    else
      for
        (left, _) <- referenceOperand.parse(List(l), pos, RangeName)
        (right, _) <- referenceOperand.parse(List(r), pos, RangeName)
      yield foldCells(left, right).getOrElse(TExpr.Call(range, (left, right)))

  private def foldCells(left: Operand, right: Operand): Option[TExpr[?]] =
    def cellText(range: CellRange): String = range.toA1Anchored.takeWhile(_ != ':')
    def spelled(a: CellRange, b: CellRange): Option[(String, CellRange)] =
      val text = s"${cellText(a)}:${cellText(b)}"
      CellRange.parse(text).toOption.map((text, _))
    (left, right) match
      case (
            Left(TExpr.RangeLocation.Local(a, RangeForm.Cell)),
            Left(TExpr.RangeLocation.Local(b, RangeForm.Cell))
          ) =>
        spelled(a, b).map((text, cells) => TExpr.RangeRef(cells, RangeForm.of(text, cells)))
      case (
            Left(TExpr.RangeLocation.CrossSheet(sa, a, RangeForm.Cell)),
            Left(TExpr.RangeLocation.CrossSheet(sb, b, RangeForm.Cell))
          ) if sa.value.equalsIgnoreCase(sb.value) =>
        spelled(a, b).map((text, cells) => TExpr.SheetRange(sa, cells, RangeForm.of(text, cells)))
      case _ => None

  /**
   * The areas an operand denotes, left to right — the ONLY multi-area resolution. A location is one
   * area (its errors propagate: `#REF!`, a missing sheet, a name bound to a non-reference). A union
   * concatenates its operands' areas, duplicates and overlaps kept (`SUM((A1,A1))` is 2×A1; nested
   * unions flatten), all on one sheet. An intersection intersects every pair of its operands'
   * areas, left-major, keeping the non-empty ones; none is `#NULL!`. Any other expression is the
   * reference it evaluates to, or `#VALUE!` when it evaluates to a value. The first error wins.
   */
  private[formula] def areas(
    op: Operand,
    ctx: EvalContext
  ): Either[EvalError, Vector[RangeOperand]] =
    op match
      case Left(location) =>
        Evaluator
          .resolveRangeLocation(location, ctx.sheet, ctx.workbook)
          .map((sheet, range) => Vector(RangeOperand(sheet, range)))
      case Right(OperatorCall(call)) if call.spec.name == UnionName =>
        call.args
          .asInstanceOf[List[Operand]]
          .foldLeft[Either[EvalError, Vector[RangeOperand]]](Right(Vector.empty)) {
            (acc, operand) =>
              acc.flatMap(found => areas(operand, ctx).map(found ++ _))
          }
          .flatMap { found =>
            if found.map(_.sheet.name).distinct.size <= 1 then Right(found)
            else
              Left(
                EvalError.ErrorValue(CellError.Value, Some("a union's areas must be on one sheet"))
              )
          }
      case Right(OperatorCall(call)) =>
        val (left, right) = call.args.asInstanceOf[(Operand, Operand)]
        intersect(left, right, ctx)
      case Right(other) =>
        ctx.evalReference(other).flatMap {
          case operand: RangeOperand => Right(Vector(operand))
          case _ => Left(EvalError.ErrorValue(CellError.Value, Some("not a reference")))
        }

  private def intersect(
    left: Operand,
    right: Operand,
    ctx: EvalContext
  ): Either[EvalError, Vector[RangeOperand]] =
    for
      leftAreas <- areas(left, ctx)
      rightAreas <- areas(right, ctx)
      pairs = for l <- leftAreas; r <- rightAreas yield (l, r)
      _ <-
        if pairs.forall((l, r) => l.sheet.name == r.sheet.name) then Right(())
        else
          Left(
            EvalError.ErrorValue(
              CellError.Value,
              Some("an intersection's references must be on one sheet")
            )
          )
      shared = pairs.flatMap((l, r) => l.range.intersect(r.range).map(RangeOperand(l.sheet, _)))
      result <-
        if shared.nonEmpty then Right(shared)
        else Left(EvalError.ErrorValue(CellError.Null, Some("the intersection is empty")))
    yield result

  /**
   * Where a location lies, for the dependency graph: the sheet (None: the formula's own) and the
   * range. The sheet-level graph knows the static locations only ([[staticShape]]); the workbook
   * graph also resolves defined names.
   */
  type ShapeOf = TExpr.RangeLocation => Option[(Option[SheetName], CellRange)]

  /** A static location's shape; None for a name, an error or an external location. */
  val staticShape: ShapeOf = {
    case TExpr.RangeLocation.Local(range, _) => Some((None, range))
    case TExpr.RangeLocation.CrossSheet(sheet, range, _) => Some((Some(sheet), range))
    case _ => None
  }

  /**
   * The static dependency edges of an operator call: an intersection of two static locations on one
   * sheet depends only on the cells they share (whole-operand edges would falsely make
   * `C3: =SUM(A1:C3 A1:A3)` circular), nothing when they share none. GH-713: a range whose operands
   * all denote fixed cells (`A1:B1:C3`, `Start:Finish`) also depends on its whole bounding range,
   * which neither operand names (B2 here). Anything else — names, dynamic operands, nesting — keeps
   * every operand's edges, a sound over-approximation; a computed range those edges do not cover is
   * a dynamic reference ([[isComputedRangeDynamic]]).
   */
  def dependencyValues(specName: String, values: List[ArgValue], shapeOf: ShapeOf): List[ArgValue] =
    if specName == RangeName then values ++ boundingEdges(values, shapeOf)
    else if specName != IntersectionName then values
    else
      values match
        case List(ArgValue.Range(l), ArgValue.Range(r)) =>
          (l, r) match
            case (TExpr.RangeLocation.Local(a, _), TExpr.RangeLocation.Local(b, _)) =>
              a.intersect(b).map(s => ArgValue.Range(TExpr.RangeLocation.Local(s))).toList
            case (
                  TExpr.RangeLocation.CrossSheet(sa, a, _),
                  TExpr.RangeLocation.CrossSheet(sb, b, _)
                ) if sa.value.equalsIgnoreCase(sb.value) =>
              a.intersect(b).map(s => ArgValue.Range(TExpr.RangeLocation.CrossSheet(sa, s))).toList
            case _ => values
        case _ => values

  private def operandOf(value: ArgValue): Option[Operand] = value match
    case ArgValue.Range(location) => Some(Left(location))
    case ArgValue.Expr(expr) => Some(Right(expr.asInstanceOf[TExpr[Any]]))
    case ArgValue.Cells(_) => None

  /** The bounding range of a range whose operands denote fixed cells, on each sheet it may be. */
  private def boundingEdges(values: List[ArgValue], shapeOf: ShapeOf): List[ArgValue] =
    values.flatMap(operandOf) match
      case List(l, r) =>
        (envelope(l, shapeOf), envelope(r, shapeOf)) match
          case (el, er) if fixed(el) && fixed(er) =>
            val leaves = fixedLeaves(el) ++ fixedLeaves(er)
            val qualified = leaves.flatMap(_._1).distinctBy(sheetKey)
            leaves.headOption match
              case Some((_, first)) if qualified.size <= 1 =>
                val box = leaves.foldLeft(first)((acc, leaf) => hull(acc, leaf._2))
                val local =
                  if leaves.exists(_._1.isEmpty) then List(TExpr.RangeLocation.Local(box)) else Nil
                (local ++ qualified.map(TExpr.RangeLocation.CrossSheet(_, box)))
                  .map(ArgValue.Range(_))
              case _ => Nil
          case _ => Nil
      case _ => Nil

  /**
   * GH-713: the cells an operand of `:` may denote, for the dependency graph.
   *
   *   - `Exact(leaves, covered)`: the operand denotes exactly these locations (a static location, a
   *     name the graph resolves, a union, a range of such); the range reads their bounding range.
   *   - `Within(sheet, range, covered)`: the operand denotes some reference inside `range` (INDEX's
   *     array, XLOOKUP's return array, the branches IF, IFS, CHOOSE and SWITCH select from).
   *   - `NoCells`: it denotes none (an error literal, an INDEX into an array constant).
   *   - `Unknown`: the graph cannot bound it (OFFSET, INDIRECT, a LET name, an unresolved name).
   *
   * `covered`: the operand's own dependency edges already include every cell of its bounding range.
   */
  private[formula] enum Envelope derives CanEqual:
    case Exact(leaves: Vector[(Option[SheetName], CellRange)], covered: Boolean)
    case Within(sheet: Option[SheetName], range: CellRange, covered: Boolean)
    case NoCells
    case Unknown

  private def sheetKey(sheet: SheetName): String = sheet.value.toLowerCase(java.util.Locale.ROOT)

  private def sameSheet(a: Option[SheetName], b: Option[SheetName]): Boolean =
    a.map(sheetKey) == b.map(sheetKey)

  private def hull(a: CellRange, b: CellRange): CellRange = a.expand(b.start).expand(b.end)

  private def holds(outer: CellRange, inner: CellRange): Boolean =
    outer.colStart.index0 <= inner.colStart.index0 && outer.colEnd.index0 >= inner.colEnd.index0 &&
      outer.rowStart.index0 <= inner.rowStart.index0 && outer.rowEnd.index0 >= inner.rowEnd.index0

  private def fixed(e: Envelope): Boolean = e match
    case Envelope.Exact(_, _) | Envelope.NoCells => true
    case _ => false

  private def fixedLeaves(e: Envelope): Vector[(Option[SheetName], CellRange)] = e match
    case Envelope.Exact(leaves, _) => leaves
    case _ => Vector.empty

  /** An envelope as one area on one sheet: None when it spans sheets or is not bounded. */
  private def asArea(e: Envelope): Option[(Option[SheetName], CellRange, Boolean)] = e match
    case Envelope.Within(sheet, range, covered) => Some((sheet, range, covered))
    case Envelope.Exact(leaves, covered) =>
      leaves.headOption.flatMap { case (sheet, first) =>
        Option.when(leaves.forall(leaf => sameSheet(leaf._1, sheet)))(
          (sheet, leaves.foldLeft(first)((acc, leaf) => hull(acc, leaf._2)), covered)
        )
      }
    case _ => None

  /** The bounding envelope of two (a range's operands, a union's areas). */
  private def join(a: Envelope, b: Envelope, covered: Boolean): Envelope = (a, b) match
    case (Envelope.Unknown, _) | (_, Envelope.Unknown) => Envelope.Unknown
    case (Envelope.NoCells, other) => other
    case (other, Envelope.NoCells) => other
    case (Envelope.Exact(la, _), Envelope.Exact(lb, _)) => Envelope.Exact(la ++ lb, covered)
    case _ =>
      (asArea(a), asArea(b)) match
        case (Some((sa, ra, _)), Some((sb, rb, _))) if sameSheet(sa, sb) =>
          Envelope.Within(sa, hull(ra, rb), covered)
        case _ => Envelope.Unknown

  /** `e` as the bound of a reference chosen inside it (INDEX's array, XLOOKUP's return array). */
  private def within(e: Envelope, covered: Boolean): Envelope = e match
    case Envelope.Unknown | Envelope.NoCells => e
    case _ =>
      asArea(e).fold[Envelope](Envelope.Unknown) { case (sheet, range, own) =>
        Envelope.Within(sheet, range, covered && own)
      }

  private val selectingFunctions = Set("IF", "IFS", "CHOOSE", "SWITCH")

  /** The envelope of one operand of `:` (see [[Envelope]]). */
  private[formula] def envelope(op: Operand, shapeOf: ShapeOf): Envelope = op match
    case Left(TExpr.RangeLocation.Error(_, _)) => Envelope.NoCells
    case Left(location @ (_: TExpr.RangeLocation.Local | _: TExpr.RangeLocation.CrossSheet)) =>
      shapeOf(location).fold[Envelope](Envelope.Unknown)(leaf => Envelope.Exact(Vector(leaf), true))
    case Left(location) =>
      shapeOf(location).fold[Envelope](Envelope.Unknown)(leaf =>
        Envelope.Exact(Vector(leaf), false)
      )
    case Right(expr) => expressionEnvelope(expr, shapeOf)

  private def expressionEnvelope(expr: TExpr[?], shapeOf: ShapeOf): Envelope = expr match
    case TExpr.Coerced(inner, _) => expressionEnvelope(inner, shapeOf)
    case TExpr.ErrorLit(_, _) => Envelope.NoCells
    case RangeOperatorCall(call) =>
      val (l, r) = call.args.asInstanceOf[(Operand, Operand)]
      // an all-fixed range carries its bounding edges (dependencyValues); otherwise only a
      // covering operand's edges cover it
      val (el, er) = (envelope(l, shapeOf), envelope(r, shapeOf))
      join(el, er, covered = (fixed(el) && fixed(er)) || coveredBy(el, er))
    case OperatorCall(call) if call.spec.name == UnionName =>
      call.args
        .asInstanceOf[List[Operand]]
        .map(envelope(_, shapeOf))
        .foldLeft[Envelope](Envelope.NoCells)((acc, e) => join(acc, e, covered = false))
    case OperatorCall(call) =>
      val (l, r) = call.args.asInstanceOf[(Operand, Operand)]
      (envelope(l, shapeOf), envelope(r, shapeOf)) match
        case (Envelope.Unknown, _) | (_, Envelope.Unknown) => Envelope.Unknown
        case (Envelope.NoCells, _) | (_, Envelope.NoCells) => Envelope.NoCells
        case (el, er) =>
          (asArea(el), asArea(er)) match
            case (Some((sa, ra, ca)), Some((sb, rb, cb))) if sameSheet(sa, sb) =>
              ra.intersect(rb).fold[Envelope](Envelope.NoCells)(Envelope.Within(sa, _, ca || cb))
            case _ => Envelope.Unknown
    case call: TExpr.Call[?] if call.spec.name == "INDEX" =>
      call.args.asInstanceOf[IndexArgs]._1 match
        case Right(TExpr.Lit(_: ArrayResult)) => Envelope.NoCells
        case array => within(envelope(array, shapeOf), covered = true)
    case call: TExpr.Call[?] if call.spec.name == "XLOOKUP" =>
      within(envelope(Left(call.args.asInstanceOf[XLookupArgs]._3), shapeOf), covered = true)
    case call: TExpr.Call[?] if selectingFunctions.contains(call.spec.name) =>
      val branches: List[Operand] = call.spec.argSpec.toValues(call.args).collect {
        case ArgValue.Range(location) => Left(location)
        case ArgValue.Expr(e) if isReferenceShape(e) => Right(e.asInstanceOf[TExpr[Any]])
      }
      within(
        branches
          .map(envelope(_, shapeOf))
          .foldLeft[Envelope](Envelope.NoCells)((acc, e) => join(acc, e, covered = false)),
        covered = false
      )
    case _ => Envelope.Unknown

  /** Whether one operand's own covered edges hold the whole bounding range of both. */
  private def coveredBy(el: Envelope, er: Envelope): Boolean =
    join(el, er, covered = false) match
      case Envelope.NoCells => true
      case whole =>
        asArea(whole).exists { case (sheet, box, _) =>
          List(el, er).exists(e =>
            asArea(e).exists((s, range, covered) =>
              covered && sameSheet(s, sheet) && holds(range, box)
            )
          )
        }

  /**
   * GH-713: whether a range-operator call reads cells its dependency edges may miss, so the graph
   * treats it as a dynamic reference (evaluated last, always recalculated). A range of fixed
   * operands carries its bounding edges, and one whose bounding range lies inside a single
   * operand's covered envelope needs nothing more — `A1:INDEX(A:A,COUNTA(A:A))`,
   * `INDEX(r,1):INDEX(r,2)`, `B1:XLOOKUP(k,A1:A3,B1:B3)` stay static and non-volatile. Anything
   * unbounded (OFFSET, INDIRECT, a LET or unresolved name) or uncovered (`A1:INDEX(C:C,5)`) is
   * dynamic.
   */
  def isComputedRangeDynamic(call: TExpr.Call[?], shapeOf: ShapeOf): Boolean =
    call.spec.name == RangeName && {
      val (l, r) = call.args.asInstanceOf[(Operand, Operand)]
      val (el, er) = (envelope(l, shapeOf), envelope(r, shapeOf))
      el == Envelope.Unknown || er == Envelope.Unknown ||
      !((fixed(el) && fixed(er)) || coveredBy(el, er))
    }

  /**
   * GH-713: the dynamic-cell pre-filter for computed ranges — a linear scan, outside string
   * literals and quoted sheet names, for a `:` that is not one lexical range: one whose neighbours
   * are not a cell, column or row pair (`)` before it, `(` after it, a name, a qualifier, a call's
   * name followed by `(`). Law: every formula whose parse [[isComputedRangeDynamic]] flags passes,
   * since a dynamic range always has a call, a parenthesis or a name beside one of its `:`.
   */
  def mayContainComputedRange(text: String): Boolean =
    def tokenChar(c: Char): Boolean =
      c.isLetterOrDigit || c == '$' || c == '_' || c == '.' || c == '\\'
    def skipQuoted(open: Int, quote: Char): Int =
      @annotation.tailrec
      def loop(j: Int): Int =
        if j >= text.length then j
        else if text.charAt(j) == quote then
          if text.lift(j + 1).contains(quote) then loop(j + 2) else j + 1
        else loop(j + 1)
      loop(open + 1)
    // the token run ending before `end` (exclusive) and the one starting at `start`
    def runBefore(end: Int): String =
      @annotation.tailrec
      def back(j: Int): Int = if j > 0 && tokenChar(text.charAt(j - 1)) then back(j - 1) else j
      text.substring(back(end), end)
    def runFrom(start: Int): String =
      @annotation.tailrec
      def ahead(j: Int): Int =
        if j < text.length && tokenChar(text.charAt(j)) then ahead(j + 1) else j
      text.substring(start, ahead(start))
    @annotation.tailrec
    def scan(i: Int): Boolean =
      if i >= text.length then false
      else
        text.charAt(i) match
          case '"' => scan(skipQuoted(i, '"'))
          case '\'' => scan(skipQuoted(i, '\''))
          case ':' =>
            val left = runBefore(i)
            val right = runFrom(i + 1)
            val next = text.lift(i + 1 + right.length)
            val lexical = left.nonEmpty && right.nonEmpty &&
              !next.exists(c => c == '(' || c == '!') &&
              CellRange.parse(s"$left:$right").isRight
            !lexical || scan(i + 1)
          case _ => scan(i + 1)
    text.indexOf(':') >= 0 && scan(0)
